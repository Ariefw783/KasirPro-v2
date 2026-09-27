/* KasirPro AT-09 — Database Operational Incremental Sync & Quota Optimization
 * Master statis: IndexedDB + Master Snapshot Firestore (tetap AT-08).
 * Operasional: IndexedDB cache lokal + incremental Firestore sync.
 * Full seed hanya saat cache belum tersedia / recovery manual.
 * Write operasional tetap fail-closed dan memakai delta/fingerprint.
 */

import * as legacy from "./database-store-core.js";
import { firebaseDb } from "./firebase-client.js";
import { waitForFirebaseUser } from "./auth.js";
import {
    collectionSegments,
    documentSegments,
    readableDocumentId
} from "./database-paths.js";
import {
    collection,
    doc,
    getDocsFromServer,
    limit,
    orderBy,
    query,
    runTransaction,
    where,
    Timestamp,
    serverTimestamp,
    writeBatch
} from "https://www.gstatic.com/firebasejs/12.18.0/firebase-firestore.js";
import {
    readMasterSnapshot,
    readMasterVersion,
    installMasterSnapshot
} from "../local/master-repository.js";
import {
    ensureLocalMaster,
    inspectMasterSyncState,
    updateLocalMasterFromCentral
} from "../local/master-sync.js";
import { FIRESTORE_MASTER_SNAPSHOT_ENABLED } from "../local/master-firestore-config.js";
import {
    readOperationalRecords,
    replaceOperationalRecords,
    mergeOperationalRecords,
    getOperationalSyncMeta,
    setOperationalSyncMeta,
    resetOperationalCollection
} from "../local/operational-indexeddb.js";

export * from "./database-store-core.js";
export const STORE_KEYS = legacy.STORE_KEYS;
const ACTIVE_STOCK_KEY = "__kasirpro_active_stock_v1";

const OPERATIONAL_COLLECTIONS = Object.freeze({
    [STORE_KEYS.invoices]: "purchaseInvoices",
    [STORE_KEYS.sales]: "sales",
    [STORE_KEYS.movements]: "stockMovements",
    [STORE_KEYS.opnames]: "stockOpnames",
    [ACTIVE_STOCK_KEY]: "activeStocks"
});

const OPERATIONAL_CACHE_STORES = Object.freeze({
    [STORE_KEYS.invoices]: "purchaseInvoices",
    [STORE_KEYS.sales]: "sales",
    [STORE_KEYS.movements]: "stockMovements",
    [STORE_KEYS.opnames]: "stockOpnames",
    [ACTIVE_STOCK_KEY]: "stockLevels"
});

const INITIAL_READ_LIMITS = Object.freeze({
    [STORE_KEYS.invoices]: 500,
    [STORE_KEYS.sales]: 1000,
    [STORE_KEYS.movements]: 1500,
    [STORE_KEYS.opnames]: 250
});

let initializationPromise = null;
let localMasterMode = false;
let masterCache = null;
let operationalRemoteReady = false;
let lastMasterSyncState = null;
let operationalState = new Map();
let operationalFingerprints = new Map();
let operationalWriteQueue = Promise.resolve();
let stockIndex = new Map();
let materializedMasterCache = null;
let masterRevision = 0;
let stockRevision = 0;
let materializedRevision = "";
let operationalSyncState = {
    mode: "not-initialized",
    page: "login",
    ready: false,
    collections: {}
};

function clone(value) {
    if (value === undefined) return undefined;
    return typeof structuredClone === "function"
        ? structuredClone(value)
        : JSON.parse(JSON.stringify(value));
}

function norm(value) {
    return String(value ?? "").trim().toLowerCase();
}

function num(value) {
    return Number(String(value ?? 0).replace(/[^0-9.-]/g, "")) || 0;
}

function fingerprint(value) {
    return JSON.stringify(value);
}

function emptyMaster() {
    return { produk: [], supplier: [], kategori: [], pengguna: [], pengaturan_toko: [] };
}

function pageMode() {
    const path = location.pathname.toLowerCase();
    if (path.includes("/management/")) return "management";
    if (path.includes("/pos/")) return "pos";
    return "login";
}

function emitReady(master, syncState = null) {
    const products = Array.isArray(master?.produk) ? master.produk.length : 0;
    const recovered = syncState?.recovered === true || syncState?.state === "recovered";
    window.dispatchEvent(new CustomEvent("kasirpro:database-preload-start", {
        detail: { message: recovered ? "Memulihkan Master Aktif lokal…" : "Membuka Master Aktif lokal…", total: 1 }
    }));
    window.dispatchEvent(new CustomEvent("kasirpro:database-preload-progress", {
        detail: {
            loaded: 1,
            total: 1,
            percent: 100,
            message: recovered ? `Master lokal dipulihkan (${products} produk)` : `Master lokal siap (${products} produk)`,
            collection: "indexeddb-master",
            status: "success"
        }
    }));
    window.dispatchEvent(new CustomEvent("kasirpro:database-idle"));
}

function emitSyncState(state) {
    lastMasterSyncState = state || null;
    window.dispatchEvent(new CustomEvent("kasirpro:master-sync-state", {
        detail: clone(lastMasterSyncState)
    }));
}

function emitOperationalSync(detail) {
    window.dispatchEvent(new CustomEvent("kasirpro:operational-sync", {
        detail: clone(detail)
    }));
}

async function refreshMasterCacheFromIndexedDb() {
    masterCache = await readMasterSnapshot();
    localMasterMode = true;
    masterRevision += 1;
    materializedMasterCache = null;
    return masterCache;
}

async function prepareLocalMaster() {
    const networkOptions = { allowNetwork: FIRESTORE_MASTER_SNAPSHOT_ENABLED };
    let syncState;
    try {
        syncState = await ensureLocalMaster(networkOptions);
    } catch (error) {
        console.warn("Recovery Master lokal gagal:", error);
        syncState = {
            state: "recovery-error",
            hasLocal: false,
            localVersion: 0,
            centralVersion: 0,
            recovered: false,
            error: error?.message || String(error)
        };
    }

    if (syncState?.hasLocal && syncState?.state === "update-available") {
        try {
            const updated = await updateLocalMasterFromCentral(networkOptions);
            if (updated?.updated) syncState = updated;
        } catch (error) {
            console.warn("Sinkronisasi Master pusat gagal; tetap memakai Master lokal:", error);
            syncState = {
                ...syncState,
                state: "update-failed-local-kept",
                error: error?.message || String(error)
            };
        }
    }

    emitSyncState(syncState);
    return syncState;
}

function operationalKeysForPage(mode) {
    if (mode === "pos") return [ACTIVE_STOCK_KEY, STORE_KEYS.sales, STORE_KEYS.movements];
    if (mode === "management") return [ACTIVE_STOCK_KEY, STORE_KEYS.invoices, STORE_KEYS.sales, STORE_KEYS.movements, STORE_KEYS.opnames];
    return [];
}

function operationalDocumentId(storeKey, record) {
    if (record?._firestoreDocumentId) return String(record._firestoreDocumentId);
    if (storeKey === STORE_KEYS.invoices) return readableDocumentId("faktur", record?.id || record?.number);
    if (storeKey === STORE_KEYS.sales) return readableDocumentId("penjualan", record?.id || record?.number);
    if (storeKey === STORE_KEYS.opnames) return readableDocumentId("opname", record?.id || record?.at);
    if (storeKey === STORE_KEYS.movements) {
        return readableDocumentId("mutasi", [record?.reference, record?.productCode, record?.at, record?.id].filter(Boolean).join("-"));
    }
    if (storeKey === ACTIVE_STOCK_KEY) return readableDocumentId("stok", record?.productCode);
    throw new Error(`Store operasional tidak dikenal: ${storeKey}`);
}

function cleanOperationalRecord(record) {
    const clean = clone(record) || {};
    delete clean._firestoreDocumentId;
    delete clean.syncUpdatedAt;
    delete clean._syncUpdatedAtMillis;
    Object.keys(clean).forEach(key => {
        if (key.startsWith("_local") || clean[key] === undefined) delete clean[key];
    });
    return clean;
}

function snapshotRecord(item) {
    const data = item.data() || {};
    const stamp = data.syncUpdatedAt;
    const syncMillis = stamp && typeof stamp.toMillis === "function"
        ? Number(stamp.toMillis()) || 0
        : Number(data._syncUpdatedAtMillis) || 0;
    const record = { ...data, _firestoreDocumentId: item.id };
    delete record.syncUpdatedAt;
    if (syncMillis > 0) record._syncUpdatedAtMillis = syncMillis;
    return record;
}

function maxSyncMillis(records, fallback = 0) {
    return (Array.isArray(records) ? records : []).reduce((max, record) => {
        const value = Number(record?._syncUpdatedAtMillis) || 0;
        return Math.max(max, value);
    }, Number(fallback) || 0);
}

function safeCursorMillis(value) {
    const n = Number(value) || 0;
    return n > 1 ? Math.max(0, n - 1) : 0;
}

function installFingerprints(storeKey, rows) {
    (Array.isArray(rows) ? rows : []).forEach(record => {
        try {
            const id = operationalDocumentId(storeKey, record);
            operationalFingerprints.set(`${storeKey}/${id}`, fingerprint(cleanOperationalRecord(record)));
        } catch (error) {
            console.warn("Fingerprint operasional dilewati:", error);
        }
    });
}

function mergeRuntimeRecords(storeKey, current, incoming) {
    const map = new Map();
    (Array.isArray(current) ? current : []).forEach(record => {
        try { map.set(operationalDocumentId(storeKey, record), record); }
        catch {}
    });
    (Array.isArray(incoming) ? incoming : []).forEach(record => {
        try { map.set(operationalDocumentId(storeKey, record), record); }
        catch {}
    });
    return [...map.values()];
}

function movementTime(record) {
    return new Date(record?.at || 0).getTime() || Number(record?._syncUpdatedAtMillis) || 0;
}

function rebuildStockIndex(
    movementList = operationalState.get(STORE_KEYS.movements) || [],
    activeStockList = operationalState.get(ACTIVE_STOCK_KEY) || []
) {
    if (Array.isArray(activeStockList) && activeStockList.length) {
        stockIndex = new Map(activeStockList
            .filter(record => norm(record?.productCode))
            .map(record => [norm(record.productCode), num(record.quantity)]));
        stockRevision += 1;
        materializedMasterCache = null;
        return;
    }
    const next = new Map();
    for (const movement of Array.isArray(movementList) ? movementList : []) {
        const code = norm(movement?.productCode);
        if (!code) continue;
        const previous = next.get(code) || { quantity: 0, latestAt: -1, hasAbsolute: false };
        const at = movementTime(movement);
        const hasAbsolute = Number.isFinite(Number(movement?.stockAfter));
        if (hasAbsolute && (!previous.hasAbsolute || at >= previous.latestAt)) {
            next.set(code, { quantity: num(movement.stockAfter), latestAt: at, hasAbsolute: true });
        } else if (!previous.hasAbsolute) {
            previous.quantity += num(movement?.delta);
            previous.latestAt = Math.max(previous.latestAt, at);
            next.set(code, previous);
        }
    }
    stockIndex = new Map([...next].map(([code, value]) => [code, num(value.quantity)]));
    stockRevision += 1;
    materializedMasterCache = null;
}

async function readFullOperationalCollection(storeKey) {
    const collectionKey = OPERATIONAL_COLLECTIONS[storeKey];
    if (!collectionKey) return [];
    const ref = collection(firebaseDb, ...collectionSegments(collectionKey));
    const hasCentralStock = (operationalState.get(ACTIVE_STOCK_KEY) || []).length > 0;
    const cap = storeKey === ACTIVE_STOCK_KEY
        ? null
        : (storeKey === STORE_KEYS.movements && !hasCentralStock ? null : INITIAL_READ_LIMITS[storeKey]);
    const request = cap
        ? query(ref, orderBy("syncUpdatedAt", "desc"), limit(cap))
        : ref;
    const snapshot = await getDocsFromServer(request);
    return snapshot.docs.map(snapshotRecord);
}

async function readIncrementalOperationalCollection(storeKey, lastSyncMillis) {
    const collectionKey = OPERATIONAL_COLLECTIONS[storeKey];
    if (!collectionKey) return [];
    const cursor = Math.max(0, Number(lastSyncMillis) || 0);
    const ref = collection(firebaseDb, ...collectionSegments(collectionKey));
    const request = query(
        ref,
        where("syncUpdatedAt", ">", Timestamp.fromMillis(cursor)),
        orderBy("syncUpdatedAt", "asc")
    );
    const snapshot = await getDocsFromServer(request);
    return snapshot.docs.map(snapshotRecord);
}

async function loadOperationalCache(keys) {
    const result = {};
    for (const storeKey of keys) {
        const cacheStore = OPERATIONAL_CACHE_STORES[storeKey];
        const rows = await readOperationalRecords(cacheStore).catch(error => {
            console.warn(`Cache operasional ${cacheStore} tidak dapat dibaca:`, error);
            return [];
        });
        operationalState.set(storeKey, rows);
        installFingerprints(storeKey, rows);
        if (storeKey === STORE_KEYS.movements || storeKey === ACTIVE_STOCK_KEY) rebuildStockIndex();
        result[storeKey] = rows.length;
    }
    window.dispatchEvent(new CustomEvent("kasirpro:operational-cache-ready", {
        detail: { page: pageMode(), counts: result }
    }));
    return result;
}

async function syncOneOperationalCollection(storeKey, { forceFull = false } = {}) {
    const cacheStore = OPERATIONAL_CACHE_STORES[storeKey];
    const meta = await getOperationalSyncMeta(cacheStore);
    const cached = operationalState.get(storeKey) || [];
    const cacheLooksMissing = meta.seeded === true && Number(meta.recordCount) > 0 && cached.length === 0;
    const fullSeed = forceFull || meta.seeded !== true || cacheLooksMissing;

    if (fullSeed) {
        const rows = await readFullOperationalCollection(storeKey);
        await replaceOperationalRecords(cacheStore, rows);
        operationalState.set(storeKey, rows);
        installFingerprints(storeKey, rows);
        if (storeKey === STORE_KEYS.movements) rebuildStockIndex(rows);
        const maxMillis = maxSyncMillis(rows, 0);
        const nextMeta = await setOperationalSyncMeta(cacheStore, {
            seeded: true,
            lastSyncMillis: safeCursorMillis(maxMillis),
            lastFullSeedAt: new Date().toISOString(),
            lastIncrementalSyncAt: null,
            recordCount: rows.length
        });
        return {
            storeKey,
            cacheStore,
            mode: "full-seed",
            readCount: rows.length,
            changedCount: rows.length,
            recordCount: rows.length,
            lastSyncMillis: nextMeta.lastSyncMillis
        };
    }

    const incoming = await readIncrementalOperationalCollection(storeKey, meta.lastSyncMillis || 0);
    const merged = mergeRuntimeRecords(storeKey, cached, incoming);
    if (incoming.length) await mergeOperationalRecords(cacheStore, incoming);
    operationalState.set(storeKey, merged);
    installFingerprints(storeKey, incoming);
    if ((storeKey === STORE_KEYS.movements || storeKey === ACTIVE_STOCK_KEY) && incoming.length) rebuildStockIndex();
    const maxMillis = maxSyncMillis(incoming, meta.lastSyncMillis || 0);
    const nextMeta = await setOperationalSyncMeta(cacheStore, {
        seeded: true,
        lastSyncMillis: safeCursorMillis(maxMillis),
        lastIncrementalSyncAt: new Date().toISOString(),
        recordCount: merged.length
    });
    return {
        storeKey,
        cacheStore,
        mode: "incremental",
        readCount: incoming.length,
        changedCount: incoming.length,
        recordCount: merged.length,
        lastSyncMillis: nextMeta.lastSyncMillis
    };
}

async function initializeOperationalState() {
    const mode = pageMode();
    const keys = operationalKeysForPage(mode);
    operationalSyncState = { mode: "initializing", page: mode, ready: false, collections: {} };

    if (!keys.length) {
        operationalRemoteReady = false;
        operationalSyncState = { mode: "login-page", page: mode, ready: false, collections: {} };
        return { ready: false, reason: "login-page", sync: clone(operationalSyncState) };
    }

    operationalFingerprints.clear();
    await loadOperationalCache(keys);

    const user = await waitForFirebaseUser();
    if (!user) {
        operationalRemoteReady = false;
        operationalSyncState = { mode: "cache-only", page: mode, ready: false, reason: "firebase-user-unavailable", collections: {} };
        window.dispatchEvent(new CustomEvent("kasirpro:operational-unavailable", {
            detail: { page: mode, reason: "firebase-user-unavailable", cacheAvailable: true }
        }));
        return { ready: false, page: mode, reason: "firebase-user-unavailable", cacheAvailable: true, sync: clone(operationalSyncState) };
    }

    const results = {};
    const failures = [];
    for (const key of keys) {
        try {
            const result = await syncOneOperationalCollection(key);
            results[key] = result;
            emitOperationalSync({ phase: result.mode, page: mode, ...result });
        } catch (error) {
            const failure = { storeKey: key, error: error?.message || String(error) };
            failures.push(failure);
            results[key] = { mode: "cache-kept", ...failure };
            console.warn(`Sinkronisasi incremental ${key} gagal; cache lokal dipertahankan:`, error);
        }
    }

    if (!failures.length) {
        try {
            const migration = await seedCentralStockFromMovements();
            results.activeStockMigration = migration;
        } catch (error) {
            const failure = { storeKey: ACTIVE_STOCK_KEY, error: error?.message || String(error) };
            failures.push(failure);
            results.activeStockMigration = { mode: "failed", ...failure };
            console.warn("Migrasi StokAktif gagal; operasi stok dinonaktifkan:", error);
        }
    }

    operationalRemoteReady = failures.length === 0;
    operationalSyncState = {
        mode: operationalRemoteReady ? "incremental-ready" : "cache-only",
        page: mode,
        ready: operationalRemoteReady,
        collections: results,
        failures
    };

    if (operationalRemoteReady) {
        window.dispatchEvent(new CustomEvent("kasirpro:operational-ready", {
            detail: { page: mode, collections: keys.length, syncMode: "indexeddb-incremental" }
        }));
    } else {
        window.dispatchEvent(new CustomEvent("kasirpro:operational-unavailable", {
            detail: { page: mode, cacheAvailable: true, failures }
        }));
    }

    return {
        ready: operationalRemoteReady,
        page: mode,
        collections: keys.length,
        cacheAvailable: true,
        sync: clone(operationalSyncState)
    };
}

function movementStock(code, movementList = operationalState.get(STORE_KEYS.movements) || []) {
    return stockIndex.get(norm(code)) || 0;
}

export function readCurrentStock(code) {
    return movementStock(code);
}

export async function reloadMasterCache() {
    await refreshMasterCacheFromIndexedDb();
    rebuildStockIndex();
    return masterWithOperationalStock(masterCache);
}

function masterWithOperationalStock(master) {
    const revision = `${masterRevision}:${stockRevision}`;
    if (materializedMasterCache && materializedRevision === revision) return clone(materializedMasterCache);
    const out = clone(master || emptyMaster());
    out.produk = (Array.isArray(out.produk) ? out.produk : []).map(product => ({
        ...product,
        "Stok Awal": movementStock(product?.["Kode Produk"])
    }));
    materializedMasterCache = out;
    materializedRevision = revision;
    return clone(materializedMasterCache);
}

function stripOperationalStockFromMaster(incoming, currentStatic) {
    const next = clone(incoming || emptyMaster());
    const currentProducts = Array.isArray(currentStatic?.produk) ? currentStatic.produk : [];
    const currentMap = new Map(currentProducts.map(product => [norm(product?.["Kode Produk"]), product]));
    next.produk = (Array.isArray(next.produk) ? next.produk : []).map(product => {
        const clean = { ...product };
        const existing = currentMap.get(norm(clean?.["Kode Produk"]));
        clean["Stok Awal"] = existing ? existing["Stok Awal"] : 0;
        return clean;
    });
    return next;
}

function collectOperationalChanges(entries) {
    const changes = [];
    for (const entry of entries) {
        const rows = Array.isArray(entry.value) ? entry.value : [];
        const collectionKey = OPERATIONAL_COLLECTIONS[entry.key];
        if (!collectionKey) continue;
        for (const record of rows) {
            const clean = cleanOperationalRecord(record);
            const id = operationalDocumentId(entry.key, record);
            const fpKey = `${entry.key}/${id}`;
            const nextFingerprint = fingerprint(clean);
            if (operationalFingerprints.get(fpKey) === nextFingerprint) continue;
            changes.push({ storeKey: entry.key, collectionKey, id, clean, fpKey, nextFingerprint });
        }
    }
    return changes;
}

function attachDocumentIds(storeKey, rows) {
    return (Array.isArray(rows) ? rows : []).map(record => {
        if (record?._firestoreDocumentId) return record;
        try { return { ...record, _firestoreDocumentId: operationalDocumentId(storeKey, record) }; }
        catch { return record; }
    });
}

async function cacheCommittedChanges(changes) {
    const grouped = new Map();
    changes.forEach(change => {
        const list = grouped.get(change.storeKey) || [];
        list.push({ ...change.clean, _firestoreDocumentId: change.id });
        grouped.set(change.storeKey, list);
    });

    for (const [storeKey, rows] of grouped.entries()) {
        const cacheStore = OPERATIONAL_CACHE_STORES[storeKey];
        await mergeOperationalRecords(cacheStore, rows);
        const current = attachDocumentIds(storeKey, operationalState.get(storeKey) || []);
        operationalState.set(storeKey, current);
        const meta = await getOperationalSyncMeta(cacheStore);
        await setOperationalSyncMeta(cacheStore, {
            seeded: meta.seeded === true,
            recordCount: current.length
        });
    }
}

async function seedCentralStockFromMovements() {
    const existing = operationalState.get(ACTIVE_STOCK_KEY) || [];
    const movements = operationalState.get(STORE_KEYS.movements) || [];
    if (existing.length || !movements.length) return { seeded: false, count: existing.length };

    rebuildStockIndex(movements, []);
    const productNames = new Map(movements.map(record => [norm(record?.productCode), String(record?.productName || "").trim()]));
    const now = new Date().toISOString();
    const rows = [...stockIndex.entries()].map(([code, quantity]) => ({
        productCode: code,
        productName: productNames.get(code) || code,
        quantity,
        updatedAt: now
    }));

    for (let start = 0; start < rows.length; start += 400) {
        const batch = writeBatch(firebaseDb);
        rows.slice(start, start + 400).forEach(record => batch.set(
            doc(firebaseDb, ...documentSegments("activeStocks", operationalDocumentId(ACTIVE_STOCK_KEY, record))),
            { ...record, syncUpdatedAt: serverTimestamp() },
            { merge: true }
        ));
        await batch.commit();
    }

    const cached = attachDocumentIds(ACTIVE_STOCK_KEY, rows);
    await replaceOperationalRecords(OPERATIONAL_CACHE_STORES[ACTIVE_STOCK_KEY], cached);
    operationalState.set(ACTIVE_STOCK_KEY, cached);
    installFingerprints(ACTIVE_STOCK_KEY, cached);
    await setOperationalSyncMeta(OPERATIONAL_CACHE_STORES[ACTIVE_STOCK_KEY], {
        seeded: true,
        lastFullSeedAt: now,
        recordCount: cached.length
    });
    rebuildStockIndex();
    return { seeded: true, count: cached.length };
}

function guardedOperationalRecords(entries) {
    const guards = [];
    for (const entry of entries) {
        for (const record of Array.isArray(entry?.value) ? entry.value : []) {
            if (entry.key === STORE_KEYS.sales && norm(record?.status) === "void") {
                guards.push({
                    kind: "void-sale",
                    ref: doc(firebaseDb, ...documentSegments("sales", operationalDocumentId(entry.key, record)))
                });
            }
            if (entry.key === STORE_KEYS.invoices && record?.stockApplied === true) {
                guards.push({
                    kind: "confirm-invoice",
                    ref: doc(firebaseDb, ...documentSegments("purchaseInvoices", operationalDocumentId(entry.key, record)))
                });
            }
        }
    }
    return guards;
}

/**
 * Menulis perubahan stok dan dokumen operasional dalam satu transaksi Firestore.
 * Mutasi stok menjadi sumber audit, sedangkan StokAktif menjadi sumber kuantitas
 * terserialisasi untuk mencegah overselling lintas perangkat.
 */
export async function writeStockTransaction(entries) {
    if (!operationalRemoteReady) {
        throw new Error("Data operasional pusat belum siap. Perubahan stok dibatalkan.");
    }

    const sourceEntries = (Array.isArray(entries) ? entries : []).map(entry => ({
        key: entry?.key,
        value: clone(Array.isArray(entry?.records) ? entry.records : (Array.isArray(entry?.value) ? entry.value : []))
    })).filter(entry => Object.prototype.hasOwnProperty.call(OPERATIONAL_COLLECTIONS, entry.key) && entry.key !== ACTIVE_STOCK_KEY);
    const sourceMovements = sourceEntries.find(entry => entry.key === STORE_KEYS.movements)?.value || [];
    if (!sourceMovements.length) return writeOperationalDelta(entries);

    let remoteCommitted = false;
    operationalWriteQueue = operationalWriteQueue.catch(() => {}).then(async () => {
        let committedEntries = null;
        let committedStocks = null;
        let committedChanges = null;

        await runTransaction(firebaseDb, async transaction => {
            const workingEntries = clone(sourceEntries);
            const movements = workingEntries.find(entry => entry.key === STORE_KEYS.movements)?.value || [];
            const codes = [...new Set(movements.map(record => norm(record?.productCode)).filter(Boolean))];
            const stockRefs = new Map(codes.map(code => [
                code,
                doc(firebaseDb, ...documentSegments("activeStocks", readableDocumentId("stok", code)))
            ]));
            const guards = guardedOperationalRecords(workingEntries);
            const reads = await Promise.all([
                ...codes.map(code => transaction.get(stockRefs.get(code))),
                ...guards.map(guard => transaction.get(guard.ref))
            ]);
            const stockSnapshots = new Map(codes.map((code, index) => [code, reads[index]]));
            const guardSnapshots = reads.slice(codes.length);

            guardSnapshots.forEach((snapshot, index) => {
                const guard = guards[index];
                const data = snapshot.exists() ? snapshot.data() : null;
                if (guard.kind === "void-sale" && (!data || norm(data.status) === "void")) {
                    throw new Error("Transaksi penjualan sudah di-VOID atau tidak lagi tersedia.");
                }
                if (guard.kind === "confirm-invoice" && data?.stockApplied === true) {
                    throw new Error("Barang Masuk faktur ini sudah dikonfirmasi dari perangkat lain.");
                }
            });

            const workingStock = new Map();
            const stockNames = new Map();
            for (const movement of movements) {
                const code = norm(movement?.productCode);
                if (!code) throw new Error("Mutasi stok tidak memiliki Kode Produk.");
                const snapshot = stockSnapshots.get(code);
                const before = workingStock.has(code)
                    ? workingStock.get(code)
                    : (snapshot?.exists() ? num(snapshot.data()?.quantity) : num(movement?.stockBefore));
                const isOpname = norm(movement?.type) === "opname";
                const isOpening = norm(movement?.type) === "opening";
                if ((isOpname || isOpening) && before !== num(movement?.stockBefore)) {
                    throw new Error(`Stok ${movement.productCode} berubah dari acuan. Muat ulang data sebelum mencoba lagi.`);
                }
                const after = isOpname ? num(movement?.stockAfter) : before + num(movement?.delta);
                if (!Number.isFinite(after) || after < 0) {
                    throw new Error(`Stok ${movement.productName || movement.productCode} tidak mencukupi.`);
                }
                movement.stockBefore = before;
                movement.stockAfter = after;
                workingStock.set(code, after);
                stockNames.set(code, String(movement?.productName || snapshot?.data()?.productName || code).trim());
            }

            const changes = collectOperationalChanges(workingEntries);
            if (changes.length + workingStock.size > 450) {
                throw new Error("Operasi memiliki terlalu banyak item. Bagi menjadi maksimal 200 produk per transaksi.");
            }
            const now = new Date().toISOString();
            const stockRows = [...workingStock.entries()].map(([code, quantity]) => ({
                productCode: code,
                productName: stockNames.get(code) || code,
                quantity,
                updatedAt: now
            }));

            stockRows.forEach(record => transaction.set(
                stockRefs.get(norm(record.productCode)),
                { ...record, syncUpdatedAt: serverTimestamp() },
                { merge: true }
            ));
            changes.forEach(change => transaction.set(
                doc(firebaseDb, ...documentSegments(change.collectionKey, change.id)),
                { ...change.clean, syncUpdatedAt: serverTimestamp() },
                { merge: true }
            ));

            committedEntries = workingEntries;
            committedStocks = stockRows;
            committedChanges = changes;
        });

        remoteCommitted = true;
        for (const entry of committedEntries || []) {
            const current = operationalState.get(entry.key) || [];
            operationalState.set(entry.key, mergeRuntimeRecords(entry.key, current, attachDocumentIds(entry.key, entry.value)));
        }
        const currentStocks = operationalState.get(ACTIVE_STOCK_KEY) || [];
        const attachedStocks = attachDocumentIds(ACTIVE_STOCK_KEY, committedStocks || []);
        operationalState.set(ACTIVE_STOCK_KEY, mergeRuntimeRecords(ACTIVE_STOCK_KEY, currentStocks, attachedStocks));

        const stockChanges = (committedStocks || []).map(record => {
            const id = operationalDocumentId(ACTIVE_STOCK_KEY, record);
            const clean = cleanOperationalRecord(record);
            return {
                storeKey: ACTIVE_STOCK_KEY,
                collectionKey: "activeStocks",
                id,
                clean,
                fpKey: `${ACTIVE_STOCK_KEY}/${id}`,
                nextFingerprint: fingerprint(clean)
            };
        });
        const allChanges = [...(committedChanges || []), ...stockChanges];
        if (allChanges.length) await cacheCommittedChanges(allChanges);
        allChanges.forEach(change => operationalFingerprints.set(change.fpKey, change.nextFingerprint));
        rebuildStockIndex();
        return {
            mode: "firestore-stock-transaction",
            changedRecords: committedChanges?.length || 0,
            stockDocuments: committedStocks?.length || 0,
            cacheUpdated: true
        };
    });

    return operationalWriteQueue.catch((error) => {
        if (remoteCommitted && error && typeof error === "object") error.remoteCommitted = true;
        throw error;
    });
}

async function commitOperationalEntries(entries) {
    if (!operationalRemoteReady) {
        throw new Error("Data operasional pusat belum siap. Perubahan dibatalkan agar tidak hanya tersimpan di perangkat ini.");
    }

    const list = Array.isArray(entries) ? entries : [];
    const previousState = new Map(list.map(entry => [entry.key, clone(operationalState.get(entry.key) || [])]));
    list.forEach(entry => operationalState.set(entry.key, clone(Array.isArray(entry.value) ? entry.value : [])));
    const changes = collectOperationalChanges(list);

    operationalWriteQueue = operationalWriteQueue.then(async () => {
        try {
            for (let start = 0; start < changes.length; start += 400) {
                const chunk = changes.slice(start, start + 400);
                const batch = writeBatch(firebaseDb);
                chunk.forEach(change => {
                    batch.set(
                        doc(firebaseDb, ...documentSegments(change.collectionKey, change.id)),
                        { ...change.clean, syncUpdatedAt: serverTimestamp() },
                        { merge: true }
                    );
                });
                await batch.commit();
                chunk.forEach(change => operationalFingerprints.set(change.fpKey, change.nextFingerprint));
            }
            if (changes.length) await cacheCommittedChanges(changes);
            return {
                mode: "firestore-operational-delta-indexeddb",
                changedRecords: changes.length,
                batches: Math.ceil(changes.length / 400),
                cacheUpdated: true
            };
        } catch (error) {
            previousState.forEach((value, key) => operationalState.set(key, value));
            throw error;
        }
    });

    return operationalWriteQueue;
}

export async function writeOperationalDelta(entries) {
    if (!operationalRemoteReady) {
        throw new Error("Data operasional pusat belum siap. Perubahan dibatalkan agar tidak hanya tersimpan di perangkat ini.");
    }
    const list = (Array.isArray(entries) ? entries : []).map(entry => ({
        key: entry?.key,
        value: Array.isArray(entry?.records) ? entry.records : []
    })).filter(entry => Object.prototype.hasOwnProperty.call(OPERATIONAL_COLLECTIONS, entry.key) && entry.value.length);
    const changes = collectOperationalChanges(list);
    operationalWriteQueue = operationalWriteQueue.then(async () => {
        for (let start = 0; start < changes.length; start += 400) {
            const chunk = changes.slice(start, start + 400);
            const batch = writeBatch(firebaseDb);
            chunk.forEach(change => batch.set(
                doc(firebaseDb, ...documentSegments(change.collectionKey, change.id)),
                { ...change.clean, syncUpdatedAt: serverTimestamp() },
                { merge: true }
            ));
            await batch.commit();
            chunk.forEach(change => operationalFingerprints.set(change.fpKey, change.nextFingerprint));
        }
        for (const entry of list) {
            const current = operationalState.get(entry.key) || [];
            operationalState.set(entry.key, mergeRuntimeRecords(entry.key, current, attachDocumentIds(entry.key, entry.value)));
        }
        if (changes.length) await cacheCommittedChanges(changes);
        if (list.some(entry => entry.key === STORE_KEYS.movements)) rebuildStockIndex(operationalState.get(STORE_KEYS.movements) || []);
        return { mode: "firestore-operational-delta-at13", changedRecords: changes.length, batches: Math.ceil(changes.length / 400) };
    });
    return operationalWriteQueue;
}

export function initializeDatabase() {
    if (initializationPromise) return initializationPromise;
    initializationPromise = (async () => {
        const syncState = await prepareLocalMaster();
        if (syncState?.hasLocal) await refreshMasterCacheFromIndexedDb();
        else {
            masterCache = emptyMaster();
            localMasterMode = true;
        }
        emitReady(masterCache, syncState);

        const operational = await initializeOperationalState();
        return {
            mode: syncState?.hasLocal ? "indexeddb-master-operational-cache" : "indexeddb-master-empty-operational-cache",
            remoteEnabled: operational.ready,
            operationalRemoteReady: operational.ready,
            operationalCacheEnabled: true,
            operationalSyncMode: operational.sync?.mode || "unknown",
            masterSnapshotNetworkEnabled: FIRESTORE_MASTER_SNAPSHOT_ENABLED,
            masterVersion: await readMasterVersion(),
            centralMasterVersion: Number(syncState?.centralVersion) || 0,
            syncState: syncState?.state || "recovery-unavailable",
            productCount: masterCache.produk?.length || 0,
            operational: operational.sync
        };
    })();
    return initializationPromise;
}

export function readStore(key, fallback) {
    if (key === STORE_KEYS.master && localMasterMode) {
        return masterWithOperationalStock(masterCache ?? fallback ?? emptyMaster());
    }
    if (Object.prototype.hasOwnProperty.call(OPERATIONAL_COLLECTIONS, key)) {
        return clone(operationalState.has(key) ? operationalState.get(key) : (fallback ?? []));
    }
    return legacy.readStore(key, fallback);
}

export function writeStore(key, value) {
    if (key === STORE_KEYS.master && localMasterMode) {
        const staticMaster = stripOperationalStockFromMaster(value, masterCache);
        masterCache = clone(staticMaster);
        masterRevision += 1;
        materializedMasterCache = null;
        return (async () => {
            const currentVersion = Number(await readMasterVersion()) || 0;
            const installed = await installMasterSnapshot(masterCache, {
                version: currentVersion + 1,
                source: "database-facade-write"
            });
            masterCache = clone(installed.master);
            masterRevision += 1;
            materializedMasterCache = null;
            window.dispatchEvent(new CustomEvent("kasirpro:local-master-updated", {
                detail: {
                    version: installed.version,
                    products: masterCache.produk?.length || 0,
                    suppliers: masterCache.supplier?.length || 0,
                    categories: masterCache.kategori?.length || 0
                }
            }));
            return { mode: "indexeddb-master", version: installed.version };
        })();
    }

    if (Object.prototype.hasOwnProperty.call(OPERATIONAL_COLLECTIONS, key)) {
        return commitOperationalEntries([{ key, value }]);
    }
    return legacy.writeStore(key, value);
}

export async function writeStoreBundle(entries) {
    const list = Array.isArray(entries) ? entries : [];
    const masterEntry = list.find(entry => entry?.key === STORE_KEYS.master);
    const operationalEntries = list.filter(entry => Object.prototype.hasOwnProperty.call(OPERATIONAL_COLLECTIONS, entry?.key));

    const previousMaster = clone(masterCache);
    try {
        const results = [];
        if (operationalEntries.length) results.push(await commitOperationalEntries(operationalEntries));
        if (masterEntry) {
            const staticMaster = stripOperationalStockFromMaster(masterEntry.value, masterCache);
            results.push(await writeStore(STORE_KEYS.master, staticMaster));
        }
        return { mode: "indexeddb-master-firestore-operational-at09", results };
    } catch (error) {
        masterCache = previousMaster;
        throw error;
    }
}

export function isLocalMasterMode() { return localMasterMode; }
export function isOperationalRemoteReady() { return operationalRemoteReady; }
export function getMasterSyncState() { return clone(lastMasterSyncState); }
export function getOperationalSyncState() { return clone(operationalSyncState); }

export async function recheckMasterSync() {
    const state = await inspectMasterSyncState({ allowNetwork: FIRESTORE_MASTER_SNAPSHOT_ENABLED });
    emitSyncState(state);
    return state;
}

export async function forceOperationalFullResync() {
    const mode = pageMode();
    const keys = operationalKeysForPage(mode);
    const user = await waitForFirebaseUser();
    if (!user) throw new Error("Sesi Firebase belum tersedia untuk full resync operasional.");

    const results = [];
    for (const key of keys) {
        const cacheStore = OPERATIONAL_CACHE_STORES[key];
        await resetOperationalCollection(cacheStore);
        operationalState.set(key, []);
        const result = await syncOneOperationalCollection(key, { forceFull: true });
        results.push(result);
    }
    operationalRemoteReady = true;
    operationalSyncState = {
        mode: "manual-full-resync",
        page: mode,
        ready: true,
        collections: Object.fromEntries(results.map(item => [item.storeKey, item])),
        failures: []
    };
    window.dispatchEvent(new CustomEvent("kasirpro:operational-ready", {
        detail: { page: mode, collections: keys.length, syncMode: "manual-full-resync" }
    }));
    return clone(operationalSyncState);
}
