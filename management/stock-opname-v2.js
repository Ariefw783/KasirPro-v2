/* KasirPro - Stock Opname V2
 * Kontrak:
 * - Pencocokan hanya berdasarkan Kode Produk.
 * - Nama Produk tidak pernah dipakai sebagai fallback pencocokan.
 * - Stok sistem berasal dari MutasiStok, bukan Stok Awal Master.
 * - Preview menyimpan stok acuan; sebelum commit stok direvalidasi.
 * - Adjustment menulis MutasiStok + riwayat StockOpname, tanpa mengubah Master statis.
 */

import {
    STORE_KEYS,
    readStore,
    writeStockTransaction,
    readCurrentStock
} from "../modules/database/database-store.js";

let opnamePreviewV2 = [];
let opnameIssuesV2 = [];

const $ = (id) => document.getElementById(id);
const text = (value) => String(value ?? "").trim();
const norm = (value) => text(value).toLowerCase();
const num = (value) => Number(String(value ?? 0).replace(/[^0-9.-]/g, "")) || 0;
function physicalQuantity(value) {
    if (typeof value === "number") return Number.isSafeInteger(value) && value >= 0 ? value : null;
    const raw = text(value);
    if (!/^\d+$/.test(raw)) return null;
    const parsed = Number(raw);
    return Number.isSafeInteger(parsed) ? parsed : null;
}
const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
}[char]));
const nowIso = () => new Date().toISOString();
const uid = (prefix) => `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
const rupiah = (value) => new Intl.NumberFormat("id-ID", {
    style: "currency",
    currency: "IDR",
    maximumFractionDigits: 0
}).format(num(value));

function getMaster() {
    return readStore(STORE_KEYS.master, {
        produk: [], supplier: [], kategori: [], pengguna: [], pengaturan_toko: []
    });
}

function getMovements() {
    return readStore(STORE_KEYS.movements, []);
}

function getOpnames() {
    return readStore(STORE_KEYS.opnames, []);
}

function products(master = getMaster()) {
    return Array.isArray(master?.produk) ? master.produk : [];
}

function findProductByCode(code, master = getMaster()) {
    const key = norm(code);
    if (!key) return null;
    return products(master).find((product) => norm(product?.["Kode Produk"]) === key) || null;
}

function currentStockByCode(code) {
    return readCurrentStock(code);
}

function dateOnly(value) {
    if (!value) return "";
    const date = value instanceof Date ? value : new Date(value);
    return Number.isNaN(date.getTime()) ? text(value) : date.toISOString().slice(0, 10);
}

function sheetRows(workbook, names) {
    const wanted = names.map((name) => norm(name).replace(/\s+/g, "_"));
    const sheetName = workbook.SheetNames.find((name) =>
        wanted.includes(norm(name).replace(/\s+/g, "_"))
    );
    return sheetName
        ? window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" })
        : [];
}

async function readWorkbook(file) {
    if (!file) throw new Error("Pilih file Stock Opname terlebih dahulu.");
    if (!window.XLSX?.read) throw new Error("SheetJS/XLSX belum tersedia.");
    const buffer = await file.arrayBuffer();
    return window.XLSX.read(buffer, { type: "array", cellDates: true });
}

function exportStockOpnameV2() {
    if (!window.XLSX?.utils || typeof window.XLSX.writeFile !== "function") {
        return alert("SheetJS/XLSX belum tersedia.");
    }

    const master = getMaster();
    const rows = products(master).map((product) => ({
        "Kode Produk": text(product?.["Kode Produk"]),
        "Nama Produk": text(product?.["Nama Produk"]),
        "Kategori": text(product?.["Kategori"]),
        "Supplier": text(product?.["Supplier"]),
        "Satuan": text(product?.["Satuan"]),
        "Stok Sistem": currentStockByCode(product?.["Kode Produk"]),
        "Stok Fisik": "",
        "Selisih": "",
        "Harga Beli": num(product?.["Harga Beli"]),
        "Nilai Selisih": "",
        "Lokasi Rak": text(product?.["Lokasi Rak"]),
        "Batch": text(product?.["Batch"]),
        "Tanggal Expired": dateOnly(product?.["Tanggal Expired"]),
        "Keterangan": ""
    }));

    const workbook = window.XLSX.utils.book_new();
    const sheet = window.XLSX.utils.json_to_sheet(rows);
    sheet["!cols"] = [
        { wch: 16 }, { wch: 32 }, { wch: 18 }, { wch: 16 }, { wch: 12 },
        { wch: 13 }, { wch: 13 }, { wch: 12 }, { wch: 14 }, { wch: 16 },
        { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 30 }
    ];
    window.XLSX.utils.book_append_sheet(workbook, sheet, "STOCK_OPNAME");
    window.XLSX.writeFile(workbook, `Stock_Opname_KasirPro_${new Date().toISOString().slice(0, 10)}.xlsx`, { compression: true });
}

async function previewOpnameImportV2() {
    const workbook = await readWorkbook($("opname-import-file")?.files?.[0]);
    const rows = sheetRows(workbook, ["stock_opname", "stock opname"]);
    const master = getMaster();
    const preview = [];
    const issues = [];

    rows.forEach((row, index) => {
        const code = text(row?.["Kode Produk"]);
        const physicalRaw = row?.["Stok Fisik"];

        if (!code) {
            issues.push(`Baris ${index + 2}: Kode Produk wajib diisi.`);
            return;
        }

        const product = findProductByCode(code, master);
        if (!product) {
            issues.push(`Baris ${index + 2}: Kode Produk ${code} tidak ditemukan pada Master Aktif.`);
            return;
        }

        if (physicalRaw === "" || physicalRaw === null || physicalRaw === undefined) {
            return;
        }

        const physical = physicalQuantity(physicalRaw);
        if (physical === null) {
            issues.push(`Baris ${index + 2}: Stok Fisik ${code} harus bilangan bulat tidak negatif (tanpa pemisah ribuan).`);
            return;
        }

        const system = currentStockByCode(code);
        const diff = physical - system;
        const buyPrice = num(product?.["Harga Beli"]);
        preview.push({
            sourceRow: index + 2,
            productCode: text(product?.["Kode Produk"]),
            productName: text(product?.["Nama Produk"]),
            system,
            physical,
            diff,
            buyPrice,
            valueDiff: diff * buyPrice,
            note: text(row?.["Keterangan"]),
            previewedAt: nowIso()
        });
    });

    const duplicateCodes = new Map();
    preview.forEach((item) => {
        const key = norm(item.productCode);
        duplicateCodes.set(key, (duplicateCodes.get(key) || 0) + 1);
    });
    duplicateCodes.forEach((count, key) => {
        if (count > 1) issues.push(`Kode Produk ${key.toUpperCase()} muncul ${count} kali pada file opname.`);
    });

    opnamePreviewV2 = preview;
    opnameIssuesV2 = issues;

    const summary = $("opname-import-summary");
    if (summary) {
        summary.hidden = false;
        summary.innerHTML = `<strong>${preview.length} produk</strong> cocok berdasarkan Kode Produk. ` +
            `Total selisih qty: ${preview.reduce((sum, item) => sum + item.diff, 0)}.` +
            (issues.length ? ` <strong>${issues.length} masalah harus diperbaiki.</strong>` : "");
    }

    const body = $("opname-import-preview-body");
    if (body) {
        body.innerHTML = preview.length
            ? preview.map((item) => `<tr>
                <td>${escapeHtml(item.productCode)}</td>
                <td>${escapeHtml(item.productName)}</td>
                <td>${item.system}</td>
                <td>${item.physical}</td>
                <td class="${item.diff >= 0 ? "qty-plus" : "qty-minus"}">${item.diff >= 0 ? "+" : ""}${item.diff}</td>
                <td>${rupiah(item.valueDiff)}</td>
            </tr>`).join("")
            : '<tr><td colspan="6" class="empty-table-state">Tidak ada baris opname valid.</td></tr>';
    }

    if ($("opname-import-preview")) $("opname-import-preview").hidden = false;
    if ($("opname-import-actions")) $("opname-import-actions").hidden = false;

    if (issues.length) {
        alert(`Preview Stock Opname menemukan masalah:\n\n${issues.slice(0, 10).join("\n")}${issues.length > 10 ? `\nDan ${issues.length - 10} masalah lainnya.` : ""}`);
    }
}

async function confirmOpnameImportV2() {
    if (!opnamePreviewV2.length) return alert("Baca dan preview file Stock Opname terlebih dahulu.");
    if (opnameIssuesV2.length) return alert("Stock Opname belum dapat diterapkan karena masih ada masalah validasi.");

    const currentMovements = getMovements();
    const stale = opnamePreviewV2.find((item) =>
        currentStockByCode(item.productCode) !== item.system
    );
    if (stale) {
        opnamePreviewV2 = [];
        return alert(
            `Stok ${stale.productCode} berubah sejak preview.\n` +
            `Stock Opname dibatalkan agar adjustment tidak memakai stok lama.\n\n` +
            `Silakan Baca & Preview ulang file.`
        );
    }

    const changed = opnamePreviewV2.filter((item) => item.diff !== 0);
    if (changed.length > 200) return alert("Terlalu banyak perubahan stok dalam satu transaksi. Bagi hasil opname menjadi maksimal 200 produk per file.");
    if (!changed.length) {
        opnamePreviewV2 = [];
        opnameIssuesV2 = [];
        return alert("Tidak ada selisih stok yang perlu disesuaikan.");
    }

    if (!confirm(
        `Konfirmasi Stock Opname?\n\n` +
        `${changed.length} produk akan disesuaikan ke Stok Fisik.\n` +
        `Pencocokan hanya berdasarkan Kode Produk.`
    )) return;

    const movements = [...currentMovements];
    const history = getOpnames();
    const reference = `OPN-${Date.now()}`;
    const at = nowIso();

    changed.forEach((item) => {
        movements.unshift({
            id: uid("MOV"),
            at,
            type: "opname",
            reference,
            productCode: item.productCode,
            productName: item.productName,
            delta: item.diff,
            stockBefore: item.system,
            stockAfter: item.physical,
            note: item.note || "Adjustment Stock Opname"
        });
    });

    history.unshift({
        id: reference,
        at,
        productCount: changed.length,
        qtyDiff: changed.reduce((sum, item) => sum + item.diff, 0),
        valueDiff: changed.reduce((sum, item) => sum + item.valueDiff, 0),
        source: "stock-opname-v2",
        matching: "kode-produk-only"
    });

    try {
        await writeStockTransaction([
            { key: STORE_KEYS.movements, records: movements.slice(0, changed.length) },
            { key: STORE_KEYS.opnames, records: [history[0]] }
        ]);
    } catch (error) {
        console.error("Stock Opname V2 gagal disimpan:", error);
        return alert("Stock Opname belum dapat disimpan. Tidak ada Master statis yang diubah.");
    }

    opnamePreviewV2 = [];
    opnameIssuesV2 = [];
    alert(`Stock Opname ${reference} berhasil diterapkan melalui MutasiStok.`);
    document.querySelector('[data-view="stock-opname"]')?.click();
}

async function captureOpnameActions(event) {
    const exportButton = event.target.closest("#opname-export-button");
    const readButton = event.target.closest("#read-opname-import");
    const confirmButton = event.target.closest("#confirm-opname-import");

    if (exportButton) {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        exportStockOpnameV2();
        return;
    }

    if (readButton) {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        readButton.disabled = true;
        try { await previewOpnameImportV2(); }
        catch (error) { console.error(error); alert(error?.message || "Gagal membaca Stock Opname."); }
        finally { readButton.disabled = false; }
        return;
    }

    if (confirmButton) {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        confirmButton.disabled = true;
        try { await confirmOpnameImportV2(); }
        catch (error) { console.error(error); alert(error?.message || "Stock Opname gagal."); }
        finally { confirmButton.disabled = false; }
    }
}

document.addEventListener("click", captureOpnameActions, true);
document.getElementById("opname-import-file")?.addEventListener("change", () => {
    opnamePreviewV2 = [];
    opnameIssuesV2 = [];
    if ($("opname-import-preview")) $("opname-import-preview").hidden = true;
    if ($("opname-import-actions")) $("opname-import-actions").hidden = true;
});

window.KasirProStockOpnameV2 = Object.freeze({
    exportStockOpnameV2,
    previewOpnameImportV2,
    confirmOpnameImportV2,
    currentStockByCode
});
