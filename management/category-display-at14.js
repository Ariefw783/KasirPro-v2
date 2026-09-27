/* KasirPro AT-14A — Category Display Reliability
 * Display-only patch.
 * - Internal category codes remain unchanged.
 * - Product table and category filter show Nama Kategori after local/central master readiness.
 * - No database writes and no schema changes.
 */

import { readStore } from "../modules/database/database-store.js";

const MASTER_KEY = "kasirpro_master_store_v1";
let refreshTimer = null;

const text = value => String(value ?? "").trim();
const norm = value => text(value).toLowerCase();

function categoryMap() {
    const master = readStore(MASTER_KEY, {}) || {};
    const rows = Array.isArray(master?.kategori) ? master.kategori : [];
    const map = new Map();

    for (const row of rows) {
        const code = text(row?.["Kode Kategori"] || row?.["Kategori"]);
        const name = text(row?.["Nama Kategori"] || row?.["Kategori"] || row?.["Kode Kategori"]);
        if (!name) continue;
        if (code) map.set(norm(code), name);
        map.set(norm(name), name);
    }
    return map;
}

function categoryColumnIndex(table) {
    const headers = [...table.querySelectorAll("thead th")];
    return headers.findIndex(th => {
        const label = norm(th.textContent);
        return label.includes("kategori") || label.includes("category");
    });
}

function updateProductTable(map) {
    const body = document.getElementById("products-table-body");
    const table = body?.closest("table");
    if (!body || !table || !map.size) return;

    const index = categoryColumnIndex(table);
    if (index < 0) return;

    body.querySelectorAll("tr").forEach(row => {
        const cells = [...row.children].filter(cell => cell.tagName === "TD");
        const cell = cells[index];
        if (!cell || cell.hasAttribute("colspan")) return;

        const current = text(cell.textContent);
        const display = map.get(norm(current));
        if (!display || display === current) return;

        if (cell.childElementCount === 0) {
            cell.textContent = display;
            cell.title = display;
            return;
        }

        const leaf = [...cell.querySelectorAll("span,strong,small,div")]
            .find(node => node.childElementCount === 0 && map.has(norm(node.textContent)));
        if (leaf) {
            const name = map.get(norm(leaf.textContent));
            leaf.textContent = name;
            leaf.title = name;
        }
    });
}

function updateCategoryFilters(map) {
    if (!map.size) return;
    document.querySelectorAll("select").forEach(select => {
        const identity = norm(`${select.id} ${select.name} ${select.getAttribute("aria-label") || ""}`);
        const categoryLike = identity.includes("kategori") || identity.includes("category") ||
            [...select.options].some(option => map.has(norm(option.value)));
        if (!categoryLike) return;

        [...select.options].forEach(option => {
            const display = map.get(norm(option.value)) || map.get(norm(option.textContent));
            if (display) option.textContent = display;
        });
    });
}

function refreshCategoryDisplay() {
    const map = categoryMap();
    if (!map.size) return;
    updateProductTable(map);
    updateCategoryFilters(map);
}

function scheduleRefresh(delay = 60) {
    clearTimeout(refreshTimer);
    refreshTimer = setTimeout(() => {
        refreshTimer = null;
        refreshCategoryDisplay();
    }, delay);
}

function scheduleBurst() {
    scheduleRefresh(0);
    [120, 350, 900].forEach(delay => setTimeout(refreshCategoryDisplay, delay));
}

[
    "kasirpro:database-preload-progress",
    "kasirpro:master-sync-state",
    "kasirpro:local-master-updated",
    "kasirpro:database-ready",
    "kasirpro:database-synced",
    "kasirpro:operational-cache-ready",
    "kasirpro:operational-ready"
].forEach(name => window.addEventListener(name, scheduleBurst));

document.addEventListener("click", event => {
    if (event.target?.closest?.("#products-table-body, select, button")) scheduleRefresh();
}, true);

window.addEventListener("pageshow", scheduleBurst);

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", scheduleBurst, { once: true });
} else {
    scheduleBurst();
}

