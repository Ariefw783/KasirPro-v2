/* KasirPro - Integrasi Faktur Pembelian V2
 * Kontrak:
 * - Supplier diterima berdasarkan nama perusahaan atau ID/kode Master Supplier.
 * - Import faktur selalu Draft dan TIDAK mengubah stok.
 * - Produk existing dicocokkan hanya dengan Kode Produk.
 * - Produk tidak dikenal ditandai Produk Baru dan wajib direview Admin saat Barang Masuk.
 * - Produk baru hanya wajib Nama Produk + Supplier; kode kosong dibuat ADD0001 dst.
 * - Stok operasional dihitung dari MutasiStok, bukan Stok Awal Master.
 */

import {
    STORE_KEYS,
    readStore,
    writeStore,
    writeStockTransaction,
    readCurrentStock
} from "../modules/database/database-store.js";

let invoicePreviewV2 = [];
let invoiceValidationV2 = [];

const $ = (id) => document.getElementById(id);
const text = (value) => String(value ?? "").trim();
const norm = (value) => text(value).toLowerCase();
const num = (value) => Number(String(value ?? 0).replace(/[^0-9.-]/g, "")) || 0;
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

function dateOnly(value) {
    if (!value) return "";
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? text(value) : parsed.toISOString().slice(0, 10);
}

function getMaster() {
    return readStore(STORE_KEYS.master, {
        produk: [], supplier: [], kategori: [], pengguna: [], pengaturan_toko: []
    });
}

function getInvoices() {
    return readStore(STORE_KEYS.invoices, []);
}

function getMovements() {
    return readStore(STORE_KEYS.movements, []);
}

function products(master = getMaster()) {
    return Array.isArray(master?.produk) ? master.produk : [];
}

function suppliers(master = getMaster()) {
    return Array.isArray(master?.supplier) ? master.supplier : [];
}

function canonicalSupplier(value, master = getMaster()) {
    const raw = text(value);
    const label = norm(raw);
    if (!label) return "";

    const match = suppliers(master).find((row) =>
        norm(row?.["Nama Supplier"]) === label ||
        norm(row?.["Nama Perusahaan"]) === label ||
        norm(row?.["Supplier"]) === label ||
        norm(row?.["Kode Supplier"]) === label
    );
    return text(match?.["Supplier"] || match?.["Kode Supplier"]);
}

function supplierName(label, master = getMaster()) {
    const row = suppliers(master).find((item) => norm(item?.["Supplier"]) === norm(label) || norm(item?.["Kode Supplier"]) === norm(label));
    return text(row?.["Nama Supplier"] || row?.["Nama Perusahaan"]) || label;
}
function invoiceIdentity(invoice){return `${norm(invoice?.supplierCode||invoice?.supplierName)}::${norm(invoice?.number)}`;}

function findProductByCode(code, master = getMaster()) {
    const key = norm(code);
    if (!key) return null;
    return products(master).find((product) => norm(product?.["Kode Produk"]) === key) || null;
}

function findPossibleProduct(name, supplier, master = getMaster()) {
    const n = norm(name);
    const s = norm(supplier);
    if (!n || !s) return null;
    const matches = products(master).filter((product) =>
        norm(product?.["Nama Produk"]) === n && norm(product?.["Supplier"]) === s
    );
    return matches.length === 1 ? matches[0] : null;
}

function addNumber(code) {
    const match = text(code).toUpperCase().match(/^ADD(\d+)$/);
    return match ? Number(match[1]) : 0;
}

function nextAddNumber(master = getMaster()) {
    return products(master).reduce((max, product) =>
        Math.max(max, addNumber(product?.["Kode Produk"])), 0) + 1;
}

function nextAddCode(master, usedCodes, cursor) {
    let number = Math.max(1, Number(cursor.value) || 1);
    let code = `ADD${String(number).padStart(4, "0")}`;
    while (usedCodes.has(norm(code))) {
        number += 1;
        code = `ADD${String(number).padStart(4, "0")}`;
    }
    cursor.value = number + 1;
    usedCodes.add(norm(code));
    return code;
}

function invoiceSubtotal(item) {
    return Math.max(0, num(item.qty) * num(item.buyPrice) - num(item.discount));
}

function invoiceTotal(invoice) {
    const itemTotal = (invoice.items || []).reduce((sum, item) => sum + invoiceSubtotal(item), 0);
    const discounted = Math.max(0, itemTotal - num(invoice.discount));
    const perItem = norm(invoice.taxMethod) === "per item" || norm(invoice.taxMethod) === "per_product";
    const tax = perItem
        ? (invoice.items || []).reduce((sum, item) => sum + invoiceSubtotal(item) * num(item.taxPercent) / 100, 0)
        : discounted * num(invoice.taxPercent) / 100;
    return discounted + tax + num(invoice.otherCost);
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
    if (!file) throw new Error("Pilih file Excel faktur terlebih dahulu.");
    if (!window.XLSX?.read) throw new Error("SheetJS/XLSX belum tersedia.");
    const buffer = await file.arrayBuffer();
    return window.XLSX.read(buffer, { type: "array", cellDates: true });
}

function invoiceSupplierFromHeader(row, master) {
    return canonicalSupplier(
        row?.["Supplier"] || row?.["Kode Supplier"] || row?.["Nama Supplier"],
        master
    );
}

function statusCell(invoice) {
    const newCount = (invoice.items || []).filter((item) => item.productStatus === "new").length;
    if (!newCount) return '<span class="ops-badge warn">Draft</span>';
    return `<span class="ops-badge warn">Draft</span> <span class="ops-badge neutral">${newCount} Produk Baru</span>`;
}

async function previewInvoiceImportV2() {
    const master = getMaster();
    const workbook = await readWorkbook($("invoice-import-file")?.files?.[0]);
    const headers = sheetRows(workbook, ["faktur_pembelian", "faktur pembelian"]);
    const details = sheetRows(workbook, ["detail_faktur", "detail faktur"]);
    const map = new Map();
    const issues = [];

    headers.forEach((row, index) => {
        const number = text(row?.["Nomor Faktur"]) || `IMPORT-${index + 1}`;
        const supplier = invoiceSupplierFromHeader(row, master);
        if (!supplier) {
            issues.push(`Faktur ${number}: nama perusahaan supplier tidak ditemukan pada Master Supplier.`);
        }
        map.set(norm(number), {
            id: uid("INV"),
            number,
            date: dateOnly(row?.["Tanggal Faktur"]),
            supplierCode: supplier,
            supplierName: supplierName(supplier, master),
            paymentType: text(row?.["Jenis Pembayaran"]) || "Cash",
            dueDate: dateOnly(row?.["Tanggal Jatuh Tempo"]),
            taxMethod: text(row?.["Metode PPN"]),
            taxPercent: num(row?.["PPN Global (%)"]),
            discount: num(row?.["Diskon Faktur"]),
            otherCost: num(row?.["Biaya Lain"]),
            notes: text(row?.["Catatan"]),
            status: "draft",
            stockApplied: false,
            source: "excel-v2",
            items: [],
            createdAt: nowIso()
        });
    });

    details.forEach((row, index) => {
        const number = text(row?.["Nomor Faktur"]);
        const key = norm(number);
        if (!key || !map.has(key)) {
            issues.push(`Detail baris ${index + 2}: Nomor Faktur tidak memiliki header faktur yang cocok.`);
            return;
        }

        const invoice = map.get(key);
        const rawCode = text(row?.["Kode Produk"]);
        const rawName = text(row?.["Nama Produk"]);
        const existing = findProductByCode(rawCode, master);
        const resolvedName = text(existing?.["Nama Produk"]) || rawName;

        if (!existing && !resolvedName) {
            issues.push(`Faktur ${number}, detail baris ${index + 2}: Produk Baru wajib memiliki Nama Produk.`);
            return;
        }

        invoice.items.push({
            code: existing ? text(existing?.["Kode Produk"]) : rawCode,
            originalCode: rawCode,
            name: resolvedName,
            qty: num(row?.["Qty"]),
            unit: text(row?.["Satuan"]) || text(existing?.["Satuan"]),
            buyPrice: num(row?.["Harga Beli"]),
            discount: num(row?.["Diskon Item"]),
            taxPercent: num(row?.["PPN Item (%)"]),
            sellPrice: num(row?.["Harga Jual"]),
            batch: text(row?.["Batch"]),
            expired: dateOnly(row?.["Tanggal Expired"]),
            notes: text(row?.["Catatan"]),
            supplier: invoice.supplierCode,
            productStatus: existing ? "existing" : "new",
            reviewRequired: !existing
        });
    });

    for (const invoice of map.values()) {
        if (!invoice.items.length) issues.push(`Faktur ${invoice.number}: tidak memiliki item valid.`);
        for (const item of invoice.items) {
            if (num(item.qty) <= 0) issues.push(`Faktur ${invoice.number}: Qty ${item.name || item.code} harus lebih dari 0.`);
        }
    }

    invoicePreviewV2 = [...map.values()];
    invoiceValidationV2 = issues;

    const newProducts = invoicePreviewV2.reduce((sum, invoice) =>
        sum + invoice.items.filter((item) => item.productStatus === "new").length, 0);

    const summary = $("invoice-import-summary");
    if (summary) {
        summary.hidden = false;
        summary.innerHTML = `<strong>${invoicePreviewV2.length} faktur</strong> terbaca dengan ${details.length} baris detail. ` +
            `Semua akan disimpan sebagai Draft. <strong>${newProducts} item Produk Baru</strong> memerlukan review Admin sebelum Barang Masuk.` +
            (issues.length ? `<br><strong>${issues.length} masalah validasi harus diperbaiki sebelum import.</strong>` : "");
    }

    const body = $("invoice-import-preview-body");
    if (body) {
        body.innerHTML = invoicePreviewV2.length
            ? invoicePreviewV2.map((invoice) => `<tr>
                <td>${escapeHtml(invoice.number || "—")}</td>
                <td>${escapeHtml(invoice.date || "—")}</td>
                <td>${escapeHtml(invoice.supplierName || invoice.supplierCode || "—")}</td>
                <td>${invoice.items.length}</td>
                <td>${rupiah(invoiceTotal(invoice))}</td>
                <td>${statusCell(invoice)}</td>
            </tr>`).join("")
            : '<tr><td colspan="6" class="empty-table-state">Tidak ada faktur valid.</td></tr>';
    }

    if ($("invoice-import-preview")) $("invoice-import-preview").hidden = false;
    if ($("invoice-import-actions")) $("invoice-import-actions").hidden = false;

    if (issues.length) {
        alert(`Preview Faktur menemukan masalah:\n\n${issues.slice(0, 10).join("\n")}${issues.length > 10 ? `\nDan ${issues.length - 10} masalah lainnya.` : ""}`);
    }
}

async function applyInvoiceImportV2() {
    if (!invoicePreviewV2.length) return alert("Baca dan preview file faktur terlebih dahulu.");
    if (invoiceValidationV2.length) return alert("Import belum dapat dilakukan karena masih ada masalah validasi pada preview.");

    const list = getInvoices();
    let added = 0;
    let updated = 0;
    let skipped = 0;

    for (const incoming of invoicePreviewV2) {
        const index = list.findIndex((invoice) => invoiceIdentity(invoice) === invoiceIdentity(incoming));
        if (index < 0) {
            list.unshift(incoming);
            added += 1;
            continue;
        }

        const old = list[index];
        if (old.stockApplied || norm(old.status) === "confirmed") {
            await window.KasirProDialog?.error?.("Faktur Sudah Dikonfirmasi", `Faktur ${incoming.number} dari ${incoming.supplierName || incoming.supplierCode} tidak dapat diimpor ulang karena stok sudah pernah diterapkan.`);
            skipped += 1;
            continue;
        }
        const replace = await window.KasirProDialog?.confirm?.("Ganti Draft Lama?", `Seluruh isi Draft ${incoming.number} akan diganti oleh file terbaru. Item tidak akan digabungkan.`, {confirmText:"Ganti Draft"});
        if (!replace) { skipped += 1; continue; }
        incoming.id = old.id;
        incoming.createdAt = old.createdAt;
        incoming.updatedAt = nowIso();
        list[index] = incoming;
        updated += 1;
    }

    await writeStore(STORE_KEYS.invoices, list);
    invoicePreviewV2 = [];
    invoiceValidationV2 = [];
    alert(`Import Faktur selesai sebagai Draft.\nBaru: ${added}\nDiperbarui: ${updated}\nDilewati: ${skipped}\n\nStok belum berubah.`);
    document.querySelector('[data-view="purchase-invoices"]')?.click();
}

function currentStockByCode(code) {
    return readCurrentStock(code);
}

async function reviewInvoiceProducts(invoice, master) {
    const productList = products(master);
    const usedCodes = new Set(productList.map((product) => norm(product?.["Kode Produk"])).filter(Boolean));
    const cursor = { value: nextAddNumber(master) };
    const created = [];

    for (const item of invoice.items || []) {
        const exact = findProductByCode(item.code, master);
        if (exact) {
            item.code = text(exact["Kode Produk"]);
            item.name = text(exact["Nama Produk"]);
            item.productStatus = "existing";
            item.reviewRequired = false;
            continue;
        }

        const possible = findPossibleProduct(item.name, invoice.supplierCode, master);
        if (possible) {
            const useExisting = confirm(
                `Review Produk Faktur ${invoice.number}\n\n` +
                `Produk: ${item.name}\nSupplier: ${invoice.supplierCode}\n\n` +
                `Ditemukan produk existing ${possible["Kode Produk"]} — ${possible["Nama Produk"]}.\n` +
                `OK = gunakan produk existing.\nBatal = perlakukan sebagai Produk Baru.`
            );
            if (useExisting) {
                item.code = text(possible["Kode Produk"]);
                item.name = text(possible["Nama Produk"]);
                item.productStatus = "existing";
                item.reviewRequired = false;
                continue;
            }
        }

        const manualCode = text(item.originalCode || item.code);
        let proposedCode = manualCode;
        if (proposedCode && usedCodes.has(norm(proposedCode))) proposedCode = "";
        if (!proposedCode) proposedCode = nextAddCode(master, usedCodes, cursor);
        else usedCodes.add(norm(proposedCode));

        const approved = confirm(
            `Produk Baru pada Faktur ${invoice.number}\n\n` +
            `Nama: ${item.name}\nSupplier: ${invoice.supplierCode}\nKode: ${proposedCode}\n\n` +
            `Tambahkan produk ini ke Master?\nStok awal produk tetap 0; stok baru bertambah setelah Barang Masuk dikonfirmasi.`
        );
        if (!approved) return { approved: false, created: [] };

        const product = {
            "Kode Produk": proposedCode,
            "Nama Produk": text(item.name),
            "Supplier": invoice.supplierCode,
            "Kategori": "",
            "Satuan": text(item.unit),
            "Harga Beli": num(item.buyPrice) || "",
            "Harga Jual": num(item.sellPrice) || "",
            "Stok Minimum": "",
            "Lokasi Rak": "",
            "Batch": text(item.batch),
            "Tanggal Expired": text(item.expired),
            "Status Produk": "Aktif",
            "Catatan": `Dibuat dari review Faktur ${invoice.number}`
        };
        productList.push(product);
        created.push(product);
        item.code = proposedCode;
        item.productStatus = "existing";
        item.reviewRequired = false;
    }

    master.produk = productList;
    return { approved: true, created };
}

async function confirmInvoiceGoodsInV2(invoiceId) {
    const invoices = getInvoices();
    const invoice = invoices.find((item) => String(item.id) === String(invoiceId));
    if (!invoice) return alert("Faktur tidak ditemukan.");
    if (invoice.stockApplied) return alert("Barang Masuk faktur ini sudah pernah diterapkan.");

    const supplier = canonicalSupplier(invoice.supplierCode || invoice.supplierName);
    if (!supplier) return alert("Nama perusahaan supplier pada faktur tidak ditemukan di Master Supplier.");
    invoice.supplierCode = supplier;
    invoice.supplierName = supplierName(supplier);

    const master = getMaster();
    const review = await reviewInvoiceProducts(invoice, master);
    if (!review.approved) return alert("Barang Masuk dibatalkan. Tidak ada stok yang berubah.");

    if (review.created.length) {
        try {
            await writeStore(STORE_KEYS.master, master);
        } catch (error) {
            console.error("Produk Baru gagal disimpan:", error);
            return alert("Produk Baru belum dapat disimpan ke Master. Barang Masuk dibatalkan.");
        }
    }

    const unresolved = (invoice.items || []).find((item) => !findProductByCode(item.code, master));
    if (unresolved) {
        return alert(`Produk ${unresolved.name || unresolved.code || "faktur"} belum memiliki Kode Produk valid. Barang Masuk dibatalkan.`);
    }

    if (!confirm(`Konfirmasi Barang Masuk untuk faktur ${invoice.number}?\n\n${invoice.items.length} item akan menambah stok operasional.`)) return;

    const movements = getMovements();
    const stockWorking = new Map();
    const at = nowIso();

    for (const item of invoice.items || []) {
        const code = text(item.code);
        const product = findProductByCode(code, master);
        if (!product) return alert(`Produk ${code} tidak ditemukan. Barang Masuk dibatalkan.`);
        const delta = num(item.qty);
        if (delta <= 0) return alert(`Qty ${product["Nama Produk"]} harus lebih dari 0.`);

        const key = norm(code);
        const before = stockWorking.has(key)
            ? stockWorking.get(key)
            : currentStockByCode(code);
        const after = before + delta;
        stockWorking.set(key, after);
        movements.unshift({
            id: uid("MOV"),
            at,
            type: "purchase",
            reference: invoice.number,
            productCode: code,
            productName: text(product["Nama Produk"]),
            delta,
            stockBefore: before,
            stockAfter: after,
            supplier: invoice.supplierCode,
            note: `Barang masuk faktur ${invoice.number}`
        });
    }

    invoice.status = "confirmed";
    invoice.stockApplied = true;
    invoice.confirmedAt = at;
    invoice.updatedAt = at;

    try {
        await writeStockTransaction([
            { key: STORE_KEYS.movements, records: movements.slice(0, (invoice.items || []).length) },
            { key: STORE_KEYS.invoices, records: [invoice] }
        ]);
    } catch (error) {
        console.error("Barang Masuk V2 gagal disimpan:", error);
        invoice.status = "draft";
        invoice.stockApplied = false;
        delete invoice.confirmedAt;
        return alert("Barang Masuk belum dapat disimpan. Faktur tetap Draft; periksa koneksi/database lalu coba kembali.");
    }

    alert(`Barang Masuk faktur ${invoice.number} berhasil.\nStok operasional diperbarui melalui MutasiStok.${review.created.length ? `\nProduk Baru ditambahkan ke Master: ${review.created.length}.` : ""}`);
    window.KasirProInvoiceFinalAT08?.render?.();
    document.querySelector('[data-view="goods-in"]')?.click();
}

async function captureInvoiceActions(event) {
    const readButton = event.target.closest("#read-invoice-import");
    const applyButton = event.target.closest("#apply-invoice-import");
    const confirmButton = event.target.closest("[data-invoice-confirm]");
    const editorConfirm = event.target.closest("#confirm-invoice");

    if (readButton) {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        readButton.disabled = true;
        try { await previewInvoiceImportV2(); }
        catch (error) { console.error(error); alert(error?.message || "Gagal membaca faktur."); }
        finally { readButton.disabled = false; }
        return;
    }

    if (applyButton) {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        applyButton.disabled = true;
        try { await applyInvoiceImportV2(); }
        catch (error) { console.error(error); alert(error?.message || "Import Faktur gagal."); }
        finally { applyButton.disabled = false; }
        return;
    }

    if (confirmButton) {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        await confirmInvoiceGoodsInV2(confirmButton.dataset.invoiceConfirm);
        return;
    }

    if (editorConfirm) {
        const invoiceNumber = text($("invoice-number")?.value);
        const invoice = getInvoices().find((item) => norm(item.number) === norm(invoiceNumber));
        if (!invoice) return;
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        await confirmInvoiceGoodsInV2(invoice.id);
    }
}

document.addEventListener("click", captureInvoiceActions, true);

window.KasirProInvoiceV2 = Object.freeze({
    previewInvoiceImportV2,
    applyInvoiceImportV2,
    confirmInvoiceGoodsInV2,
    currentStockByCode
});
