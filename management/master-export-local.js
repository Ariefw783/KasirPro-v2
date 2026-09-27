/* KasirPro Management - Export Master Aktif dari IndexedDB ke Excel
 * OFFLINE-SAFE: tidak mengakses Firebase/Firestore.
 */

import { readMasterSnapshot, readMasterVersion } from "../modules/local/master-repository.js";

const STANDARD_HEADERS = Object.freeze({
    SUPPLIER: ["Supplier","Alamat","Telepon","Email","Kontak Person","NPWP","Termin Default","Status","Catatan"],
    KATEGORI: ["Kode Kategori","Nama Kategori","Deskripsi","Status"],
    PRODUK: ["Kode Produk","Nama Produk","Kategori","Satuan","Supplier","Harga Beli","Harga Jual","Stok Awal","Stok Minimum","Lokasi Rak","Batch","Tanggal Expired","Status Produk","Catatan"]
});

const PRIVATE_FIELDS = new Set(["_id","_localKey","_firestoreDocumentId"]);

function cleanRecords(records) {
    return (Array.isArray(records) ? records : []).map(record => {
        const out = {};
        Object.entries(record || {}).forEach(([key, value]) => {
            if (!PRIVATE_FIELDS.has(key)) out[key] = value ?? "";
        });
        return out;
    });
}

function sheetFromRecords(records, standardHeaders, kind = "") {
    const rows = cleanRecords(records);
    const headers = standardHeaders;
    const matrix = [headers, ...rows.map(row => headers.map(header => {
        if (kind === "PRODUK" && header === "Stok Awal") return "";
        if (kind === "SUPPLIER" && header === "Supplier") return row.Supplier || row["Nama Supplier"] || "";
        return row?.[header] ?? "";
    }))];
    const sheet = window.XLSX.utils.aoa_to_sheet(matrix);
    sheet["!cols"] = headers.map(header => ({ wch: Math.min(42, Math.max(12, String(header).length + 2)) }));
    if (matrix.length > 1) sheet["!autofilter"] = { ref: `A1:${window.XLSX.utils.encode_col(headers.length - 1)}${matrix.length}` };
    return sheet;
}

function instructionSheet(version, snapshot) {
    const rows = [
        ["KASIRPRO - MASTER AKTIF"],
        ["Versi Master", version],
        ["Tanggal Export", new Date().toLocaleString("id-ID")],
        ["Jumlah Produk", snapshot.produk?.length || 0],
        ["Jumlah Supplier", snapshot.supplier?.length || 0],
        ["Jumlah Kategori", snapshot.kategori?.length || 0],
        [],
        ["ATURAN EDIT & IMPORT"],
        ["1", "Produk baru wajib hanya Nama Produk dan Supplier."],
        ["2", "Kode Produk boleh dikosongkan. KasirPro akan membuat ADD0001, ADD0002, dan seterusnya."],
        ["3", "Import Master mengganti daftar Produk, Kategori, dan Supplier. Pertahankan semua baris yang masih digunakan."],
        ["4", "Kolom Supplier berisi nama perusahaan yang sama pada sheet SUPPLIER dan PRODUK."],
        ["5", "Pengguna dan Pengaturan Toko dikelola melalui aplikasi, bukan file Master."],
        ["6", "Stok Awal produk lama kosong karena stok aktual tetap tersimpan. Isi hanya untuk produk baru jika diperlukan."],
        [],
        ["WORKFLOW"],
        ["Export Master Aktif → Edit Excel → Import Master → Preview/Validasi → Konfirmasi"]
    ];
    const sheet = window.XLSX.utils.aoa_to_sheet(rows);
    sheet["!cols"] = [{ wch: 18 }, { wch: 90 }];
    return sheet;
}

function fileName(version) {
    const date = new Date();
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const dd = String(date.getDate()).padStart(2, "0");
    return `Master_Aktif_KasirPro_v${Number(version) || 0}_${yyyy}-${mm}-${dd}.xlsx`;
}

export async function exportActiveMasterToExcel() {
    if (!window.XLSX?.utils || typeof window.XLSX.writeFile !== "function") {
        throw new Error("SheetJS/XLSX belum tersedia pada halaman Management.");
    }

    const snapshot = await readMasterSnapshot();
    const version = await readMasterVersion();
    if (!snapshot) throw new Error("Master Aktif belum tersedia di IndexedDB.");

    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, instructionSheet(version, snapshot), "PETUNJUK_IMPORT");
    window.XLSX.utils.book_append_sheet(workbook, sheetFromRecords(snapshot.supplier, STANDARD_HEADERS.SUPPLIER, "SUPPLIER"), "SUPPLIER");
    window.XLSX.utils.book_append_sheet(workbook, sheetFromRecords(snapshot.kategori, STANDARD_HEADERS.KATEGORI), "KATEGORI");
    window.XLSX.utils.book_append_sheet(workbook, sheetFromRecords(snapshot.produk, STANDARD_HEADERS.PRODUK, "PRODUK"), "PRODUK");

    const name = fileName(version);
    window.XLSX.writeFile(workbook, name, { compression: true });
    return {
        fileName: name,
        version,
        counts: {
            products: snapshot.produk?.length || 0,
            suppliers: snapshot.supplier?.length || 0,
            categories: snapshot.kategori?.length || 0
        }
    };
}

function findMasterTemplateAction() {
    return [...document.querySelectorAll("a,button")].find(element =>
        element.textContent?.trim().toLowerCase().includes("download template")
    ) || null;
}

function installExportButton() {
    if (document.getElementById("export-active-master")) return;
    const templateAction = findMasterTemplateAction();
    if (!templateAction?.parentElement) return;

    const button = document.createElement("button");
    button.type = "button";
    button.id = "export-active-master";
    button.className = templateAction.className;
    button.innerHTML = '<i class="fa-solid fa-file-export" aria-hidden="true"></i> Export Master Aktif';
    button.style.marginLeft = "8px";
    templateAction.insertAdjacentElement("afterend", button);
}

async function handleExportClick(event) {
    const button = event.target.closest("#export-active-master");
    if (!button) return;

    const previous = button.innerHTML;
    button.disabled = true;
    button.textContent = "Menyiapkan Excel...";

    try {
        const result = await exportActiveMasterToExcel();
        alert(`Export Master Aktif selesai.\n\nFile: ${result.fileName}\nMaster v${result.version}\nProduk: ${result.counts.products}\nSupplier: ${result.counts.suppliers}\nKategori: ${result.counts.categories}\n\nFile siap diedit lalu diimport kembali.`);
    } catch (error) {
        console.error("Export Master Aktif gagal:", error);
        alert(`Export Master Aktif gagal: ${error?.message || error}`);
    } finally {
        button.disabled = false;
        button.innerHTML = previous;
    }
}

document.addEventListener("click", handleExportClick);
if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", installExportButton, { once: true });
} else {
    installExportButton();
}

window.KasirProMasterExport = Object.freeze({ exportActiveMasterToExcel });
