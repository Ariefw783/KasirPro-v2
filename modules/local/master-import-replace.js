/* KasirPro — Full Replace Master Import
 * Master V2 hanya memuat produk, kategori, dan supplier. Stok produk lama tetap
 * operasional; stok awal produk baru dicatat sebagai mutasi saldo awal.
 */
import {
  masterReady,
  readMasterSnapshot,
  readMasterVersion,
  installMasterSnapshot
} from "./master-repository.js";
import {
  STORE_KEYS,
  isOperationalRemoteReady,
  reloadMasterCache,
  writeStockTransaction
} from "../database/database-store.js";

const REQUIRED_SHEETS = ["SUPPLIER", "KATEGORI", "PRODUK"];
const MAX_OPENING_PRODUCTS = 200;

const text = v => String(v ?? "").trim();
const norm = v => text(v).toLowerCase();
const headerKey = v => norm(v).replace(/[\s_\-]+/g, "");
const clone = v => typeof structuredClone === "function" ? structuredClone(v) : JSON.parse(JSON.stringify(v));

function normalizeMaster(master) {
  const src = master || {};
  return {
    produk: Array.isArray(src.produk) ? src.produk : [],
    supplier: Array.isArray(src.supplier) ? src.supplier : [],
    kategori: Array.isArray(src.kategori) ? src.kategori : [],
    pengguna: [],
    pengaturan_toko: Array.isArray(src.pengaturan_toko) ? src.pengaturan_toko : []
  };
}

function sheetByName(workbook, wanted) {
  const actual = (workbook?.sheetNames || []).find(name => norm(name) === norm(wanted));
  return actual ? workbook.sheets?.[actual] : null;
}

function rowsAsObjects(sheet) {
  if (!sheet) return [];
  const headers = Array.isArray(sheet.headers) ? sheet.headers : [];
  return (sheet.rows || []).map(row => {
    const record = {};
    headers.forEach((header, i) => { record[header] = row.values?.[i] ?? ""; });
    return { sourceRow: row.sourceRow, record };
  }).filter(({record}) => Object.values(record).some(v => text(v) !== ""));
}

function getField(record, wanted) {
  const key = Object.keys(record || {}).find(k => headerKey(k) === headerKey(wanted));
  return key ? record[key] : "";
}

function sanitizeRecord(record) {
  const out = { ...(record || {}) };
  delete out._firestoreDocumentId;
  delete out._localKey;
  return out;
}

function sanitizeProduct(record) {
  const out = sanitizeRecord(record);
  out["Kode Produk"] = text(getField(record, "Kode Produk"));
  out["Nama Produk"] = text(getField(record, "Nama Produk"));
  out["Supplier"] = text(getField(record, "Supplier"));
  delete out["Kode Supplier"];
  delete out["Nama Supplier"];
  delete out["Stok Awal"];
  return out;
}

function sanitizeSupplier(record) {
  const out = sanitizeRecord(record);
  out["Supplier"] = text(getField(record, "Supplier"));
  out["Nama Supplier"] = text(getField(record, "Nama Supplier"));
  delete out["Kode Supplier"];
  return out;
}

function addNumber(code) {
  const m = text(code).toUpperCase().match(/^ADD(\d+)$/);
  return m ? Number(m[1]) : 0;
}

function findUniqueCurrentByNameSupplier(current, name, supplier) {
  const n = norm(name), s = norm(supplier);
  const matches = (current?.produk || []).filter(p => norm(p?.["Nama Produk"]) === n && norm(p?.["Supplier"]) === s);
  return matches.length === 1 ? matches[0] : null;
}

function validateStructure(workbook) {
  const issues = [];
  const names = (workbook?.sheetNames || []).map(norm);
  for (const required of REQUIRED_SHEETS) {
    if (!names.includes(norm(required))) issues.push(`Sheet ${required} wajib tersedia.`);
  }
  const productSheet = sheetByName(workbook, "PRODUK");
  if (productSheet) {
    const headers = new Set((productSheet.headers || []).map(headerKey));
    for (const required of ["Nama Produk", "Supplier"]) {
      if (!headers.has(headerKey(required))) issues.push(`Sheet PRODUK wajib memiliki kolom ${required}.`);
    }
  }
  return issues;
}

function buildReplacementMaster(workbook, current) {
  const issues = validateStructure(workbook);
  const supplierRows = rowsAsObjects(sheetByName(workbook, "SUPPLIER"));
  const categoryRows = rowsAsObjects(sheetByName(workbook, "KATEGORI"));
  const productRows = rowsAsObjects(sheetByName(workbook, "PRODUK"));
  const suppliers = supplierRows.map(({record}) => sanitizeSupplier(record));
  const supplierLabels = new Set();
  suppliers.forEach((row, i) => {
    const label = text(row["Supplier"]);
    if (!label) issues.push(`SUPPLIER baris ${supplierRows[i].sourceRow || i + 2}: isi nama perusahaan supplier.`);
    if (supplierLabels.has(norm(label))) issues.push(`SUPPLIER: ${label || "label kosong"} tercatat lebih dari sekali.`);
    if (label) supplierLabels.add(norm(label));
  });

  const products = productRows.map(({record}) => sanitizeProduct(record));
  const incomingCodes = new Set();
  const nameSupplierPairs = new Set();
  let maxAdd = 0;
  (current?.produk || []).forEach(p => { maxAdd = Math.max(maxAdd, addNumber(p?.["Kode Produk"])); });
  products.forEach(p => { maxAdd = Math.max(maxAdd, addNumber(p["Kode Produk"])); });
  let nextAdd = maxAdd + 1;

  products.forEach((product, i) => {
    const rowNo = productRows[i].sourceRow || i + 2;
    const name = text(product["Nama Produk"]);
    const supplier = text(product["Supplier"]);
    if (!name) issues.push(`PRODUK baris ${rowNo}: Nama Produk wajib diisi.`);
    if (!supplier) issues.push(`PRODUK baris ${rowNo}: Supplier wajib diisi.`);
    else if (!supplierLabels.has(norm(supplier))) issues.push(`PRODUK baris ${rowNo}: nama perusahaan ${supplier} tidak tersedia pada sheet SUPPLIER.`);

    const pair = `${norm(name)}|${norm(supplier)}`;
    if (name && supplier) {
      if (nameSupplierPairs.has(pair)) issues.push(`PRODUK baris ${rowNo}: kombinasi Nama Produk + Supplier tercatat lebih dari sekali.`);
      nameSupplierPairs.add(pair);
    }

    let code = text(product["Kode Produk"]);
    if (!code && name && supplier) {
      const existing = findUniqueCurrentByNameSupplier(current, name, supplier);
      const existingCode = text(existing?.["Kode Produk"]);
      if (existingCode && !incomingCodes.has(norm(existingCode))) code = existingCode;
      else {
        do { code = `ADD${String(nextAdd++).padStart(4, "0")}`; }
        while (incomingCodes.has(norm(code)));
      }
      product["Kode Produk"] = code;
    }
    if (!code) issues.push(`PRODUK baris ${rowNo}: Kode Produk tidak dapat dibuat karena Nama Produk/Supplier belum lengkap.`);
    else if (incomingCodes.has(norm(code))) issues.push(`PRODUK baris ${rowNo}: Kode Produk ${code} digunakan lebih dari sekali.`);
    else incomingCodes.add(norm(code));
  });

  const kategori = categoryRows.map(({record}) => sanitizeRecord(record));
  const currentCodes = new Set((current?.produk || []).map(p => norm(p?.["Kode Produk"])).filter(Boolean));
  const openingStock = [];
  products.forEach((product, i) => {
    if (currentCodes.has(norm(product["Kode Produk"]))) return;
    const raw = getField(productRows[i].record, "Stok Awal");
    if (text(raw) === "") return;
    const quantity = Number(raw);
    if (!Number.isFinite(quantity) || quantity < 0 || !Number.isInteger(quantity)) {
      issues.push(`PRODUK baris ${productRows[i].sourceRow || i + 2}: Stok Awal harus bilangan bulat tidak negatif.`);
    } else if (quantity > 0) {
      openingStock.push({ code: product["Kode Produk"], name: product["Nama Produk"], quantity });
    }
  });
  if (openingStock.length > MAX_OPENING_PRODUCTS) {
    issues.push(`Stok awal hanya dapat diinisialisasi untuk maksimal ${MAX_OPENING_PRODUCTS} produk baru per impor. Bagi file menjadi beberapa tahap.`);
  }

  const master = { produk: products, supplier: suppliers, kategori, pengguna: [], pengaturan_toko: current.pengaturan_toko };
  return { master, openingStock, issues };
}

export async function prepareMasterReplacement(workbookData) {
  const current = (await masterReady()) ? normalizeMaster(await readMasterSnapshot()) : normalizeMaster({});
  const workbook = clone(workbookData || {sheetNames:[], sheets:{}});
  const built = buildReplacementMaster(workbook, current);
  const currentCodes = new Set((current.produk || []).map(p => norm(p?.["Kode Produk"])).filter(Boolean));
  const nextCodes = new Set((built.master.produk || []).map(p => norm(p?.["Kode Produk"])).filter(Boolean));
  return {
    current,
    workbook,
    master: built.master,
    openingStock: built.openingStock,
    validation: { issues: built.issues, canContinue: built.issues.length === 0 },
    summary: {
      products: built.master.produk.length,
      suppliers: built.master.supplier.length,
      categories: built.master.kategori.length,
      initializedStock: built.openingStock.length,
      removedProducts: [...currentCodes].filter(code => !nextCodes.has(code)).length
    }
  };
}

export async function commitMasterReplacement(workbookData) {
  const prepared = await prepareMasterReplacement(workbookData);
  if (!prepared.validation.canContinue) {
    const error = new Error(prepared.validation.issues.join("\n"));
    error.code = "MASTER_REPLACE_VALIDATION_FAILED";
    error.validation = prepared.validation;
    throw error;
  }
  const currentVersion = Number(await readMasterVersion()) || 0;
  const version = currentVersion + 1 || 1;
  if (prepared.openingStock.length && !isOperationalRemoteReady()) {
    throw new Error("Stok awal produk baru memerlukan koneksi data operasional pusat. Import dibatalkan; coba lagi saat koneksi siap.");
  }

  if (prepared.openingStock.length) {
    // Master lokal diperlukan untuk indeks produk; jangan publikasikan sebelum stok berhasil.
    await installMasterSnapshot(prepared.master, { version, source: "master-full-replace", queuePublish: false });
    const reference = `AWAL-MASTER-${version}-${Date.now()}`;
    const at = new Date().toISOString();
    const movements = prepared.openingStock.map((item, index) => ({
      id: `${reference}-${index + 1}`,
      at,
      type: "opening",
      reference,
      productCode: item.code,
      productName: item.name,
      delta: item.quantity,
      stockBefore: 0,
      stockAfter: item.quantity,
      note: "Stok awal produk baru dari Import Master"
    }));
    try {
      await writeStockTransaction([{ key: STORE_KEYS.movements, records: movements }]);
    } catch (error) {
      if (error?.remoteCommitted) {
        // Firestore sudah menyimpan stok; kegagalan hanya terjadi saat cache lokal diperbarui.
        // Jangan rollback Master, agar stok pusat tidak menjadi yatim.
      } else {
        try {
          await installMasterSnapshot(prepared.current, { version: currentVersion, source: "master-import-rollback", queuePublish: false });
        } catch (rollbackError) {
          throw new Error(`Stok awal gagal dan Master lokal tidak dapat dipulihkan: ${rollbackError.message}. Hentikan impor dan periksa data sebelum mencoba lagi.`, { cause: error });
        }
        throw new Error(`Import dibatalkan karena stok awal gagal disimpan: ${error.message}`);
      }
    }
  }

  // Stok operasional sudah beres; kini antrekan snapshot Master dapat dipublikasikan.
  let installed;
  try {
    installed = await installMasterSnapshot(prepared.master, { version, source: "master-full-replace" });
  } catch (error) {
    installed = {
      version,
      master: await readMasterSnapshot(),
      publish: { published: false, queued: false, reason: "publish-queue-error", error: error?.message || String(error) }
    };
  }
  await reloadMasterCache();
  return {
    version: installed.version,
    master: installed.master,
    publish: installed.publish,
    summary: prepared.summary
  };
}
