import {
    initializeDatabase,
    readStore,
    writeStore,
    writeStoreBundle,
    writeOperationalDelta
} from "../modules/database/database-store.js";

const databaseInitialization = initializeDatabase();

const MASTER_KEY = "kasirpro_master_store_v1";
const INVOICE_KEY = "kasirpro_purchase_invoices_v1";
const MOVEMENT_KEY = "kasirpro_stock_movements_v1";
const OPNAME_KEY = "kasirpro_stock_opname_v1";
const SALES_KEY = "kasirpro_sales_v1";

let importedInvoicesPreview = [];
let importedOpnamePreview = [];
let editingInvoiceId = null;
let editorItems = [];
const STOCK_PAGE_SIZE = 50;
let stockCurrentPage = 1;
let stockRenderToken = 0;

const $ = (id) => document.getElementById(id);
const num = (v) => Number(String(v ?? 0).replace(/[^0-9.-]/g, "")) || 0;
const text = (v) => String(v ?? "").trim();
const norm = (v) => text(v).toLowerCase();
const rupiah = (v) => new Intl.NumberFormat("id-ID", { style: "currency", currency: "IDR", maximumFractionDigits: 0 }).format(num(v));
const nowIso = () => new Date().toISOString();
const uid = (prefix) => `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
const dateOnly = (v) => {
    if (!v) return "";
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    const d = new Date(v);
    return Number.isNaN(d.getTime()) ? text(v) : d.toISOString().slice(0, 10);
};
const fmtDateTime = (v) => {
    const d = new Date(v);
    return Number.isNaN(d.getTime()) ? text(v) || "—" : d.toLocaleString("id-ID");
};

function getMaster() {
    return readStore(MASTER_KEY, {});
}
function saveMaster(store) { return writeStore(MASTER_KEY, store); }
function getInvoices() { return readStore(INVOICE_KEY, []); }
function saveInvoices(v) { return writeStore(INVOICE_KEY, v); }
function getMovements() { return readStore(MOVEMENT_KEY, []); }
function saveMovements(v) { return writeStore(MOVEMENT_KEY, v); }
function getOpnameHistory() { return readStore(OPNAME_KEY, []); }
function saveOpnameHistory(v) { return writeStore(OPNAME_KEY, v); }
function getSales() { return readStore(SALES_KEY, []); }
function saveSales(v) { return writeStore(SALES_KEY, v); }

function products() { const s = getMaster(); return Array.isArray(s.produk) ? s.produk : []; }
function suppliers() { const s = getMaster(); return Array.isArray(s.supplier) ? s.supplier : []; }

function findProduct(ref = {}) {
    const list = products();
    const code = norm(ref["Kode Produk"] ?? ref.code);
    const name = norm(ref["Nama Produk"] ?? ref.name);
    return list.find((p) => code && norm(p["Kode Produk"]) === code)
        || list.find((p) => name && norm(p["Nama Produk"]) === name)
        || null;
}

function discountAmount(method, value, base) {
    const gross = Math.max(0, num(base));
    const entered = Math.max(0, num(value));
    return Math.min(gross, method === "percentage" ? gross * entered / 100 : entered);
}
function invoiceSubtotal(item) {
    const gross = num(item.qty) * num(item.buyPrice);
    return Math.max(0, gross - num(item.discount));
}
function invoiceTotal(inv) {
    const items = (inv.items || []).reduce((a, i) => a + invoiceSubtotal(i), 0);
    const afterDiscount = Math.max(0, items - num(inv.discount));
    const perItem = norm(inv.taxMethod) === "per item" || norm(inv.taxMethod) === "per_product";
    const tax = perItem
        ? (inv.items || []).reduce((a, i) => a + invoiceSubtotal(i) * num(i.taxPercent) / 100, 0)
        : afterDiscount * num(inv.taxPercent) / 100;
    return afterDiscount + tax + num(inv.otherCost);
}

function navigate(view) {
    const btn = document.querySelector(`[data-view="${view}"]`);
    btn?.click();
}

function statusBadge(status) {
    const s = norm(status);
    if (s === "confirmed" || s === "terkonfirmasi") return '<span class="ops-badge ok">Terkonfirmasi</span>';
    if (s === "draft") return '<span class="ops-badge warn">Draft</span>';
    return `<span class="ops-badge neutral">${text(status) || "Belum Diatur"}</span>`;
}
function movementBadge(type) {
    const map = {
        purchase: ["Barang Masuk", "ok"],
        opening: ["Stok Awal", "neutral"],
        opname: ["Stock Opname", "warn"],
        sale: ["Penjualan", "danger"],
        sale_void: ["VOID Penjualan", "neutral"],
        reversal: ["Pembalikan", "neutral"]
    };
    const x = map[type] || [type || "Mutasi", "neutral"];
    return `<span class="ops-badge ${x[1]}">${x[0]}</span>`;
}

function renderInvoices() {
    const all = getInvoices();
    const q = norm($("invoice-search")?.value);
    const sf = norm($("invoice-status-filter")?.value);
    const filtered = all.filter((i) => (!sf || norm(i.status) === sf) && (!q || [i.number, i.supplierCode, i.supplierName].some((x) => norm(x).includes(q))));
    $("invoice-count") && ($("invoice-count").textContent = all.length);
    $("invoice-draft-count") && ($("invoice-draft-count").textContent = all.filter((i) => norm(i.status) === "draft").length);
    $("invoice-confirmed-count") && ($("invoice-confirmed-count").textContent = all.filter((i) => norm(i.status) === "confirmed").length);
    $("invoice-value-total") && ($("invoice-value-total").textContent = rupiah(all.reduce((a, i) => a + invoiceTotal(i), 0)));
    const body = $("invoice-table-body");
    if (!body) return;
    body.innerHTML = filtered.length ? filtered.map((i) => `<tr>
        <td>${text(i.number) || "—"}</td><td>${dateOnly(i.date) || "—"}</td><td>${text(i.supplierName || i.supplierCode) || "—"}</td><td>${(i.items || []).length}</td><td>${rupiah(invoiceTotal(i))}</td><td>${statusBadge(i.status)}</td>
        <td class="ops-actions"><button class="icon-button" data-invoice-edit="${i.id}" title="Lihat/Edit"><i class="fa-solid fa-pen"></i></button>${norm(i.status) === "draft" ? `<button class="icon-button" data-invoice-confirm="${i.id}" title="Konfirmasi Barang Masuk"><i class="fa-solid fa-box-open"></i></button>` : ""}</td>
    </tr>`).join("") : '<tr><td colspan="7" class="empty-table-state">Belum ada faktur.</td></tr>';
    body.querySelectorAll("[data-invoice-edit]").forEach((b) => b.addEventListener("click", () => openInvoiceEditor(b.dataset.invoiceEdit)));
    body.querySelectorAll("[data-invoice-confirm]").forEach((b) => b.addEventListener("click", () => confirmExistingInvoice(b.dataset.invoiceConfirm)));
}

function fillInvoiceSuppliers() {
    const sel = $("invoice-supplier");
    if (!sel) return;
    sel.innerHTML = '<option value="">Pilih supplier</option>' + suppliers().map((s) => `<option value="${text(s["Kode Supplier"])}">${text(s["Kode Supplier"])} — ${text(s["Nama Supplier"])}</option>`).join("");
}
function fillInvoiceProducts() {
    const sel = $("invoice-item-product");
    if (!sel) return;
    sel.innerHTML = '<option value="">Pilih produk</option>' + products().map((p, idx) => `<option value="${idx}">${text(p["Kode Produk"])} — ${text(p["Nama Produk"])}</option>`).join("");
}
function openInvoiceEditor(id = null) {
    editingInvoiceId = id;
    editorItems = [];
    fillInvoiceSuppliers();
    fillInvoiceProducts();
    const inv = id ? getInvoices().find((x) => x.id === id) : null;
    $("invoice-editor-title").textContent = inv ? `Faktur ${inv.number || ""}` : "Faktur Baru";
    $("invoice-number").value = inv?.number || "";
    $("invoice-date").value = dateOnly(inv?.date) || new Date().toISOString().slice(0, 10);
    $("invoice-supplier").value = inv?.supplierCode || "";
    $("invoice-payment-type").value = inv?.paymentType || "Cash";
    $("invoice-due-date").value = dateOnly(inv?.dueDate);
    $("invoice-tax").value = num(inv?.taxPercent);
    $("invoice-discount-method") && ($("invoice-discount-method").value = inv?.discountMethod || "nominal");
    $("invoice-discount").value = num(inv?.discountValue ?? inv?.discount);
    $("invoice-other-cost").value = num(inv?.otherCost);
    $("invoice-notes").value = inv?.notes || "";
    editorItems = (inv?.items || []).map((i) => ({ ...i }));
    $("invoice-item-discount-method") && ($("invoice-item-discount-method").value = "nominal");
    const locked = inv && norm(inv.status) === "confirmed";
    ["invoice-number","invoice-date","invoice-supplier","invoice-payment-type","invoice-due-date","invoice-tax","invoice-discount-method","invoice-discount","invoice-other-cost","invoice-notes","invoice-item-product","invoice-item-qty","invoice-item-buy","invoice-item-sell","invoice-item-discount-method","invoice-item-discount","add-invoice-item","save-invoice-draft"].forEach((id2) => { if ($(id2)) $(id2).disabled = !!locked; });
    $("confirm-invoice").hidden = !!locked;
    renderEditorItems();
    $("invoice-editor-overlay").hidden = false;
    document.body.classList.add("product-detail-open");
}
function closeInvoiceEditor() { $("invoice-editor-overlay") && ($("invoice-editor-overlay").hidden = true); document.body.classList.remove("product-detail-open"); }
function readEditorHeader() {
    const sc = $("invoice-supplier").value;
    const sup = suppliers().find((s) => norm(s["Kode Supplier"]) === norm(sc));
    const discountMethod = $("invoice-discount-method")?.value || "nominal";
    const discountValue = num($("invoice-discount").value);
    const subtotal = editorItems.reduce((sum, item) => sum + invoiceSubtotal(item), 0);
    return {
        number: text($("invoice-number").value), date: $("invoice-date").value, supplierCode: sc,
        supplierName: text(sup?.["Nama Supplier"]), paymentType: $("invoice-payment-type").value,
        dueDate: $("invoice-due-date").value, taxMethod: editingInvoiceId
            ? (getInvoices().find((invoice) => invoice.id === editingInvoiceId)?.taxMethod || "Global")
            : "Global",
        taxPercent: num($("invoice-tax").value),
        discount: discountAmount(discountMethod, discountValue, subtotal),
        discountMethod, discountValue,
        otherCost: num($("invoice-other-cost").value), notes: text($("invoice-notes").value)
    };
}
function renderEditorItems() {
    const body = $("invoice-editor-items"); if (!body) return;
    body.innerHTML = editorItems.length ? editorItems.map((i, idx) => `<tr><td>${text(i.code)} — ${text(i.name)}</td><td>${num(i.qty)}</td><td>${rupiah(i.buyPrice)}</td><td>${rupiah(i.discount)}</td><td>${rupiah(invoiceSubtotal(i))}</td><td><button class="icon-button danger" data-remove-item="${idx}" type="button"><i class="fa-solid fa-trash"></i></button></td></tr>`).join("") : '<tr><td colspan="6" class="empty-table-state">Belum ada item.</td></tr>';
    body.querySelectorAll("[data-remove-item]").forEach((b) => b.addEventListener("click", () => { editorItems.splice(Number(b.dataset.removeItem), 1); renderEditorItems(); }));
    const temp = { ...readEditorHeader(), items: editorItems };
    $("invoice-editor-total").textContent = rupiah(invoiceTotal(temp));
}
function addEditorItem() {
    const idx = Number($("invoice-item-product").value);
    const p = products()[idx]; if (!p) return alert("Pilih produk terlebih dahulu.");
    const qty = num($("invoice-item-qty").value);
    const buyPrice = num($("invoice-item-buy").value) || num(p["Harga Beli"]);
    const discountMethod = $("invoice-item-discount-method")?.value || "nominal";
    const discountValue = num($("invoice-item-discount").value);
    if (qty <= 0) return alert("Qty harus lebih dari 0.");
    if (discountValue < 0 || (discountMethod === "percentage" && discountValue > 100)) return alert("Diskon item harus valid (Persen 0–100).");
    editorItems.push({ code: text(p["Kode Produk"]), name: text(p["Nama Produk"]), unit: text(p["Satuan"]), qty, buyPrice, sellPrice: num($("invoice-item-sell").value) || num(p["Harga Jual"]), discount: discountAmount(discountMethod, discountValue, qty * buyPrice), discountMethod, discountValue, batch: text(p["Batch"]), expired: dateOnly(p["Tanggal Expired"]) });
    renderEditorItems();
}
function saveInvoiceFromEditor(confirmNow = false) {
    const head = readEditorHeader();
    if (!head.number) return alert("Nomor Faktur wajib diisi untuk penyimpanan Management.");
    if (!editorItems.length) return alert("Tambahkan minimal satu item faktur.");
    if (head.paymentType === "Tempo" && !head.dueDate) return alert("Tanggal Jatuh Tempo wajib untuk pembayaran Tempo.");
    if (head.discountValue < 0 || (head.discountMethod === "percentage" && head.discountValue > 100)) return alert("Diskon faktur harus valid (Persen 0–100).");
    const list = getInvoices();
    const duplicate = list.find((x) => norm(x.number) === norm(head.number) && x.id !== editingInvoiceId);
    if (duplicate) return alert("Nomor faktur sudah ada.");
    let inv = editingInvoiceId ? list.find((x) => x.id === editingInvoiceId) : null;
    if (!inv) { inv = { id: uid("INV"), createdAt: nowIso(), stockApplied: false }; list.unshift(inv); }
    Object.assign(inv, head, { items: editorItems.map((i) => ({ ...i })), status: inv.status || "draft", updatedAt: nowIso() });
    saveInvoices(list);
    if (confirmNow) applyInvoiceStock(inv.id); else { closeInvoiceEditor(); renderInvoices(); }
}
async function applyInvoiceStock(invoiceId) {
    const invoices = getInvoices(); const inv = invoices.find((x) => x.id === invoiceId); if (!inv) return;
    if (inv.stockApplied) return alert("Stok dari faktur ini sudah pernah diterapkan.");
    if (!confirm(`Konfirmasi Barang Masuk untuk faktur ${inv.number}? Stok akan bertambah.`)) return;
    const store = getMaster(); const plist = Array.isArray(store.produk) ? store.produk : []; const movements = getMovements();
    for (const item of inv.items || []) {
        const p = plist.find((x) => (item.code && norm(x["Kode Produk"]) === norm(item.code)) || (item.name && norm(x["Nama Produk"]) === norm(item.name)));
        if (!p) continue;
        const before = num(p["Stok Awal"]); const delta = num(item.qty); const after = before + delta;
        p["Stok Awal"] = after;
        if (num(item.buyPrice) > 0) p["Harga Beli"] = num(item.buyPrice);
        if (num(item.sellPrice) > 0) p["Harga Jual"] = num(item.sellPrice);
        if (text(item.batch)) p["Batch"] = text(item.batch);
        if (text(item.expired)) p["Tanggal Expired"] = text(item.expired);
        movements.unshift({ id: uid("MOV"), at: nowIso(), type: "purchase", reference: inv.number, productCode: text(p["Kode Produk"]), productName: text(p["Nama Produk"]), delta, stockAfter: after, note: `Barang masuk faktur ${inv.number}` });
    }
    store.produk = plist;
    inv.status = "confirmed"; inv.stockApplied = true; inv.confirmedAt = nowIso();
    try {
        await writeOperationalDelta([
            { key: MOVEMENT_KEY, records: movements.slice(0, (inv.items || []).length) },
            { key: INVOICE_KEY, records: [inv] }
        ]);
    } catch (error) {
        console.error("Konfirmasi faktur gagal disimpan:", error);
        return alert("Barang masuk belum dapat disimpan. Periksa koneksi/database lalu coba kembali.");
    }
    closeInvoiceEditor(); renderAllOperational();
}
function confirmExistingInvoice(id) { applyInvoiceStock(id); }

async function readWorkbookFile(file) {
    if (!file) throw new Error("Pilih file Excel terlebih dahulu.");
    const buf = await file.arrayBuffer();
    return XLSX.read(buf, { type: "array", cellDates: true });
}
function sheetRows(wb, names) {
    const sheetName = wb.SheetNames.find((n) => names.includes(norm(n).replace(/\s+/g,"_")));
    return sheetName ? XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" }) : [];
}
async function previewInvoiceImport() {
    try {
        const wb = await readWorkbookFile($("invoice-import-file").files?.[0]);
        const headers = sheetRows(wb, ["faktur_pembelian", "faktur pembelian"]);
        const details = sheetRows(wb, ["detail_faktur", "detail faktur"]);
        const map = new Map();
        headers.forEach((h, idx) => {
            const number = text(h["Nomor Faktur"]) || `IMPORT-${idx+1}`;
            map.set(norm(number), { id: uid("INV"), number, date: dateOnly(h["Tanggal Faktur"]), supplierCode: text(h["Kode Supplier"]), supplierName: text(h["Nama Supplier"]), paymentType: text(h["Jenis Pembayaran"]), dueDate: dateOnly(h["Tanggal Jatuh Tempo"]), taxMethod: text(h["Metode PPN"]), taxPercent: num(h["PPN Global (%)"]), discount: num(h["Diskon Faktur"]), otherCost: num(h["Biaya Lain"]), notes: text(h["Catatan"]), status: "draft", stockApplied: false, source: "excel", items: [], createdAt: nowIso() });
        });
        details.forEach((d) => {
            const key = norm(d["Nomor Faktur"]); if (!map.has(key)) map.set(key, { id: uid("INV"), number: text(d["Nomor Faktur"]), status:"draft", stockApplied:false, source:"excel", items:[], createdAt:nowIso() });
            map.get(key).items.push({ code:text(d["Kode Produk"]), name:text(d["Nama Produk"]), qty:num(d["Qty"]), unit:text(d["Satuan"]), buyPrice:num(d["Harga Beli"]), discount:num(d["Diskon Item"]), taxPercent:num(d["PPN Item (%)"]), sellPrice:num(d["Harga Jual"]), batch:text(d["Batch"]), expired:dateOnly(d["Tanggal Expired"]), notes:text(d["Catatan"]) });
        });
        importedInvoicesPreview = [...map.values()];
        $("invoice-import-summary").hidden = false; $("invoice-import-summary").innerHTML = `<strong>${importedInvoicesPreview.length} faktur</strong> terbaca dengan ${details.length} baris detail. Semua akan masuk sebagai Draft.`;
        const body = $("invoice-import-preview-body"); body.innerHTML = importedInvoicesPreview.map((i)=>`<tr><td>${i.number}</td><td>${i.date||"—"}</td><td>${i.supplierName||i.supplierCode||"—"}</td><td>${i.items.length}</td><td>${rupiah(invoiceTotal(i))}</td><td>${statusBadge("draft")}</td></tr>`).join("");
        $("invoice-import-preview").hidden = false; $("invoice-import-actions").hidden = false;
    } catch (e) { alert(e.message || "Gagal membaca faktur."); }
}
async function applyInvoiceImport() {
    if (!importedInvoicesPreview.length) return;
    let list = getInvoices(); let added=0, skipped=0, updated=0;
    for (const incoming of importedInvoicesPreview) {
        const identity=(item)=>`${norm(item.supplierCode||item.supplierName)}::${norm(item.number)}`;
        const idx = list.findIndex((x)=>identity(x)===identity(incoming));
        if (idx < 0) { list.unshift(incoming); added++; continue; }
        const old = list[idx];
        if (old.stockApplied || norm(old.status)==="confirmed") {
            await window.KasirProDialog?.error?.("Faktur Sudah Dikonfirmasi",`Faktur ${old.number} sudah menyesuaikan stok dan tidak dapat diimpor ulang.`); skipped++; continue;
        }
        const replace=await window.KasirProDialog?.confirm?.("Ganti Draft Lama?",`Seluruh isi Draft ${old.number} akan diganti oleh file terbaru. Item tidak akan digabungkan.`,{confirmText:"Ganti Draft"});
        if(!replace){skipped++;continue;}
        incoming.id=old.id;incoming.createdAt=old.createdAt;incoming.updatedAt=nowIso();list[idx]=incoming;updated++;
    }
    saveInvoices(list); alert(`Import selesai. Baru: ${added}, diperbarui: ${updated}, dilewati: ${skipped}.`); importedInvoicesPreview=[]; renderInvoices(); navigate("purchase-invoices");
}

function renderStock() {
    const q = norm($("stock-search")?.value);
    const list = products();
    const supplierNames = new Map(suppliers().map((supplier) => [norm(supplier["Kode Supplier"]), text(supplier["Supplier"] || supplier["Nama Supplier"])]));
    const supplierLabel = (product) => supplierNames.get(norm(product["Supplier"] || product["Kode Supplier"])) || text(product["Supplier"] || product["Kode Supplier"]);
    const updateFilterOptions = (id, labels) => {
        const select = $(id); if (!select) return;
        const selected = select.value;
        const options = [...new Map(labels.filter(Boolean).map((label) => [norm(label), label])).entries()]
            .sort((a, b) => a[1].localeCompare(b[1], "id-ID"));
        select.replaceChildren(select.options[0], ...options.map(([value, label]) => new Option(label, value)));
        select.value = selected;
    };
    updateFilterOptions("stock-category-filter", list.map((p) => text(p["Kategori"])));
    updateFilterOptions("stock-supplier-filter", list.map(supplierLabel));
    const category = $("stock-category-filter")?.value || "";
    const supplier = $("stock-supplier-filter")?.value || "";
    const status = $("stock-status-filter")?.value || "";
    const stockStatus = (p) => {
        const stock = num(p["Stok Awal"]), minimum = num(p["Stok Minimum"]);
        return stock <= 0 ? "empty" : minimum > 0 && stock <= minimum ? "low" : "safe";
    };
    const filtered = list.filter((p) =>
        (!q || [p["Kode Produk"], p["Nama Produk"]].some((x) => norm(x).includes(q)))
        && (!category || norm(p["Kategori"]) === category)
        && (!supplier || norm(supplierLabel(p)) === supplier)
        && (!status || stockStatus(p) === status)
    );
    if ($("stock-reset-filter")) $("stock-reset-filter").hidden = !(q || category || supplier || status);
    $("stock-product-count") && ($("stock-product-count").textContent = list.length);
    $("stock-unit-count") && ($("stock-unit-count").textContent = list.reduce((a,p)=>a+num(p["Stok Awal"]),0));
    $("stock-alert-count") && ($("stock-alert-count").textContent = list.filter((p)=>stockStatus(p)!=="safe").length);
    const body=$("stock-table-body"); if(!body)return;
    const totalPages=Math.max(1,Math.ceil(filtered.length/STOCK_PAGE_SIZE));
    stockCurrentPage=Math.min(Math.max(1,stockCurrentPage),totalPages);
    const first=(stockCurrentPage-1)*STOCK_PAGE_SIZE;
    const visible=filtered.slice(first,first+STOCK_PAGE_SIZE);
    body.innerHTML=visible.length?visible.map((p)=>{const st=num(p["Stok Awal"]), mn=num(p["Stok Minimum"]); const badge=st<=0?'<span class="stock-badge stock-empty">Habis</span>':mn>0&&st<=mn?'<span class="stock-badge stock-low">Menipis</span>':'<span class="stock-badge stock-safe">Aman</span>'; return `<tr><td>${text(p["Kode Produk"])||"—"}</td><td>${text(p["Nama Produk"])||"—"}</td><td>${text(p["Kategori"])||"—"}</td><td><div class="product-stock"><strong>${st}</strong>${badge}</div></td><td>${mn}</td><td>${rupiah(p["Harga Beli"])}</td><td>${rupiah(st*num(p["Harga Beli"]))}</td><td>${text(p["Lokasi Rak"])||"—"}</td></tr>`}).join(""):'<tr><td colspan="8" class="empty-table-state">Tidak ada produk.</td></tr>';
    const info=$("stock-page-info"); if(info) info.textContent=filtered.length?`Menampilkan ${first+1}–${first+visible.length} dari ${filtered.length.toLocaleString("id-ID")} produk`:'Tidak ada produk yang cocok';
    if($("stock-prev-page")) $("stock-prev-page").disabled=stockCurrentPage<=1;
    if($("stock-next-page")) $("stock-next-page").disabled=stockCurrentPage>=totalPages;
}

function scheduleStockRender() {
    const token=++stockRenderToken;
    const status=$("stock-loading-status");
    if(status) status.hidden=false;
    requestAnimationFrame(()=>requestAnimationFrame(()=>{
        if(token!==stockRenderToken)return;
        try { renderStock(); }
        finally { if(status) status.hidden=true; }
    }));
}
function renderMovements() {
    const q=norm($("movement-search")?.value), tf=$("movement-type-filter")?.value||""; const rows=getMovements().filter((m)=>(!tf||m.type===tf)&&(!q||[m.reference,m.productCode,m.productName,m.type,m.note].some((x)=>norm(x).includes(q))));
    const body=$("movement-table-body"); if(!body)return; body.innerHTML=rows.length?rows.map((m)=>`<tr><td>${fmtDateTime(m.at)}</td><td>${movementBadge(m.type)}</td><td>${text(m.reference)||"—"}</td><td>${text(m.productCode)||"—"}</td><td>${text(m.productName)||"—"}</td><td class="${num(m.delta)>=0?'qty-plus':'qty-minus'}">${num(m.delta)>=0?'+':''}${num(m.delta)}</td><td>${num(m.stockAfter)}</td><td>${text(m.note)||"—"}</td></tr>`).join(""):'<tr><td colspan="8" class="empty-table-state">Belum ada mutasi stok.</td></tr>';
}
function exportStockOpname() {
    const data=products().map((p)=>({"Kode Produk":text(p["Kode Produk"]),"Nama Produk":text(p["Nama Produk"]),"Kategori":text(p["Kategori"]),"Supplier":text(p["Kode Supplier"]),"Satuan":text(p["Satuan"]),"Stok Sistem":num(p["Stok Awal"]),"Stok Fisik":"","Selisih":"","Harga Beli":num(p["Harga Beli"]),"Nilai Selisih":"","Lokasi Rak":text(p["Lokasi Rak"]),"Batch":text(p["Batch"]),"Tanggal Expired":dateOnly(p["Tanggal Expired"]),"Keterangan":""}));
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(data),"STOCK_OPNAME"); XLSX.writeFile(wb,`Stock_Opname_${new Date().toISOString().slice(0,10)}.xlsx`);
}
async function previewOpnameImport() {
    try {
        const wb=await readWorkbookFile($("opname-import-file").files?.[0]); const rows=sheetRows(wb,["stock_opname","stock opname"]); importedOpnamePreview=[];
        for(const r of rows){ const p=findProduct({code:r["Kode Produk"],name:r["Nama Produk"]}); if(!p) continue; const system=num(p["Stok Awal"]); const physical=num(r["Stok Fisik"]); const diff=physical-system; importedOpnamePreview.push({productCode:text(p["Kode Produk"]),productName:text(p["Nama Produk"]),system,physical,diff,buyPrice:num(p["Harga Beli"]),valueDiff:diff*num(p["Harga Beli"]),note:text(r["Keterangan"])}); }
        $("opname-import-summary").hidden=false; $("opname-import-summary").innerHTML=`<strong>${importedOpnamePreview.length} produk</strong> cocok dengan master. Total selisih qty: ${importedOpnamePreview.reduce((a,x)=>a+x.diff,0)}.`;
        $("opname-import-preview-body").innerHTML=importedOpnamePreview.map((x)=>`<tr><td>${x.productCode}</td><td>${x.productName}</td><td>${x.system}</td><td>${x.physical}</td><td class="${x.diff>=0?'qty-plus':'qty-minus'}">${x.diff>=0?'+':''}${x.diff}</td><td>${rupiah(x.valueDiff)}</td></tr>`).join(""); $("opname-import-preview").hidden=false; $("opname-import-actions").hidden=false;
    } catch(e){alert(e.message||"Gagal membaca stock opname.");}
}
async function confirmOpnameImport() {
    if(!importedOpnamePreview.length)return; if(!confirm("Konfirmasi adjustment Stock Opname? Stok produk akan disesuaikan ke Stok Fisik."))return;
    const store=getMaster(); const plist=Array.isArray(store.produk)?store.produk:[]; const moves=getMovements(); const ref=`OPN-${Date.now()}`;
    importedOpnamePreview.forEach((x)=>{const p=plist.find((p)=>norm(p["Kode Produk"])===norm(x.productCode)); if(!p||x.diff===0)return; p["Stok Awal"]=x.physical; moves.unshift({id:uid("MOV"),at:nowIso(),type:"opname",reference:ref,productCode:x.productCode,productName:x.productName,delta:x.diff,stockAfter:x.physical,note:x.note||"Adjustment stock opname"});});
    store.produk=plist;
    const hist=getOpnameHistory();
    hist.unshift({id:ref,at:nowIso(),productCount:importedOpnamePreview.filter((x)=>x.diff!==0).length,qtyDiff:importedOpnamePreview.reduce((a,x)=>a+x.diff,0),valueDiff:importedOpnamePreview.reduce((a,x)=>a+x.valueDiff,0)});
    try {
        await writeOperationalDelta([
            {key:MOVEMENT_KEY,records:moves.slice(0, importedOpnamePreview.filter((x)=>x.diff!==0).length)},
            {key:OPNAME_KEY,records:[hist[0]]}
        ]);
    } catch (error) {
        console.error("Stock opname gagal disimpan:",error);
        return alert("Stock Opname belum dapat disimpan. Periksa koneksi/database lalu coba kembali.");
    }
    importedOpnamePreview=[]; alert("Stock Opname berhasil diterapkan."); renderAllOperational(); navigate("stock-opname");
}
function renderOpnameHistory(){const body=$("opname-history-body");if(!body)return;const h=getOpnameHistory();body.innerHTML=h.length?h.map((x)=>`<tr><td>${fmtDateTime(x.at)}</td><td>${x.id}</td><td>${x.productCount}</td><td class="${num(x.qtyDiff)>=0?'qty-plus':'qty-minus'}">${num(x.qtyDiff)>=0?'+':''}${num(x.qtyDiff)}</td><td>${rupiah(x.valueDiff)}</td></tr>`).join(""):'<tr><td colspan="5" class="empty-table-state">Belum ada riwayat Stock Opname.</td></tr>'}

const SALES_PAGE_SIZE = 20;
let salesCurrentPage = 1;
let selectedManagementSaleId = null;
let filteredManagementSales = [];

function getOperationalSession() {
    try { return JSON.parse(sessionStorage.getItem("kasirpro_session") || "null") || {}; }
    catch { return {}; }
}

function getStoreSettings() {
    const store = getMaster();
    for (const key of ["pengaturanToko", "pengaturan_toko", "pengaturan"]) {
        if (Array.isArray(store[key])) return store[key][0] || {};
        if (store[key] && typeof store[key] === "object") return store[key];
    }
    return {};
}

function managementSaleStatus(sale) {
    return norm(sale?.status) === "void" ? "VOID" : "SELESAI";
}

function managementSalePeriod(sale, period) {
    const date = new Date(sale?.at);
    if (Number.isNaN(date.getTime())) return period === "all";
    if (period === "all") return true;
    const now = new Date();
    if (period === "today") return date.getFullYear() === now.getFullYear() && date.getMonth() === now.getMonth() && date.getDate() === now.getDate();
    if (period === "month") return date.getFullYear() === now.getFullYear() && date.getMonth() === now.getMonth();
    const days = Number(period);
    return days ? date.getTime() >= now.getTime() - days * 24 * 60 * 60 * 1000 : true;
}

function managementSaleMatches(sale, query) {
    if (!query) return true;
    const items = (Array.isArray(sale.items) ? sale.items : []).map((item) => `${item.code || ""} ${item.name || ""}`).join(" ");
    return [sale.number, sale.cashier, sale.paymentMethod, sale.status, items].some((value) => norm(value).includes(query));
}

function salesPeriodLabel(period) {
    return ({ today: "Hari ini", "7": "7 hari terakhir", "30": "30 hari terakhir", month: "Bulan ini", all: "Semua data" })[period] || "Periode aktif";
}

function getFilteredManagementSales() {
    const query = norm($("sales-search")?.value);
    const period = $("sales-period")?.value || "month";
    const status = $("sales-status-filter")?.value || "";
    const payment = norm($("sales-payment-filter")?.value);
    return getSales()
        .filter((sale) => managementSalePeriod(sale, period))
        .filter((sale) => !status || managementSaleStatus(sale) === status)
        .filter((sale) => !payment || norm(sale.paymentMethod) === payment)
        .filter((sale) => managementSaleMatches(sale, query))
        .sort((a, b) => (new Date(b.at).getTime() || 0) - (new Date(a.at).getTime() || 0));
}

function renderManagementSales() {
    if (!$("sales-table-body")) return;
    filteredManagementSales = getFilteredManagementSales();
    const completed = filteredManagementSales.filter((sale) => managementSaleStatus(sale) === "SELESAI");
    const voided = filteredManagementSales.filter((sale) => managementSaleStatus(sale) === "VOID");
    const units = completed.reduce((sum, sale) => sum + (sale.items || []).reduce((itemSum, item) => itemSum + num(item.qty), 0), 0);
    const revenue = completed.reduce((sum, sale) => sum + num(sale.total), 0);

    $("sales-total-count").textContent = filteredManagementSales.length;
    $("sales-completed-count").textContent = completed.length;
    $("sales-unit-count").textContent = units.toLocaleString("id-ID");
    $("sales-revenue-total").textContent = rupiah(revenue);
    $("sales-void-count").textContent = voided.length;
    $("sales-result-count").textContent = `${filteredManagementSales.length} data`;

    const totalPages = Math.max(1, Math.ceil(filteredManagementSales.length / SALES_PAGE_SIZE));
    salesCurrentPage = Math.max(1, Math.min(salesCurrentPage, totalPages));
    const start = (salesCurrentPage - 1) * SALES_PAGE_SIZE;
    const rows = filteredManagementSales.slice(start, start + SALES_PAGE_SIZE);

    $("sales-table-body").innerHTML = rows.length ? rows.map((sale) => {
        const status = managementSaleStatus(sale);
        const units = (sale.items || []).reduce((sum, item) => sum + num(item.qty), 0);
        const discount = num(sale.itemDiscount) + num(sale.transDiscount);
        const identifier = reportEscape(sale.id || sale.number);
        return `<tr><td>${fmtDateTime(sale.at)}</td><td>${reportEscape(sale.number || "—")}</td><td>${reportEscape(sale.cashier || "—")}</td><td>${reportEscape(sale.paymentMethod || "—")}</td><td>${units}</td><td>${rupiah(sale.subtotal)}</td><td>${rupiah(discount)}</td><td>${rupiah(sale.tax)}</td><td>${rupiah(sale.total)}</td><td><span class="ops-badge ${status === "VOID" ? "danger" : "ok"}">${status}</span></td><td><div class="sales-action-group"><button class="icon-button" type="button" data-sale-detail="${identifier}" title="Lihat Detail"><i class="fa-solid fa-eye"></i></button><button class="icon-button" type="button" data-sale-print="${identifier}" title="Cetak Ulang"><i class="fa-solid fa-print"></i></button></div></td></tr>`;
    }).join("") : '<tr><td colspan="11" class="empty-table-state">Tidak ada transaksi pada filter yang dipilih.</td></tr>';

    $("sales-page-info").textContent = filteredManagementSales.length ? `Menampilkan ${start + 1}–${Math.min(start + SALES_PAGE_SIZE, filteredManagementSales.length)} dari ${filteredManagementSales.length} · Halaman ${salesCurrentPage}/${totalPages}` : "Tidak ada data";
    $("sales-prev-page").disabled = salesCurrentPage <= 1;
    $("sales-next-page").disabled = salesCurrentPage >= totalPages;
    const query = text($("sales-search")?.value);
    $("sales-filter-label").textContent = `${salesPeriodLabel($("sales-period").value)}${query ? ` · Pencarian: “${query}”` : ""}`;
    $("sales-last-updated").textContent = `Diperbarui ${new Date().toLocaleString("id-ID")}`;

    $("sales-table-body").querySelectorAll("[data-sale-detail]").forEach((button) => button.addEventListener("click", () => openManagementSaleDetail(button.dataset.saleDetail)));
    $("sales-table-body").querySelectorAll("[data-sale-print]").forEach((button) => button.addEventListener("click", () => { openManagementSaleDetail(button.dataset.salePrint); printManagementSaleDetail(); }));
}

function findManagementSale(identifier) {
    return getSales().find((sale) => String(sale.id || sale.number) === String(identifier)) || null;
}

function openManagementSaleDetail(identifier) {
    const sale = findManagementSale(identifier);
    if (!sale) return alert("Transaksi tidak ditemukan.");
    selectedManagementSaleId = sale.id || sale.number;
    const status = managementSaleStatus(sale);
    const items = Array.isArray(sale.items) ? sale.items : [];
    const settings = getStoreSettings();
    const storeName = text(settings["Nama Toko"]) || "KasirPro";
    const storeContact = [settings["Alamat"], settings["Telepon"]].filter(Boolean).join("\n");
    $("sales-detail-title").textContent = sale.number || "Transaksi";
    $("sales-detail-subtitle").textContent = `${fmtDateTime(sale.at)} · ${sale.cashier || "Kasir"}`;
    $("sales-detail-print-area").innerHTML = `
        <div class="sales-detail-store"><h3>${reportEscape(storeName)}</h3><p>${reportEscape(storeContact)}</p></div>
        <div class="sales-detail-summary"><div><span>Status</span><strong><span class="ops-badge ${status === "VOID" ? "danger" : "ok"}">${status}</span></strong></div><div><span>Waktu</span><strong>${fmtDateTime(sale.at)}</strong></div><div><span>Kasir</span><strong>${reportEscape(sale.cashier || "—")}</strong></div><div><span>Pembayaran</span><strong>${reportEscape(sale.paymentMethod || "—")}</strong></div></div>
        <div class="sales-detail-items">${items.length ? items.map((item) => `<div class="sales-detail-item"><div><strong>${reportEscape(item.name || "Produk")}</strong><small>${reportEscape(item.code || "Tanpa kode")} · ${num(item.qty)} × ${rupiah(item.price)}${num(item.discount) > 0 ? ` · Diskon ${rupiah(item.discount)}` : ""}</small></div><strong>${rupiah(num(item.qty) * num(item.price) - num(item.discount))}</strong></div>`).join("") : '<div class="report-empty">Tidak ada item transaksi.</div>'}</div>
        <div class="sales-detail-totals"><div><span>Subtotal</span><strong>${rupiah(sale.subtotal)}</strong></div><div><span>Diskon</span><strong>-${rupiah(num(sale.itemDiscount) + num(sale.transDiscount))}</strong></div><div><span>Pajak</span><strong>${rupiah(sale.tax)}</strong></div><div class="total"><span>Total</span><strong>${rupiah(sale.total)}</strong></div><div><span>Dibayar (${reportEscape(sale.paymentMethod || "—")})</span><strong>${rupiah(sale.paid)}</strong></div><div><span>Kembalian</span><strong>${rupiah(sale.change)}</strong></div></div>
        ${status === "VOID" ? `<div class="sales-void-note"><strong>Transaksi telah di-VOID.</strong><br>Alasan: ${reportEscape(sale.voidReason || "—")}<br>Oleh: ${reportEscape(sale.voidBy || "Admin")}<br>${sale.voidAt ? fmtDateTime(sale.voidAt) : ""}</div>` : ""}`;
    $("sales-detail-void").hidden = status === "VOID";
    $("sales-detail-overlay").hidden = false;
    document.body.classList.add("product-detail-open");
}

function closeManagementSaleDetail() {
    $("sales-detail-overlay") && ($("sales-detail-overlay").hidden = true);
    document.body.classList.remove("product-detail-open");
    document.body.classList.remove("sales-detail-printing");
}

function printManagementSaleDetail() {
    if (!selectedManagementSaleId || $("sales-detail-overlay")?.hidden) return;
    document.body.classList.add("sales-detail-printing");
    window.print();
    setTimeout(() => document.body.classList.remove("sales-detail-printing"), 100);
}

async function voidManagementSale() {
    const saleList = getSales();
    const sale = saleList.find((item) => String(item.id || item.number) === String(selectedManagementSaleId));
    if (!sale) return alert("Transaksi tidak ditemukan.");
    if (managementSaleStatus(sale) === "VOID") return alert("Transaksi ini sudah di-VOID.");
    const reason = text(prompt("Masukkan alasan VOID transaksi:"));
    if (!reason) return;
    if (!confirm(`VOID transaksi ${sale.number}? Stok seluruh item akan dikembalikan.`)) return;

    const store = getMaster();
    const productList = Array.isArray(store.produk) ? store.produk : [];
    const items = Array.isArray(sale.items) ? sale.items : [];
    const resolved = items.map((item) => {
        const code = norm(item.code);
        const name = norm(item.name);
        const product = productList.find((entry) => code && norm(entry["Kode Produk"]) === code)
            || productList.find((entry) => name && norm(entry["Nama Produk"]) === name)
            || null;
        return { item, product };
    });
    const missing = resolved.find((entry) => !entry.product);
    if (missing) return alert(`Produk ${missing.item.name || missing.item.code || "transaksi"} tidak ditemukan. VOID dibatalkan agar stok tidak berubah sebagian.`);

    const movements = getMovements();
    const voidAt = nowIso();
    const session = getOperationalSession();
    const voidBy = session.name || session.nama || session.username || "Administrator";
    resolved.forEach(({ item, product }) => {
        const after = num(product["Stok Awal"]) + num(item.qty);
        product["Stok Awal"] = after;
        movements.unshift({ id: uid("MOV"), at: voidAt, type: "sale_void", reference: sale.number, productCode: text(product["Kode Produk"] || item.code), productName: text(product["Nama Produk"] || item.name), delta: num(item.qty), stockAfter: after, note: `VOID penjualan ${sale.number}: ${reason}` });
    });
    sale.status = "VOID";
    sale.voidAt = voidAt;
    sale.voidBy = voidBy;
    sale.voidReason = reason;
    store.produk = productList;
    try {
        await writeOperationalDelta([
            {key:MOVEMENT_KEY,records:movements.slice(0,(sale.items||[]).length)},
            {key:SALES_KEY,records:[sale]}
        ]);
    } catch (error) {
        console.error("VOID gagal disimpan:",error);
        return alert("VOID belum dapat disimpan. Periksa koneksi/database lalu coba kembali.");
    }
    renderManagementSales();
    renderStock();
    renderMovements();
    renderReports();
    openManagementSaleDetail(sale.id || sale.number);
}

function sanitizeBackupValue(value) {
    if (Array.isArray(value)) return value.map(sanitizeBackupValue);
    if (!value || typeof value !== "object") return value;
    return Object.fromEntries(Object.entries(value)
        .filter(([key]) => !["Password Awal", "password", "_firestoreDocumentId", "_localSyncStatus"].includes(key))
        .map(([key, item]) => [key, sanitizeBackupValue(item)]));
}

function exportDatabaseBackup() {
    const session = getOperationalSession();
    if (session?.role !== "admin") return alert("Backup database hanya dapat dibuat oleh Administrator.");

    const backup = {
        application: "Kasir Pro V2",
        store: "Toko Utama",
        schemaVersion: 2,
        exportedAt: nowIso(),
        exportedBy: session.username || session.name || "admin",
        note: "Backup operasional. Password dan metadata lokal tidak disertakan.",
        data: sanitizeBackupValue({
            master: getMaster(),
            fakturPembelian: getInvoices(),
            transaksiPenjualan: getSales(),
            mutasiStok: getMovements(),
            stockOpname: getOpnameHistory()
        })
    };
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `KasirPro_V2_Backup_${reportDateInput(new Date())}.json`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

let activeReportTab = "summary";
let reportExportRows = {};

function reportEscape(value) {
    return String(value ?? "").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[char]);
}

function reportDate(value) {
    const result = new Date(value);
    return Number.isNaN(result.getTime()) ? null : result;
}

function reportStartOfDay(value) {
    const result = new Date(value);
    result.setHours(0, 0, 0, 0);
    return result;
}

function reportEndOfDay(value) {
    const result = new Date(value);
    result.setHours(23, 59, 59, 999);
    return result;
}

function reportDateInput(value) {
    const date = new Date(value);
    const offset = date.getTimezoneOffset() * 60000;
    return new Date(date.getTime() - offset).toISOString().slice(0, 10);
}

function reportRange() {
    const mode = $("report-period")?.value || "month";
    const now = new Date();
    let start = null;
    let end = null;
    let label = "Semua data";

    if (mode === "today") {
        start = reportStartOfDay(now);
        end = reportEndOfDay(now);
        label = "Hari ini";
    } else if (mode === "7" || mode === "30") {
        const days = Number(mode);
        start = reportStartOfDay(now);
        start.setDate(start.getDate() - (days - 1));
        end = reportEndOfDay(now);
        label = `${days} hari terakhir`;
    } else if (mode === "month") {
        start = new Date(now.getFullYear(), now.getMonth(), 1);
        end = reportEndOfDay(now);
        label = now.toLocaleDateString("id-ID", { month: "long", year: "numeric" });
    } else if (mode === "custom") {
        const startValue = $("report-start-date")?.value;
        const endValue = $("report-end-date")?.value;
        start = startValue ? reportStartOfDay(`${startValue}T00:00:00`) : null;
        end = endValue ? reportEndOfDay(`${endValue}T00:00:00`) : null;
        label = startValue || endValue
            ? `${startValue ? new Date(`${startValue}T00:00:00`).toLocaleDateString("id-ID") : "Awal"} – ${endValue ? new Date(`${endValue}T00:00:00`).toLocaleDateString("id-ID") : "Sekarang"}`
            : "Tanggal khusus belum dipilih";
    }

    return { mode, start, end, label };
}

function reportInRange(value, range) {
    if (!range.start && !range.end) return true;
    const date = reportDate(value);
    if (!date) return false;
    if (range.start && date < range.start) return false;
    if (range.end && date > range.end) return false;
    return true;
}

function reportSaleStatus(sale) {
    return norm(sale?.status) === "void" ? "VOID" : "SELESAI";
}

function reportSaleCost(sale) {
    return (Array.isArray(sale?.items) ? sale.items : []).reduce((sum, item) => sum + num(item.buyPrice) * num(item.qty), 0);
}

function reportSaleNetBeforeTax(sale) {
    return Math.max(0, num(sale?.subtotal) - num(sale?.itemDiscount) - num(sale?.transDiscount));
}

function reportSaleProfit(sale) {
    return reportSaleNetBeforeTax(sale) - reportSaleCost(sale);
}

function reportSaleMatches(sale, query) {
    if (!query) return true;
    const itemText = (Array.isArray(sale.items) ? sale.items : []).map((item) => `${item.code || ""} ${item.name || ""}`).join(" ");
    return [sale.number, sale.cashier, sale.paymentMethod, sale.status, itemText].some((value) => norm(value).includes(query));
}

function reportProductRows(sales) {
    const map = new Map();

    sales.forEach((sale) => {
        const items = Array.isArray(sale.items) ? sale.items : [];
        const itemNetTotal = items.reduce((sum, item) => {
            const gross = num(item.price) * num(item.qty);
            return sum + Math.max(0, gross - Math.min(num(item.discount), gross));
        }, 0);

        items.forEach((item) => {
            const code = text(item.code);
            const name = text(item.name) || "Produk";
            const key = code ? `code:${norm(code)}` : `name:${norm(name)}`;
            const gross = num(item.price) * num(item.qty);
            const afterItemDiscount = Math.max(0, gross - Math.min(num(item.discount), gross));
            const transactionShare = itemNetTotal > 0 ? num(sale.transDiscount) * afterItemDiscount / itemNetTotal : 0;
            const revenue = Math.max(0, afterItemDiscount - transactionShare);
            const cost = num(item.buyPrice) * num(item.qty);
            const current = map.get(key) || { code, name, qty: 0, revenue: 0, cost: 0, profit: 0 };
            current.qty += num(item.qty);
            current.revenue += revenue;
            current.cost += cost;
            current.profit += revenue - cost;
            map.set(key, current);
        });
    });

    return [...map.values()].sort((a, b) => b.qty - a.qty || b.revenue - a.revenue);
}

function reportData() {
    const range = reportRange();
    const query = norm($("report-search")?.value);
    const allSales = getSales().filter((sale) => reportInRange(sale.at, range));
    const sales = allSales.filter((sale) => reportSaleMatches(sale, query));
    const completedSales = sales.filter((sale) => reportSaleStatus(sale) === "SELESAI");
    const voidSales = sales.filter((sale) => reportSaleStatus(sale) === "VOID");
    const invoices = getInvoices().filter((invoice) => reportInRange(invoice.confirmedAt || invoice.date || invoice.createdAt, range))
        .filter((invoice) => !query || [invoice.number, invoice.supplierCode, invoice.supplierName, invoice.status].some((value) => norm(value).includes(query)));
    const movements = getMovements().filter((movement) => movement.type === "purchase" && reportInRange(movement.at, range))
        .filter((movement) => !query || [movement.reference, movement.productCode, movement.productName, movement.note].some((value) => norm(value).includes(query)));
    const stock = products().filter((product) => !query || [product["Kode Produk"], product["Nama Produk"], product["Kategori"], product["Lokasi Rak"]].some((value) => norm(value).includes(query)));

    return { range, query, sales, completedSales, voidSales, invoices, movements, stock };
}

function renderReportSummary(data, productRows) {
    const revenue = data.completedSales.reduce((sum, sale) => sum + num(sale.total), 0);
    const units = productRows.reduce((sum, item) => sum + item.qty, 0);
    const profit = data.completedSales.reduce((sum, sale) => sum + reportSaleProfit(sale), 0);
    const confirmedInvoices = data.invoices.filter((invoice) => norm(invoice.status) === "confirmed");
    const purchaseValue = confirmedInvoices.reduce((sum, invoice) => sum + invoiceTotal(invoice), 0);
    const stockValue = data.stock.reduce((sum, product) => sum + num(product["Stok Awal"]) * num(product["Harga Beli"]), 0);
    const lowStock = data.stock.filter((product) => {
        const stock = num(product["Stok Awal"]);
        const minimum = num(product["Stok Minimum"]);
        return stock <= 0 || (minimum > 0 && stock <= minimum);
    }).length;

    $("report-revenue").textContent = rupiah(revenue);
    $("report-transaction-count").textContent = data.completedSales.length;
    $("report-void-count").textContent = `${data.voidSales.length} transaksi VOID`;
    $("report-unit-count").textContent = units.toLocaleString("id-ID");
    $("report-profit").textContent = rupiah(profit);
    $("report-purchase-value").textContent = rupiah(purchaseValue);
    $("report-invoice-count").textContent = `${data.invoices.length} faktur; ${confirmedInvoices.length} terkonfirmasi`;
    $("report-stock-value").textContent = rupiah(stockValue);
    $("report-low-stock-count").textContent = `${lowStock} produk menipis/habis`;

    const dailyMap = new Map();
    data.completedSales.forEach((sale) => {
        const date = reportDate(sale.at);
        if (!date) return;
        const key = reportDateInput(date);
        const row = dailyMap.get(key) || { date: key, transactions: 0, units: 0, revenue: 0, profit: 0 };
        row.transactions += 1;
        row.units += (sale.items || []).reduce((sum, item) => sum + num(item.qty), 0);
        row.revenue += num(sale.total);
        row.profit += reportSaleProfit(sale);
        dailyMap.set(key, row);
    });
    const dailyRows = [...dailyMap.values()].sort((a, b) => b.date.localeCompare(a.date));
    $("report-daily-body").innerHTML = dailyRows.length
        ? dailyRows.map((row) => `<tr><td>${new Date(`${row.date}T00:00:00`).toLocaleDateString("id-ID", { weekday: "short", day: "2-digit", month: "short", year: "numeric" })}</td><td>${row.transactions}</td><td>${row.units}</td><td>${rupiah(row.revenue)}</td><td class="${row.profit >= 0 ? "qty-plus" : "qty-minus"}">${rupiah(row.profit)}</td></tr>`).join("")
        : '<tr><td colspan="5" class="empty-table-state">Belum ada transaksi selesai pada periode ini.</td></tr>';

    const paymentMap = new Map();
    data.completedSales.forEach((sale) => {
        const method = text(sale.paymentMethod) || "Tidak diketahui";
        const row = paymentMap.get(method) || { method, count: 0, revenue: 0 };
        row.count += 1;
        row.revenue += num(sale.total);
        paymentMap.set(method, row);
    });
    const payments = [...paymentMap.values()].sort((a, b) => b.revenue - a.revenue);
    const paymentTotal = payments.reduce((sum, row) => sum + row.revenue, 0);
    $("report-payment-list").innerHTML = payments.length
        ? payments.map((row) => { const percent = paymentTotal > 0 ? row.revenue / paymentTotal * 100 : 0; return `<div class="report-breakdown-item"><div class="report-breakdown-head"><span>${reportEscape(row.method)} · ${row.count} transaksi</span><strong>${rupiah(row.revenue)}</strong></div><div class="report-progress"><span style="width:${Math.max(2, percent)}%"></span></div></div>`; }).join("")
        : '<div class="report-empty">Belum ada data pembayaran.</div>';

    reportExportRows.summary = dailyRows.map((row) => ({ "Tanggal": row.date, "Transaksi": row.transactions, "Produk Terjual": row.units, "Omzet": row.revenue, "Estimasi Laba": row.profit }));
}

function renderReportSales(data) {
    $("report-sales-count").textContent = `${data.sales.length} data`;
    $("report-sales-body").innerHTML = data.sales.length
        ? data.sales.sort((a, b) => (reportDate(b.at)?.getTime() || 0) - (reportDate(a.at)?.getTime() || 0)).map((sale) => {
            const units = (sale.items || []).reduce((sum, item) => sum + num(item.qty), 0);
            const discount = num(sale.itemDiscount) + num(sale.transDiscount);
            const status = reportSaleStatus(sale);
            return `<tr><td>${fmtDateTime(sale.at)}</td><td>${reportEscape(sale.number || "—")}</td><td>${reportEscape(sale.cashier || "—")}</td><td>${reportEscape(sale.paymentMethod || "—")}</td><td>${units}</td><td>${rupiah(sale.subtotal)}</td><td>${rupiah(discount)}</td><td>${rupiah(sale.tax)}</td><td>${rupiah(sale.total)}</td><td><span class="ops-badge ${status === "VOID" ? "danger" : "ok"}">${status}</span></td></tr>`;
        }).join("")
        : '<tr><td colspan="10" class="empty-table-state">Tidak ada transaksi pada periode atau pencarian ini.</td></tr>';
    reportExportRows.sales = data.sales.map((sale) => ({ "Waktu": fmtDateTime(sale.at), "Nomor Transaksi": sale.number, "Kasir": sale.cashier, "Pembayaran": sale.paymentMethod, "Jumlah Produk": (sale.items || []).reduce((sum, item) => sum + num(item.qty), 0), "Subtotal": num(sale.subtotal), "Diskon": num(sale.itemDiscount) + num(sale.transDiscount), "Pajak": num(sale.tax), "Total": num(sale.total), "Status": reportSaleStatus(sale) }));
}

function renderReportProducts(productRows) {
    $("report-products-count").textContent = `${productRows.length} produk`;
    $("report-products-body").innerHTML = productRows.length
        ? productRows.map((row) => `<tr><td>${reportEscape(row.code || "—")}</td><td>${reportEscape(row.name)}</td><td>${row.qty}</td><td>${rupiah(row.revenue)}</td><td>${rupiah(row.cost)}</td><td class="${row.profit >= 0 ? "qty-plus" : "qty-minus"}">${rupiah(row.profit)}</td></tr>`).join("")
        : '<tr><td colspan="6" class="empty-table-state">Belum ada produk terjual pada periode ini.</td></tr>';
    reportExportRows.products = productRows.map((row) => ({ "Kode Produk": row.code, "Nama Produk": row.name, "Qty Terjual": row.qty, "Omzet Bersih": row.revenue, "HPP": row.cost, "Estimasi Laba": row.profit }));
}

function renderReportStock(data) {
    $("report-stock-count").textContent = `${data.stock.length} produk`;
    $("report-stock-body").innerHTML = data.stock.length
        ? data.stock.map((product) => {
            const stock = num(product["Stok Awal"]);
            const minimum = num(product["Stok Minimum"]);
            const state = stock <= 0 ? ["Habis", "danger"] : minimum > 0 && stock <= minimum ? ["Menipis", "warn"] : ["Aman", "ok"];
            return `<tr><td>${reportEscape(product["Kode Produk"] || "—")}</td><td>${reportEscape(product["Nama Produk"] || "—")}</td><td>${reportEscape(product["Kategori"] || "—")}</td><td>${stock}</td><td>${minimum}</td><td>${rupiah(product["Harga Beli"])}</td><td>${rupiah(stock * num(product["Harga Beli"]))}</td><td><span class="ops-badge ${state[1]}">${state[0]}</span></td></tr>`;
        }).join("")
        : '<tr><td colspan="8" class="empty-table-state">Tidak ada produk yang sesuai.</td></tr>';
    reportExportRows.stock = data.stock.map((product) => ({ "Kode Produk": product["Kode Produk"], "Nama Produk": product["Nama Produk"], "Kategori": product["Kategori"], "Stok": num(product["Stok Awal"]), "Stok Minimum": num(product["Stok Minimum"]), "Harga Beli": num(product["Harga Beli"]), "Nilai Stok": num(product["Stok Awal"]) * num(product["Harga Beli"]) }));
}

function renderReportGoods(data) {
    const productMap = new Map(products().map((product) => [norm(product["Kode Produk"]), product]));
    $("report-goods-count").textContent = `${data.movements.length} mutasi`;
    $("report-goods-body").innerHTML = data.movements.length
        ? data.movements.map((movement) => {
            const product = productMap.get(norm(movement.productCode));
            const value = num(movement.delta) * num(product?.["Harga Beli"]);
            return `<tr><td>${fmtDateTime(movement.at)}</td><td>${reportEscape(movement.reference || "—")}</td><td>${reportEscape(movement.productCode || "—")}</td><td>${reportEscape(movement.productName || "—")}</td><td class="qty-plus">+${num(movement.delta)}</td><td>${num(movement.stockAfter)}</td><td>${rupiah(value)}</td></tr>`;
        }).join("")
        : '<tr><td colspan="7" class="empty-table-state">Tidak ada barang masuk pada periode ini.</td></tr>';
    reportExportRows.goods = data.movements.map((movement) => { const product = productMap.get(norm(movement.productCode)); return { "Waktu": fmtDateTime(movement.at), "Referensi": movement.reference, "Kode Produk": movement.productCode, "Nama Produk": movement.productName, "Qty Masuk": num(movement.delta), "Stok Setelah": num(movement.stockAfter), "Estimasi Nilai": num(movement.delta) * num(product?.["Harga Beli"]) }; });
}

function renderReportInvoices(data) {
    $("report-invoices-count").textContent = `${data.invoices.length} faktur`;
    $("report-invoices-body").innerHTML = data.invoices.length
        ? data.invoices.map((invoice) => `<tr><td>${dateOnly(invoice.date || invoice.createdAt) || "—"}</td><td>${reportEscape(invoice.number || "—")}</td><td>${reportEscape(invoice.supplierName || invoice.supplierCode || "—")}</td><td>${(invoice.items || []).length}</td><td>${rupiah(invoiceTotal(invoice))}</td><td>${statusBadge(invoice.status)}</td></tr>`).join("")
        : '<tr><td colspan="6" class="empty-table-state">Tidak ada faktur pada periode ini.</td></tr>';
    reportExportRows.invoices = data.invoices.map((invoice) => ({ "Tanggal": dateOnly(invoice.date || invoice.createdAt), "Nomor Faktur": invoice.number, "Supplier": invoice.supplierName || invoice.supplierCode, "Jumlah Item": (invoice.items || []).length, "Nilai": invoiceTotal(invoice), "Status": invoice.status }));
}

function renderReportCashiers(data) {
    const map = new Map();
    data.sales.forEach((sale) => {
        const cashier = text(sale.cashier) || "Tidak diketahui";
        const row = map.get(norm(cashier)) || { cashier, completed: 0, voided: 0, units: 0, revenue: 0, profit: 0 };
        if (reportSaleStatus(sale) === "VOID") row.voided += 1;
        else {
            row.completed += 1;
            row.units += (sale.items || []).reduce((sum, item) => sum + num(item.qty), 0);
            row.revenue += num(sale.total);
            row.profit += reportSaleProfit(sale);
        }
        map.set(norm(cashier), row);
    });
    const rows = [...map.values()].sort((a, b) => b.revenue - a.revenue);
    $("report-cashiers-count").textContent = `${rows.length} kasir`;
    $("report-cashiers-body").innerHTML = rows.length
        ? rows.map((row) => `<tr><td>${reportEscape(row.cashier)}</td><td>${row.completed}</td><td>${row.voided}</td><td>${row.units}</td><td>${rupiah(row.revenue)}</td><td>${rupiah(row.completed ? row.revenue / row.completed : 0)}</td><td class="${row.profit >= 0 ? "qty-plus" : "qty-minus"}">${rupiah(row.profit)}</td></tr>`).join("")
        : '<tr><td colspan="7" class="empty-table-state">Belum ada aktivitas kasir pada periode ini.</td></tr>';
    reportExportRows.cashiers = rows.map((row) => ({ "Kasir": row.cashier, "Transaksi Selesai": row.completed, "Transaksi VOID": row.voided, "Produk Terjual": row.units, "Omzet": row.revenue, "Rata-rata Transaksi": row.completed ? row.revenue / row.completed : 0, "Estimasi Laba": row.profit }));
}

function renderReports() {
    if (!$("report-period")) return;
    const data = reportData();
    const productRows = reportProductRows(data.completedSales);
    reportExportRows = {};
    renderReportSummary(data, productRows);
    renderReportSales(data);
    renderReportProducts(productRows);
    renderReportStock(data);
    renderReportGoods(data);
    renderReportInvoices(data);
    renderReportCashiers(data);
    $("report-period-label").textContent = `Periode: ${data.range.label}${data.query ? ` · Pencarian: “${text($("report-search").value)}”` : ""}`;
    $("report-last-updated").textContent = `Diperbarui ${new Date().toLocaleString("id-ID")}`;
}

function setReportTab(tab) {
    activeReportTab = tab;
    document.querySelectorAll("[data-report-tab]").forEach((button) => button.classList.toggle("active", button.dataset.reportTab === tab));
    document.querySelectorAll("[data-report-panel]").forEach((panel) => { panel.hidden = panel.dataset.reportPanel !== tab; });
}

function updateReportPeriodInputs() {
    const custom = $("report-period")?.value === "custom";
    if ($("report-start-date")) $("report-start-date").disabled = !custom;
    if ($("report-end-date")) $("report-end-date").disabled = !custom;
    renderReports();
}

function renderAllOperational(){renderInvoices();renderStock();renderMovements();renderOpnameHistory();renderManagementSales();renderReports();}

function renderOperationalView(view) {
    if(view==="stock") { scheduleStockRender(); return; }
    if(view==="purchase-invoices" || view==="import-faktur") renderInvoices();
    if(view==="goods-in") renderMovements();
    if(view==="stock-opname") renderOpnameHistory();
    if(view==="sales") renderManagementSales();
    if(view==="reports") renderReports();
}

function bindOperational(){
    $("new-invoice")?.addEventListener("click",()=>showOpsNotice("Faktur pembelian dikelola melalui Import Excel.","info")); $("close-invoice-editor")?.addEventListener("click",closeInvoiceEditor); $("invoice-editor-overlay")?.addEventListener("click",e=>{if(e.target===$("invoice-editor-overlay"))closeInvoiceEditor()});
    $("invoice-item-product")?.addEventListener("change",()=>{const p=products()[Number($("invoice-item-product").value)];if(p){$("invoice-item-buy").value=num(p["Harga Beli"]);$("invoice-item-sell").value=num(p["Harga Jual"]);}});
    $("add-invoice-item")?.addEventListener("click",addEditorItem); $("save-invoice-draft")?.addEventListener("click",()=>saveInvoiceFromEditor(false)); $("confirm-invoice")?.addEventListener("click",()=>saveInvoiceFromEditor(true));
    $("invoice-discount")?.addEventListener("input",renderEditorItems);
    $("invoice-discount-method")?.addEventListener("change",renderEditorItems);
    $("invoice-tax")?.addEventListener("input",renderEditorItems);
    $("invoice-search")?.addEventListener("input",renderInvoices); $("invoice-status-filter")?.addEventListener("change",renderInvoices); $("go-import-invoice")?.addEventListener("click",()=>navigate("import-faktur"));
    $("read-invoice-import")?.addEventListener("click",previewInvoiceImport); $("apply-invoice-import")?.addEventListener("click",applyInvoiceImport);
    $("refresh-stock")?.addEventListener("click",()=>{stockCurrentPage=1;scheduleStockRender()}); $("stock-search")?.addEventListener("input",()=>{stockCurrentPage=1;scheduleStockRender()}); ["stock-category-filter","stock-supplier-filter","stock-status-filter"].forEach((id)=>$(id)?.addEventListener("change",()=>{stockCurrentPage=1;scheduleStockRender()})); $("stock-reset-filter")?.addEventListener("click",()=>{["stock-search","stock-category-filter","stock-supplier-filter","stock-status-filter"].forEach((id)=>{if($(id))$(id).value=""});stockCurrentPage=1;scheduleStockRender()}); $("stock-prev-page")?.addEventListener("click",()=>{if(stockCurrentPage>1){stockCurrentPage--;scheduleStockRender()}}); $("stock-next-page")?.addEventListener("click",()=>{stockCurrentPage++;scheduleStockRender()}); $("opname-export-button")?.addEventListener("click",exportStockOpname); $("go-import-opname")?.addEventListener("click",()=>navigate("import-opname")); $("export-database-backup")?.addEventListener("click",exportDatabaseBackup);
    $("read-opname-import")?.addEventListener("click",previewOpnameImport); $("confirm-opname-import")?.addEventListener("click",confirmOpnameImport); $("refresh-goods-in")?.addEventListener("click",renderMovements); $("movement-search")?.addEventListener("input",renderMovements); $("movement-type-filter")?.addEventListener("change",renderMovements);
    $("sales-search")?.addEventListener("input",()=>{salesCurrentPage=1;renderManagementSales()}); $("sales-period")?.addEventListener("change",()=>{salesCurrentPage=1;renderManagementSales()}); $("sales-status-filter")?.addEventListener("change",()=>{salesCurrentPage=1;renderManagementSales()}); $("sales-payment-filter")?.addEventListener("change",()=>{salesCurrentPage=1;renderManagementSales()}); $("sales-refresh")?.addEventListener("click",renderManagementSales); $("sales-prev-page")?.addEventListener("click",()=>{if(salesCurrentPage>1){salesCurrentPage--;renderManagementSales()}}); $("sales-next-page")?.addEventListener("click",()=>{if(salesCurrentPage*SALES_PAGE_SIZE<filteredManagementSales.length){salesCurrentPage++;renderManagementSales()}});
    $("close-sales-detail")?.addEventListener("click",closeManagementSaleDetail); $("sales-detail-close-button")?.addEventListener("click",closeManagementSaleDetail); $("sales-detail-print")?.addEventListener("click",printManagementSaleDetail); $("sales-detail-void")?.addEventListener("click",voidManagementSale); $("sales-detail-overlay")?.addEventListener("click",(event)=>{if(event.target===$("sales-detail-overlay"))closeManagementSaleDetail()});
    $("report-period")?.addEventListener("change",updateReportPeriodInputs); $("report-start-date")?.addEventListener("change",renderReports); $("report-end-date")?.addEventListener("change",renderReports); $("report-search")?.addEventListener("input",renderReports); $("report-refresh")?.addEventListener("click",renderReports); $("report-print")?.addEventListener("click",()=>window.print());
    document.querySelectorAll("[data-report-tab]").forEach((button)=>button.addEventListener("click",()=>setReportTab(button.dataset.reportTab)));
    document.querySelectorAll("[data-view]").forEach((b)=>b.addEventListener("click",()=>setTimeout(()=>renderOperationalView(b.dataset.view),0)));
    document.addEventListener("keydown",e=>{if(e.key==="Escape"&&!$("invoice-editor-overlay")?.hidden)closeInvoiceEditor();if(e.key==="Escape"&&!$("sales-detail-overlay")?.hidden)closeManagementSaleDetail()});
    const today = new Date(); if ($("report-start-date")) $("report-start-date").value = reportDateInput(new Date(today.getFullYear(), today.getMonth(), 1)); if ($("report-end-date")) $("report-end-date").value = reportDateInput(today); setReportTab(activeReportTab); renderOperationalView(document.querySelector(".view-section:not([hidden])")?.dataset.viewSection);
}

async function startOperational() {
    try {
        await databaseInitialization;
        bindOperational();
    } catch (error) {
        console.error("Modul operasional tidak dapat membuka database:", error);
    }
}

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", startOperational, { once: true });
} else {
    startOperational();
}
