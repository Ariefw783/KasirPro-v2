/* KasirPro V2 — Form faktur manual dengan diskon global atau per item. */
import { STORE_KEYS, readStore, writeStore } from '../modules/database/database-store.js';

const $ = id => document.getElementById(id);
const text = value => String(value ?? '').trim();
const norm = value => text(value).toLowerCase();
const num = value => Number(String(value ?? 0).replace(/[^0-9.-]/g, '')) || 0;
const rupiah = value => new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(num(value));
const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
const today = () => { const d = new Date(); return [d.getFullYear(), String(d.getMonth()+1).padStart(2,'0'), String(d.getDate()).padStart(2,'0')].join('-'); };
let items = [], editingId = null, editingItem = -1;

function master(){ return readStore(STORE_KEYS.master,{produk:[],supplier:[]}); }
function products(){ const rows=master().produk; return Array.isArray(rows)?rows:[]; }
function suppliers(){ const rows=master().supplier; return Array.isArray(rows)?rows:[]; }
function invoices(){ return readStore(STORE_KEYS.invoices,[]) || []; }
function selected(id){ return $(id)?.value ?? ''; }
function amount(method,value,base){ return method==='percentage' ? base*num(value)/100 : num(value); }
function lineBase(item){ return num(item.qty)*num(item.buyPrice); }
function lineNet(item){ return Math.max(0,lineBase(item)-num(item.discount)); }
function scope(){ return selected('manual-discount-scope'); }
function taxScope(){ return selected('manual-tax-method'); }
function total(){
  const subtotal=items.reduce((sum,item)=>sum+lineNet(item),0);
  const discount=scope()==='global'?amount(selected('manual-global-discount-method'),selected('manual-global-discount'),subtotal):0;
  const base=Math.max(0,subtotal-discount);
  const tax=taxScope()==='Per Item'
    ?items.reduce((sum,item)=>sum+lineNet(item)*num(item.taxPercent)/100,0)
    :base*num(selected('manual-tax-global'))/100;
  return {subtotal,discount,tax,total:base+tax+num(selected('manual-other-cost'))};
}

function installStyle(){
  if($('kp-manual-invoice-style'))return;
  const style=document.createElement('style');
  style.id='kp-manual-invoice-style';
  style.textContent=`
    .kp-manual{padding:18px;display:grid;gap:18px}
    .kp-manual [hidden]{display:none!important}
    .kp-manual h3{font-size:14px;font-weight:800;color:#0f172a;margin:0 0 10px}
    .kp-manual-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
    .kp-manual-item-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
    .kp-manual-extra{border:1px solid #e2e8f0;border-radius:9px;padding:10px 12px}
    .kp-manual-extra summary{cursor:pointer;font-size:12px;font-weight:700;color:#475569}
    .kp-manual-extra .kp-manual-grid{margin-top:12px}
    #kp-invoice-final>.kp-final-card[hidden]{display:none!important}
    .kp-manual label{display:grid;gap:5px;font-size:11px;font-weight:700;color:#475569;min-width:0}
    .kp-manual input,.kp-manual select{width:100%;height:38px;padding:0 10px;border:1px solid #cbd5e1;border-radius:9px;background:#fff;color:#0f172a;font-size:12px}
    .kp-manual input:disabled,.kp-manual select:disabled{background:#f1f5f9;color:#94a3b8}
    .kp-manual .wide{grid-column:span 2}
    .kp-manual-section{border-top:1px solid #e2e8f0;padding-top:16px}
    .kp-manual-table{width:100%;border-collapse:collapse;font-size:12px}
    .kp-manual-table th,.kp-manual-table td{padding:10px;border-bottom:1px solid #e2e8f0;text-align:left}
    .kp-manual-table th{background:#f8fafc;color:#475569}
    .kp-manual-table button{border:0;background:none;color:#0f766e;font-weight:700;cursor:pointer;margin-right:8px}
    .kp-manual-summary{display:flex;justify-content:space-between;align-items:center;gap:14px;background:#f8fafc;padding:16px;border-radius:11px}
    .kp-manual-summary small{display:block;color:#64748b;font-size:11px;margin-bottom:4px}
    .kp-manual-summary strong{font-size:20px;color:#0f766e}
    .kp-manual-actions{display:flex;gap:8px;justify-content:flex-end}
    .kp-manual-note{font-size:11px;color:#64748b;margin:0}
    @media(max-width:900px){.kp-manual-grid,.kp-manual-item-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
    @media(max-width:600px){.kp-manual-grid,.kp-manual-item-grid{grid-template-columns:1fr}.kp-manual .wide{grid-column:auto}.kp-manual-summary{align-items:flex-start;flex-direction:column}.kp-manual-actions{width:100%}.kp-manual-actions button{flex:1}}
  `;
  document.head.appendChild(style);
}

function build(){
  const root=$('kp-invoice-final');
  if(!root||$('kp-manual-invoice'))return;
  installStyle();
  const card=document.createElement('section');
  card.id='kp-manual-invoice';
  card.className='kp-final-card';
  card.innerHTML=`
    <div class="kp-final-card-head"><div><h3>Input Faktur Pembelian</h3><p>Isi header, tambahkan produk ke antrean, periksa total, lalu simpan sebagai Draft.</p></div></div>
    <div class="kp-manual">
      <div><h3>1. Header Faktur</h3><div class="kp-manual-grid">
        <label>Nomor Faktur *<input id="manual-number" autocomplete="off" placeholder="Nomor dari supplier"></label>
        <label>Tanggal Faktur *<input id="manual-date" type="date"></label>
        <label>Supplier *<select id="manual-supplier"><option value="">Pilih supplier</option></select></label>
        <label>Jenis Pembayaran *<select id="manual-payment"><option value="Cash">Cash</option><option value="Tempo">Kredit / Tempo</option></select></label>
        <label id="manual-due-wrap" hidden>Tanggal Jatuh Tempo *<input id="manual-due" type="date"></label>
        <label>Metode PPN<select id="manual-tax-method"><option value="Global">PPN Global</option><option value="Per Item">PPN Per Item</option></select></label>
        <label id="manual-tax-global-wrap">PPN Global (%)<input id="manual-tax-global" type="number" min="0" max="100" step="0.01" value="0"></label>
        <label>Metode Diskon *<select id="manual-discount-scope"><option value="global">Diskon Global</option><option value="per_item">Diskon Per Item</option></select></label>
        <label id="manual-global-discount-wrap">Diskon Global<div style="display:flex;gap:6px"><select id="manual-global-discount-method" style="width:110px"><option value="nominal">Rupiah</option><option value="percentage">Persen</option></select><input id="manual-global-discount" type="number" min="0" step="0.01" value="0"></div></label>
      </div><details class="kp-manual-extra"><summary>Biaya lain dan catatan (opsional)</summary><div class="kp-manual-grid">
        <label>Biaya Lain (Rp)<input id="manual-other-cost" type="number" min="0" step="1" value="0"></label>
        <label class="wide">Catatan<input id="manual-notes"></label>
      </div></details></div>
      <div class="kp-manual-section"><h3>2. Input Produk</h3><div class="kp-manual-item-grid">
        <label class="wide">Produk *<input id="manual-product" list="manual-product-options" autocomplete="off" placeholder="Ketik nama atau kode, lalu pilih produk"><datalist id="manual-product-options"></datalist></label>
        <label>Qty *<input id="manual-qty" type="number" min="0.01" step="0.01" value="1"></label>
        <label>Harga Beli (Rp) *<input id="manual-buy" type="number" min="0" step="1"></label>
        <label>Harga Jual (Rp)<input id="manual-sell" type="number" min="0" step="1"></label>
        <label id="manual-item-discount-wrap" hidden>Diskon Item<div style="display:flex;gap:6px"><select id="manual-item-discount-method" style="width:110px"><option value="nominal">Rupiah</option><option value="percentage">Persen</option></select><input id="manual-item-discount" type="number" min="0" step="0.01" value="0"></div></label>
        <label id="manual-item-tax-wrap" hidden>PPN Item (%)<input id="manual-item-tax" type="number" min="0" max="100" step="0.01" value="0"></label>
      </div><div style="margin-top:12px"><button type="button" id="manual-add" class="kp-final-btn primary">Tambahkan Item</button></div></div>
      <div class="kp-manual-section"><h3>3. Antrean Item Faktur</h3><div class="kp-final-tablewrap"><table class="kp-manual-table"><thead><tr><th>Produk</th><th>Qty</th><th>Harga Beli</th><th>Diskon</th><th>PPN</th><th>Subtotal</th><th>Aksi</th></tr></thead><tbody id="manual-items"></tbody></table></div></div>
      <div class="kp-manual-summary"><div><small>Subtotal <span id="manual-subtotal">Rp0</span> · Diskon <span id="manual-discount-total">Rp0</span> · PPN <span id="manual-tax-total">Rp0</span></small><strong id="manual-total">Rp0</strong></div><div class="kp-manual-actions"><button type="button" id="manual-reset" class="kp-final-btn">Bersihkan</button><button type="button" id="manual-save" class="kp-final-btn primary">Simpan Draft</button></div></div>
      <p class="kp-manual-note">Stok baru bertambah setelah kamu memilih Barang Masuk pada faktur Draft.</p>
    </div>`;
  const excelCard=root.querySelector('.kp-final-card');
  root.insertBefore(card,excelCard);
  if(excelCard){
    const toggle=document.createElement('button');
    toggle.id='manual-toggle-excel';toggle.type='button';toggle.className='kp-final-btn';
    toggle.textContent='Tampilkan impor Excel (opsional)';
    root.insertBefore(toggle,excelCard);excelCard.hidden=true;
    toggle.addEventListener('click',()=>{
      excelCard.hidden=!excelCard.hidden;
      toggle.textContent=excelCard.hidden?'Tampilkan impor Excel (opsional)':'Sembunyikan impor Excel';
    });
  }
  reset();
  bind(card);
}

function refreshSuppliers(){
  const sel=$('manual-supplier'),current=sel.value;
  sel.innerHTML='<option value="">Pilih supplier</option>'+Array.from({length:6},(_,i)=>{
    const code=`Supplier ${i+1}`;
    const row=suppliers().find(s=>norm(s.Supplier)===norm(code));
    return `<option value="${code}">${esc(text(row?.['Nama Supplier'])||code)}</option>`;
  }).join('');
  sel.value=current;
}
function availableProducts(){
  const supplier=norm(selected('manual-supplier'));
  return products().filter(p=>!supplier||norm(p.Supplier||p['Kode Supplier'])===supplier);
}
function productLabel(p){return `${text(p['Kode Produk'])} — ${text(p['Nama Produk'])}`;}
function productForInput(){
  const value=norm(selected('manual-product'));
  return availableProducts().find(p=>norm(productLabel(p))===value||norm(p['Kode Produk'])===value);
}
function refreshProducts(){
  $('manual-product-options').innerHTML=availableProducts().map(p=>`<option value="${esc(productLabel(p))}"></option>`).join('');
}
function fillProductPrices(){
  const p=productForInput();
  if(p){$('manual-buy').value=num(p['Harga Beli']);$('manual-sell').value=num(p['Harga Jual']);}
}
function refreshFields(){
  $('manual-due-wrap').hidden=selected('manual-payment')!=='Tempo';
  $('manual-tax-global-wrap').hidden=taxScope()!=='Global';
  $('manual-item-tax-wrap').hidden=taxScope()!=='Per Item';
  $('manual-global-discount-wrap').hidden=scope()!=='global';
  $('manual-item-discount-wrap').hidden=scope()!=='per_item';
  render();
}
function render(){
  const body=$('manual-items');if(!body)return;
  body.innerHTML=items.length?items.map((item,index)=>`<tr><td>${esc(item.code)} — ${esc(item.name)}</td><td>${item.qty}</td><td>${rupiah(item.buyPrice)}</td><td>${rupiah(item.discount)}</td><td>${num(item.taxPercent)}%</td><td>${rupiah(lineNet(item))}</td><td><button type="button" data-manual-edit="${index}">Edit</button><button type="button" data-manual-remove="${index}">Hapus</button></td></tr>`).join(''):'<tr><td colspan="7">Belum ada item.</td></tr>';
  const sum=total();
  $('manual-subtotal').textContent=rupiah(sum.subtotal);
  $('manual-discount-total').textContent=rupiah(sum.discount);
  $('manual-tax-total').textContent=rupiah(sum.tax);
  $('manual-total').textContent=rupiah(sum.total);
}
function validRate(value,label){if(value<0||value>100)throw new Error(`${label} harus antara 0 dan 100.`);}
function addItem(){
  const p=productForInput();
  if(!p)throw new Error('Pilih produk dari daftar.');
  const qty=num(selected('manual-qty')),buyPrice=num(selected('manual-buy')),sellPrice=num(selected('manual-sell'));
  if(qty<=0)throw new Error('Qty harus lebih dari 0.');
  if(buyPrice<0||sellPrice<0)throw new Error('Harga tidak boleh negatif.');
  const method=selected('manual-item-discount-method'),value=scope()==='per_item'?num(selected('manual-item-discount')):0;
  if(value<0)throw new Error('Diskon tidak boleh negatif.');
  if(method==='percentage')validRate(value,'Diskon item Persen');
  const gross=qty*buyPrice,discount=amount(method,value,gross);
  if(discount>gross)throw new Error('Diskon item melebihi nilai baris.');
  const taxPercent=taxScope()==='Per Item'?num(selected('manual-item-tax')):0;
  validRate(taxPercent,'PPN item');
  const item={code:text(p['Kode Produk']),name:text(p['Nama Produk']),unit:text(p.Satuan),qty,buyPrice,sellPrice,
    discount,discountMethod:method,discountValue:value,taxPercent,supplier:selected('manual-supplier'),productStatus:'existing',reviewRequired:false};
  if(editingItem>=0)items[editingItem]=item;else items.push(item);
  editingItem=-1;$('manual-add').textContent='Tambahkan Item';
  $('manual-qty').value=1;$('manual-item-discount').value=0;$('manual-item-tax').value=0;
  render();
}
function reset(){
  items=[];editingId=null;editingItem=-1;
  $('manual-number').value='';$('manual-date').value=today();$('manual-payment').value='Cash';
  $('manual-due').value='';$('manual-tax-method').value='Global';$('manual-tax-global').value=0;
  $('manual-discount-scope').value='global';$('manual-global-discount-method').value='nominal';$('manual-global-discount').value=0;
  $('manual-other-cost').value=0;$('manual-notes').value='';$('manual-product').value='';
  $('manual-buy').value='';$('manual-sell').value='';$('manual-qty').value=1;
  $('manual-item-discount').value=0;$('manual-item-tax').value=0;$('manual-add').textContent='Tambahkan Item';
  refreshSuppliers();$('manual-supplier').value='';refreshProducts();refreshFields();
}
async function save(){
  const number=text(selected('manual-number')),date=selected('manual-date'),supplier=selected('manual-supplier');
  if(!number||!date||!supplier)throw new Error('Nomor, tanggal, dan supplier wajib diisi.');
  if(selected('manual-payment')==='Tempo'&&!selected('manual-due'))throw new Error('Tanggal jatuh tempo wajib untuk Kredit.');
  if(!items.length)throw new Error('Tambahkan minimal satu item.');
  const globalValue=scope()==='global'?num(selected('manual-global-discount')):0;
  const globalMethod=selected('manual-global-discount-method');
  if(globalValue<0)throw new Error('Diskon tidak boleh negatif.');
  if(globalMethod==='percentage')validRate(globalValue,'Diskon global Persen');
  const taxPercent=taxScope()==='Global'?num(selected('manual-tax-global')):0;
  validRate(taxPercent,'PPN global');
  const otherCost=num(selected('manual-other-cost'));
  if(otherCost<0)throw new Error('Biaya lain tidak boleh negatif.');
  const subtotal=items.reduce((sum,item)=>sum+lineNet(item),0);
  const discount=amount(globalMethod,globalValue,subtotal);
  if(discount>subtotal)throw new Error('Diskon global melebihi subtotal.');
  const list=invoices(),idx=editingId?list.findIndex(inv=>inv.id===editingId):-1;
  const duplicate=list.find(inv=>norm(inv.number)===norm(number)&&inv.id!==editingId);
  if(duplicate)throw new Error('Nomor faktur sudah dipakai.');
  if(idx>=0&&list[idx].stockApplied)throw new Error('Faktur yang sudah Barang Masuk tidak dapat diubah.');
  const supplierRow=suppliers().find(row=>norm(row.Supplier)===norm(supplier));
  const at=new Date().toISOString();
  const invoice={
    id:idx>=0?list[idx].id:`INV-${Date.now()}-${Math.random().toString(36).slice(2,8)}`,
    number,date,supplierCode:supplier,supplierName:text(supplierRow?.['Nama Supplier'])||supplier,
    paymentType:selected('manual-payment'),dueDate:selected('manual-payment')==='Tempo'?selected('manual-due'):'',
    taxMethod:taxScope(),taxPercent,discountScope:scope(),discount,
    discountMethod:scope()==='global'?globalMethod:'nominal',
    discountValue:globalValue,otherCost,notes:text(selected('manual-notes')),
    status:'draft',stockApplied:false,source:'manual-v2',
    items:items.map(item=>({...item})),createdAt:idx>=0?list[idx].createdAt:at,updatedAt:at
  };
  if(idx>=0)list[idx]=invoice;else list.unshift(invoice);
  await writeStore(STORE_KEYS.invoices,list);
  reset();window.KasirProInvoiceFinalAT08?.render?.();
  await window.KasirProDialog?.success?.('Faktur Draft Tersimpan',`Faktur ${number} tersimpan. Stok belum berubah.`,{confirmText:'Selesai'});
}
function edit(id){
  const inv=invoices().find(row=>String(row.id)===String(id));
  if(!inv||inv.source!=='manual-v2'||inv.stockApplied)return false;
  reset();editingId=inv.id;items=(inv.items||[]).map(item=>({...item}));
  $('manual-number').value=inv.number;$('manual-date').value=inv.date;
  $('manual-supplier').value=inv.supplierCode;$('manual-payment').value=inv.paymentType;
  $('manual-due').value=inv.dueDate||'';$('manual-tax-method').value=inv.taxMethod||'Global';
  $('manual-tax-global').value=num(inv.taxPercent);
  $('manual-discount-scope').value=inv.discountScope||((inv.items||[]).some(item=>num(item.discount)>0)?'per_item':'global');
  $('manual-global-discount-method').value=inv.discountMethod||'nominal';
  $('manual-global-discount').value=num(inv.discountValue??inv.discount);
  $('manual-other-cost').value=num(inv.otherCost);$('manual-notes').value=inv.notes||'';
  refreshProducts();refreshFields();$('kp-manual-invoice').scrollIntoView({behavior:'smooth'});
  return true;
}
function bind(card){
  card.addEventListener('input',event=>{
    if(event.target.id==='manual-product')fillProductPrices();
    if(['manual-global-discount','manual-tax-global','manual-other-cost'].includes(event.target.id))render();
  });
  card.addEventListener('change',event=>{
    if(event.target.id==='manual-supplier'){items=[];$('manual-product').value='';refreshProducts();render();}
    if(event.target.id==='manual-product')fillProductPrices();
    if(event.target.id==='manual-discount-scope'){
      if(scope()==='global'){
        items=items.map(item=>({...item,discount:0,discountValue:0}));
      }else{
        $('manual-global-discount').value=0;
      }
    }
    if(['manual-payment','manual-tax-method','manual-discount-scope'].includes(event.target.id))refreshFields();
    if(event.target.id==='manual-global-discount-method')render();
  });
  card.addEventListener('click',async event=>{
    const target=event.target;
    try{
      if(target.id==='manual-add')addItem();
      if(target.id==='manual-reset')reset();
      if(target.id==='manual-save')await save();
      const remove=target.closest('[data-manual-remove]');
      if(remove){items.splice(Number(remove.dataset.manualRemove),1);render();}
      const editButton=target.closest('[data-manual-edit]');
      if(editButton){
        editingItem=Number(editButton.dataset.manualEdit);
        const item=items[editingItem];
        $('manual-product').value=`${item.code} — ${item.name}`;$('manual-qty').value=item.qty;
        $('manual-buy').value=item.buyPrice;$('manual-sell').value=item.sellPrice;
        $('manual-item-discount-method').value=item.discountMethod||'nominal';
        $('manual-item-discount').value=item.discountValue??item.discount;
        $('manual-item-tax').value=item.taxPercent||0;$('manual-add').textContent='Simpan Perubahan Item';
      }
    }catch(error){window.KasirProDialog?.warning?.('Periksa Faktur',error.message,{confirmText:'Perbaiki'})||alert(error.message);}
  });
}
function install(){build();}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',install,{once:true});else install();
window.addEventListener('kasirpro:operational-ready',install);
window.KasirProInvoiceManualV2=Object.freeze({edit,install});
