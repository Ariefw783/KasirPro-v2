/* KasirPro Acceptance — Import Master UX FINAL + direct XLSX download */
function norm(v){return String(v??"").replace(/\s+/g," ").trim().toLowerCase()}
function replaceExact(from,to){[...document.querySelectorAll("p,small,span,strong,h2,h3,h4")].forEach(el=>{if(norm(el.textContent)===norm(from))el.textContent=to})}

function installStyle(){
  if(document.getElementById("kp-master-replace-style")) return;
  const s=document.createElement("style");
  s.id="kp-master-replace-style";
  s.textContent=`
    .kp-replace-wrap{padding:18px 0 4px}
    .kp-replace-title{font-size:1rem;font-weight:850;color:var(--text-main,#0f172a);margin:0 0 6px}
    .kp-replace-sub{font-size:.68rem;color:var(--text-muted,#64748b);line-height:1.55;margin:0 0 14px}
    .kp-replace-warning{padding:14px;border:1px solid #f59e0b;border-radius:12px;background:#fffbeb;color:#92400e;margin-bottom:14px}
    .kp-replace-warning strong{display:block;font-size:.72rem;margin-bottom:6px}
    .kp-replace-warning p{margin:0;font-size:.63rem;line-height:1.55}
    .kp-replace-safe{padding:12px 14px;border:1px solid #bfdbfe;border-radius:12px;background:#eff6ff;color:#1e40af;font-size:.62rem;line-height:1.55;margin-bottom:14px}
    .kp-replace-list{margin:6px 0 0 18px;padding:0}
    .kp-replace-list li{margin:3px 0}
    .kp-replace-consent{display:flex;gap:10px;align-items:flex-start;padding:13px;border:1px solid #cbd5e1;border-radius:12px;background:#fff;margin:0 0 16px;cursor:pointer}
    .kp-replace-consent input{margin-top:2px;width:18px;height:18px;accent-color:var(--primary,#2563eb)}
    .kp-replace-consent span{font-size:.63rem;line-height:1.5;color:var(--text-main,#334155)}
    #confirm-master-import:disabled{opacity:.5;cursor:not-allowed}
  `;
  document.head.appendChild(s);
}

function rebuildStep4(){
  const confirm=document.getElementById("confirm-master-import");
  if(!confirm||document.getElementById("kp-master-replace-consent"))return;
  let stage=confirm.parentElement;
  while(stage&&!stage.querySelector('input[name="existingDataMode"]'))stage=stage.parentElement;
  if(!stage)return;
  const allButtons=[...stage.querySelectorAll("button")];
  const back=allButtons.find(b=>b!==confirm&&/kembali/i.test(b.textContent));
  const wrap=document.createElement("div");
  wrap.className="kp-replace-wrap";
  wrap.innerHTML=`
    <h3 class="kp-replace-title">4. Simpan Master Baru</h3>
    <p class="kp-replace-sub">Langkah terakhir. Pastikan file Excel sudah berisi seluruh Master yang ingin digunakan.</p>
    <div class="kp-replace-warning"><strong>Perhatian: Master lama akan diganti seluruhnya</strong><p>Setelah disimpan, Supplier, Kategori, dan Produk akan mengikuti isi file Excel ini. Pengaturan Toko tetap dikelola di halaman V2; stok berjalan produk lama tidak diubah.</p></div>
    <div class="kp-replace-safe"><strong>Yang tidak ikut dihapus:</strong><ul class="kp-replace-list"><li>Transaksi penjualan</li><li>Faktur pembelian</li><li>Riwayat Mutasi Stok</li><li>Riwayat Stock Opname</li><li>Akun Admin dan Kasir</li></ul></div>
    <label class="kp-replace-consent"><input id="kp-master-replace-consent" type="checkbox"><span>Saya sudah memeriksa file Excel dan memahami bahwa Master lama akan diganti dengan isi file ini.</span></label>`;
  stage.innerHTML="";
  stage.appendChild(wrap);
  const actions=document.createElement("div");
  actions.style.cssText="display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap";
  if(back)actions.appendChild(back);
  confirm.textContent="Simpan Master Baru";
  confirm.disabled=true;
  actions.appendChild(confirm);
  stage.appendChild(actions);
  document.getElementById("kp-master-replace-consent")?.addEventListener("change",e=>{confirm.disabled=!e.target.checked});
}

function setWidths(ws,widths){ws["!cols"]=widths.map(w=>({wch:w}))}
function downloadMasterTemplate(){
  if(!window.XLSX?.utils||typeof window.XLSX.writeFile!=="function"){
    alert("Pembuatan Template Master belum siap. Muat ulang halaman saat koneksi internet tersedia lalu coba kembali.");
    return;
  }
  const wb=window.XLSX.utils.book_new();
  const add=(name,rows,widths)=>{const ws=window.XLSX.utils.aoa_to_sheet(rows);setWidths(ws,widths);window.XLSX.utils.book_append_sheet(wb,ws,name)};

  add("PETUNJUK_IMPORT",[
    ["MASTER KASIRPRO — PETUNJUK IMPORT"],
    ["Setiap Import Master mengganti seluruh Master lama dengan isi file ini."],
    ["Data transaksi, faktur pembelian, Mutasi Stok, Stock Opname, serta akun Admin/Kasir tidak ikut dihapus."],
    [],
    ["Sheet","Kolom","Status","Keterangan"],
    ["SUPPLIER","Supplier","Wajib","Gunakan label Supplier 1 sampai Supplier 6."],
    ["SUPPLIER","Nama Supplier","Wajib","Nama supplier yang ditampilkan aplikasi."],
    ["KATEGORI","Kode Kategori","Disarankan","Kode unik kategori."],
    ["KATEGORI","Nama Kategori","Wajib","Nama kategori produk."],
    ["PRODUK","Kode Produk","Opsional","Boleh kosong. KasirPro membuat ADD0001, ADD0002, dst."],
    ["PRODUK","Nama Produk","Wajib","Nama produk."],
    ["PRODUK","Supplier","Wajib","Harus memakai Supplier 1 sampai Supplier 6."],
    ["PRODUK","Kategori","Opsional","Gunakan kategori yang tersedia pada sheet KATEGORI."],
    ["PRODUK","Stok Aktual","Tidak ada di Master","Stok dikelola melalui Barang Masuk, transaksi, Mutasi Stok, dan Stock Opname."],
    ["PENGATURAN_TOKO","Nama Toko","Wajib","Identitas toko."],
    [],
    ["ATURAN","File ini adalah Master lengkap. Data Master yang tidak ada di file baru tidak akan dipakai lagi setelah import."]
  ],[22,24,18,70]);

  add("SUPPLIER",[
    ["Supplier","Nama Supplier","Alamat","Telepon","Email","Kontak Person","NPWP","Termin Default","Status","Catatan"],
    ["Supplier 1","PT. Rosa Nugraha Abadi","","","","","",30,"Aktif",""],
    ["Supplier 2","PT. Mondy Inti Persada","","","","","",30,"Aktif",""],
    ["Supplier 3","Liwa Herbal","","","","","",30,"Aktif",""],
    ["Supplier 4","PT. Bima Sakti Medica Palembang","","","","","",30,"Aktif",""],
    ["Supplier 5","Hayati Herbal","","","","","",30,"Aktif",""],
    ["Supplier 6","Umum 2","","","","","",30,"Aktif",""]
  ],[14,32,32,16,24,18,18,16,12,24]);

  add("KATEGORI",[
    ["Kode Kategori","Nama Kategori","Deskripsi","Status"],
    ["KAT001","Obat Bebas","Contoh kategori; boleh diubah.","Aktif"]
  ],[16,24,36,14]);

  add("PRODUK",[
    ["Kode Produk","Nama Produk","Kategori","Satuan","Supplier","Harga Beli","Harga Jual","Stok Minimum","Lokasi Rak","Batch","Tanggal Expired","Status Produk","Catatan"],
    ["","Paracetamol 500 mg","Obat Bebas","Strip","Supplier 1",8500,11000,5,"A-03","B26001","30/06/2028","Aktif","Contoh; boleh dihapus"]
  ],[16,32,20,12,14,14,14,14,14,14,16,14,28]);

  add("PENGATURAN_TOKO",[
    ["Nama Toko","Alamat","Telepon","Email","NPWP","Logo","Footer Struk","Prefix Transaksi","Prefix Faktur","Format Nomor Transaksi","Format Nomor Faktur","Mata Uang","Zona Waktu","Pajak Default (%)","Ukuran Struk","Printer Default","Lebar Kertas","Tampilkan Logo di Struk","Tampilkan Nama Kasir di Struk","Tampilkan Pajak di Struk","Tampilkan Diskon di Struk","Status Toko","Catatan"],
    ["Apotek Doa Ibu","Kota Gajah","","apotekdoaibu.v2@gmail.com","","","Terima kasih atas kunjungan Anda","TRX","INV","{YYYY}{MM}{DD}-{SEQ}","{YYYY}{MM}-{SEQ}","IDR","Asia/Jakarta",0,"80mm","",80,"Ya","Ya","Ya","Ya","Aktif","Konfigurasi utama"]
  ],Array(23).fill(18));

  window.XLSX.writeFile(wb,"Template_Master_KasirPro.xlsx",{compression:true});
}

function patchMasterDownload(){
  const candidates=[...document.querySelectorAll("a,button")].filter(el=>norm(el.textContent)==="download template");
  if(!candidates.length)return;
  const masterButton=candidates[0];
  masterButton.dataset.kpMasterTemplate="true";
  if(masterButton.tagName==="A")masterButton.setAttribute("href","#download-master-template");
}

document.addEventListener("click",event=>{
  const target=event.target.closest('[data-kp-master-template="true"]');
  if(!target)return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation();
  downloadMasterTemplate();
},true);

function refreshCopy(){
  replaceExact("Semua kolom bersifat opsional. Data parsial tetap dapat diproses dengan peringatan.","File Excel ini adalah Master lengkap. Setiap import akan mengganti seluruh Master lama.");
  replaceExact("Import Simulasi Selesai","Master Baru Tersimpan");
  replaceExact("Data telah diproses.","Master lama telah diganti dengan isi file terbaru.");
}

function install(){installStyle();refreshCopy();patchMasterDownload();rebuildStep4()}
install();
let n=0;const timer=setInterval(()=>{install();if(++n>24)clearInterval(timer)},250);
const obs=new MutationObserver(install);obs.observe(document.body,{childList:true,subtree:true});
setTimeout(()=>obs.disconnect(),10000);
