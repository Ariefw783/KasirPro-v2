/* KasirPro V4 — reset data pusat is intentionally Console-only. */
function installResetSafety(){
  const candidates=[...document.querySelectorAll('button,a')].filter(node=>/preview\s*reset|reset\s*data/i.test(node.textContent||''));
  for(const control of candidates){
    if(control.dataset.kpResetSafe)return;
    control.dataset.kpResetSafe='1';
    control.disabled=true;
    control.setAttribute('aria-disabled','true');
    control.title='Reset data pusat dilakukan melalui Firebase Console.';
    const host=control.closest('section,.ops-card,.import-panel,.panel')||control.parentElement;
    if(host&&!host.querySelector('.kp-reset-console-note')){
      const note=document.createElement('div');
      note.className='kp-reset-console-note';
      note.innerHTML='<i class="fa-solid fa-shield-halved"></i><div><strong>Reset data pusat dinonaktifkan di V2</strong><p>Untuk mencegah penghapusan tidak disengaja dan kegagalan izin, reset dilakukan melalui Firebase Console oleh pemilik proyek. Tidak ada preview atau penghapusan Firestore yang dijalankan dari halaman ini.</p></div>';
      control.parentElement?.insertAdjacentElement('afterend',note);
    }
  }
}
const style=document.createElement('style');style.textContent='.kp-reset-console-note{display:flex;gap:11px;margin-top:12px;padding:13px;border:1px solid #fed7aa;border-radius:11px;background:#fff7ed;color:#9a3412}.kp-reset-console-note i{margin-top:2px}.kp-reset-console-note strong,.kp-reset-console-note p{display:block;margin:0}.kp-reset-console-note p{margin-top:3px;color:#7c2d12;font-size:12px;line-height:1.5}';document.head.appendChild(style);
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',installResetSafety,{once:true});else installResetSafety();
new MutationObserver(installResetSafety).observe(document.documentElement,{childList:true,subtree:true});
