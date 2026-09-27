/* KasirPro AT-14B/C/D — IndexedDB & Firestore Loading UI
 * UI-only loading lifecycle.
 * - IndexedDB/local loading has a neutral local-device visual.
 * - Firestore/cloud synchronization has a distinct cloud visual.
 * - Guards prevent flicker, overlap, and stuck loading state.
 */

const MIN_VISIBLE_MS = 420;
const FAILSAFE_MS = 30000;

let overlay = null;
let titleNode = null;
let messageNode = null;
let detailNode = null;
let visibleSince = 0;
let activeMode = "";
let hideTimer = null;
let failsafeTimer = null;
let cloudCollectionsDone = 0;
let cloudCollectionsTotal = 0;

function install() {
    if (overlay) return;

    const style = document.createElement("style");
    style.id = "kasirpro-at14-loading-style";
    style.textContent = `
      .kp-data-loading{position:fixed;inset:0;z-index:2147483645;display:grid;place-items:center;padding:20px}
      .kp-data-loading[hidden]{display:none!important}
      .kp-data-loading__backdrop{position:absolute;inset:0;background:rgba(15,23,42,.36);backdrop-filter:blur(2px)}
      .kp-data-loading__card{position:relative;width:min(92vw,430px);border-radius:18px;padding:22px;background:#fff;border:1px solid rgba(148,163,184,.28);box-shadow:0 28px 80px rgba(15,23,42,.24);display:grid;grid-template-columns:58px minmax(0,1fr);gap:16px;align-items:center}
      .kp-data-loading__icon{width:56px;height:56px;border-radius:16px;display:grid;place-items:center;position:relative}
      .kp-data-loading__icon::after{content:"";position:absolute;inset:-5px;border:2px solid transparent;border-top-color:currentColor;border-radius:999px;animation:kp-at14-spin .9s linear infinite;opacity:.72}
      .kp-data-loading__copy h2{margin:0 0 6px;font-size:17px;line-height:1.3;color:#0f172a}
      .kp-data-loading__copy p{margin:0;color:#475569;font-size:13px;line-height:1.5}
      .kp-data-loading__detail{margin-top:7px!important;font-size:11px!important;font-weight:700;letter-spacing:.01em}
      .kp-data-loading[data-mode="local"] .kp-data-loading__card{border-color:#cbd5e1}
      .kp-data-loading[data-mode="local"] .kp-data-loading__icon{background:#eef2f7;color:#475569}
      .kp-data-loading[data-mode="local"] .kp-data-loading__detail{color:#64748b}
      .kp-data-loading[data-mode="cloud"] .kp-data-loading__card{border-color:#bfdbfe;background:linear-gradient(180deg,#fff,#f8fbff)}
      .kp-data-loading[data-mode="cloud"] .kp-data-loading__icon{background:#dbeafe;color:#2563eb}
      .kp-data-loading[data-mode="cloud"] .kp-data-loading__detail{color:#2563eb}
      .kp-data-loading__glyph{font-size:25px;line-height:1}
      @keyframes kp-at14-spin{to{transform:rotate(360deg)}}
      @media(max-width:520px){.kp-data-loading__card{grid-template-columns:48px minmax(0,1fr);padding:18px}.kp-data-loading__icon{width:46px;height:46px;border-radius:13px}.kp-data-loading__glyph{font-size:21px}}
      @media(prefers-reduced-motion:reduce){.kp-data-loading__icon::after{animation-duration:1.8s}}
    `;
    document.head.appendChild(style);

    overlay = document.createElement("div");
    overlay.className = "kp-data-loading";
    overlay.hidden = true;
    overlay.setAttribute("role", "status");
    overlay.setAttribute("aria-live", "polite");
    overlay.setAttribute("aria-busy", "true");
    overlay.innerHTML = `
      <div class="kp-data-loading__backdrop"></div>
      <div class="kp-data-loading__card">
        <div class="kp-data-loading__icon"><span class="kp-data-loading__glyph" data-kp-loading-glyph>▣</span></div>
        <div class="kp-data-loading__copy">
          <h2 data-kp-loading-title>Memuat data…</h2>
          <p data-kp-loading-message>Mohon tunggu.</p>
          <p class="kp-data-loading__detail" data-kp-loading-detail></p>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    titleNode = overlay.querySelector("[data-kp-loading-title]");
    messageNode = overlay.querySelector("[data-kp-loading-message]");
    detailNode = overlay.querySelector("[data-kp-loading-detail]");
}

function clearTimers() {
    clearTimeout(hideTimer);
    clearTimeout(failsafeTimer);
    hideTimer = null;
    failsafeTimer = null;
}

function show(mode, { title, message, detail = "" }) {
    install();
    clearTimers();

    activeMode = mode;
    visibleSince = Date.now();
    overlay.dataset.mode = mode;
    overlay.querySelector("[data-kp-loading-glyph]").textContent = mode === "cloud" ? "☁" : "▣";
    titleNode.textContent = title;
    messageNode.textContent = message;
    detailNode.textContent = detail;
    overlay.hidden = false;

    failsafeTimer = setTimeout(() => hide(mode, true), FAILSAFE_MS);
}

function update(mode, { title, message, detail } = {}) {
    if (!overlay || overlay.hidden || activeMode !== mode) return;
    if (title) titleNode.textContent = title;
    if (message) messageNode.textContent = message;
    if (detail !== undefined) detailNode.textContent = detail;
}

function hide(mode, immediate = false) {
    if (!overlay || overlay.hidden || (mode && activeMode !== mode)) return;

    const close = () => {
        if (!overlay || (mode && activeMode !== mode)) return;
        clearTimers();
        overlay.hidden = true;
        overlay.removeAttribute("data-mode");
        activeMode = "";
    };

    const remaining = immediate ? 0 : Math.max(0, MIN_VISIBLE_MS - (Date.now() - visibleSince));
    clearTimeout(hideTimer);
    hideTimer = setTimeout(close, remaining);
}

function showLocal(detail = "") {
    show("local", {
        title: "Membuka data lokal…",
        message: "Mengambil data KasirPro dari penyimpanan perangkat.",
        detail: detail || "IndexedDB • Penyimpanan lokal"
    });
}

function showCloud(detail = "") {
    cloudCollectionsDone = 0;
    show("cloud", {
        title: "Sinkronisasi Firebase…",
        message: "Memeriksa perubahan terbaru dari database pusat.",
        detail: detail || "Firestore • Database cloud"
    });
}

window.addEventListener("kasirpro:database-preload-start", event => {
    const message = String(event.detail?.message || "").trim();
    showLocal(message ? `IndexedDB • ${message}` : "");
});

window.addEventListener("kasirpro:database-preload-progress", event => {
    const loaded = Number(event.detail?.loaded || 0);
    const total = Number(event.detail?.total || 0);
    const message = String(event.detail?.message || "").trim();
    const progress = total > 0 ? `${loaded}/${total}` : "";
    update("local", {
        message: message || "Membuka Master KasirPro dari penyimpanan perangkat.",
        detail: ["IndexedDB", progress].filter(Boolean).join(" • ")
    });
});

window.addEventListener("kasirpro:database-idle", () => hide("local"));

window.addEventListener("kasirpro:operational-cache-ready", event => {
    const counts = event.detail?.counts || {};
    cloudCollectionsTotal = Object.keys(counts).length;
    showCloud(cloudCollectionsTotal ? `Firestore • 0/${cloudCollectionsTotal} koleksi` : "");
});

window.addEventListener("kasirpro:operational-sync", event => {
    cloudCollectionsDone += 1;
    const mode = String(event.detail?.phase || "").trim();
    const storeKey = String(event.detail?.cacheStore || event.detail?.storeKey || "").trim();
    const detail = [
        "Firestore",
        cloudCollectionsTotal ? `${Math.min(cloudCollectionsDone, cloudCollectionsTotal)}/${cloudCollectionsTotal} koleksi` : "",
        storeKey || mode
    ].filter(Boolean).join(" • ");

    update("cloud", {
        message: mode === "full-seed"
            ? "Mengambil data awal dari Firebase."
            : "Menerapkan perubahan terbaru dari Firebase.",
        detail
    });
});

window.addEventListener("kasirpro:operational-ready", () => {
    update("cloud", {
        message: "Sinkronisasi Firebase selesai.",
        detail: "Firestore • Data terbaru siap digunakan"
    });
    hide("cloud");
});

window.addEventListener("kasirpro:operational-unavailable", () => {
    update("cloud", {
        message: "Firebase tidak tersedia. Data lokal yang tersimpan tetap digunakan.",
        detail: "Firestore • Menggunakan cache lokal"
    });
    hide("cloud");
});

window.addEventListener("kasirpro:database-error", () => hide(activeMode, true));

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
} else {
    install();
}

