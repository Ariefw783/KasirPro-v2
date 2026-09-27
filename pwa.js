(function () {
    "use strict";

    const scriptUrl = new URL(document.currentScript.src);
    const appRoot = new URL("./", scriptUrl);
    let installPrompt = null;
    let registration = null;
    let refreshing = false;

    const style = document.createElement("style");
    style.textContent = `
        .kp-pwa-control{position:fixed;right:max(16px,env(safe-area-inset-right));bottom:max(16px,env(safe-area-inset-bottom));z-index:100000;display:flex;max-width:min(360px,calc(100vw - 32px));align-items:center;gap:10px;padding:11px 14px;border:0;border-radius:14px;background:#0f2a43;color:#fff;box-shadow:0 12px 32px rgba(15,42,67,.28);font:600 14px/1.35 system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;cursor:pointer}
        .kp-pwa-control[hidden]{display:none!important}.kp-pwa-control:focus-visible{outline:3px solid #60a5fa;outline-offset:3px}.kp-pwa-control svg{width:20px;height:20px;flex:0 0 auto;fill:currentColor}
        .kp-pwa-toast{position:fixed;left:50%;bottom:max(18px,env(safe-area-inset-bottom));z-index:100001;transform:translate(-50%,18px);max-width:min(520px,calc(100vw - 32px));padding:11px 16px;border-radius:12px;background:#172033;color:#fff;box-shadow:0 10px 28px rgba(15,23,42,.28);font:500 14px/1.45 system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;opacity:0;pointer-events:none;transition:.2s ease}
        .kp-pwa-toast.is-visible{opacity:1;transform:translate(-50%,0)}
        @media (max-width:640px){.kp-pwa-control{left:16px;right:16px;justify-content:center;max-width:none}.kp-pwa-toast{bottom:76px}}
        @media print{.kp-pwa-control,.kp-pwa-toast{display:none!important}}
    `;
    document.head.append(style);

    const installButton = document.createElement("button");
    installButton.type = "button";
    installButton.className = "kp-pwa-control";
    installButton.hidden = true;
    installButton.setAttribute("aria-label", "Instal KasirPro");
    installButton.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 16 7 11l1.4-1.4 2.6 2.6V3h2v9.2l2.6-2.6L17 11l-5 5Zm-7 5v-6h2v4h10v-4h2v6H5Z"/></svg><span>Instal KasirPro</span>';

    const toast = document.createElement("div");
    toast.className = "kp-pwa-toast";
    toast.setAttribute("role", "status");
    toast.setAttribute("aria-live", "polite");
    let toastTimer = null;

    function showToast(message, duration = 3500) {
        toast.textContent = message;
        toast.classList.add("is-visible");
        clearTimeout(toastTimer);
        toastTimer = setTimeout(() => toast.classList.remove("is-visible"), duration);
    }

    function showUpdateButton(worker) {
        installButton.hidden = false;
        installButton.querySelector("span").textContent = "Perbarui KasirPro";
        installButton.setAttribute("aria-label", "Perbarui KasirPro");
        installButton.onclick = () => worker.postMessage({ type: "SKIP_WAITING" });
    }

    window.addEventListener("beforeinstallprompt", (event) => {
        event.preventDefault();
        installPrompt = event;
        installButton.hidden = false;
    });

    installButton.onclick = async () => {
        if (!installPrompt) return;
        installButton.hidden = true;
        await installPrompt.prompt();
        const choice = await installPrompt.userChoice;
        if (choice.outcome === "accepted") showToast("KasirPro sedang dipasang.");
        installPrompt = null;
    };

    window.addEventListener("appinstalled", () => {
        installButton.hidden = true;
        installPrompt = null;
        showToast("KasirPro berhasil dipasang.");
    });

    window.addEventListener("online", () => showToast("Koneksi internet kembali aktif."));
    window.addEventListener("offline", () => showToast("KasirPro memerlukan internet untuk membaca dan menyimpan data Firestore.", 5500));
    window.addEventListener("kasirpro:database-ready", () => showToast("Database Firestore siap digunakan."));
    window.addEventListener("kasirpro:database-synced", (event) => {
        const total = Number(event.detail?.totalRecords || 0);
        if (total > 0) showToast(`${total.toLocaleString("id-ID")} perubahan tersimpan ke Firestore.`);
    });
    window.addEventListener("kasirpro:database-error", () => showToast("Operasi Firestore gagal. Data tidak disimpan sebagai cadangan lokal.", 6000));

    document.addEventListener("DOMContentLoaded", () => {
        document.body.append(installButton, toast);
        const isIos = /iphone|ipad|ipod/i.test(navigator.userAgent);
        const isStandalone = window.matchMedia("(display-mode: standalone)").matches || navigator.standalone;
        if (isIos && !isStandalone) {
            setTimeout(() => showToast("Untuk instal di iPhone/iPad: pilih Bagikan, lalu Tambahkan ke Layar Utama.", 6500), 1200);
        }
    });

    if (!("serviceWorker" in navigator)) return;

    window.addEventListener("load", async () => {
        try {
            registration = await navigator.serviceWorker.register(new URL("sw.js", appRoot), {
                scope: appRoot.pathname,
                updateViaCache: "none"
            });
            await registration.update().catch(() => {});
            if (registration.waiting) showUpdateButton(registration.waiting);
            registration.addEventListener("updatefound", () => {
                const worker = registration.installing;
                if (!worker) return;
                worker.addEventListener("statechange", () => {
                    if (worker.state === "installed" && navigator.serviceWorker.controller) {
                        showUpdateButton(worker);
                    }
                });
            });
            navigator.serviceWorker.addEventListener("controllerchange", () => {
                if (refreshing) return;
                refreshing = true;
                window.location.reload();
            });
        } catch (error) {
            console.error("PWA KasirPro gagal diaktifkan:", error);
        }
    });
})();

