import {
    readExcelWorkbook,
    getWorkbookStats,
    searchSheetRows,
    formatFileSize,
    analyzeMasterWorkbook,
    createEmptyMasterStore,
    buildMasterImportPlan,
    applyMasterImportPlan,
    validateMasterWorkbookHeaders,
    validateSupplierReferences
} from "../modules/excel/excel-import.js";

import {
    initializeDatabase,
    readStore,
    writeStore
} from "../modules/database/database-store.js";

import { signOutKasirPro, sendKasirProPasswordReset, usernameToFirebaseEmail } from "../modules/database/auth.js";
import { firebaseDb } from "../modules/database/firebase-client.js";
import { collection, doc, getDocs, setDoc } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-firestore.js";
import { documentSegments } from "../modules/database/database-paths.js";

const databaseInitialization = initializeDatabase();


const SESSION_KEY =
    "kasirpro_session";

const MASTER_STORE_KEY =
    "kasirpro_master_store_v1";
const SALES_STORE_KEY = "kasirpro_sales_v1";
const PURCHASE_INVOICE_STORE_KEY = "kasirpro_purchase_invoices_v1";

let dashboardPeriod = "today";


let isRedirecting =
    false;

let selectedMasterExcelFile =
    null;

let masterWorkbookData =
    null;

let masterAnalysis =
    null;

let selectedUserRecord = null;
let resetUserPasswordButton = null;
let userDirectoryCache = [];
let userDirectoryLoaded = false;

let masterImportPlan =
    null;

const MASTER_IMPORT_ALLOWED_SHEETS = new Set([
    "petunjuk_import",
    "produk",
    "kategori",
    "supplier"
]);

let activeMasterSheet =
    null;

let productCurrentPage =
    1;

let productViewCache = null;
let productFilterOptionsReady = false;
let productSearchDebounceTimer = null;
const EMPTY_CELL_TEXT = "\u2014";

let categoryCurrentPage =
    1;

let supplierCurrentPage =
    1;

let userCurrentPage =
    1;

const PRODUCT_PAGE_SIZE =
    20;

const VIEW_CONFIG = {

    dashboard: {
        title: "Dashboard",
        subtitle: "Ringkasan operasional toko"
    },

    "import-master": {
        title: "Import Master",
        subtitle: "Import data master dari workbook Excel"
    },

    "import-faktur": {
        title: "Import Faktur",
        subtitle: "Import faktur pembelian dari Excel"
    },

    "import-opname": {
        title: "Import Stock Opname",
        subtitle: "Import hasil perhitungan stok fisik"
    },

    products: {
        title: "Produk",
        subtitle: "Kelola master produk"
    },

    categories: {
        title: "Kategori",
        subtitle: "Kelola kategori produk"
    },

    suppliers: {
        title: "Supplier",
        subtitle: "Kelola data supplier"
    },

    stock: {
        title: "Stok Saat Ini",
        subtitle: "Lihat posisi stok produk"
    },

    "goods-in": {
        title: "Barang Masuk",
        subtitle: "Riwayat penambahan persediaan"
    },

    "stock-opname": {
        title: "Stock Opname",
        subtitle: "Penyesuaian stok fisik"
    },

    "purchase-invoices": {
        title: "Faktur Pembelian",
        subtitle: "Kelola faktur pembelian"
    },

    sales: {
        title: "Transaksi Penjualan",
        subtitle: "Riwayat transaksi dari POS"
    },

    reports: {
        title: "Laporan",
        subtitle: "Analisis dan laporan operasional"
    },

    users: {
        title: "Pengguna",
        subtitle: "Kelola akun admin dan kasir"
    },

    settings: {
        title: "Pengaturan Toko",
        subtitle: "Konfigurasi identitas toko"
    }

};


const sidebar =
    document.getElementById("sidebar");

const sidebarToggle =
    document.getElementById("sidebar-toggle");

const sidebarOverlay =
    document.getElementById("sidebar-overlay");

const logoutButton =
    document.getElementById("logout-button");

const pageTitle =
    document.getElementById("page-title");

const pageSubtitle =
    document.getElementById("page-subtitle");

const currentUserName =
    document.getElementById("current-user-name");

const navParents =
    document.querySelectorAll(".nav-parent");

const navItems =
    document.querySelectorAll(
        ".nav-item[data-view], .nav-subitem[data-view]"
    );


const masterDropZone =
    document.getElementById("master-drop-zone");

const masterFileInput =
    document.getElementById("master-file-input");

const browseMasterFileButton =
    document.getElementById("browse-master-file");

const selectedMasterFile =
    document.getElementById("selected-master-file");

const selectedFileName =
    document.getElementById("selected-file-name");

const selectedFileSize =
    document.getElementById("selected-file-size");

const removeMasterFileButton =
    document.getElementById("remove-master-file");

const readMasterFileButton =
    document.getElementById("read-master-file");

const masterSheetTabs =
    document.getElementById("master-sheet-tabs");

const masterPreviewHead =
    document.getElementById("master-preview-head");

const masterPreviewBody =
    document.getElementById("master-preview-body");

const masterPreviewSearch =
    document.getElementById("master-preview-search");

const summarySheetCount =
    document.getElementById("summary-sheet-count");

const summaryRowCount =
    document.getElementById("summary-row-count");

const summaryWarningCount =
    document.getElementById("summary-warning-count");

const previewVisibleCount =
    document.getElementById("preview-visible-count");


const warningTotalCount =
    document.getElementById("warning-total-count");

const warningEmptyRowCount =
    document.getElementById("warning-empty-row-count");

const warningPartialRowCount =
    document.getElementById("warning-partial-row-count");

const masterWarningBody =
    document.getElementById("master-warning-body");


const matchNewCount =
    document.getElementById("match-new-count");

const matchExactCount =
    document.getElementById("match-exact-count");

const matchReviewCount =
    document.getElementById("match-review-count");

const matchDuplicateCount =
    document.getElementById("match-duplicate-count");

const matchTotalCount =
    document.getElementById("match-total-count");

const masterMatchBody =
    document.getElementById("master-match-body");


const masterImportResult =
    document.getElementById("master-import-result");

const masterImportResultText =
    document.getElementById("master-import-result-text");

const resultAddedCount =
    document.getElementById("result-added-count");

const resultUpdatedCount =
    document.getElementById("result-updated-count");

const resultSkippedCount =
    document.getElementById("result-skipped-count");

const resultReviewCount =
    document.getElementById("result-review-count");

const refreshProductsButton =
    document.getElementById("refresh-products");

const productTotalCount =
    document.getElementById("product-total-count");

const productTotalStock =
    document.getElementById("product-total-stock");

const productLowStockCount =
    document.getElementById("product-low-stock-count");

const productSearch =
    document.getElementById("product-search");

const productCategoryFilter =
    document.getElementById("product-category-filter");

const productSupplierFilter =
    document.getElementById("product-supplier-filter");

const productStatusFilter =
    document.getElementById("product-status-filter");

const productsTableBody =
    document.getElementById("products-table-body");

const productsVisibleCount =
    document.getElementById("products-visible-count");

const productsTotalResultCount =
    document.getElementById("products-total-result-count");

const productsPrevPage =
    document.getElementById("products-prev-page");

const productsNextPage =
    document.getElementById("products-next-page");

const productDetailOverlay =
    document.getElementById("product-detail-overlay");

const closeProductDetailButton =
    document.getElementById("close-product-detail");

const closeProductDetailFooterButton =
    document.getElementById("close-product-detail-footer");

const productDetailTitle = document.getElementById("product-detail-title");
const productDetailCode = document.getElementById("product-detail-code");
const detailProductCode = document.getElementById("detail-product-code");
const detailProductName = document.getElementById("detail-product-name");
const detailProductCategory = document.getElementById("detail-product-category");
const detailProductUnit = document.getElementById("detail-product-unit");
const detailProductSupplier = document.getElementById("detail-product-supplier");
const detailProductBuyPrice = document.getElementById("detail-product-buy-price");
const detailProductSellPrice = document.getElementById("detail-product-sell-price");
const detailProductStock = document.getElementById("detail-product-stock");
const detailProductMinStock = document.getElementById("detail-product-min-stock");
const detailProductLocation = document.getElementById("detail-product-location");
const detailProductStatus = document.getElementById("detail-product-status");
const detailProductBatch = document.getElementById("detail-product-batch");
const detailProductExpired = document.getElementById("detail-product-expired");
const detailProductNotes = document.getElementById("detail-product-notes");

const refreshCategoriesButton = document.getElementById("refresh-categories");
const categoryTotalCount = document.getElementById("category-total-count");
const categoryActiveCount = document.getElementById("category-active-count");
const categorySearch = document.getElementById("category-search");
const categoryStatusFilter = document.getElementById("category-status-filter");
const categoriesTableBody = document.getElementById("categories-table-body");
const categoriesVisibleCount = document.getElementById("categories-visible-count");
const categoriesTotalResultCount = document.getElementById("categories-total-result-count");
const categoriesPrevPage = document.getElementById("categories-prev-page");
const categoriesNextPage = document.getElementById("categories-next-page");
const categoryDetailOverlay = document.getElementById("category-detail-overlay");
const closeCategoryDetailButton = document.getElementById("close-category-detail");
const closeCategoryDetailFooterButton = document.getElementById("close-category-detail-footer");
const categoryDetailTitle = document.getElementById("category-detail-title");
const categoryDetailCode = document.getElementById("category-detail-code");
const detailCategoryCode = document.getElementById("detail-category-code");
const detailCategoryName = document.getElementById("detail-category-name");
const detailCategoryStatus = document.getElementById("detail-category-status");
const detailCategoryDescription = document.getElementById("detail-category-description");

const refreshSuppliersButton = document.getElementById("refresh-suppliers");
const supplierTotalCount = document.getElementById("supplier-total-count");
const supplierActiveCount = document.getElementById("supplier-active-count");
const supplierLinkedProductCount = document.getElementById("supplier-linked-product-count");
const supplierSearch = document.getElementById("supplier-search");
const supplierStatusFilter = document.getElementById("supplier-status-filter");
const suppliersTableBody = document.getElementById("suppliers-table-body");
const suppliersVisibleCount = document.getElementById("suppliers-visible-count");
const suppliersTotalResultCount = document.getElementById("suppliers-total-result-count");
const suppliersPrevPage = document.getElementById("suppliers-prev-page");
const suppliersNextPage = document.getElementById("suppliers-next-page");
const supplierDetailOverlay = document.getElementById("supplier-detail-overlay");
const closeSupplierDetailButton = document.getElementById("close-supplier-detail");
const closeSupplierDetailFooterButton = document.getElementById("close-supplier-detail-footer");
const supplierDetailTitle = document.getElementById("supplier-detail-title");
const supplierDetailCode = document.getElementById("supplier-detail-code");
const detailSupplierCode = document.getElementById("detail-supplier-code");
const detailSupplierName = document.getElementById("detail-supplier-name");
const detailSupplierStatus = document.getElementById("detail-supplier-status");
const detailSupplierContact = document.getElementById("detail-supplier-contact");
const detailSupplierPhone = document.getElementById("detail-supplier-phone");
const detailSupplierEmail = document.getElementById("detail-supplier-email");
const detailSupplierNpwp = document.getElementById("detail-supplier-npwp");
const detailSupplierTerm = document.getElementById("detail-supplier-term");
const detailSupplierProductCount = document.getElementById("detail-supplier-product-count");
const detailSupplierAddress = document.getElementById("detail-supplier-address");
const detailSupplierNotes = document.getElementById("detail-supplier-notes");

const refreshUsersButton = document.getElementById("refresh-users");
const userTotalCount = document.getElementById("user-total-count");
const userAdminCount = document.getElementById("user-admin-count");
const userCashierCount = document.getElementById("user-cashier-count");
const userActiveCount = document.getElementById("user-active-count");
const userSearch = document.getElementById("user-search");
const userRoleFilter = document.getElementById("user-role-filter");
const userStatusFilter = document.getElementById("user-status-filter");
const usersTableBody = document.getElementById("users-table-body");
const usersVisibleCount = document.getElementById("users-visible-count");
const usersTotalResultCount = document.getElementById("users-total-result-count");
const usersPrevPage = document.getElementById("users-prev-page");
const usersNextPage = document.getElementById("users-next-page");
const userDetailOverlay = document.getElementById("user-detail-overlay");
const closeUserDetailButton = document.getElementById("close-user-detail");
const closeUserDetailFooterButton = document.getElementById("close-user-detail-footer");
const userDetailTitle = document.getElementById("user-detail-title");
const userDetailId = document.getElementById("user-detail-id");
const detailUserId = document.getElementById("detail-user-id");
const detailUserName = document.getElementById("detail-user-name");
const detailUserRole = document.getElementById("detail-user-role");
const detailUserUsername = document.getElementById("detail-user-username");
const detailUserStatus = document.getElementById("detail-user-status");
const detailUserPasswordState = document.getElementById("detail-user-password-state");
const detailUserPhone = document.getElementById("detail-user-phone");
const detailUserEmail = document.getElementById("detail-user-email");
const detailUserNotes = document.getElementById("detail-user-notes");
const editUserButton = document.getElementById("edit-user");
const userEditorOverlay = document.getElementById("user-editor-overlay");
const closeUserEditorButton = document.getElementById("close-user-editor");
const cancelUserEditorButton = document.getElementById("cancel-user-editor");
const userEditorForm = document.getElementById("user-editor-form");
const editUserName = document.getElementById("edit-user-name");
const editUserPhone = document.getElementById("edit-user-phone");
const editUserStatus = document.getElementById("edit-user-status");


const storeSettingsForm = document.getElementById("store-settings-form");
const reloadStoreSettingsButton = document.getElementById("reload-store-settings");
const settingsSaveNotice = document.getElementById("settings-save-notice");
const settingsDataSource = document.getElementById("settings-data-source");
const settingsLogoPreview = document.getElementById("settings-logo-preview");
const settingsLogoImage = document.getElementById("settings-logo-image");

const settingStoreName = document.getElementById("setting-store-name");
const settingStoreStatus = document.getElementById("setting-store-status");
const settingStoreAddress = document.getElementById("setting-store-address");
const settingStorePhone = document.getElementById("setting-store-phone");
const settingStoreEmail = document.getElementById("setting-store-email");
const settingStoreNpwp = document.getElementById("setting-store-npwp");
const settingStoreLogo = document.getElementById("setting-store-logo");
const settingTransactionPrefix = document.getElementById("setting-transaction-prefix");
const settingCurrency = document.getElementById("setting-currency");
const settingTimezone = document.getElementById("setting-timezone");
const settingReceiptSize = document.getElementById("setting-receipt-size");
const settingShowLogo = document.getElementById("setting-show-logo");
const settingShowCashier = document.getElementById("setting-show-cashier");
const settingShowDiscount = document.getElementById("setting-show-discount");
const settingReceiptFooter = document.getElementById("setting-receipt-footer");
const settingStoreNotes = document.getElementById("setting-store-notes");



async function startManagement() {

    try {

        setAppLoading(true);

        await databaseInitialization;
        init();

    } catch (error) {

        console.error("Management tidak dapat membuka database:", error);
        const message = error?.message || "Database Firebase belum dapat dibuka.";
        alert(`Management gagal memuat data.\n\n${message}\n\nSilakan periksa koneksi lalu muat ulang halaman.`);

    } finally {

        setAppLoading(false);

    }

}

function setAppLoading(state, message = "", progress = null) {
    const loading = document.getElementById("app-loading");
    if (!loading) return;
    loading.hidden = !state;
    let style = document.getElementById("kasirpro-loading-progress-style");
    if (!style) {
        style = document.createElement("style");
        style.id = "kasirpro-loading-progress-style";
        style.textContent = "#app-loading .app-loading-progress{width:min(300px,100%);margin:14px auto 0;text-align:center}#app-loading .app-loading-progress__value{font-size:1.75rem;line-height:1;font-weight:800;color:#174a8b;margin-bottom:9px;font-variant-numeric:tabular-nums}#app-loading .app-loading-progress__track{height:8px;border-radius:999px;background:#dbe8f7;overflow:hidden;box-shadow:inset 0 1px 2px rgba(20,62,110,.12)}#app-loading .app-loading-progress__bar{height:100%;width:0;border-radius:inherit;background:linear-gradient(90deg,#2563eb,#39a9ff);transition:width .22s ease}";
        document.head.appendChild(style);
    }
    const messageElement = loading.querySelector("[data-app-loading-message]") || loading.querySelector("p");
    if (message && messageElement) messageElement.textContent = message;
    let progressElement = loading.querySelector("[data-app-loading-progress]");
    if (!progressElement) {
        progressElement = document.createElement("div");
        progressElement.className = "app-loading-progress";
        progressElement.dataset.appLoadingProgress = "";
        progressElement.innerHTML = '<div class="app-loading-progress__value" data-app-loading-percent>0%</div><div class="app-loading-progress__track" role="progressbar" aria-label="Kemajuan pemuatan" aria-valuemin="0" aria-valuemax="100"><div class="app-loading-progress__bar" data-app-loading-bar></div></div>';
        const anchor = messageElement || loading.firstElementChild;
        if (anchor?.parentNode) anchor.parentNode.insertBefore(progressElement, anchor.nextSibling);
        else loading.appendChild(progressElement);
    }
    const value = Number(progress);
    const hasProgress = Number.isFinite(value);
    progressElement.hidden = !hasProgress;
    if (hasProgress) {
        const percent = Math.max(0, Math.min(100, Math.round(value)));
        progressElement.querySelector("[data-app-loading-percent]").textContent = `${percent}%`;
        const track = progressElement.querySelector("[role=progressbar]");
        track.setAttribute("aria-valuenow", String(percent));
        progressElement.querySelector("[data-app-loading-bar]").style.width = `${percent}%`;
    }
}

window.addEventListener("kasirpro:database-preload-start", (event) => {
    setAppLoading(true, event.detail?.message || "Menyiapkan data toko\u2026");
});
window.addEventListener("kasirpro:database-preload-progress", (event) => {
    const detail = event.detail || {};
    const percent = Number.isFinite(Number(detail.percent)) ? Number(detail.percent) : null;
    const loaded = Number(detail.loaded || 0).toLocaleString("id-ID");
    const total = Number(detail.total || 0).toLocaleString("id-ID");
    const baseMessage = detail.message || "Memuat data dari Firestore";
    const label = detail.total
        ? `${baseMessage} \u2014 selesai ${loaded} dari ${total}`
        : baseMessage;
    setAppLoading(true, label, percent);
});
window.addEventListener("kasirpro:database-loading", (event) => setAppLoading(true, event.detail?.message || "Mengambil data dari Firebase\u2026"));
window.addEventListener("kasirpro:database-idle", () => setAppLoading(false));
window.addEventListener("kasirpro:database-sync-start", (event) => setAppLoading(true, event.detail?.message || "Menyinkronkan data ke Firebase\u2026"));
window.addEventListener("kasirpro:database-sync-progress", (event) => {
    const detail = event.detail || {};
    setAppLoading(true, `Menyinkronkan data ${detail.currentBatch || 0}/${detail.totalBatches || 0}\u2026`);
});
window.addEventListener("kasirpro:database-synced", () => setAppLoading(false));
window.addEventListener("kasirpro:database-error", () => setAppLoading(false));


if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", startManagement, { once: true });
} else {
    startManagement();
}


function init() {

    const session =
        validateAdminSession();


    if (!session) {
        return;
    }


    setupCurrentUser(session);
    updateManagementBrand();

    bindNavigation();

    bindSidebar();

    bindDashboard();

    bindLogout();

    bindImportMaster();

    bindWizard();

    bindProducts();

    bindCategories();

    bindSuppliers();

    bindUsers();

    bindStoreSettings();

    window.addEventListener("kasirpro:database-synced", renderDashboard);
    window.addEventListener("kasirpro:database-ready", renderDashboard);

    openDashboard();
}

function updateManagementBrand() {
    const node = document.getElementById("management-store-name");
    if (!node) return;
    const store = loadStore();
    const rows = store?.pengaturan_toko || store?.pengaturanToko || store?.pengaturan || [];
    const settings = Array.isArray(rows) ? rows[0] : rows;
    const name = String(settings?.["Nama Toko"] || "").trim().replace(/\s+v\.?\s*2(?:\.0)?$/i, "").trim();
    node.textContent = name || "Management";
}


function validateAdminSession() {

    const raw =
        sessionStorage.getItem(
            SESSION_KEY
        );


    if (!raw) {

        redirectToLogin();

        return null;

    }


    try {

        const session =
            JSON.parse(raw);


        if (
            !session ||
            session.role !== "admin"
        ) {

            sessionStorage.removeItem(
                SESSION_KEY
            );

            redirectToLogin();

            return null;

        }


        return session;

    } catch {

        sessionStorage.removeItem(
            SESSION_KEY
        );

        redirectToLogin();

        return null;

    }

}


function redirectToLogin() {

    if (isRedirecting) {
        return;
    }


    isRedirecting =
        true;


    window.location.replace(
        "../index.html"
    );

}


function setupCurrentUser(
    session
) {

    currentUserName.textContent =
        session.name ||
        session.nama ||
        session.username ||
        "Administrator";

}


function bindNavigation() {

    navParents.forEach(
        (button) => {

            button.addEventListener(
                "click",
                () => {

                    const group =
                        button.dataset.group;


                    const wasOpen =
                        button.getAttribute(
                            "aria-expanded"
                        ) === "true";


                    closeAllSubmenus();


                    if (!wasOpen) {

                        openSubmenu(
                            button,
                            group
                        );

                    }

                }
            );

        }
    );


    navItems.forEach(
        (button) => {

            button.addEventListener(
                "click",
                () => {

                    openView(
                        button.dataset.view,
                        button
                    );

                }
            );

        }
    );

}


function closeAllSubmenus() {

    navParents.forEach(
        (parent) => {

            parent.setAttribute(
                "aria-expanded",
                "false"
            );


            const submenu =
                document.querySelector(
                    `[data-submenu="${parent.dataset.group}"]`
                );


            if (submenu) {

                submenu.hidden =
                    true;

            }

        }
    );

}


function openSubmenu(
    parent,
    group
) {

    const submenu =
        document.querySelector(
            `[data-submenu="${group}"]`
        );


    if (!submenu) {
        return;
    }


    parent.setAttribute(
        "aria-expanded",
        "true"
    );


    submenu.hidden =
        false;

}


function openView(
    view,
    button
) {

    document
        .querySelectorAll(
            ".nav-item.active, .nav-subitem.active"
        )
        .forEach(
            (item) =>
                item.classList.remove(
                    "active"
                )
        );


    button?.classList.add(
        "active"
    );


    const submenu =
        button?.closest(
            ".nav-submenu"
        );


    if (submenu) {

        const group =
            submenu.dataset.submenu;


        const parent =
            document.querySelector(
                `.nav-parent[data-group="${group}"]`
            );


        if (parent) {

            closeAllSubmenus();

            openSubmenu(
                parent,
                group
            );

        }

    }


    document
        .querySelectorAll(
            "[data-view-section]"
        )
        .forEach(
            (section) => {

                section.hidden =
                    section.dataset.viewSection !==
                    view;

            }
        );


    const config =
        VIEW_CONFIG[view];


    if (config) {

        pageTitle.textContent =
            config.title;


        pageSubtitle.textContent =
            config.subtitle;

    }


    createPlaceholder(
    view,
    config
);


if (
    view === "products"
) {

    renderProductsPage();

}

if (view === "dashboard") renderDashboard();


if (
    view === "categories"
) {

    renderCategoriesPage();

}


if (
    view === "suppliers"
) {

    renderSuppliersPage();

}


if (
    view === "users"
) {

    renderUsersPage();

}


if (
    view === "settings"
) {

    renderStoreSettingsPage();

}


closeMobileSidebar();

}


function createPlaceholder(
    view,
    config
) {

    const section =
        document.querySelector(
            `[data-view-section="${view}"]`
        );


    if (
        !section ||
        section.innerHTML.trim()
    ) {
        return;
    }


    section.innerHTML =
        `
            <div class="content-placeholder">

                <i class="fa-solid fa-screwdriver-wrench"></i>

                <h2>
                    ${escapeHTML(config?.title || "Modul")}
                </h2>

                <p>
                    Modul akan dibangun pada tahap berikutnya.
                </p>

            </div>
        `;

}


function openDashboard() {

    const button =
        document.querySelector(
            '[data-view="dashboard"]'
        );


    openView(
        "dashboard",
        button
    );

}

function bindDashboard() {
    document.querySelectorAll("[data-dashboard-period]").forEach((button) => {
        button.addEventListener("click", () => {
            dashboardPeriod = button.dataset.dashboardPeriod || "today";
            renderDashboard();
        });
    });
}

function dashboardNumber(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;
    const text = String(value ?? "").trim().replace(/rp/gi, "").replace(/\s/g, "");
    if (!text) return 0;
    const normalized = text.includes(".") && text.includes(",")
        ? text.replace(/\./g, "").replace(",", ".")
        : text.includes(",")
            ? text.replace(",", ".")
            : text.replace(/[^0-9.-]/g, "");
    const number = Number(normalized);
    return Number.isFinite(number) ? number : 0;
}

function dashboardCurrency(value) {
    return new Intl.NumberFormat("id-ID", {
        style: "currency",
        currency: "IDR",
        maximumFractionDigits: 0
    }).format(dashboardNumber(value));
}

function dashboardDateMatches(value, period = dashboardPeriod) {
    if (period === "all") return true;
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return false;
    const now = new Date();
    const startToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const endToday = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
    if (period === "today") return date >= startToday && date < endToday;
    if (period === "7days") {
        const start = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 6);
        return date >= start && date < endToday;
    }
    const startMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    return date >= startMonth && date < endToday;
}

function dashboardInvoiceTotal(invoice) {
    const itemTotal = (Array.isArray(invoice?.items) ? invoice.items : []).reduce((sum, item) => {
        const subtotal = dashboardNumber(item?.qty) * dashboardNumber(item?.buyPrice);
        return sum + Math.max(0, subtotal - dashboardNumber(item?.discount));
    }, 0);
    const afterDiscount = Math.max(0, itemTotal - dashboardNumber(invoice?.discount));
    return afterDiscount + (afterDiscount * dashboardNumber(invoice?.taxPercent) / 100) + dashboardNumber(invoice?.otherCost);
}

function dashboardSaleProfit(sale, productByCode) {
    const salesBeforeTax = Math.max(0, dashboardNumber(sale?.subtotal) - dashboardNumber(sale?.itemDiscount) - dashboardNumber(sale?.transDiscount));
    const cost = (Array.isArray(sale?.items) ? sale.items : []).reduce((sum, item) => {
        const fallback = productByCode.get(String(item?.code || "").trim().toLowerCase());
        const buyPrice = dashboardNumber(item?.buyPrice) || dashboardNumber(fallback?.["Harga Beli"]);
        return sum + (buyPrice * dashboardNumber(item?.qty));
    }, 0);
    return salesBeforeTax - cost;
}

function dashboardPeriodLabel() {
    return ({ today: "Hari Ini", "7days": "7 Hari Terakhir", month: "Bulan Ini", all: "Sepanjang Waktu" })[dashboardPeriod] || "Hari Ini";
}

function renderDashboard() {
    const store = loadStore();
    const products = Array.isArray(store.produk) ? store.produk : [];
    const suppliers = Array.isArray(store.supplier) ? store.supplier : [];
    const sales = Array.isArray(readStore(SALES_STORE_KEY, [])) ? readStore(SALES_STORE_KEY, []) : [];
    const invoices = Array.isArray(readStore(PURCHASE_INVOICE_STORE_KEY, [])) ? readStore(PURCHASE_INVOICE_STORE_KEY, []) : [];
    const productByCode = new Map(products.map((product) => [String(product?.["Kode Produk"] || "").trim().toLowerCase(), product]));
    const completed = sales.filter((sale) => String(sale?.status || "SELESAI").trim().toUpperCase() !== "VOID");
    const filteredSales = completed.filter((sale) => dashboardDateMatches(sale?.at));
    const voidSales = sales.filter((sale) => String(sale?.status || "").trim().toUpperCase() === "VOID" && dashboardDateMatches(sale?.voidAt || sale?.at));
    const filteredInvoices = invoices.filter((invoice) => dashboardDateMatches(invoice?.confirmedAt || invoice?.date || invoice?.createdAt));
    const now = new Date();
    now.setHours(23, 59, 59, 999);
    const dueInvoices = invoices.filter((invoice) => {
        const due = new Date(invoice?.dueDate);
        return String(invoice?.status || "").trim().toLowerCase() === "confirmed"
            && String(invoice?.paymentType || "").trim().toLowerCase() === "tempo"
            && !Number.isNaN(due.getTime())
            && due <= now;
    });
    const lowStock = products.filter((product) => {
        const stock = dashboardNumber(product?.["Stok Awal"]);
        const minimum = dashboardNumber(product?.["Stok Minimum"]);
        return stock <= 0 || (minimum > 0 && stock <= minimum);
    }).length;
    const draftInvoices = invoices.filter((invoice) => String(invoice?.status || "").trim().toLowerCase() === "draft").length;
    const totalRevenue = filteredSales.reduce((sum, sale) => sum + dashboardNumber(sale?.total), 0);
    const salesProfit = filteredSales.reduce((sum, sale) => sum + dashboardSaleProfit(sale, productByCode), 0);
    const purchaseInvoiceValue = filteredInvoices.reduce((sum, invoice) => sum + dashboardInvoiceTotal(invoice), 0);
    const stockValue = products.reduce((sum, product) => sum + (dashboardNumber(product?.["Stok Awal"]) * dashboardNumber(product?.["Harga Beli"])), 0);
    const totalItems = products.reduce((sum, product) => sum + dashboardNumber(product?.["Stok Awal"]), 0);
    const todaySales = completed.filter((sale) => dashboardDateMatches(sale?.at, "today")).length;
    const set = (id, value) => { const element = document.getElementById(id); if (element) element.textContent = String(value); };

    set("dashboard-period-label", dashboardPeriodLabel());
    document.querySelectorAll("[data-dashboard-period]").forEach((button) => button.classList.toggle("active", button.dataset.dashboardPeriod === dashboardPeriod));
    set("dashboard-total-revenue", dashboardCurrency(totalRevenue));
    set("dashboard-sales-profit", dashboardCurrency(salesProfit));
    set("dashboard-purchase-invoice-value", dashboardCurrency(purchaseInvoiceValue));
    set("dashboard-supplier-due", dashboardCurrency(dueInvoices.reduce((sum, invoice) => sum + dashboardInvoiceTotal(invoice), 0)));
    set("dashboard-supplier-due-note", `${dueInvoices.length} faktur tempo jatuh tempo; pembayaran supplier belum dicatat di aplikasi.`);
    set("dashboard-stock-value", dashboardCurrency(stockValue));
    set("dashboard-total-products", products.length);
    set("dashboard-total-stock", new Intl.NumberFormat("id-ID").format(totalItems));
    set("dashboard-total-suppliers", suppliers.length);
    set("dashboard-sales-today", todaySales);
    set("dashboard-low-stock", lowStock);
    set("dashboard-draft-invoices", draftInvoices);
    set("dashboard-void-sales", voidSales.length);
    set("dashboard-review-count", `${lowStock + draftInvoices + voidSales.length} perhatian`);
}


function bindSidebar() {

    sidebarToggle?.addEventListener(
        "click",
        () => {

            sidebar.classList.toggle(
                "open"
            );


            sidebarOverlay.classList.toggle(
                "show",
                sidebar.classList.contains(
                    "open"
                )
            );

        }
    );


    sidebarOverlay?.addEventListener(
        "click",
        closeMobileSidebar
    );

}


function closeMobileSidebar() {

    sidebar?.classList.remove(
        "open"
    );


    sidebarOverlay?.classList.remove(
        "show"
    );

}


function bindLogout() {

    logoutButton?.addEventListener(
        "click",
        () => showAppModal({
            title: "Keluar dari Management?",
            message: "Sesi Administrator akan diakhiri.",
            icon: "fa-right-from-bracket",
            confirmLabel: "Keluar",
            onConfirm: async () => {
                await signOutKasirPro().catch((error) =>
                    console.warn("Logout Firebase belum tuntas:", error)
                );
                redirectToLogin();
            }
        })
    );

}

function showAppModal({ title, message, icon = "fa-circle-info", confirmLabel = "Mengerti", onConfirm, cancellable = true }) {
    const modal = document.getElementById("app-modal");
    const titleElement = document.getElementById("app-modal-title");
    const messageElement = document.getElementById("app-modal-message");
    const iconElement = document.getElementById("app-modal-icon");
    const confirmButton = document.getElementById("app-modal-confirm");
    const cancelButton = document.getElementById("app-modal-cancel");
    if (!modal || !confirmButton || !cancelButton) return;

    titleElement.textContent = title;
    messageElement.textContent = message;
    iconElement.className = `fa-solid ${icon}`;
    confirmButton.textContent = confirmLabel;
    cancelButton.hidden = !cancellable;
    const close = () => { modal.hidden = true; };
    cancelButton.onclick = close;
    modal.onclick = (event) => { if (event.target === modal && cancellable) close(); };
    confirmButton.onclick = async () => { close(); await onConfirm?.(); };
    modal.hidden = false;
    confirmButton.focus();
}


function bindImportMaster() {

    browseMasterFileButton?.addEventListener(
        "click",
        (event) => {

            event.stopPropagation();

            masterFileInput.click();

        }
    );


    masterDropZone?.addEventListener(
        "click",
        (event) => {

            if (
                event.target.closest(
                    "#browse-master-file"
                )
            ) {
                return;
            }


            masterFileInput.click();

        }
    );


    masterDropZone?.addEventListener(
        "keydown",
        (event) => {

            if (
                event.key === "Enter" ||
                event.key === " "
            ) {

                event.preventDefault();

                masterFileInput.click();

            }

        }
    );


    masterFileInput?.addEventListener(
        "change",
        (event) => {

            const file =
                event.target.files?.[0];


            if (file) {

                selectMasterFile(
                    file
                );

            }

        }
    );


    removeMasterFileButton?.addEventListener(
        "click",
        resetImport
    );


    readMasterFileButton?.addEventListener(
        "click",
        readMasterFile
    );


    masterPreviewSearch?.addEventListener(
        "input",
        searchPreview
    );


    document
        .getElementById(
            "review-master-warnings"
        )
        ?.addEventListener(
            "click",
            () => {

                renderWarnings();

                showImportStep(
                    3
                );

            }
        );


    document
        .getElementById(
            "continue-master-confirm"
        )
        ?.addEventListener(
            "click",
            () => {

                prepareImportPlan();

                showImportStep(
                    4
                );

            }
        );


    document
        .getElementById(
            "confirm-master-import"
        )
        ?.addEventListener(
            "click",
            executeImport
        );


    document
        .getElementById(
            "import-another-master"
        )
        ?.addEventListener(
            "click",
            resetImport
        );


    setupDragDrop();

}


function selectMasterFile(
    file
) {

    const extension =
        file.name
            .split(".")
            .pop()
            ?.toLowerCase();


    if (
        extension !== "xlsx" &&
        extension !== "xls"
    ) {

        alert(
            "Gunakan file XLSX atau XLS."
        );

        return;

    }


    selectedMasterExcelFile =
        file;


    selectedFileName.textContent =
        file.name;


    selectedFileSize.textContent =
        formatFileSize(
            file.size
        );


    selectedMasterFile.hidden =
        false;


    readMasterFileButton.disabled =
        false;

}


function setupDragDrop() {

    [
        "dragenter",
        "dragover"
    ].forEach(
        (eventName) => {

            masterDropZone?.addEventListener(
                eventName,
                (event) => {

                    event.preventDefault();


                    masterDropZone.classList.add(
                        "drag-over"
                    );

                }
            );

        }
    );


    [
        "dragleave",
        "drop"
    ].forEach(
        (eventName) => {

            masterDropZone?.addEventListener(
                eventName,
                (event) => {

                    event.preventDefault();


                    masterDropZone.classList.remove(
                        "drag-over"
                    );

                }
            );

        }
    );


    masterDropZone?.addEventListener(
        "drop",
        (event) => {

            const file =
                event.dataTransfer
                    ?.files?.[0];


            if (file) {

                selectMasterFile(
                    file
                );

            }

        }
    );

}


async function readMasterFile() {

    if (!selectedMasterExcelFile) {
        return;
    }


    const original =
        readMasterFileButton.innerHTML;


    try {

        readMasterFileButton.disabled =
            true;


        readMasterFileButton.innerHTML =
            `
                <i class="fa-solid fa-spinner fa-spin"></i>
                Membaca...
            `;


        masterWorkbookData = sanitizeMasterWorkbook(
            await readExcelWorkbook(
                selectedMasterExcelFile
            )
        );


        masterAnalysis =
            analyzeMasterWorkbook(
                masterWorkbookData
            );

        const headerIssues = validateMasterWorkbookHeaders(masterWorkbookData);
        if (headerIssues.length) {
            const issue = headerIssues[0];
            throw new Error(issue.message || `Header sheet ${issue.sheet} belum sesuai. Kolom wajib: ${issue.missing.join(", ")}.`);
        }
        const supplierIssues = validateSupplierReferences(masterWorkbookData, loadStore());
        if (supplierIssues.length) {
            const issue = supplierIssues[0];
            const label = issue.label || "kosong";
            throw new Error(`Supplier produk tidak ditemukan pada baris ${issue.sourceRow}: ${label}. Tambahkan label yang sama pada sheet SUPPLIER terlebih dahulu.`);
        }


        const stats =
            getWorkbookStats(
                masterWorkbookData
            );


        summarySheetCount.textContent =
            stats.sheetCount;


        summaryRowCount.textContent =
            masterAnalysis.totalDataRows;


        summaryWarningCount.textContent =
            masterAnalysis.totalWarnings;


        renderSheetTabs();


        const firstSheet =
            masterWorkbookData
                .sheetNames
                .find(
                    (name) =>
                        normalizeText(name) !==
                        "petunjuk_import"
                ) ||
            masterWorkbookData
                .sheetNames[0];


        if (firstSheet) {

            selectSheet(
                firstSheet
            );

        }


        showImportStep(
            2
        );

    } catch (error) {

        console.error(error);


        alert(
            error.message ||
            "Gagal membaca file Excel."
        );

    } finally {

        readMasterFileButton.disabled =
            false;


        readMasterFileButton.innerHTML =
            original;

    }

}


function renderSheetTabs() {

    masterSheetTabs.innerHTML =
        "";


    masterWorkbookData
        .sheetNames
        .forEach(
            (sheetName) => {

                const button =
                    document.createElement(
                        "button"
                    );


                button.type =
                    "button";


                button.className =
                    "sheet-tab";


                button.dataset.sheet =
                    sheetName;


                button.textContent =
                    sheetName;


                button.addEventListener(
                    "click",
                    () => {

                        selectSheet(
                            sheetName
                        );

                    }
                );


                masterSheetTabs.appendChild(
                    button
                );

            }
        );

}


function selectSheet(
    sheetName
) {

    activeMasterSheet =
        sheetName;


    masterSheetTabs
        .querySelectorAll(
            ".sheet-tab"
        )
        .forEach(
            (button) => {

                button.classList.toggle(
                    "active",
                    button.dataset.sheet ===
                        sheetName
                );

            }
        );


    masterPreviewSearch.value =
        "";


    renderPreview(
        masterWorkbookData
            .sheets[
                sheetName
            ]
    );

}


function renderPreview(
    sheet,
    rows = sheet.rows
) {

    masterPreviewHead.innerHTML =
        "";


    const headerRow =
        document.createElement(
            "tr"
        );


    sheet.headers.forEach(
        (header) => {

            const th =
                document.createElement(
                    "th"
                );


            th.textContent =
                header;


            headerRow.appendChild(
                th
            );

        }
    );


    masterPreviewHead.appendChild(
        headerRow
    );


    masterPreviewBody.innerHTML =
        "";


    if (!rows.length) {

        masterPreviewBody.innerHTML =
            `
                <tr>
                    <td class="empty-table-state">
                        Tidak ada data.
                    </td>
                </tr>
            `;

    } else {

        rows
            .slice(
                0,
                100
            )
            .forEach(
                (rowData) => {

                    const tr =
                        document.createElement(
                            "tr"
                        );


                    rowData.values.forEach(
                        (value) => {

                            const td =
                                document.createElement(
                                    "td"
                                );


                            td.textContent = value || EMPTY_CELL_TEXT;


                            tr.appendChild(
                                td
                            );

                        }
                    );


                    masterPreviewBody.appendChild(
                        tr
                    );

                }
            );

    }


    previewVisibleCount.textContent =
        rows.length;

}


function searchPreview() {

    if (
        !masterWorkbookData ||
        !activeMasterSheet
    ) {
        return;
    }


    const sheet =
        masterWorkbookData
            .sheets[
                activeMasterSheet
            ];


    const rows =
        searchSheetRows(
            sheet,
            masterPreviewSearch.value
        );


    renderPreview(
        sheet,
        rows
    );

}


function renderWarnings() {

    warningTotalCount.textContent =
        masterAnalysis.totalWarnings;


    warningEmptyRowCount.textContent =
        masterAnalysis.emptyRows;


    warningPartialRowCount.textContent =
        masterAnalysis.partialRows;


    masterWarningBody.innerHTML =
        "";


    if (!masterAnalysis.items.length) {

        masterWarningBody.innerHTML =
            `
                <tr>
                    <td colspan="5" class="empty-table-state">
                        Tidak ditemukan peringatan.
                    </td>
                </tr>
            `;

        return;

    }


    masterAnalysis.items.forEach(
        (item) => {

            const tr =
                document.createElement(
                    "tr"
                );


            [
                item.severity === "warning"
                    ? "Peringatan"
                    : "Informasi",

                item.sheet,
                item.title,
                item.count,
                item.message

            ].forEach(
                (value) => {

                    const td =
                        document.createElement(
                            "td"
                        );


                    td.textContent =
                        value;


                    tr.appendChild(
                        td
                    );

                }
            );


            masterWarningBody.appendChild(
                tr
            );

        }
    );

}


function prepareImportPlan() {

    const store =
        loadStore();


    masterImportPlan =
        buildMasterImportPlan(
            masterWorkbookData,
            store
        );


    matchNewCount.textContent =
        masterImportPlan.summary.new;


    matchExactCount.textContent =
        masterImportPlan.summary.exact;


    matchReviewCount.textContent =
        masterImportPlan.summary.review;


    matchDuplicateCount.textContent =
        masterImportPlan.summary.duplicate;


    matchTotalCount.textContent =
        masterImportPlan.summary.total;


    renderMatchTable();

}


function renderMatchTable() {

    masterMatchBody.innerHTML =
        "";


    if (!masterImportPlan.items.length) {

        masterMatchBody.innerHTML =
            `
                <tr>
                    <td colspan="5" class="empty-table-state">
                        Tidak ada data untuk diproses.
                    </td>
                </tr>
            `;

        return;

    }


    masterImportPlan.items
        .slice(
            0,
            200
        )
        .forEach(
            (item) => {

                const tr =
                    document.createElement(
                        "tr"
                    );


                [
                    item.sheet,
                    item.sourceRow,
                    item.identity || EMPTY_CELL_TEXT,
                    getStatusLabel(
                        item.status
                    ),
                    item.matchBy || EMPTY_CELL_TEXT

                ].forEach(
                    (value) => {

                        const td =
                            document.createElement(
                                "td"
                            );


                        td.textContent =
                            value;


                        tr.appendChild(
                            td
                        );

                    }
                );


                masterMatchBody.appendChild(
                    tr
                );

            }
        );

}


function executeImport() {

    if (!masterImportPlan) {
        return;
    }


    const existingMode =
        document.querySelector(
            'input[name="existingDataMode"]:checked'
        )?.value ||
        "update";


    const currentStore =
        loadStore();


    const result =
        applyMasterImportPlan(
            masterImportPlan,
            currentStore,
            {
                existingMode
            }
        );

    result.store = preserveOperationalMasterData(
        currentStore,
        result.store
    );


    saveStore(
        result.store
    );


    resultAddedCount.textContent =
        result.added;


    resultUpdatedCount.textContent =
        result.updated;


    resultSkippedCount.textContent =
        result.skipped;


    resultReviewCount.textContent =
        result.review;


    masterImportResultText.textContent =
        `${result.total} data diperiksa. ${result.review} data ditahan untuk ditinjau Admin.`;


    document
        .querySelectorAll(
            "[data-import-step]"
        )
        .forEach(
            (panel) => {

                panel.hidden =
                    true;

            }
        );


    masterImportResult.hidden =
        false;

}


function loadStore() {

    return {
        ...createEmptyMasterStore(),
        ...readStore(
            MASTER_STORE_KEY,
            createEmptyMasterStore()
        )
    };

}


function saveStore(
    store
) {

    invalidateProductViewCache();

    writeStore(
        MASTER_STORE_KEY,
        store
    ).catch((error) =>
        console.error("Master data gagal disinkronkan:", error)
    );

}


function resetImport() {

    selectedMasterExcelFile =
        null;


    masterWorkbookData =
        null;


    masterAnalysis =
        null;


    masterImportPlan =
        null;


    activeMasterSheet =
        null;


    masterFileInput.value =
        "";


    selectedMasterFile.hidden =
        true;


    selectedFileName.textContent =
        "-";


    selectedFileSize.textContent =
        "-";


    readMasterFileButton.disabled =
        true;


    masterSheetTabs.innerHTML =
        "";


    masterPreviewHead.innerHTML =
        "";


    masterPreviewBody.innerHTML =
        `
            <tr>
                <td class="empty-table-state">
                    Belum ada data.
                </td>
            </tr>
        `;


    masterImportResult.hidden =
        true;


    showImportStep(
        1
    );

}

function invalidateProductViewCache() {
    productViewCache = null;
    productFilterOptionsReady = false;
}
function getProductViewCache() {
    if (productViewCache) return productViewCache;
    const store = loadStore();
    productViewCache = {
        products: Array.isArray(store.produk) ? store.produk : [],
        suppliers: Array.isArray(store.supplier) ? store.supplier : []
    };
    return productViewCache;
}
function renderProductsPage() {
    const view = getProductViewCache();
    updateProductSummary(view.products);
    if (!productFilterOptionsReady) {
        populateProductFilters(view.products, view.suppliers);
        productFilterOptionsReady = true;
    }
    renderProductsTable(view);
}
function updateProductSummary(products) {
    const total = products.length;
    const stock = products.reduce((sum, product) => sum + parseProductNumber(product["Stok Awal"]), 0);
    const lowStock = products.filter((product) => {
        const stockValue = parseProductNumber(product["Stok Awal"]);
        const minimum = parseProductNumber(product["Stok Minimum"]);
        return minimum > 0 && stockValue <= minimum;
    }).length;
    if (productTotalCount) productTotalCount.textContent = total;
    if (productTotalStock) productTotalStock.textContent = formatProductNumber(stock);
    if (productLowStockCount) productLowStockCount.textContent = lowStock;
}
function populateProductFilters(products, suppliers) {
    populateSelectOptions(productCategoryFilter, products.map((product) => product["Kategori"]));
    populateSupplierFilter(products, suppliers);
}
function getSupplierLabel(supplier) {
    return String(supplier?.["Supplier"] ?? supplier?.["Kode Supplier"] ?? "").trim();
}
function getSupplierCompanyName(supplier) {
    return String(supplier?.["Nama Supplier"] ?? supplier?.["Supplier"] ?? "").trim();
}
function resolveSupplierForProduct(product, suppliers) {
    const label = String(product?.["Supplier"] ?? "").trim().toLowerCase();
    const legacyCode = String(product?.["Kode Supplier"] ?? "").trim().toLowerCase();
    return (suppliers || []).find((supplier) => getSupplierLabel(supplier).toLowerCase() === label)
        || (suppliers || []).find((supplier) => String(supplier?.["Kode Supplier"] ?? "").trim().toLowerCase() === legacyCode)
        || null;
}
function populateSupplierFilter(products, suppliers) {
    if (!productSupplierFilter) return;
    const current = productSupplierFilter.value;
    const labels = [...new Set(products.map((product) => {
        const supplier = resolveSupplierForProduct(product, suppliers);
        return getSupplierLabel(supplier) || String(product["Supplier"] ?? product["Kode Supplier"] ?? "").trim();
    }).filter(Boolean))].sort((a, b) => a.localeCompare(b, "id"));
    productSupplierFilter.innerHTML = '<option value="">Semua Supplier</option>';
    labels.forEach((label) => {
        const supplier = (suppliers || []).find((item) => getSupplierLabel(item).toLowerCase() === label.toLowerCase());
        const name = getSupplierCompanyName(supplier);
        const option = document.createElement("option");
        option.value = label;
        option.textContent = name || label;
        productSupplierFilter.appendChild(option);
    });
    if (labels.includes(current)) productSupplierFilter.value = current;
}
function populateSelectOptions(element, values) {
    if (!element) return;
    const current = element.value;
    const defaultLabel = element.options[0]?.textContent || "Semua";
    element.innerHTML = "";
    const defaultOption = document.createElement("option");
    defaultOption.value = "";
    defaultOption.textContent = defaultLabel;
    element.appendChild(defaultOption);
    [...new Set(values.map((value) => String(value ?? "").trim()).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, "id"))
        .forEach((value) => {
            const option = document.createElement("option");
            option.value = value;
            option.textContent = value;
            element.appendChild(option);
        });
    if ([...element.options].some((option) => option.value === current)) element.value = current;
}
function renderProductsTable(view = getProductViewCache()) {
    const filtered = filterProducts(view.products);
    const totalPages = Math.max(1, Math.ceil(filtered.length / PRODUCT_PAGE_SIZE));
    productCurrentPage = Math.min(Math.max(productCurrentPage, 1), totalPages);
    const start = (productCurrentPage - 1) * PRODUCT_PAGE_SIZE;
    const visible = filtered.slice(start, start + PRODUCT_PAGE_SIZE);
    renderProductRows(visible, view.suppliers);
    updateProductPagination(filtered.length, visible.length, totalPages);
}
function filterProducts(products) {
    const keyword = String(productSearch?.value ?? "").trim().toLowerCase();
    const category = String(productCategoryFilter?.value ?? "").trim().toLowerCase();
    const supplier = String(productSupplierFilter?.value ?? "").trim().toLowerCase();
    const status = String(productStatusFilter?.value ?? "").trim().toLowerCase();
    return products.filter((product) => {
        const code = String(product["Kode Produk"] ?? "").toLowerCase();
        const name = String(product["Nama Produk"] ?? "").toLowerCase();
        const productCategory = String(product["Kategori"] ?? "").toLowerCase();
        const linkedSupplier = resolveSupplierForProduct(product, getProductViewCache().suppliers);
        const productSupplier = (getSupplierLabel(linkedSupplier) || String(product["Supplier"] ?? product["Kode Supplier"] ?? "")).toLowerCase();
        const productStatus = String(product["Status Produk"] ?? "").trim().toLowerCase();
        return (!keyword || code.includes(keyword) || name.includes(keyword)) &&
            (!category || productCategory === category) &&
            (!supplier || productSupplier === supplier) &&
            (!status || productStatus === status);
    });
}
function renderProductRows(products, suppliers) {
    if (!productsTableBody) return;

    productsTableBody.innerHTML = "";

    if (!products.length) {
        productsTableBody.innerHTML =
            '<tr><td colspan="10" class="empty-table-state">Tidak ada produk yang sesuai.</td></tr>';
        return;
    }

    products.forEach((product) => {
        const row = document.createElement("tr");
        row.classList.add("product-row");
        row.tabIndex = 0;
        row.title = "Klik untuk melihat detail produk";

        const supplier = resolveSupplierForProduct(product, suppliers);
        const supplierDisplay = getSupplierCompanyName(supplier)
            || String(product["Supplier"] ?? product["Kode Supplier"] ?? "").trim();
        const stock = parseProductNumber(product["Stok Awal"]);
        const minimumStock = parseProductNumber(product["Stok Minimum"]);

        [
            product["Kode Produk"],
            product["Nama Produk"],
            product["Kategori"],
            supplierDisplay,
            product["Satuan"],
            formatProductCurrency(product["Harga Beli"]),
            formatProductCurrency(product["Harga Jual"])
        ].forEach((value) => {
            const cell = document.createElement("td");
            cell.textContent =
                value === null || value === undefined || String(value).trim() === ""
                    ? EMPTY_CELL_TEXT : value;
            row.appendChild(cell);
        });

        const stockCell = document.createElement("td");
        const stockWrapper = document.createElement("div");
        const stockValue = document.createElement("strong");
        const stockBadge = document.createElement("span");
        stockWrapper.className = "product-stock";
        stockValue.textContent = formatProductNumber(stock);

        if (stock <= 0) {
            stockBadge.className = "stock-badge stock-empty";
            stockBadge.textContent = "Habis";
        } else if (minimumStock > 0 && stock <= minimumStock) {
            stockBadge.className = "stock-badge stock-low";
            stockBadge.textContent = "Menipis";
        } else {
            stockBadge.className = "stock-badge stock-safe";
            stockBadge.textContent = "Aman";
        }

        stockWrapper.append(stockValue, stockBadge);
        stockCell.appendChild(stockWrapper);
        row.appendChild(stockCell);

        const minimumCell = document.createElement("td");
        minimumCell.textContent = formatProductNumber(minimumStock);
        row.appendChild(minimumCell);

        const statusCell = document.createElement("td");
        const statusBadge = document.createElement("span");
        const status = String(product["Status Produk"] ?? "").trim().toLowerCase();

        if (status === "aktif") {
            statusBadge.className = "product-status product-status-active";
            statusBadge.textContent = "Aktif";
        } else if (status === "nonaktif") {
            statusBadge.className = "product-status product-status-inactive";
            statusBadge.textContent = "Nonaktif";
        } else {
            statusBadge.className = "product-status product-status-unknown";
            statusBadge.textContent = product["Status Produk"] || "Belum Diatur";
        }

        statusCell.appendChild(statusBadge);
        row.appendChild(statusCell);

        const openDetail = () => openProductDetail(product, suppliers);
        row.addEventListener("click", openDetail);
        row.addEventListener("keydown", (event) => {
            if (event.key === "Enter" || event.key === " ") {
                event.preventDefault();
                openDetail();
            }
        });

        productsTableBody.appendChild(row);
    });
}
function updateProductPagination(t,v,p){if(productsVisibleCount)productsVisibleCount.textContent=v;if(productsTotalResultCount)productsTotalResultCount.textContent=t;if(productsPrevPage)productsPrevPage.disabled=productCurrentPage<=1;if(productsNextPage)productsNextPage.disabled=productCurrentPage>=p||t===0;}
function parseProductNumber(v){if(typeof v==="number")return Number.isFinite(v)?v:0;let t=String(v??"").trim().replace(/rp/gi,"").replace(/\s/g,"");if(!t)return 0;if(t.includes(".")&&t.includes(","))t=t.replace(/\./g,"").replace(",",".");else if(t.includes(","))t=t.replace(",",".");else t=t.replace(/[^0-9.-]/g,"");const n=Number(t);return Number.isFinite(n)?n:0;}
function formatProductNumber(v){return new Intl.NumberFormat("id-ID",{maximumFractionDigits:2}).format(Number(v)||0);}
function formatProductCurrency(v){return new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",maximumFractionDigits:0}).format(parseProductNumber(v));}

function getProductText(value) {
    return String(value ?? "").trim();
}

function getProductDisplay(value) {
    return getProductText(value) || EMPTY_CELL_TEXT;
}

function openProductDetail(product, suppliers) {
    if (!productDetailOverlay) return;

    const supplier = resolveSupplierForProduct(product, suppliers);
    const supplierDisplay = getSupplierCompanyName(supplier)
        || getProductText(product["Supplier"])
        || getProductText(product["Kode Supplier"])
        || EMPTY_CELL_TEXT;
    const stock = parseProductNumber(product["Stok Awal"]);
    const minimumStock = parseProductNumber(product["Stok Minimum"]);
    const status = getProductText(product["Status Produk"]).toLowerCase();

    productDetailTitle.textContent = getProductText(product["Nama Produk"]) || "Produk Tanpa Nama";
    productDetailCode.textContent = getProductText(product["Kode Produk"]) || "Kode produk belum tersedia";
    detailProductCode.textContent = getProductDisplay(product["Kode Produk"]);
    detailProductName.textContent = getProductDisplay(product["Nama Produk"]);
    detailProductCategory.textContent = getProductDisplay(product["Kategori"]);
    detailProductUnit.textContent = getProductDisplay(product["Satuan"]);
    detailProductSupplier.textContent = supplierDisplay;
    detailProductBuyPrice.textContent = formatProductCurrency(product["Harga Beli"]);
    detailProductSellPrice.textContent = formatProductCurrency(product["Harga Jual"]);
    detailProductMinStock.textContent = formatProductNumber(minimumStock);
    detailProductLocation.textContent = getProductDisplay(product["Lokasi Rak"]);
    detailProductBatch.textContent = getProductDisplay(product["Batch"]);
    detailProductExpired.textContent = formatProductDate(product["Tanggal Expired"]);
    detailProductNotes.textContent = getProductText(product["Catatan"]) || "Tidak ada catatan.";

    renderDetailStock(stock, minimumStock);
    renderDetailStatus(status, product["Status Produk"]);

    productDetailOverlay.hidden = false;
    document.body.classList.add("product-detail-open");
    closeProductDetailButton?.focus();
}

function closeProductDetail() {
    if (!productDetailOverlay) return;
    productDetailOverlay.hidden = true;
    document.body.classList.remove("product-detail-open");
}

function renderDetailStock(stock, minimumStock) {
    if (!detailProductStock) return;
    detailProductStock.innerHTML = "";

    const value = document.createElement("strong");
    const badge = document.createElement("span");
    value.textContent = formatProductNumber(stock);

    if (stock <= 0) {
        badge.className = "stock-badge stock-empty";
        badge.textContent = "Habis";
    } else if (minimumStock > 0 && stock <= minimumStock) {
        badge.className = "stock-badge stock-low";
        badge.textContent = "Menipis";
    } else {
        badge.className = "stock-badge stock-safe";
        badge.textContent = "Aman";
    }

    detailProductStock.append(value, badge);
}

function renderDetailStatus(status, originalStatus) {
    if (!detailProductStatus) return;
    detailProductStatus.innerHTML = "";

    const badge = document.createElement("span");
    if (status === "aktif") {
        badge.className = "product-status product-status-active";
        badge.textContent = "Aktif";
    } else if (status === "nonaktif") {
        badge.className = "product-status product-status-inactive";
        badge.textContent = "Nonaktif";
    } else {
        badge.className = "product-status product-status-unknown";
        badge.textContent = getProductText(originalStatus) || "Belum Diatur";
    }

    detailProductStatus.appendChild(badge);
}

function formatProductDate(value) {
    if (!value) return EMPTY_CELL_TEXT;
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return new Intl.DateTimeFormat("id-ID", { day: "2-digit", month: "long", year: "numeric" }).format(value);
    }

    const text = getProductText(value);
    if (!text) return EMPTY_CELL_TEXT;
    const parsed = new Date(text);
    if (!Number.isNaN(parsed.getTime()) && /[-/T]/.test(text)) {
        return new Intl.DateTimeFormat("id-ID", { day: "2-digit", month: "long", year: "numeric" }).format(parsed);
    }
    return text;
}

function bindProducts() {

    refreshProductsButton?.addEventListener(
        "click",
        () => {

            productCurrentPage =
                1;


            renderProductsPage();

        }
    );


    productSearch?.addEventListener(
        "input",
        () => {
            window.clearTimeout(productSearchDebounceTimer);
            productSearchDebounceTimer = window.setTimeout(() => {
                productCurrentPage = 1;
                renderProductsTable();
            }, 250);
        }
    );


    productCategoryFilter?.addEventListener(
        "change",
        () => {

            productCurrentPage =
                1;


            renderProductsTable();

        }
    );


    productSupplierFilter?.addEventListener(
        "change",
        () => {

            productCurrentPage =
                1;


            renderProductsTable();

        }
    );


    productStatusFilter?.addEventListener(
        "change",
        () => {

            productCurrentPage =
                1;


            renderProductsTable();

        }
    );


    productsPrevPage?.addEventListener(
        "click",
        () => {

            if (
                productCurrentPage <= 1
            ) {
                return;
            }


            productCurrentPage--;


            renderProductsTable();

        }
    );


    productsNextPage?.addEventListener(
        "click",
        () => {

            productCurrentPage++;


            renderProductsTable();

        }
    );


    closeProductDetailButton?.addEventListener("click", closeProductDetail);
    closeProductDetailFooterButton?.addEventListener("click", closeProductDetail);

    productDetailOverlay?.addEventListener("click", (event) => {
        if (event.target === productDetailOverlay) closeProductDetail();
    });

    document.addEventListener("keydown", (event) => {
        if (event.key === "Escape" && productDetailOverlay && !productDetailOverlay.hidden) {
            closeProductDetail();
        }
    });

}


function bindCategories() {
    refreshCategoriesButton?.addEventListener("click", () => {
        categoryCurrentPage = 1;
        renderCategoriesPage();
    });

    categorySearch?.addEventListener("input", () => {
        categoryCurrentPage = 1;
        renderCategoriesTable();
    });

    categoryStatusFilter?.addEventListener("change", () => {
        categoryCurrentPage = 1;
        renderCategoriesTable();
    });

    categoriesPrevPage?.addEventListener("click", () => {
        if (categoryCurrentPage <= 1) return;
        categoryCurrentPage--;
        renderCategoriesTable();
    });

    categoriesNextPage?.addEventListener("click", () => {
        categoryCurrentPage++;
        renderCategoriesTable();
    });

    closeCategoryDetailButton?.addEventListener("click", closeCategoryDetail);
    closeCategoryDetailFooterButton?.addEventListener("click", closeCategoryDetail);

    categoryDetailOverlay?.addEventListener("click", (event) => {
        if (event.target === categoryDetailOverlay) closeCategoryDetail();
    });

    document.addEventListener("keydown", (event) => {
        if (event.key === "Escape" && categoryDetailOverlay && !categoryDetailOverlay.hidden) {
            closeCategoryDetail();
        }
    });
}


function renderCategoriesPage() {
    const store = loadStore();
    const categories = Array.isArray(store.kategori) ? store.kategori : [];

    if (categoryTotalCount) categoryTotalCount.textContent = categories.length;

    if (categoryActiveCount) {
        categoryActiveCount.textContent = categories.filter(
            (category) => String(category["Status"] ?? "").trim().toLowerCase() === "aktif"
        ).length;
    }

    renderCategoriesTable();
}


function renderCategoriesTable() {
    if (!categoriesTableBody) return;

    const store = loadStore();
    const categories = Array.isArray(store.kategori) ? store.kategori : [];
    const keyword = String(categorySearch?.value ?? "").trim().toLowerCase();
    const status = String(categoryStatusFilter?.value ?? "").trim().toLowerCase();

    const filtered = categories.filter((category) => {
        const code = String(category["Kode Kategori"] ?? "").toLowerCase();
        const name = String(category["Nama Kategori"] ?? "").toLowerCase();
        const description = String(category["Deskripsi"] ?? "").toLowerCase();
        const categoryStatus = String(category["Status"] ?? "").trim().toLowerCase();

        return (
            (!keyword || code.includes(keyword) || name.includes(keyword) || description.includes(keyword)) &&
            (!status || categoryStatus === status)
        );
    });

    const totalPages = Math.max(1, Math.ceil(filtered.length / PRODUCT_PAGE_SIZE));
    if (categoryCurrentPage > totalPages) categoryCurrentPage = totalPages;
    if (categoryCurrentPage < 1) categoryCurrentPage = 1;

    const startIndex = (categoryCurrentPage - 1) * PRODUCT_PAGE_SIZE;
    const pageCategories = filtered.slice(startIndex, startIndex + PRODUCT_PAGE_SIZE);

    categoriesTableBody.innerHTML = "";

    if (!pageCategories.length) {
        categoriesTableBody.innerHTML =
            '<tr><td colspan="4" class="empty-table-state">Tidak ada kategori yang sesuai.</td></tr>';
    } else {
        pageCategories.forEach((category) => {
            const row = document.createElement("tr");
            row.classList.add("product-row");
            row.tabIndex = 0;
            row.title = "Klik untuk melihat detail kategori";

            [category["Kode Kategori"], category["Nama Kategori"], category["Deskripsi"]].forEach((value) => {
                const cell = document.createElement("td");
                cell.textContent = getProductDisplay(value);
                row.appendChild(cell);
            });

            const statusCell = document.createElement("td");
            statusCell.appendChild(createCategoryStatusBadge(category["Status"]));
            row.appendChild(statusCell);

            const openDetail = () => openCategoryDetail(category);
            row.addEventListener("click", openDetail);
            row.addEventListener("keydown", (event) => {
                if (event.key === "Enter" || event.key === " ") {
                    event.preventDefault();
                    openDetail();
                }
            });

            categoriesTableBody.appendChild(row);
        });
    }

    if (categoriesVisibleCount) categoriesVisibleCount.textContent = pageCategories.length;
    if (categoriesTotalResultCount) categoriesTotalResultCount.textContent = filtered.length;
    if (categoriesPrevPage) categoriesPrevPage.disabled = categoryCurrentPage <= 1;
    if (categoriesNextPage) {
        categoriesNextPage.disabled =
            categoryCurrentPage >= totalPages || filtered.length === 0;
    }
}


function createCategoryStatusBadge(originalStatus) {
    const status = String(originalStatus ?? "").trim().toLowerCase();
    const badge = document.createElement("span");

    if (status === "aktif") {
        badge.className = "product-status product-status-active";
        badge.textContent = "Aktif";
    } else if (status === "nonaktif") {
        badge.className = "product-status product-status-inactive";
        badge.textContent = "Nonaktif";
    } else {
        badge.className = "product-status product-status-unknown";
        badge.textContent = getProductText(originalStatus) || "Belum Diatur";
    }

    return badge;
}


function openCategoryDetail(category) {
    if (!categoryDetailOverlay) return;

    const code = getProductText(category["Kode Kategori"]);
    const name = getProductText(category["Nama Kategori"]);

    categoryDetailTitle.textContent = name || "Kategori Tanpa Nama";
    categoryDetailCode.textContent = code || "Kode kategori belum tersedia";
    detailCategoryCode.textContent = code || EMPTY_CELL_TEXT;
    detailCategoryName.textContent = name || EMPTY_CELL_TEXT;
    detailCategoryDescription.textContent =
        getProductText(category["Deskripsi"]) || "Tidak ada deskripsi.";

    detailCategoryStatus.innerHTML = "";
    detailCategoryStatus.appendChild(createCategoryStatusBadge(category["Status"]));

    categoryDetailOverlay.hidden = false;
    document.body.classList.add("product-detail-open");
}


function closeCategoryDetail() {
    if (!categoryDetailOverlay) return;

    categoryDetailOverlay.hidden = true;

    if (!productDetailOverlay || productDetailOverlay.hidden) {
        document.body.classList.remove("product-detail-open");
    }
}



function bindSuppliers() {
    refreshSuppliersButton?.addEventListener("click", () => {
        supplierCurrentPage = 1;
        renderSuppliersPage();
    });

    supplierSearch?.addEventListener("input", () => {
        supplierCurrentPage = 1;
        renderSuppliersTable();
    });

    supplierStatusFilter?.addEventListener("change", () => {
        supplierCurrentPage = 1;
        renderSuppliersTable();
    });

    suppliersPrevPage?.addEventListener("click", () => {
        if (supplierCurrentPage <= 1) return;
        supplierCurrentPage--;
        renderSuppliersTable();
    });

    suppliersNextPage?.addEventListener("click", () => {
        supplierCurrentPage++;
        renderSuppliersTable();
    });

    closeSupplierDetailButton?.addEventListener("click", closeSupplierDetail);
    closeSupplierDetailFooterButton?.addEventListener("click", closeSupplierDetail);

    supplierDetailOverlay?.addEventListener("click", (event) => {
        if (event.target === supplierDetailOverlay) closeSupplierDetail();
    });

    document.addEventListener("keydown", (event) => {
        if (event.key === "Escape" && supplierDetailOverlay && !supplierDetailOverlay.hidden) {
            closeSupplierDetail();
        }
    });
}


function renderSuppliersPage() {
    const store = loadStore();
    const suppliers = Array.isArray(store.supplier) ? store.supplier : [];
    const products = Array.isArray(store.produk) ? store.produk : [];

    if (supplierTotalCount) supplierTotalCount.textContent = suppliers.length;

    if (supplierActiveCount) {
        supplierActiveCount.textContent = suppliers.filter(
            (supplier) => String(supplier["Status"] ?? "").trim().toLowerCase() === "aktif"
        ).length;
    }

    if (supplierLinkedProductCount) {
        supplierLinkedProductCount.textContent = products.filter(
            (product) => getProductText(product["Supplier"]) || getProductText(product["Kode Supplier"])
        ).length;
    }

    renderSuppliersTable();
}


function renderSuppliersTable() {
    if (!suppliersTableBody) return;

    const store = loadStore();
    const suppliers = Array.isArray(store.supplier) ? store.supplier : [];
    const products = Array.isArray(store.produk) ? store.produk : [];
    const keyword = String(supplierSearch?.value ?? "").trim().toLowerCase();
    const status = String(supplierStatusFilter?.value ?? "").trim().toLowerCase();

    const filtered = suppliers.filter((supplier) => {
            const searchableValues = [
            supplier["Supplier"],
            supplier["Nama Supplier"],
            supplier["Kontak Person"],
            supplier["Telepon"],
            supplier["Email"],
            supplier["Alamat"],
            supplier["NPWP"]
        ].map((value) => String(value ?? "").toLowerCase());

        const supplierStatus = String(supplier["Status"] ?? "").trim().toLowerCase();

        return (
            (!keyword || searchableValues.some((value) => value.includes(keyword))) &&
            (!status || supplierStatus === status)
        );
    });

    const totalPages = Math.max(1, Math.ceil(filtered.length / PRODUCT_PAGE_SIZE));
    if (supplierCurrentPage > totalPages) supplierCurrentPage = totalPages;
    if (supplierCurrentPage < 1) supplierCurrentPage = 1;

    const startIndex = (supplierCurrentPage - 1) * PRODUCT_PAGE_SIZE;
    const pageSuppliers = filtered.slice(startIndex, startIndex + PRODUCT_PAGE_SIZE);

    suppliersTableBody.innerHTML = "";

    if (!pageSuppliers.length) {
        suppliersTableBody.innerHTML =
            '<tr><td colspan="8" class="empty-table-state">Tidak ada supplier yang sesuai.</td></tr>';
    } else {
        pageSuppliers.forEach((supplier) => {
            const row = document.createElement("tr");
            row.classList.add("product-row");
            row.tabIndex = 0;
            row.title = "Klik untuk melihat detail supplier";

            const supplierLabel = getSupplierLabel(supplier);
            const linkedCount = countProductsForSupplier(products, supplierLabel, supplier["Kode Supplier"]);

            [
                supplierLabel,
                supplier["Nama Supplier"],
                supplier["Kontak Person"],
                supplier["Telepon"],
                supplier["Email"],
                supplier["Termin Default"]
            ].forEach((value) => {
                const cell = document.createElement("td");
                cell.textContent = getProductDisplay(value);
                row.appendChild(cell);
            });

            const productCell = document.createElement("td");
            productCell.textContent = formatProductNumber(linkedCount);
            row.appendChild(productCell);

            const statusCell = document.createElement("td");
            statusCell.appendChild(createSupplierStatusBadge(supplier["Status"]));
            row.appendChild(statusCell);

            const openDetail = () => openSupplierDetail(supplier, products);
            row.addEventListener("click", openDetail);
            row.addEventListener("keydown", (event) => {
                if (event.key === "Enter" || event.key === " ") {
                    event.preventDefault();
                    openDetail();
                }
            });

            suppliersTableBody.appendChild(row);
        });
    }

    if (suppliersVisibleCount) suppliersVisibleCount.textContent = pageSuppliers.length;
    if (suppliersTotalResultCount) suppliersTotalResultCount.textContent = filtered.length;
    if (suppliersPrevPage) suppliersPrevPage.disabled = supplierCurrentPage <= 1;
    if (suppliersNextPage) {
        suppliersNextPage.disabled =
            supplierCurrentPage >= totalPages || filtered.length === 0;
    }
}


function countProductsForSupplier(products, supplierLabel, legacyCode = "") {
    const normalizedLabel = getProductText(supplierLabel).toLowerCase();
    const normalizedCode = getProductText(legacyCode).toLowerCase();
    if (!normalizedLabel && !normalizedCode) return 0;
    return products.filter((product) => {
        const productLabel = getProductText(product["Supplier"]).toLowerCase();
        const productCode = getProductText(product["Kode Supplier"]).toLowerCase();
        return (normalizedLabel && productLabel === normalizedLabel)
            || (normalizedCode && productCode === normalizedCode);
    }).length;
}


function createSupplierStatusBadge(originalStatus) {
    const status = String(originalStatus ?? "").trim().toLowerCase();
    const badge = document.createElement("span");

    if (status === "aktif") {
        badge.className = "product-status product-status-active";
        badge.textContent = "Aktif";
    } else if (status === "nonaktif") {
        badge.className = "product-status product-status-inactive";
        badge.textContent = "Nonaktif";
    } else {
        badge.className = "product-status product-status-unknown";
        badge.textContent = getProductText(originalStatus) || "Belum Diatur";
    }

    return badge;
}


function openSupplierDetail(supplier, products) {
    if (!supplierDetailOverlay) return;

    const label = getSupplierLabel(supplier);
    const name = getProductText(supplier["Nama Supplier"]);

    supplierDetailTitle.textContent = name || "Supplier Tanpa Nama";
    supplierDetailCode.textContent = label || "Label supplier belum tersedia";
    detailSupplierCode.textContent = label || EMPTY_CELL_TEXT;
    detailSupplierName.textContent = name || EMPTY_CELL_TEXT;
    detailSupplierContact.textContent = getProductDisplay(supplier["Kontak Person"]);
    detailSupplierPhone.textContent = getProductDisplay(supplier["Telepon"]);
    detailSupplierEmail.textContent = getProductDisplay(supplier["Email"]);
    detailSupplierNpwp.textContent = getProductDisplay(supplier["NPWP"]);
    detailSupplierTerm.textContent = getProductDisplay(supplier["Termin Default"]);
    detailSupplierProductCount.textContent = formatProductNumber(countProductsForSupplier(products, label, supplier["Kode Supplier"]));
    detailSupplierAddress.textContent = getProductText(supplier["Alamat"]) || "Tidak ada alamat.";
    detailSupplierNotes.textContent = getProductText(supplier["Catatan"]) || "Tidak ada catatan.";

    detailSupplierStatus.innerHTML = "";
    detailSupplierStatus.appendChild(createSupplierStatusBadge(supplier["Status"]));

    supplierDetailOverlay.hidden = false;
    document.body.classList.add("product-detail-open");
}


function closeSupplierDetail() {
    if (!supplierDetailOverlay) return;

    supplierDetailOverlay.hidden = true;

    const productClosed = !productDetailOverlay || productDetailOverlay.hidden;
    const categoryClosed = !categoryDetailOverlay || categoryDetailOverlay.hidden;

    if (productClosed && categoryClosed) {
        document.body.classList.remove("product-detail-open");
    }
}


function bindUsers() {
    refreshUsersButton?.addEventListener("click", async () => {
        userCurrentPage = 1;
        userDirectoryLoaded = false;
        await renderUsersPage();
    });

    userSearch?.addEventListener("input", () => {
        userCurrentPage = 1;
        renderUsersTable();
    });

    userRoleFilter?.addEventListener("change", () => {
        userCurrentPage = 1;
        renderUsersTable();
    });

    userStatusFilter?.addEventListener("change", () => {
        userCurrentPage = 1;
        renderUsersTable();
    });

    usersPrevPage?.addEventListener("click", () => {
        if (userCurrentPage <= 1) return;
        userCurrentPage--;
        renderUsersTable();
    });

    usersNextPage?.addEventListener("click", () => {
        userCurrentPage++;
        renderUsersTable();
    });

    closeUserDetailButton?.addEventListener("click", closeUserDetail);
    closeUserDetailFooterButton?.addEventListener("click", closeUserDetail);
    editUserButton?.addEventListener("click", openUserEditor);
    closeUserEditorButton?.addEventListener("click", closeUserEditor);
    cancelUserEditorButton?.addEventListener("click", closeUserEditor);
    userEditorOverlay?.addEventListener("click", (event) => { if (event.target === userEditorOverlay) closeUserEditor(); });
    userEditorForm?.addEventListener("submit", saveUserEditor);

    if (editUserButton && !resetUserPasswordButton) {
        resetUserPasswordButton = document.createElement("button");
        resetUserPasswordButton.type = "button";
        resetUserPasswordButton.id = "reset-user-password";
        resetUserPasswordButton.className = editUserButton.className;
        resetUserPasswordButton.innerHTML = '<i class="fa-solid fa-envelope"></i> Reset Password via Email';
        editUserButton.insertAdjacentElement("afterend", resetUserPasswordButton);
    }
    resetUserPasswordButton?.addEventListener("click", requestUserPasswordReset);

    userDetailOverlay?.addEventListener("click", (event) => {
        if (event.target === userDetailOverlay) closeUserDetail();
    });

    document.addEventListener("keydown", (event) => {
        if (event.key === "Escape" && userDetailOverlay && !userDetailOverlay.hidden) {
            closeUserDetail();
        }
    });
}


async function renderUsersPage() {
    if (!userDirectoryLoaded) {
        const localUsers = Array.isArray(loadStore().pengguna) ? loadStore().pengguna : [];
        try {
            const snapshot = await getDocs(collection(firebaseDb, ...documentSegments("users")));
            userDirectoryCache = snapshot.docs.map((item) => {
                const data = item.data() || {};
                return {
                    "ID Pengguna": item.id,
                    "Nama": data.name || data.displayName || data.username || "",
                    "Role": data.role || "",
                    "Username": data.username || "",
                    "Nomor Telepon": data.phone || "",
                    "Email": data.email || data.authEmail || "",
                    "Status": data.status || "aktif",
                    "Catatan": data.notes || ""
                };
            });
        } catch (error) {
            console.warn("Profil pengguna Firestore belum dapat dibaca:", error);
            userDirectoryCache = localUsers;
        }

        if (!userDirectoryCache.length) {
            try {
                const session = JSON.parse(sessionStorage.getItem(SESSION_KEY) || "{}");
                if (session?.username || session?.email || session?.uid) {
                    userDirectoryCache = [{
                        "ID Pengguna": session.uid || session.userId || "admin-aktif",
                        "Nama": session.name || session.displayName || session.username || "Administrator",
                        "Role": session.role || "admin",
                        "Username": session.username || "",
                        "Nomor Telepon": "",
                        "Email": session.email || session.authEmail || "",
                        "Status": "aktif",
                        "Catatan": "Sesi admin aktif; lengkapi profil akun pada koleksi users."
                    }];
                }
            } catch (error) {
                console.warn("Sesi admin tidak dapat dijadikan profil sementara:", error);
            }
        }
        userDirectoryLoaded = true;
    }

    updateUserSummary(userDirectoryCache);
    renderUsersTable();
}


function updateUserSummary(users) {
    const active = users.filter((user) => normalizeUserValue(user["Status"]) === "aktif").length;
    const admins = users.filter((user) => normalizeUserRole(user["Role"]) === "admin").length;
    const cashiers = users.filter((user) => normalizeUserRole(user["Role"]) === "kasir").length;

    if (userTotalCount) userTotalCount.textContent = users.length;
    if (userAdminCount) userAdminCount.textContent = admins;
    if (userCashierCount) userCashierCount.textContent = cashiers;
    if (userActiveCount) userActiveCount.textContent = active;
}


function renderUsersTable() {
    const users = userDirectoryCache;
    const filtered = filterUsers(users);
    const totalPages = Math.max(1, Math.ceil(filtered.length / PRODUCT_PAGE_SIZE));

    if (userCurrentPage > totalPages) userCurrentPage = totalPages;
    if (userCurrentPage < 1) userCurrentPage = 1;

    const startIndex = (userCurrentPage - 1) * PRODUCT_PAGE_SIZE;
    const visible = filtered.slice(startIndex, startIndex + PRODUCT_PAGE_SIZE);

    renderUserRows(visible);

    if (usersVisibleCount) usersVisibleCount.textContent = visible.length;
    if (usersTotalResultCount) usersTotalResultCount.textContent = filtered.length;
    if (usersPrevPage) usersPrevPage.disabled = userCurrentPage <= 1;
    if (usersNextPage) usersNextPage.disabled = userCurrentPage >= totalPages || filtered.length === 0;
}


function filterUsers(users) {
    const keyword = normalizeUserValue(userSearch?.value);
    const role = normalizeUserValue(userRoleFilter?.value);
    const status = normalizeUserValue(userStatusFilter?.value);

    return users.filter((user) => {
        const searchable = [
            user["ID Pengguna"],
            user["Nama"],
            user["Role"],
            user["Username"],
            user["Nomor Telepon"],
            user["Email"],
            user["Catatan"]
        ].map(normalizeUserValue).join(" ");

        if (keyword && !searchable.includes(keyword)) return false;
        if (role && normalizeUserRole(user["Role"]) !== role) return false;
        if (status && normalizeUserValue(user["Status"]) !== status) return false;
        return true;
    });
}


function renderUserRows(users) {
    if (!usersTableBody) return;
    usersTableBody.innerHTML = "";

    if (users.length === 0) {
        usersTableBody.innerHTML = `
            <tr>
                <td colspan="8" class="empty-table-state">Tidak ada pengguna yang sesuai.</td>
            </tr>
        `;
        return;
    }

    users.forEach((user) => {
        const row = document.createElement("tr");
        row.classList.add("product-row");
        row.tabIndex = 0;
        row.title = "Klik untuk melihat detail pengguna";

        const openDetail = () => openUserDetail(user);
        row.addEventListener("click", openDetail);
        row.addEventListener("keydown", (event) => {
            if (event.key === "Enter" || event.key === " ") {
                event.preventDefault();
                openDetail();
            }
        });

        const values = [
            getProductDisplay(user["ID Pengguna"]),
            getProductDisplay(user["Nama"]),
            formatUserRole(user["Role"]),
            getProductDisplay(user["Username"]),
            getProductDisplay(user["Nomor Telepon"]),
            getProductDisplay(user["Email"]),
            "Dikelola Firebase"
        ];

        values.forEach((value) => {
            const cell = document.createElement("td");
            cell.textContent = value;
            row.appendChild(cell);
        });

        const statusCell = document.createElement("td");
        statusCell.appendChild(createUserStatusBadge(user["Status"]));
        row.appendChild(statusCell);
        usersTableBody.appendChild(row);
    });
}


function openUserDetail(user) {
    if (!userDetailOverlay) return;

    selectedUserRecord = user;
    const name = getProductText(user["Nama"]) || "Pengguna Tanpa Nama";
    const id = getProductText(user["ID Pengguna"]);

    if (userDetailTitle) userDetailTitle.textContent = name;
    if (userDetailId) userDetailId.textContent = id || "ID pengguna belum tersedia";
    if (detailUserId) detailUserId.textContent = id || EMPTY_CELL_TEXT;
    if (detailUserName) detailUserName.textContent = getProductDisplay(user["Nama"]);
    if (detailUserRole) detailUserRole.textContent = formatUserRole(user["Role"]);
    if (detailUserUsername) detailUserUsername.textContent = getProductDisplay(user["Username"]);
    if (detailUserPhone) detailUserPhone.textContent = getProductDisplay(user["Nomor Telepon"]);
    if (detailUserEmail) detailUserEmail.textContent = getProductDisplay(user["Email"]);
    if (detailUserPasswordState) {
        detailUserPasswordState.textContent = "Dikelola aman oleh Firebase";
    }
    if (detailUserNotes) detailUserNotes.textContent = getProductText(user["Catatan"]) || "Tidak ada catatan.";

    if (detailUserStatus) {
        detailUserStatus.innerHTML = "";
        detailUserStatus.appendChild(createUserStatusBadge(user["Status"]));
    }

    userDetailOverlay.hidden = false;
    document.body.classList.add("product-detail-open");
}

function openUserEditor() {
    if (!selectedUserRecord || !userEditorOverlay) return;
    editUserName.value = getProductText(selectedUserRecord["Nama"]);
    editUserPhone.value = getProductText(selectedUserRecord["Nomor Telepon"]);
    editUserStatus.value = normalizeUserValue(selectedUserRecord["Status"]) === "nonaktif" ? "nonaktif" : "aktif";
    userEditorOverlay.hidden = false;
}

function closeUserEditor() {
    if (userEditorOverlay) userEditorOverlay.hidden = true;
}

async function saveUserEditor(event) {
    event.preventDefault();
    if (!selectedUserRecord) return;

    const saveButton = document.getElementById("save-user-editor");
    saveButton.disabled = true;
    try {
        const currentUsername = getProductText(selectedUserRecord["Username"]);
        const userId = selectedUserRecord["ID Pengguna"];

        await setDoc(doc(firebaseDb, ...documentSegments("users", userId)), {
            name: editUserName.value.trim(),
            phone: editUserPhone.value.trim(),
            status: editUserStatus.value,
            updatedAt: new Date().toISOString()
        }, { merge: true });
        await setDoc(doc(firebaseDb, ...documentSegments("loginDirectory", `akun-${currentUsername}`)), {
            userId,
            username: currentUsername,
            name: editUserName.value.trim(),
            status: editUserStatus.value,
            role: normalizeUserRole(selectedUserRecord["Role"]) === "kasir" ? "cashier" : "admin"
        }, { merge: true });

        closeUserEditor();
        closeUserDetail();
        showAppModal({ title: "Akun diperbarui", message: "Nama Tampilan, Nomor Telepon, dan Status telah disinkronkan ke Firebase.", icon: "fa-circle-check", cancellable: false, onConfirm: () => location.reload() });
    } catch (error) {
        console.error("Perubahan akun gagal:", error);
        showAppModal({ title: "Akun belum diperbarui", message: error?.message || "Perubahan tidak dapat disimpan.", icon: "fa-triangle-exclamation", cancellable: false });
    } finally {
        saveButton.disabled = false;
    }
}

async function requestUserPasswordReset() {
    if (!selectedUserRecord || !resetUserPasswordButton) return;

    const username = getProductText(selectedUserRecord["Username"]);
    const recordedEmail = getProductText(selectedUserRecord["Email"]);
    let targetEmail;

    try {
        targetEmail = usernameToFirebaseEmail(username);
    } catch (error) {
        showAppModal({ title: "Reset belum dapat dikirim", message: error?.message || "Username akun tidak valid.", icon: "fa-triangle-exclamation", cancellable: false });
        return;
    }

    if (recordedEmail.toLowerCase() !== targetEmail.toLowerCase()) {
        showAppModal({
            title: "Akun memakai email lama",
            message: `Reset tidak dikirim. Akun ini belum memakai email login ${targetEmail}.`,
            icon: "fa-triangle-exclamation",
            cancellable: false
        });
        return;
    }

    showAppModal({
        title: "Kirim email reset password?",
        message: `Tautan reset untuk ${username} akan dikirim ke ${targetEmail}. Semua alias ini masuk ke inbox Gmail Admin.`,
        icon: "fa-envelope",
        confirmLabel: "Kirim Email",
        onConfirm: async () => {
            resetUserPasswordButton.disabled = true;
            try {
                const sentTo = await sendKasirProPasswordReset({ username, authEmail: recordedEmail });
                showAppModal({
                    title: "Email reset dikirim",
                    message: `Tautan reset password telah dikirim ke ${sentTo} dan masuk ke inbox Gmail Admin.`,
                    icon: "fa-circle-check",
                    cancellable: false
                });
            } catch (error) {
                console.error("Reset password gagal:", error);
                showAppModal({
                    title: "Email reset belum terkirim",
                    message: error?.message || "Firebase belum dapat mengirim email reset.",
                    icon: "fa-triangle-exclamation",
                    cancellable: false
                });
            } finally {
                resetUserPasswordButton.disabled = false;
            }
        }
    });
}


function closeUserDetail() {
    if (!userDetailOverlay) return;
    userDetailOverlay.hidden = true;
    document.body.classList.remove("product-detail-open");
}


function createUserStatusBadge(statusValue) {
    const badge = document.createElement("span");
    const status = normalizeUserValue(statusValue);

    if (status === "aktif") {
        badge.className = "product-status product-status-active";
        badge.textContent = "Aktif";
    } else if (status === "nonaktif") {
        badge.className = "product-status product-status-inactive";
        badge.textContent = "Nonaktif";
    } else {
        badge.className = "product-status product-status-unknown";
        badge.textContent = getProductText(statusValue) || "Belum Diatur";
    }

    return badge;
}


function normalizeUserRole(value) {
    const role = normalizeUserValue(value);
    if (role === "administrator") return "admin";
    if (role.startsWith("kasir")) return "kasir";
    return role;
}


function formatUserRole(value) {
    const role = normalizeUserRole(value);
    if (role === "admin") return "Admin";
    if (role === "kasir") return "Kasir";
    return getProductDisplay(value);
}


function normalizeUserValue(value) {
    return String(value ?? "").trim().toLowerCase();
}


function bindWizard() {

    document
        .querySelectorAll(
            "[data-import-back]"
        )
        .forEach(
            (button) => {

                button.addEventListener(
                    "click",
                    () => {

                        showImportStep(
                            Number(
                                button.dataset.importBack
                            )
                        );

                    }
                );

            }
        );

}


function showImportStep(
    step
) {

    masterImportResult.hidden =
        true;


    document
        .querySelectorAll(
            "[data-import-step]"
        )
        .forEach(
            (panel) => {

                panel.hidden =
                    Number(
                        panel.dataset.importStep
                    ) !==
                    step;

            }
        );


    document
        .querySelectorAll(
            "[data-step-indicator]"
        )
        .forEach(
            (indicator) => {

                const number =
                    Number(
                        indicator.dataset.stepIndicator
                    );


                indicator.classList.toggle(
                    "active",
                    number === step
                );


                indicator.classList.toggle(
                    "completed",
                    number < step
                );

            }
        );

}


function getStatusLabel(
    status
) {

    const labels = {
        new: "Baru",
        exact: "Sudah Ada",
        review: "Perlu Ditinjau",
        duplicate: "Duplikat Workbook"
    };


    return labels[status] ||
        status;

}


function sanitizeMasterWorkbook(workbook) {
    const sheetNames = Array.isArray(workbook?.sheetNames)
        ? workbook.sheetNames.filter((name) => MASTER_IMPORT_ALLOWED_SHEETS.has(normalizeText(name)))
        : [];
    const sheets = Object.fromEntries(
        sheetNames
            .filter((name) => workbook?.sheets?.[name])
            .map((name) => [name, workbook.sheets[name]])
    );

    if (!["produk", "kategori", "supplier"].every((required) =>
        sheetNames.some((name) => normalizeText(name) === required)
    )) {
        throw new Error("Master wajib berisi sheet PRODUK, KATEGORI, dan SUPPLIER.");
    }

    return {
        ...workbook,
        sheetNames,
        sheets
    };
}


function masterProductIdentity(product) {
    const code = normalizeText(product?.["Kode Produk"]);
    return code ? `code:${code}` : `name:${normalizeText(product?.["Nama Produk"])}`;
}


function preserveOperationalMasterData(currentStore, importedStore) {
    const protectedStore = {
        ...importedStore
    };
    const protectedKeys = new Set([
        "pengguna",
        "users",
        "pengaturan",
        "pengaturan_toko",
        "settings",
        "store_settings"
    ]);

    Object.keys(currentStore || {}).forEach((key) => {
        if (protectedKeys.has(normalizeText(key))) {
            protectedStore[key] = currentStore[key];
        }
    });

    const existingStock = new Map(
        (Array.isArray(currentStore?.produk) ? currentStore.produk : [])
            .map((product) => [masterProductIdentity(product), product?.["Stok Awal"]])
    );

    protectedStore.produk = (Array.isArray(importedStore?.produk) ? importedStore.produk : [])
        .map((product) => {
            const identity = masterProductIdentity(product);
            if (!existingStock.has(identity)) return product;
            return {
                ...product,
                "Stok Awal": existingStock.get(identity)
            };
        });

    return protectedStore;
}


function normalizeText(
    value
) {

    return String(
        value ?? ""
    )
        .trim()
        .toLowerCase()
        .replace(/\s+/g, "_");

}


function escapeHTML(
    value
) {

    return String(
        value ?? ""
    )
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");

}


/* =========================================================
   STORE SETTINGS
========================================================= */

function bindStoreSettings() {
    reloadStoreSettingsButton?.addEventListener("click", () => {
        renderStoreSettingsPage();
        showSettingsNotice("Pengaturan dimuat ulang dari Master Store.", "info");
    });

    storeSettingsForm?.addEventListener("submit", async (event) => {
        event.preventDefault();
        await saveStoreSettings();
    });
}


function renderStoreSettingsPage() {
    if (!storeSettingsForm) {
        return;
    }

    const store = loadStore();
    const source = resolveStoreSettingsSource(store);
    const settings = source.rows[0] || {};
    updateManagementBrand();

    setSettingsInput(settingStoreName, settings["Nama Toko"]);
    setSettingsInput(settingStoreStatus, settings["Status Toko"]);
    setSettingsInput(settingStoreAddress, settings["Alamat"]);
    setSettingsInput(settingStorePhone, settings["Telepon"]);
    setSettingsInput(settingStoreEmail, settings["Email"]);
    setSettingsInput(settingStoreNpwp, settings["NPWP"]);
    setSettingsInput(settingStoreLogo, settings["Logo"]);
    setSettingsInput(settingTransactionPrefix, settings["Prefix Transaksi"]);
    setSettingsInput(settingCurrency, "IDR (Rupiah)");
    setSettingsInput(settingTimezone, "Asia/Jakarta");
    setSettingsInput(settingReceiptSize, settings["Ukuran Struk"] || "80mm");
    setSettingsCheckbox(settingShowLogo, settings["Tampilkan Logo di Struk"]);
    setSettingsCheckbox(settingShowCashier, settings["Tampilkan Nama Kasir di Struk"]);
    setSettingsCheckbox(settingShowDiscount, settings["Tampilkan Diskon di Struk"]);
    setSettingsInput(settingReceiptFooter, settings["Footer Struk"]);
    setSettingsInput(settingStoreNotes, settings["Catatan"]);

    updateSettingsLogoPreview(settings["Logo"]);

    if (settingsDataSource) {
        settingsDataSource.textContent = source.rows.length
            ? `Sumber: ${source.label} — perubahan disimpan saat tombol Simpan ditekan.`
            : "Belum ada pengaturan tersimpan. Isi formulir lalu tekan Simpan Perubahan.";
    }
}


async function saveStoreSettings() {
    const saveButton = storeSettingsForm?.querySelector("button[type='submit']");
    if (saveButton) saveButton.disabled = true;
    try {
        const store = loadStore();
        const source = resolveStoreSettingsSource(store);
        const previous = source.rows[0] || {};
        const settings = {
            ...previous,
            "Nama Toko": getSettingsInput(settingStoreName),
            "Status Toko": getSettingsInput(settingStoreStatus),
            "Alamat": getSettingsInput(settingStoreAddress),
            "Telepon": getSettingsInput(settingStorePhone),
            "NPWP": getSettingsInput(settingStoreNpwp),
            "Prefix Transaksi": getSettingsInput(settingTransactionPrefix),
            "Ukuran Struk": getSettingsInput(settingReceiptSize) || "80mm",
            "Tampilkan Logo di Struk": settingShowLogo?.checked ? "Ya" : "Tidak",
            "Tampilkan Nama Kasir di Struk": settingShowCashier?.checked ? "Ya" : "Tidak",
            "Tampilkan Diskon di Struk": settingShowDiscount?.checked ? "Ya" : "Tidak",
            "Footer Struk": getSettingsInput(settingReceiptFooter),
            "Catatan": getSettingsInput(settingStoreNotes)
        };
        delete settings["Pajak Default (%)"];
        delete settings["Printer Default"];
        delete settings["Lebar Kertas"];
        delete settings["Tampilkan Pajak di Struk"];
        delete settings["Format Nomor Transaksi"];
        delete settings["Prefix Faktur"];
        delete settings["Format Nomor Faktur"];

        store.pengaturan_toko = [settings];
        if ("pengaturanToko" in store) store.pengaturanToko = [settings];
        if ("pengaturan" in store) store.pengaturan = [settings];

        await writeStore(MASTER_STORE_KEY, store);
        updateManagementBrand();
        showSettingsNotice("Pengaturan berhasil disimpan dan disinkronkan.", "success");
    } catch (error) {
        console.error("Pengaturan toko gagal disimpan:", error);
        showSettingsNotice(error?.message || "Pengaturan belum dapat disimpan.", "error");
    } finally {
        if (saveButton) saveButton.disabled = false;
    }
}


function resolveStoreSettingsSource(store) {
    const preferredKeys = [
        "pengaturanToko",
        "pengaturan_toko",
        "pengaturan",
        "settings"
    ];

    for (const key of preferredKeys) {
        if (Array.isArray(store?.[key])) {
            return {
                key,
                label: getSettingsSourceLabel(key),
                rows: store[key]
            };
        }

        if (
            store?.[key] &&
            typeof store[key] === "object"
        ) {
            return {
                key,
                label: getSettingsSourceLabel(key),
                rows: [store[key]]
            };
        }
    }

    const detectedKey = Object.keys(store || {}).find(
        (key) =>
            normalizeSettingsKey(key) === "pengaturantoko"
    );

    if (detectedKey) {
        const value = store[detectedKey];

        return {
            key: detectedKey,
            label: getSettingsSourceLabel(detectedKey),
            rows: Array.isArray(value)
                ? value
                : value && typeof value === "object"
                    ? [value]
                    : []
        };
    }

    return {
        key: "pengaturanToko",
        label: "Pengaturan Toko",
        rows: []
    };
}


function normalizeSettingsKey(value) {
    return String(value ?? "")
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
}


function getSettingsSourceLabel(key) {
    return key === "pengaturanToko" ||
        key === "pengaturan_toko"
        ? "Master Pengaturan Toko"
        : key;
}


function setSettingsInput(element, value) {
    if (!element) {
        return;
    }

    const normalized = String(value ?? "").trim();

    if (
        element.tagName === "SELECT" &&
        normalized &&
        !Array.from(element.options).some(
            (option) => option.value === normalized
        )
    ) {
        const option = document.createElement("option");
        option.value = normalized;
        option.textContent = normalized;
        element.appendChild(option);
    }

    element.value = normalized;
}


function getSettingsInput(element) {
    return String(element?.value ?? "").trim();
}


function setSettingsCheckbox(element, value) {
    if (!element) {
        return;
    }

    const normalized = String(value ?? "")
        .trim()
        .toLowerCase();

    element.checked = [
        "ya",
        "yes",
        "true",
        "1",
        "aktif"
    ].includes(normalized);
}


function updateSettingsLogoPreview(value) {
    if (!settingsLogoPreview || !settingsLogoImage) {
        return;
    }

    const logoValue = String(value ?? "").trim();
    const icon = settingsLogoPreview.querySelector("i");

    if (!logoValue) {
        settingsLogoImage.hidden = true;
        settingsLogoImage.removeAttribute("src");
        if (icon) icon.hidden = false;
        return;
    }

    const src = resolveSettingsLogoPath(logoValue);

    settingsLogoImage.onload = () => {
        settingsLogoImage.hidden = false;
        if (icon) icon.hidden = true;
    };

    settingsLogoImage.onerror = () => {
        settingsLogoImage.hidden = true;
        if (icon) icon.hidden = false;
    };

    settingsLogoImage.src = src;
}


function resolveSettingsLogoPath(value) {
    const normalized = String(value ?? "").trim();

    if (
        /^(https?:|data:|blob:|\/|\.\.?\/)/i.test(normalized)
    ) {
        return normalized;
    }

    return `../assets/${normalized}`;
}


function showSettingsNotice(message, type = "success") {
    if (!settingsSaveNotice) {
        return;
    }

    settingsSaveNotice.textContent = message;
    settingsSaveNotice.className =
        `settings-notice settings-notice-${type}`;
    settingsSaveNotice.hidden = false;

    window.clearTimeout(showSettingsNotice.timer);

    showSettingsNotice.timer = window.setTimeout(
        () => {
            settingsSaveNotice.hidden = true;
        },
        3200
    );
}
