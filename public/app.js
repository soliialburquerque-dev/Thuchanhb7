// Application State
let leads = [];
let currentPage = 1;
let totalPages = 1;
let limit = 20;
let totalLeadsCount = 0;
let stats = { total: 0, vip: 0, normal: 0, spam: 0, reviewed: 0 };

// Filters
let searchVal = "";
let statusFilter = "all";
let reviewedFilter = "all";

// Selected Lead
let selectedLead = null;

// DOM Elements
const searchInput = document.getElementById("search-input");
const filterStatus = document.getElementById("filter-status");
const filterReviewed = document.getElementById("filter-reviewed");
const leadsTableBody = document.getElementById("leads-table-body");
const paginationInfo = document.getElementById("pagination-info");
const btnPrevPage = document.getElementById("btn-prev-page");
const btnNextPage = document.getElementById("btn-next-page");
const pageIndicator = document.getElementById("page-indicator");

// Stats elements
const statTotal = document.getElementById("stat-total");
const statVip = document.getElementById("stat-vip");
const statNormal = document.getElementById("stat-normal");
const statSpam = document.getElementById("stat-spam");
const statReviewed = document.getElementById("stat-reviewed");

// Drawer Elements
const leadDrawer = document.getElementById("lead-drawer");
const drawerOverlay = document.getElementById("drawer-overlay");
const drawerBody = document.getElementById("drawer-body");
const btnCloseDrawer = document.getElementById("btn-close-drawer");
const btnCancelDrawer = document.getElementById("btn-cancel-drawer");
const btnSaveDrawer = document.getElementById("btn-save-drawer");

// Settings Modal Elements
const settingsModal = document.getElementById("settings-modal");
const modalOverlay = document.getElementById("modal-overlay");
const btnRecalculateModal = document.getElementById("btn-recalculate-modal");
const btnCloseModal = document.getElementById("btn-close-modal");
const btnCancelModal = document.getElementById("btn-cancel-modal");
const btnRunScoring = document.getElementById("btn-run-scoring");
const inputApiKey = document.getElementById("settings-api-key");
const checkOverwrite = document.getElementById("settings-overwrite");
const btnToggleApiKey = document.getElementById("btn-toggle-api-key");
const scoringStatusBox = document.getElementById("scoring-status");
const scoringStatusText = document.getElementById("scoring-status-text");

// General Action Elements
const btnExport = document.getElementById("btn-export");
const toast = document.getElementById("toast");
const toastMessage = document.getElementById("toast-message");

// API URL (same host)
const API_BASE = "";

// Initialize Application
document.addEventListener("DOMContentLoaded", () => {
    fetchLeads();
    setupEventListeners();
    initIcons();
});

function initIcons() {
    if (window.lucide) {
        window.lucide.createIcons();
    }
}

// Event Listeners
function setupEventListeners() {
    // Search input with debounce
    let debounceTimer;
    searchInput.addEventListener("input", (e) => {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
            searchVal = e.target.value;
            currentPage = 1;
            fetchLeads();
        }, 300);
    });

    // Select filters
    filterStatus.addEventListener("change", (e) => {
        statusFilter = e.target.value;
        currentPage = 1;
        fetchLeads();
    });

    filterReviewed.addEventListener("change", (e) => {
        reviewedFilter = e.target.value;
        currentPage = 1;
        fetchLeads();
    });

    // Metric card filtering
    document.querySelectorAll(".metric-card").forEach(card => {
        card.addEventListener("click", () => {
            const filter = card.getAttribute("data-filter");
            if (filter === "all") {
                filterStatus.value = "all";
                filterReviewed.value = "all";
            } else if (filter === "vip") {
                filterStatus.value = "vip";
                filterReviewed.value = "all";
            } else if (filter === "normal") {
                filterStatus.value = "normal";
                filterReviewed.value = "all";
            } else if (filter === "spam") {
                filterStatus.value = "spam";
                filterReviewed.value = "all";
            } else if (filter === "reviewed") {
                filterStatus.value = "all";
                filterReviewed.value = "true";
            }
            // Trigger change event to fetch data
            statusFilter = filterStatus.value;
            reviewedFilter = filterReviewed.value;
            currentPage = 1;
            fetchLeads();
        });
    });

    // Pagination
    btnPrevPage.addEventListener("click", () => {
        if (currentPage > 1) {
            currentPage--;
            fetchLeads();
        }
    });

    btnNextPage.addEventListener("click", () => {
        if (currentPage < totalPages) {
            currentPage++;
            fetchLeads();
        }
    });

    // Drawer toggles
    btnCloseDrawer.addEventListener("click", closeDrawer);
    btnCancelDrawer.addEventListener("click", closeDrawer);
    drawerOverlay.addEventListener("click", closeDrawer);
    btnSaveDrawer.addEventListener("click", saveLeadChanges);

    // Modal toggles
    btnRecalculateModal.addEventListener("click", openModal);
    btnCloseModal.addEventListener("click", closeModal);
    btnCancelModal.addEventListener("click", closeModal);
    modalOverlay.addEventListener("click", closeModal);
    btnRunScoring.addEventListener("click", runScoringPipeline);

    // Toggle API Key visibility
    btnToggleApiKey.addEventListener("click", () => {
        const type = inputApiKey.getAttribute("type") === "password" ? "text" : "password";
        inputApiKey.setAttribute("type", type);
        const icon = btnToggleApiKey.querySelector("i");
        if (type === "text") {
            icon.setAttribute("data-lucide", "eye-off");
        } else {
            icon.setAttribute("data-lucide", "eye");
        }
        initIcons();
    });

    // Export button
    btnExport.addEventListener("click", exportToExcel);
}

// Fetch Leads Data from Server
async function fetchLeads() {
    showLoading();
    try {
        const query = new URLSearchParams({
            search: searchVal,
            status: statusFilter,
            reviewed: reviewedFilter,
            page: currentPage,
            limit: limit
        });
        const res = await fetch(`${API_BASE}/api/leads?${query.toString()}`);
        if (!res.ok) throw new Error("HTTP error " + res.status);
        
        const data = await res.json();
        
        leads = data.leads;
        totalPages = data.pages;
        currentPage = data.page;
        totalLeadsCount = data.total;
        stats = data.stats;

        renderStats();
        renderTable();
        renderPagination();
    } catch (e) {
        console.error("Fetch error:", e);
        showToast("Không thể tải danh sách khách hàng từ máy chủ", true);
        leadsTableBody.innerHTML = `
            <tr>
                <td colspan="8" class="text-center" style="color: var(--spam-color); padding: 40px 0;">
                    <i data-lucide="alert-triangle" style="width: 32px; height: 32px; margin-bottom: 12px;"></i>
                    <p>Lỗi kết nối máy chủ. Vui lòng kiểm tra và thử lại.</p>
                </td>
            </tr>
        `;
        initIcons();
    }
}

// Loading States
function showLoading() {
    leadsTableBody.innerHTML = `
        <tr>
            <td colspan="8" class="text-center loading-placeholder">
                <div class="spinner"></div>
                <p>Đang tải dữ liệu khách hàng...</p>
            </td>
        </tr>
    `;
}

// Render Stats Headers
function renderStats() {
    statTotal.innerText = stats.total;
    statVip.innerText = stats.vip;
    statNormal.innerText = stats.normal;
    statSpam.innerText = stats.spam;
    statReviewed.innerText = stats.reviewed;
}

// Render Table Rows
function renderTable() {
    if (leads.length === 0) {
        leadsTableBody.innerHTML = `
            <tr>
                <td colspan="8" class="text-center" style="padding: 60px 0; color: var(--text-secondary);">
                    <i data-lucide="info" style="width: 28px; height: 28px; margin-bottom: 12px; opacity: 0.5;"></i>
                    <p>Không tìm thấy khách hàng nào khớp với bộ lọc.</p>
                </td>
            </tr>
        `;
        initIcons();
        return;
    }

    leadsTableBody.innerHTML = "";
    leads.forEach(lead => {
        const tr = document.createElement("tr");
        if (selectedLead && selectedLead.id === lead.id) {
            tr.classList.add("row-selected");
        }

        // Determine Status Badge
        let statusBadge = "";
        if (lead.status === "VIP") {
            statusBadge = `<span class="badge badge-vip"><i data-lucide="shield-check" style="width:12px;height:12px;"></i> VIP</span>`;
        } else if (lead.status === "SPAM") {
            statusBadge = `<span class="badge badge-spam"><i data-lucide="shield-alert" style="width:12px;height:12px;"></i> SPAM</span>`;
        } else {
            statusBadge = `<span class="badge badge-normal"><i data-lucide="user-check" style="width:12px;height:12px;"></i> NORMAL</span>`;
        }

        // Determine Review Badge
        const reviewedBadge = lead.reviewed 
            ? `<span class="badge badge-reviewed"><i data-lucide="check" style="width:12px;height:12px;"></i> Đã Duyệt</span>`
            : `<span class="badge badge-unreviewed"><i data-lucide="clock" style="width:12px;height:12px;"></i> Chờ</span>`;

        tr.innerHTML = `
            <td class="text-center" style="font-family: 'JetBrains Mono', monospace; font-weight: 500;">${lead.id}</td>
            <td style="font-weight: 600;">${lead.ten_khach}</td>
            <td class="text-center" style="font-family: 'JetBrains Mono', monospace;">${lead.sdt}</td>
            <td><div class="demand-text" title="${lead.nhu_cau_mo_ta}">${lead.nhu_cau_mo_ta}</div></td>
            <td class="text-center"><span class="score-badge" style="color: ${getScoreColor(lead.score)}">${lead.score}đ</span></td>
            <td class="text-center">${statusBadge}</td>
            <td class="text-center">${reviewedBadge}</td>
            <td class="text-center">
                <button class="btn btn-secondary btn-icon btn-view-lead" style="width:28px; height:28px;">
                    <i data-lucide="eye" style="width:14px; height:14px;"></i>
                </button>
            </td>
        `;

        // Click row or eye icon to view details
        tr.addEventListener("click", () => openDrawer(lead));
        leadsTableBody.appendChild(tr);
    });

    initIcons();
}

// Color scale based on score
function getScoreColor(score) {
    if (score >= 80) return "var(--vip-color)";
    if (score < 30) return "var(--spam-color)";
    return "var(--normal-color)";
}

// Render Pagination Information
function renderPagination() {
    const start = totalLeadsCount === 0 ? 0 : (currentPage - 1) * limit + 1;
    const end = Math.min(currentPage * limit, totalLeadsCount);
    
    paginationInfo.innerText = `Hiển thị ${start} - ${end} của ${totalLeadsCount} khách hàng`;
    pageIndicator.innerText = `Trang ${currentPage} / ${totalPages || 1}`;

    btnPrevPage.disabled = currentPage <= 1;
    btnNextPage.disabled = currentPage >= totalPages;
}

// Detail Drawer Management
function openDrawer(lead) {
    selectedLead = lead;
    
    // Highlight row
    document.querySelectorAll("#leads-table-body tr").forEach((tr, index) => {
        if (leads[index] && leads[index].id === lead.id) {
            tr.classList.add("row-selected");
        } else {
            tr.classList.remove("row-selected");
        }
    });

    // Populate drawer body
    let reasonsHtml = "";
    if (lead.reason) {
        const parts = lead.reason.split(";");
        parts.forEach(part => {
            part = part.trim();
            if (!part) return;
            let cls = "neutral";
            if (part.startsWith("+")) cls = "positive";
            if (part.startsWith("-")) cls = "negative";
            reasonsHtml += `<span class="reason-tag ${cls}">${part}</span>`;
        });
    }

    drawerBody.innerHTML = `
        <div class="detail-item">
            <span class="detail-label">Khách hàng</span>
            <span class="detail-value" style="font-weight: 700; font-size: 18px;">${lead.ten_khach}</span>
        </div>
        <div class="form-row-2">
            <div class="detail-item">
                <span class="detail-label">Số điện thoại</span>
                <span class="detail-value" style="font-family: 'JetBrains Mono', monospace;">${lead.sdt}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">Mã số</span>
                <span class="detail-value" style="font-family: 'JetBrains Mono', monospace;">#${lead.id}</span>
            </div>
        </div>
        <div class="detail-item">
            <span class="detail-label">Nhu cầu chi tiết</span>
            <p class="detail-value card-desc">${lead.nhu_cau_mo_ta}</p>
        </div>
        <div class="detail-item">
            <span class="detail-label">Lý do chấm điểm hệ thống</span>
            <div class="reason-tags">
                ${reasonsHtml || '<span class="reason-tag">Không có phân tích chi tiết</span>'}
            </div>
        </div>
        <hr style="border: 0; border-top: 1px solid var(--border-color);">
        <h3 style="font-family:'Outfit',sans-serif; font-size:16px;">Đánh giá & Điều chỉnh (Human-in-the-loop)</h3>
        
        <div class="form-row-2">
            <div class="form-group">
                <label for="drawer-score">Điểm số (0-100):</label>
                <input type="number" id="drawer-score" class="glass-select" min="0" max="100" value="${lead.score}">
            </div>
            <div class="form-group">
                <label for="drawer-status">Trạng thái:</label>
                <select id="drawer-status" class="glass-select">
                    <option value="VIP" ${lead.status === 'VIP' ? 'selected' : ''}>VIP / Siêu tiềm năng</option>
                    <option value="NORMAL" ${lead.status === 'NORMAL' ? 'selected' : ''}>NORMAL / Trung bình</option>
                    <option value="SPAM" ${lead.status === 'SPAM' ? 'selected' : ''}>SPAM / Rác</option>
                </select>
            </div>
        </div>

        <div class="form-group">
            <label for="drawer-notes">Ghi chú kiểm duyệt:</label>
            <textarea id="drawer-notes" class="glass-textarea" placeholder="Nhập ghi chú hoặc lý do thay đổi...">${lead.notes || ''}</textarea>
        </div>

        <div class="form-group checkbox-group">
            <input type="checkbox" id="drawer-reviewed" ${lead.reviewed ? 'checked' : ''}>
            <label for="drawer-reviewed" style="cursor:pointer; font-weight:600;">Đánh dấu đã kiểm duyệt kết quả</label>
        </div>
    `;

    // Dynamic state synchronizer (changing score auto-selects status)
    const scoreInput = document.getElementById("drawer-score");
    const statusSelect = document.getElementById("drawer-status");
    scoreInput.addEventListener("input", (e) => {
        let val = parseInt(e.target.value) || 0;
        val = Math.max(0, Math.min(100, val));
        e.target.value = val;
        
        if (val >= 80) {
            statusSelect.value = "VIP";
        } else if (val < 30) {
            statusSelect.value = "SPAM";
        } else {
            statusSelect.value = "NORMAL";
        }
    });

    leadDrawer.classList.add("active");
    drawerOverlay.classList.add("active");
    initIcons();
}

function closeDrawer() {
    leadDrawer.classList.remove("active");
    drawerOverlay.classList.remove("active");
    selectedLead = null;
    document.querySelectorAll("#leads-table-body tr").forEach(tr => tr.classList.remove("row-selected"));
}

// Save Manual Human-in-the-loop modifications
async function saveLeadChanges() {
    if (!selectedLead) return;

    const newScore = parseInt(document.getElementById("drawer-score").value) || 0;
    const newStatus = document.getElementById("drawer-status").value;
    const newNotes = document.getElementById("drawer-notes").value;
    const isReviewed = document.getElementById("drawer-reviewed").checked;

    btnSaveDrawer.disabled = true;
    btnSaveDrawer.innerText = "Đang lưu...";

    try {
        const res = await fetch(`${API_BASE}/api/leads/update`, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                id: selectedLead.id,
                score: newScore,
                status: newStatus,
                notes: newNotes,
                reviewed: isReviewed
            })
        });

        if (!res.ok) throw new Error("HTTP error " + res.status);
        const data = await res.json();
        
        if (data.success) {
            showToast("Đã cập nhật trạng thái khách hàng thành công!");
            closeDrawer();
            fetchLeads(); // Refresh current view
        } else {
            throw new Error(data.error || "Không rõ nguyên nhân");
        }
    } catch (e) {
        console.error("Save error:", e);
        showToast("Không thể lưu thay đổi: " + e.message, true);
    } finally {
        btnSaveDrawer.disabled = false;
        btnSaveDrawer.innerText = "Lưu Kết Quả";
    }
}

// Settings Modal Management
function openModal() {
    // Read cached API Key if exists
    const savedKey = localStorage.getItem("gemini_api_key") || "";
    inputApiKey.value = savedKey;
    
    settingsModal.classList.add("active");
    modalOverlay.classList.add("active");
    scoringStatusBox.classList.add("hidden");
    btnRunScoring.disabled = false;
    btnRunScoring.innerText = "Bắt Đầu Chấm Điểm";
}

function closeModal() {
    settingsModal.classList.remove("active");
    modalOverlay.classList.remove("active");
}

// Recalculate leads (Keyword match / Gemini)
async function runScoringPipeline() {
    const key = inputApiKey.value.trim();
    const overwrite = checkOverwrite.checked;

    // Cache API Key in local storage
    if (key) {
        localStorage.setItem("gemini_api_key", key);
    } else {
        localStorage.removeItem("gemini_api_key");
    }

    btnRunScoring.disabled = true;
    btnRunScoring.innerText = "Đang chấm...";
    scoringStatusBox.classList.remove("hidden");
    scoringStatusText.innerText = key ? "Đang truy vấn mô hình AI Gemini..." : "Đang phân tích bộ quy tắc từ khóa...";

    try {
        const res = await fetch(`${API_BASE}/api/leads/recalculate`, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                api_key: key || null,
                overwrite: overwrite
            })
        });

        if (!res.ok) throw new Error("HTTP error " + res.status);
        const data = await res.json();

        if (data.success) {
            showToast("Tiến trình chấm điểm đã hoàn tất thành công!");
            closeModal();
            currentPage = 1;
            fetchLeads();
        } else {
            throw new Error(data.error || "Không thể chạy tiến trình");
        }
    } catch (e) {
        console.error("Scoring error:", e);
        showToast("Lỗi chấm điểm: " + e.message, true);
    } finally {
        btnRunScoring.disabled = false;
        btnRunScoring.innerText = "Bắt Đầu Chấm Điểm";
        scoringStatusBox.classList.add("hidden");
    }
}

// Export data to Excel
async function exportToExcel() {
    btnExport.disabled = true;
    const oldText = btnExport.innerHTML;
    btnExport.innerHTML = `<span class="spinner spinner-small" style="margin:0;"></span> Đang xuất...`;

    try {
        const res = await fetch(`${API_BASE}/api/leads/export`, {
            method: "POST"
        });

        if (!res.ok) throw new Error("HTTP error " + res.status);
        
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "Danh_sach_khach_hang_tiem_nang.xlsx";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        showToast("Xuất Excel thành công!");
    } catch (e) {
        console.error("Export error:", e);
        showToast("Xuất Excel thất bại: " + e.message, true);
    } finally {
        btnExport.disabled = false;
        btnExport.innerHTML = oldText;
    }
}

// Toast notification handler
let toastTimer;
function showToast(message, isError = false) {
    clearTimeout(toastTimer);
    toastMessage.innerText = message;
    
    const icon = document.getElementById("toast-icon");
    if (isError) {
        toast.classList.add("toast-error");
        icon.setAttribute("data-lucide", "alert-circle");
    } else {
        toast.classList.remove("toast-error");
        icon.setAttribute("data-lucide", "check-circle");
    }
    initIcons();

    toast.classList.remove("hidden");
    
    toastTimer = setTimeout(() => {
        toast.classList.add("hidden");
    }, 4000);
}
