/**
 * نظام منشر - دوال مساعدة مشتركة
 */

const PRODUCTS = ['المطور اسود (الحديدي)', 'المطور ابيض (الحديدي)', 'الذكي اسود (الصندوق)', 'الذكي ابيض (الصندوق)', 'بارات الومنيوم'];

/** إظهار اللودينج */
function showLoading(text) {
  const el = document.getElementById('loadingOverlay');
  if (el) { el.classList.add('active'); }
}

/** إخفاء اللودينج */
function hideLoading() {
  const el = document.getElementById('loadingOverlay');
  if (el) { el.classList.remove('active'); }
}

/** Toast notification */
function showToast(message, type = 'success') {
  let container = document.getElementById('toastContainer');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toastContainer';
    container.className = 'toast-container';
    document.body.appendChild(container);
  }

  const toast = document.createElement('div');
  toast.className = 'toast' + (type === 'error' ? ' error' : type === 'warning' ? ' warning' : '');
  const icon = type === 'error' ? '❌' : type === 'warning' ? '⚠️' : '✅';
  toast.innerHTML = `<span>${icon}</span><span>${message}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.animation = 'toastOut 0.3s ease forwards';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

/** تنسيق التاريخ */
function formatDate(date) {
  if (!date) return '-';
  const d = new Date(date);
  if (isNaN(d.getTime())) return String(date);
  return d.toLocaleDateString('ar-SA', { year: 'numeric', month: 'short', day: 'numeric' });
}

/** تنسيق الرقم */
function formatNumber(num) {
  return Number(num || 0).toLocaleString('ar-SA');
}

/** تبديل تبويب */
function switchTab(tabId) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  
  const content = document.getElementById(tabId);
  if (content) content.classList.add('active');
  
  const btn = document.querySelector(`.tab-btn[onclick*="${tabId}"]`);
  if (btn) btn.classList.add('active');
}

/** إنشاء هيدر الصفحة */
function renderHeader(title, subtitle) {
  const session = Auth.getSession();
  return `
    <header class="app-header">
      <div class="header-left">
        <img src="assets/logo.png" alt="منشر" class="logo"
             onerror="this.style.display='none'">
        <div>
          <div class="title">${title}</div>
          <div class="subtitle">${subtitle || 'نظام إدارة المخزون'}</div>
        </div>
      </div>
      <div class="header-right">
        <span class="user-badge">👤 ${session ? session.displayName : ''}</span>
        <button class="btn-logout" onclick="Auth.logout()">تسجيل خروج</button>
      </div>
    </header>
  `;
}

/** إنشاء loading overlay */
function renderLoadingOverlay() {
  return `
    <div id="loadingOverlay" class="loading-overlay">
      <div class="spinner"></div>
      <div class="loading-text">جاري التحميل...</div>
      <div class="loading-sub">يرجى عدم إغلاق الصفحة</div>
    </div>
  `;
}

/** إنشاء حالة فارغة */
function renderEmpty(text = 'لا توجد بيانات') {
  return `
    <div class="empty-state">
      <div class="empty-icon">📭</div>
      <div class="empty-text">${text}</div>
    </div>
  `;
}
