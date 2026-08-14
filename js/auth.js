/**
 * نظام منشر - نظام تسجيل الدخول وحماية الصفحات
 */

const Auth = {
  SESSION_KEY: 'manshar_session',

  /** حفظ الجلسة */
  save(userData) {
    const session = {
      ...userData,
      loginTime: Date.now()
    };
    localStorage.setItem(this.SESSION_KEY, JSON.stringify(session));
  },

  /** جلب الجلسة */
  getSession() {
    try {
      const raw = localStorage.getItem(this.SESSION_KEY);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch { return null; }
  },

  /** تسجيل الخروج */
  logout() {
    localStorage.removeItem(this.SESSION_KEY);
    window.location.href = 'index.html';
  },

  /** التحقق من الصلاحية - يُستدعى في بداية كل صفحة */
  requireRole(role) {
    const session = this.getSession();
    if (!session) {
      window.location.href = 'index.html';
      return null;
    }
    // admin يقدر يدخل أي صفحة
    if (session.role !== role && session.role !== 'admin') {
      window.location.href = 'index.html';
      return null;
    }
    return session;
  },

  /** الصفحة حسب الصلاحية */
  getPageForRole(role) {
    const pages = {
      admin: 'admin.html',
      factory: 'factory.html',
      distribution: 'distribution.html',
      courier: 'courier.html'
    };
    return pages[role] || 'index.html';
  }
};
