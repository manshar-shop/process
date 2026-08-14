/**
 * نظام منشر - التواصل مع الباك إند
 */

const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbyPmlxRLYq1O0pBG4Z8NW0JS1jKwzcCeETCoVQR5hO24C4xSo8ykkgY0EOXnGnx_9X0/exec';

const API = {
  /** إرسال طلب POST */
  async post(action, data = {}) {
    showLoading();
    try {
      data.action = action;
      const session = Auth.getSession();
      if (session) data.user = session.displayName;

      const res = await fetch(WEB_APP_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(data),
        redirect: 'follow'
      });
      const result = await res.json();
      if (!result.ok && result.error) showToast(result.error, 'error');
      return result;
    } catch (err) {
      showToast('خطأ في الاتصال: ' + err.message, 'error');
      return { ok: false, error: err.message };
    } finally {
      hideLoading();
    }
  },

  /** إرسال طلب GET */
  async get(action, params = {}) {
    try {
      const session = Auth.getSession();
      const query = new URLSearchParams({ action, ...params });
      if (session) query.set('user', session.displayName);

      const res = await fetch(WEB_APP_URL + '?' + query.toString(), { redirect: 'follow' });
      return await res.json();
    } catch (err) {
      console.error('API GET error:', err);
      return { ok: false, error: err.message };
    }
  },

  /** GET بدون إخفاء - للتحديث في الخلفية */
  async silentGet(action, params = {}) {
    try {
      const query = new URLSearchParams({ action, ...params });
      const res = await fetch(WEB_APP_URL + '?' + query.toString(), { redirect: 'follow' });
      return await res.json();
    } catch (err) {
      return { ok: false, error: err.message };
    }
  }
};
