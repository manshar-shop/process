/**
 * ===================================================
 *  نظام منشر المتكامل - v2 (Transaction-Based)
 *  Google Apps Script + Google Sheets
 * ===================================================
 *
 *  كل عملية = سطر حركة في جدول حركة_المخزون
 *  الأرصدة تُحسب دائماً من مجموع الحركات
 *
 *  الإعداد:
 *  1. افتح Google Sheet → Extensions → Apps Script
 *  2. ضع هذا الكود بالكامل
 *  3. Deploy → Web App (Execute as: Me, Access: Anyone)
 *  4. انسخ الرابط واستخدمه في ملفات HTML
 */

// ==========================================
// الثوابت
// ==========================================

const PRODUCTS = ['المطور اسود (الحديدي)', 'المطور ابيض (الحديدي)', 'الذكي اسود (الصندوق)', 'الذكي ابيض (الصندوق)', 'بارات الومنيوم'];

const DB_SHEET_ID = '1LBr9OE9u6wQp-_bKUjZPHcU-dWwroJZ_g2DY2WYOQkM';
const ORDERS_SHEET_ID = '15hbK2BOumyrd_1ZvI0Bew8q2gAR3fpVOHgC0rcisNrE';

const SHEET_NAMES = {
  USERS: 'المستخدمين',
  TRANSACTIONS: 'حركة_المخزون',
  NEEDS: 'الاحتياجات',
  COLLECTIONS: 'التحصيل',
  ERROR_LOG: 'سجل_الأخطاء',
  FINANCIAL_REVIEW: 'مراجعة_المالية',
  REGION_CUSTODY: 'عهدة_المنطقة'
};

const TX_TYPES = {
  INIT_STOCK: 'init_stock',
  FACTORY_DELIVERY: 'factory_delivery',
  SORT_TO_DELEGATE: 'sort_to_delegate',
  EXTRA_CUSTODY: 'extra_custody',
  INSTALLATION: 'installation',
  EXTRA_INSTALLATION: 'extra_installation',
  RETURN: 'return_stock',
  DAMAGED: 'damaged',
  ADJUSTMENT: 'adjustment',
  FACTORY_ORDER: 'factory_order',
  ORDER_ASSIGNMENT: 'order_assignment',
  DAMAGE_RECORD: 'damage_record',
  REGION_CUSTODY_ADD: 'region_custody_add',
  REGION_CUSTODY_USE: 'region_custody_use', 
  REGION_CUSTODY_RETURN: 'region_custody_return',
  BAR_AUTO_DEDUCTION: 'bar_auto_deduction'
};

const DEFAULT_REGIONS = ['الغربيه', 'الشرقية', 'المدينة'];
const BOX_PRODUCTS = ['الذكي اسود (الصندوق)', 'الذكي ابيض (الصندوق)'];
const BARS_PER_UNIT = 1;

function getCustodyRegions_() {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName('المناطق');
    if (!sheet || sheet.getLastRow() < 2) return DEFAULT_REGIONS;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    const activeRegions = data.filter(r => String(r[1]).trim() !== 'معطلة').map(r => String(r[0]).trim());
    return activeRegions.length > 0 ? activeRegions : DEFAULT_REGIONS;
  } catch(e) {
    return DEFAULT_REGIONS;
  }
}

// ==========================================
// نقاط الدخول
// ==========================================

/** 
 * تشغيل هذه الدالة يدوياً مرة واحدة من المحرر لإنشاء الجداول وإعطاء الصلاحيات 
 */
function setup() {
  initSheets_();
  SpreadsheetApp.getUi().alert('تم إنشاء الجداول بنجاح!');
}


function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    initSheets_();
    const action = data.action;

    switch (action) {
      case 'add_region':             return json_(handleAddRegion_(data));
      case 'update_region':          return json_(handleUpdateRegion_(data));
      case 'toggle_region':          return json_(handleToggleRegion_(data));
      case 'delete_region':          return json_(handleDeleteRegion_(data));
      case 'delete_delegate':        return json_(handleDeleteDelegate_(data));
      case 'login':                  return json_(handleLogin_(data));
      case 'update_password':        return json_(handleUpdatePassword_(data));
      case 'add_delegate':           return json_(handleAddDelegate_(data));
      case 'init_stock':             return json_(handleInitStock_(data));
      case 'update_inventory':       return json_(handleUpdateInventory_(data));
      case 'create_need':            return json_(handleCreateNeed_(data));
      case 'factory_delivery':       return json_(handleFactoryDelivery_(data));
      case 'sort_to_delegate':       return json_(handleSortToDelegate_(data));
      case 'extra_custody':          return json_(handleExtraCustody_(data));
      case 'register_installation':  return json_(handleInstallation_(data));
      case 'extra_installation':     return json_(handleExtraInstallation_(data));
      case 'register_return':        return json_(handleReturn_(data));
      case 'register_damaged':       return json_(handleDamaged_(data));
      case 'clear_damaged':          return json_(handleClearDamaged_(data));
      case 'register_collection':    return json_(handleCollection_(data));
      case 'factory_order':          return json_(handleFactoryOrder_(data));
      case 'order_assignment':       return json_(handleOrderAssignment_(data));
      case 'register_damage_record': return json_(handleDamageRecord_(data));
      case 'confirm_financial_dist':    return json_(handleDistFinancialConfirm_(data));
      case 'confirm_financial_factory': return json_(handleFactoryFinancialConfirm_(data));
      case 'region_custody_add':       return json_(handleRegionCustodyAdd_(data));
      case 'region_custody_return':    return json_(handleRegionCustodyReturn_(data));
      default: return json_({ ok: false, error: 'عملية غير معروفة: ' + action });
    }
  } catch (err) {
    return json_({ ok: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    initSheets_();
    const action = e.parameter.action || 'ping';

    switch (action) {
      case 'regions':                return json_(getRegions_());
      case 'all_couriers':           return json_({ ok: true, data: getAllCouriersIncludeDeleted_() });
      case 'ping':                return json_({ ok: true, message: 'نظام منشر يعمل' });
      case 'users':               return json_(getUsers_());
      case 'inventory':           return json_(getInventoryBalance_());
      case 'delegate_balance':    return json_(getDelegateBalance_(e.parameter.delegate));
      case 'all_delegates':       return json_(getAllDelegatesBalance_());
      case 'factory_needs':       return json_(getFactoryNeeds_());
      case 'damaged_summary':     return json_(getDamagedSummary_());
      case 'dashboard':           return json_(getDashboard_(e.parameter.role, e.parameter.user));
      case 'delegate_report':     return json_(getDelegateReport_(e.parameter.delegate));
      case 'transactions':        return json_(getTransactions_(e.parameter));
      case 'validate_order':      return json_(validateOrder_(e.parameter.order, e.parameter.delegate));
      case 'courier_stats':       return json_(getCourierStats_(e.parameter.delegate));
      case 'products':            return json_({ ok: true, data: PRODUCTS });
      case 'collections':         return json_(getCollections_());
      case 'factory_overview':    return json_(getFactoryOverview_());
      case 'factory_orders':      return json_(getFactoryOrders_());
      case 'order_assignments':   return json_(getOrderAssignments_());
      case 'financial_pending_dist':    return json_(getFinancialPendingDist_());
      case 'financial_pending_factory': return json_(getFinancialPendingFactory_());
      case 'financial_all':             return json_(getFinancialAll_());
      case 'extra_custody_stats':       return json_(getExtraCustodyStats_());
      case 'region_custody_dashboard':  return json_(getRegionCustodyDashboard_());
      case 'region_custody_detail':     return json_(getRegionCustodyDetail_(e.parameter.region));
      case 'courier_region_usage':      return json_(getCourierRegionUsage_());
      case 'factory_inventory_dashboard': return json_(getFactoryInventoryDashboard_());
      case 'bars_config':               return json_({ ok: true, data: { barsPerUnit: BARS_PER_UNIT, boxProducts: BOX_PRODUCTS, custodyRegions: getCustodyRegions_() } });
      default: return json_({ ok: false, error: 'طلب غير معروف: ' + action });
    }
  } catch (err) {
    return json_({ ok: false, error: err.toString() });
  }
}

// ==========================================
// تهيئة الأوراق
// ==========================================

function initSheets_() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);

  if (!ss.getSheetByName(SHEET_NAMES.USERS)) {
    const s = ss.insertSheet(SHEET_NAMES.USERS);
    s.appendRow(['القسم', 'رمز_الدخول', 'الصلاحية']);
    s.getRange('A1:C1').setFontWeight('bold').setBackground('#3D271D').setFontColor('#fff');
    s.appendRow(['الإدارة', '111', 'admin']);
    s.appendRow(['المصنع', '333', 'factory']);
    s.appendRow(['الفرز', '222', 'distribution']);
    s.appendRow(['مندوب', '123', 'courier']);
  }

  if (!ss.getSheetByName('المناديب')) {
    const s = ss.insertSheet('المناديب');
    s.appendRow(['الاسم', 'المنطقة']);
    s.getRange('A1:B1').setFontWeight('bold').setBackground('#3D271D').setFontColor('#fff');
    const defaultCouriers = [
      ['صالح جبران', 'الرياض'], ['فتحي الشجاع', 'الرياض'], ['ابراهيم دحيمان', 'الرياض'],
      ['صالح سالم', 'الرياض'], ['محمد الزنن', 'الغربيه'], ['يذيد', 'الرياض'],
      ['أنيس باكر', 'الغربيه'], ['ابو كيان', 'تبوك'], ['امجد الحداد', 'الرياض'],
      ['رضوان الحداد', 'الجنوبية'], ['محمد نجيب', 'الغربيه'], ['محمد نجيب (الرياض)', 'الرياض'],
      ['أحمد معان', 'الغربيه'], ['علاء', 'الغربيه'], ['عمر الزبيدي', 'الرياض'],
      ['سامي عمر', 'حائل'], ['جميل محيرز', 'الرياض'], ['ابو عامر', 'الرياض'],
      ['رامي', 'الشرقية'], ['سالم', 'الشرقية'], ['ناصر', 'الشرقية'],
      ['صالح اليافعي', 'الشرقية'], ['سالم نيازي', 'الشرقية'], ['عمر هادي', 'المدينة'],
      ['علي عبد الرحمن', 'الشرقية']
    ];
    s.getRange(2, 1, defaultCouriers.length, 2).setValues(defaultCouriers);
  }

  if (!ss.getSheetByName('المناطق')) {
    const s = ss.insertSheet('المناطق');
    s.getRange(1, 1, 1, 3).setValues([['الاسم', 'الحالة', 'تاريخ_الإضافة']]).setFontWeight('bold').setBackground('#3D271D').setFontColor('white');
    // إضافة المناطق الافتراضية
    DEFAULT_REGIONS.forEach(r => s.appendRow([r, 'نشطة', new Date()]));
  }

  if (!ss.getSheetByName(SHEET_NAMES.TRANSACTIONS)) {
    const s = ss.insertSheet(SHEET_NAMES.TRANSACTIONS);
    s.appendRow(['id', 'التاريخ', 'نوع_الحركة', 'المنتج', 'الكمية', 'من', 'إلى', 'رقم_الطلب', 'ملاحظات', 'المستخدم']);
    s.getRange('A1:J1').setFontWeight('bold').setBackground('#3D271D').setFontColor('#fff');
  }

  if (!ss.getSheetByName(SHEET_NAMES.NEEDS)) {
    const s = ss.insertSheet(SHEET_NAMES.NEEDS);
    s.appendRow(['id', 'التاريخ', 'المنتج', 'الكمية_المطلوبة', 'الكمية_المستلمة', 'المتبقي', 'الحالة', 'أنشأها']);
    s.getRange('A1:H1').setFontWeight('bold').setBackground('#3D271D').setFontColor('#fff');
  }

  if (!ss.getSheetByName(SHEET_NAMES.COLLECTIONS)) {
    const s = ss.insertSheet(SHEET_NAMES.COLLECTIONS);
    s.appendRow(['التاريخ', 'رقم_الطلب', 'المندوب', 'المبلغ', 'الطريقة', 'ملاحظات']);
    s.getRange('A1:F1').setFontWeight('bold').setBackground('#3D271D').setFontColor('#fff');
  }

  if (!ss.getSheetByName(SHEET_NAMES.ERROR_LOG)) {
    const s = ss.insertSheet(SHEET_NAMES.ERROR_LOG);
    s.appendRow(['التاريخ', 'المندوب', 'رقم_الطلب', 'السبب']);
    s.getRange('A1:D1').setFontWeight('bold').setBackground('#3D271D').setFontColor('#fff');
  }

  if (!ss.getSheetByName(SHEET_NAMES.FINANCIAL_REVIEW)) {
    const s = ss.insertSheet(SHEET_NAMES.FINANCIAL_REVIEW);
    s.appendRow(['id', 'التاريخ', 'رقم_الطلب', 'المندوب', 'المبلغ', 'طريقة_الدفع', 'النوع', 'حالة_الفرز', 'مراجع_الفرز', 'تاريخ_الفرز', 'حالة_المصنع', 'مراجع_المصنع', 'تاريخ_المصنع']);
    s.getRange('A1:M1').setFontWeight('bold').setBackground('#3D271D').setFontColor('#fff');
  }

  if (!ss.getSheetByName(SHEET_NAMES.REGION_CUSTODY)) {
    const s = ss.insertSheet(SHEET_NAMES.REGION_CUSTODY);
    s.appendRow(['id', 'التاريخ', 'المنطقة', 'نوع_الحركة', 'المنتج', 'الكمية', 'المندوب', 'رقم_الطلب', 'الرصيد_قبل', 'الرصيد_بعد', 'ملاحظات', 'المستخدم']);
    s.getRange('A1:L1').setFontWeight('bold').setBackground('#3D271D').setFontColor('#fff');
  }
}

// ==========================================
// دوال مساعدة
// ==========================================

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** إضافة حركة مخزون */
function addTransaction_(type, product, qty, from, to, orderNum, notes, userName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
    const id = 'TX-' + Date.now();
    sheet.appendRow([
      id,
      new Date(),
      type,
      product,
      Number(qty),
      from || '',
      to || '',
      orderNum || '',
      notes || '',
      userName || ''
    ]);
    return id;
  } finally {
    lock.releaseLock();
  }
}

/** قراءة كل الحركات (cached per execution) */
let _txCache = null;
function getAllTransactions_() {
  if (_txCache) return _txCache;
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  if (!sheet || sheet.getLastRow() < 2) { _txCache = []; return []; }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  _txCache = data.map(r => ({
    id: r[0], date: r[1], type: r[2], product: r[3],
    qty: Number(r[4]), from: r[5], to: r[6],
    orderNumber: r[7], notes: r[8], user: r[9]
  }));
  return _txCache;
}

/** حساب رصيد المخزون الرئيسي لمنتج */
function mainBalance_(product) {
  const txs = getAllTransactions_();
  let balance = 0;
  txs.forEach(t => {
    if (t.product !== product) return;
    if (t.type === TX_TYPES.INIT_STOCK)        balance += t.qty;
    if (t.type === TX_TYPES.FACTORY_DELIVERY)   balance += t.qty;
    if (t.type === TX_TYPES.RETURN)             balance += t.qty;
    if (t.type === TX_TYPES.SORT_TO_DELEGATE)   balance -= t.qty;
    if (t.type === TX_TYPES.EXTRA_CUSTODY)      balance -= t.qty;
    if (t.type === TX_TYPES.ORDER_ASSIGNMENT)  balance -= t.qty;
    if (t.type === TX_TYPES.REGION_CUSTODY_ADD)  balance -= t.qty;
    if (t.type === TX_TYPES.BAR_AUTO_DEDUCTION)  balance -= t.qty;
  });
  return balance;
}

/** حساب عهدة المندوب (طلبات) لمنتج */
function delegateOrderBalance_(delegate, product) {
  const txs = getAllTransactions_();
  let balance = 0;
  txs.forEach(t => {
    if (t.product !== product) return;
    if (t.type === TX_TYPES.SORT_TO_DELEGATE && t.to === delegate) balance += t.qty;
    if (t.type === TX_TYPES.INSTALLATION && t.from === delegate)    balance -= t.qty;
    if (t.type === TX_TYPES.RETURN && t.from === delegate)          balance -= t.qty;
    if (t.type === TX_TYPES.DAMAGED && t.from === delegate)         balance -= t.qty;
  });
  return balance;
}

/** حساب عهدة إضافية لمنتج */
function delegateExtraBalance_(delegate, product) {
  const txs = getAllTransactions_();
  let balance = 0;
  txs.forEach(t => {
    if (t.product !== product) return;
    if (t.type === TX_TYPES.EXTRA_CUSTODY && t.to === delegate)  balance += t.qty;
    if (t.type === TX_TYPES.EXTRA_INSTALLATION && t.from === delegate) balance -= t.qty;
  });
  return balance;
}

/** إجمالي عهدة المندوب (طلبات + إضافية) */
function delegateTotalBalance_(delegate, product) {
  return delegateOrderBalance_(delegate, product) + delegateExtraBalance_(delegate, product);
}

/** جلب كل المناديب */
function getCouriers_() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName('المناديب');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const cols = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(cols, 3)).getValues();
  return data.map(r => ({ displayName: String(r[0]).trim(), region: String(r[1]).trim(), status: String(r[2] || '').trim() }))
    .filter(c => c.displayName && c.status !== 'محذوف');
}

function getAllCouriersIncludeDeleted_() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName('المناديب');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const cols = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(cols, 3)).getValues();
  return data.map((r, i) => ({ index: i, displayName: String(r[0]).trim(), region: String(r[1]).trim(), status: String(r[2] || 'نشط').trim() }))
    .filter(c => c.displayName);
}

/** حذف مندوب (Soft Delete) */
function handleDeleteDelegate_(data) {
  const delegateName = String(data.name || '').trim();
  if (!delegateName) return { ok: false, error: 'اسم المندوب مطلوب' };
  
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName('المناديب');
    if (!sheet) return { ok: false, error: 'جدول المناديب غير موجود' };
    
    const data_rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < data_rows.length; i++) {
      if (String(data_rows[i][0]).trim() === delegateName) {
        // Ensure column 3 exists
        if (sheet.getLastColumn() < 3) {
          sheet.getRange(1, 3).setValue('الحالة').setFontWeight('bold').setBackground('#3D271D').setFontColor('white');
        }
        sheet.getRange(i + 2, 3).setValue('محذوف');
        return { ok: true, message: 'تم حذف المندوب بنجاح (الحذف الناعم)' };
      }
    }
    return { ok: false, error: 'المندوب غير موجود' };
  } catch(err) {
    return { ok: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// تسجيل الدخول
// ==========================================

function handleLogin_(data) {
  const user = String(data.username).trim().toLowerCase();
  
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  
  if (sheet && sheet.getLastRow() > 1) {
    const usersData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (let i = 0; i < usersData.length; i++) {
      if (String(usersData[i][1]).trim().toLowerCase() === user) {
         return { ok: true, data: { username: usersData[i][2], role: usersData[i][2], displayName: usersData[i][0] } };
      }
    }
  } else {
    if (user === '111') {
      return { ok: true, data: { username: 'admin', role: 'admin', displayName: 'الإدارة' } };
    }
    if (user === '333') {
      return { ok: true, data: { username: 'factory', role: 'factory', displayName: 'المصنع' } };
    }
    if (user === '222') {
      return { ok: true, data: { username: 'distribution', role: 'distribution', displayName: 'الفرز' } };
    }
    if (user === '123') {
      return { ok: true, data: { username: 'courier', role: 'courier', displayName: 'مندوب' } };
    }
  }
  
  return { ok: false, error: 'رمز الدخول غير صحيح. الرجاء التأكد من الرقم.' };
}

// ==========================================
// عمليات الكتابة (POST)
// ==========================================

/** تسجيل مخزون أساسي (الإدارة) */
function handleInitStock_(data) {
  const product = data.product;
  const qty = Number(data.quantity);
  if (!PRODUCTS.includes(product)) return { ok: false, error: 'منتج غير صالح' };
  if (qty <= 0) return { ok: false, error: 'الكمية يجب أن تكون أكبر من صفر' };

  _txCache = null;
  addTransaction_(TX_TYPES.INIT_STOCK, product, qty, '', 'المخزون', '', 'تسجيل مخزون أساسي', data.user || 'admin');
  return { ok: true, message: 'تم تسجيل المخزون' };
}

/** تعديل المخزون الأساسي بشكل كلي (الفرق) */
function handleUpdateInventory_(data) {
  const items = data.items || []; // [{product, newBalance}]
  _txCache = null;

  for (const item of items) {
    if (!PRODUCTS.includes(item.product)) continue;
    const newBalance = Number(item.newBalance);
    if (isNaN(newBalance)) continue;

    const currentBalance = mainBalance_(item.product);
    const diff = newBalance - currentBalance;

    if (diff !== 0) {
      addTransaction_(TX_TYPES.ADJUSTMENT, item.product, diff, 'تعديل يدوي', 'المخزون', '', 'تعديل رصيد المخزون من ' + currentBalance + ' إلى ' + newBalance, data.user || 'admin');
      _txCache = null;
    }
  }

  return { ok: true, message: 'تم تعديل المخزون بنجاح' };
}

/** إنشاء احتياج جديد (الإدارة) */
function handleCreateNeed_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.NEEDS);
    const id = 'N-' + Date.now();

    const items = data.items || []; // [{product, quantity}]
    items.forEach(item => {
      sheet.appendRow([
        id + '-' + item.product,
        new Date(),
        item.product,
        Number(item.quantity),
        0,
        Number(item.quantity),
        'جاري',
        data.user || 'admin'
      ]);
    });
    return { ok: true, message: 'تم تسجيل الاحتياجات' };
  } finally {
    lock.releaseLock();
  }
}

/** تسليم من المصنع (عدة منتجات) */
function handleFactoryDelivery_(data) {
  const items = data.items || []; // [{product, quantity}]
  _txCache = null;

  for (const item of items) {
    const qty = Number(item.quantity);
    if (qty <= 0) continue;
    if (!PRODUCTS.includes(item.product)) continue;

    // إضافة حركة مخزون
    addTransaction_(TX_TYPES.FACTORY_DELIVERY, item.product, qty, 'المصنع', 'المخزون', '', data.notes || '', data.user || 'factory');

    // تحديث الاحتياجات
    updateNeedsAfterDelivery_(item.product, qty);
  }

  return { ok: true, message: 'تم تسجيل التسليم وتحديث المخزون والاحتياجات' };
}

/** تحديث الاحتياجات بعد تسليم المصنع */
function updateNeedsAfterDelivery_(product, qty) {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.NEEDS);
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  let remaining = qty;

  for (let i = 0; i < data.length && remaining > 0; i++) {
    if (data[i][2] === product && data[i][6] === 'جاري') {
      const needRemaining = Number(data[i][5]) || 0;
      const received = Number(data[i][4]) || 0;

      if (needRemaining > 0) {
        const toAdd = Math.min(remaining, needRemaining);
        const newReceived = received + toAdd;
        const newRemaining = needRemaining - toAdd;
        const row = i + 2;

        sheet.getRange(row, 5).setValue(newReceived);
        sheet.getRange(row, 6).setValue(newRemaining);

        if (newRemaining <= 0) {
          sheet.getRange(row, 7).setValue('مكتمل');
          sheet.getRange(row, 1, 1, 8).setBackground('#d9ead3');
        }
        remaining -= toAdd;
      }
    }
  }
}

/** فرز عهدة لمندوب (طلب) */
function handleSortToDelegate_(data) {
  const delegate = data.delegate;
  const items = data.items || []; // [{product, quantity}]
  const orderNumber = data.orderNumber || '';

  _txCache = null;

  let totalBarsNeeded = 0;
  // حساب البارات المطلوبة (السماح بالمخزون السالب)
  for (const item of items) {
    if (Number(item.quantity) <= 0) continue;
    if (BOX_PRODUCTS.includes(item.product)) {
      totalBarsNeeded += Number(item.quantity) * BARS_PER_UNIT;
    }
  }

  // تنفيذ الفرز
  for (const item of items) {
    const qty = Number(item.quantity);
    if (qty <= 0) continue;
    addTransaction_(TX_TYPES.SORT_TO_DELEGATE, item.product, qty, 'المخزون', delegate, orderNumber, 'فرز طلب', data.user || 'sorter');
    _txCache = null; // reset cache after each write
  }

  if (totalBarsNeeded > 0) {
    addTransaction_(TX_TYPES.BAR_AUTO_DEDUCTION, 'بارات الومنيوم', totalBarsNeeded, 'المخزون', 'المندوب ' + delegate, orderNumber, 'خصم بارات تلقائي مع فرز الطلب', data.user || 'sorter');
    _txCache = null;
  }

  return { ok: true, message: 'تم الفرز بنجاح لـ ' + delegate + (totalBarsNeeded > 0 ? ' (تم خصم ' + totalBarsNeeded + ' بار تلقائياً)' : '') };
}

/** عهدة إضافية */
function handleExtraCustody_(data) {
  const delegate = data.delegate;
  const items = data.items || [];

  _txCache = null;

  let totalBarsNeeded = 0;
  // حساب البارات المطلوبة (السماح بالمخزون السالب)
  for (const item of items) {
    if (Number(item.quantity) <= 0) continue;
    if (BOX_PRODUCTS.includes(item.product)) {
      totalBarsNeeded += Number(item.quantity) * BARS_PER_UNIT;
    }
  }

  // تنفيذ إضافة العهدة الإضافية
  for (const item of items) {
    const qty = Number(item.quantity);
    if (qty <= 0) continue;
    addTransaction_(TX_TYPES.EXTRA_CUSTODY, item.product, qty, 'المخزون', delegate, '', 'عهدة إضافية', data.user || 'sorter');
    _txCache = null;
  }

  if (totalBarsNeeded > 0) {
    addTransaction_(TX_TYPES.BAR_AUTO_DEDUCTION, 'بارات الومنيوم', totalBarsNeeded, 'المخزون', 'عهدة إضافية ' + delegate, '', 'خصم بارات تلقائي مع العهدة الإضافية', data.user || 'sorter');
    _txCache = null;
  }

  return { ok: true, message: 'تم حفظ العهدة الإضافية بنجاح' + (totalBarsNeeded > 0 ? ' (تم خصم ' + totalBarsNeeded + ' بار تلقائياً)' : '') };
}

/** تسجيل تركيب (مندوب) */
function handleInstallation_(data) {
  const delegate = data.delegate;
  const orderNumber = String(data.orderNumber || '').trim();
  const items = data.items || []; // [{product, quantity}]

  // Basic payment data
  const paymentStatus = data.paymentStatus || '';
  const paymentAmount = Number(data.paymentAmount) || 0;
  const isCollected = (paymentStatus === 'تحصيل' && paymentAmount > 0);

  // Extra installation data
  const hasExtra = data.hasExtra === true;
  const extraProduct = data.extraProduct || '';
  const extraQuantity = Number(data.extraQuantity) || 0;
  const extraReason = data.extraReason || '';
  const extraPayment = data.extraPayment || '';
  const extraAmount = Number(data.extraAmount) || 0;

  _txCache = null;

  // Validate order and check COD status
  if (orderNumber) {
    const valRes = validateOrder_(orderNumber, delegate);
    if (!valRes.ok) return valRes;
    
    if (valRes.isCOD && !isCollected) {
      return { ok: false, error: 'تعذر التسجيل: الطلب عليه مبلغ تحصيل (دفع عند الاستلام). يجب تسجيل أنه تم التحصيل وإدخال المبلغ.' };
    }
  }

  // تحديد منطقة المندوب وهل تطبق عهدة المنطقة
  const delegateRegion = getDelegateRegion_(delegate);
  const isRegionCustody = isRegionCustodyRegion_(delegateRegion);

  // عهدة المنطقة: تنفيذ الخصم (الشرقية والغربية فقط) - يسمح بالسالب لعدم تعطيل المناديب
  if (isRegionCustody) {
    for (const item of items) {
      const qty = Number(item.quantity);
      if (qty <= 0) continue;
      if (item.product === 'بارات الومنيوم') continue;
      deductRegionCustody_(delegateRegion, item.product, qty, delegate, orderNumber, data.user || delegate);
    }
    if (hasExtra && extraQuantity > 0 && extraProduct !== 'بارات الومنيوم') {
      deductRegionCustody_(delegateRegion, extraProduct, extraQuantity, delegate, orderNumber, data.user || delegate);
    }
  }

  // Handle Photos (رفع الصور إلى Google Drive)
  let photoLinks = '';
  if (data.photos && data.photos.length > 0) {
    try {
      const folderId = '11Kl-U7idFc5e9MrKeGCCmIaIiU_Zi6XE';
      let folder;
      try {
        folder = DriveApp.getFolderById(folderId);
      } catch(e) {
        // Fallback in case folder ID is wrong or inaccessible
        const folders = DriveApp.getFoldersByName('صور_منشر');
        folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('صور_منشر');
      }
      
      const links = [];
      for (let i = 0; i < data.photos.length; i++) {
        const p = data.photos[i];
        const base64Data = p.base64.includes(',') ? p.base64.split(',')[1] : p.base64;
        
        // Extract extension from mimeType (e.g. image/jpeg -> jpeg)
        let ext = 'jpg';
        if (p.mimeType && p.mimeType.includes('/')) {
            ext = p.mimeType.split('/')[1];
        }
        
        // Rename logic: المندوب_رقم الطلب_رقم الصورة
        const fileName = `${delegate}_${orderNumber}_${i+1}.${ext}`;
        
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), p.mimeType, fileName);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        links.push(file.getUrl());
      }
      photoLinks = links.join(' \n ');
    } catch(e) {
      Logger.log('Error uploading photos: ' + e);
      photoLinks = 'خطأ في رفع الصور';
    }
  }

  // تنفيذ التركيب الإضافي وتسجيله كحركة
  if (hasExtra && extraQuantity > 0) {
    addTransaction_(TX_TYPES.EXTRA_INSTALLATION, extraProduct, extraQuantity, delegate, 'عميل', orderNumber, 'إضافة: ' + extraReason, data.user || delegate);
    
    // تسجيل التحصيل الإضافي في شيت التحصيلات إذا لم يكن آجل
    if (extraAmount > 0 && extraPayment !== 'آجل') {
      try {
        const dbSS = SpreadsheetApp.openById(DB_SHEET_ID);
        const cSheet = dbSS.getSheetByName(SHEET_NAMES.COLLECTIONS);
        if (cSheet) {
          cSheet.appendRow([
            new Date(),
            orderNumber || 'إضافي',
            delegate,
            extraAmount,
            extraPayment || 'كاش',
            'بيع عهدة إضافية: ' + extraProduct
          ]);
        }
      } catch(e) {
        Logger.log('Error logging extra collection: ' + e);
      }
    }

    // إضافة سجل مراجعة مالية للعهدة الإضافية
    if (extraAmount > 0 && extraPayment !== 'آجل') {
      addFinancialReview_(orderNumber ? orderNumber + '-E' : 'إضافي-E', delegate, extraAmount, extraPayment || 'كاش', 'عهدة إضافية');
    }

    // تسجيل الطلب الإضافي في شيت الطلبات الخارجي (ORDERS_SHEET_ID)
    try {
      const orderSS = SpreadsheetApp.openById(ORDERS_SHEET_ID);
      let extraSheet = orderSS.getSheetByName('الطلبات_الإضافية');
      if (!extraSheet) {
        extraSheet = orderSS.insertSheet('الطلبات_الإضافية');
        extraSheet.appendRow(['التاريخ', 'اسم المندوب', 'رقم الطلب', 'نوع المنشر', 'الكمية', 'حالة الدفع', 'المبلغ المحصل', 'السبب']);
      }
      extraSheet.appendRow([
        new Date(), delegate, orderNumber, extraProduct, extraQuantity, 
        extraPayment || '', extraAmount || '', extraReason || ''
      ]);
    } catch(e) {
      Logger.log('Error logging extra installation to orders sheet: ' + e);
    }
  }

  // تحديث الشيت الرئيسي للطلبات وتسجيل في شيت التركيب
  if (orderNumber) {
    markOrderDone_(orderNumber, delegate, isCollected);

    // إضافة سجل مراجعة مالية للطلبات COD
    if (isCollected && paymentAmount > 0) {
      addFinancialReview_(orderNumber, delegate, paymentAmount, data.paymentMethod || 'كاش', 'أساسي');
    }

    try {
      const dbSS = SpreadsheetApp.openById(DB_SHEET_ID);
      const sheet1 = dbSS.getSheetByName('Sheet1');
      if (sheet1) {
        const basicQty = items.reduce((sum, item) => sum + Number(item.quantity), 0);
        const basicProduct = items.map(item => item.product).join(' + ');
        const totalCollected = paymentAmount + (extraPayment !== 'آجل' ? extraAmount : 0);

        sheet1.appendRow([
          new Date(),
          delegate,
          data.region || '',
          orderNumber,
          basicProduct,
          basicQty,
          paymentStatus,
          paymentAmount,
          hasExtra ? extraProduct : '',
          hasExtra ? extraQuantity : '',
          hasExtra ? extraReason : '',
          hasExtra ? extraPayment : '',
          hasExtra ? extraAmount : '',
          totalCollected,
          photoLinks // روابط الصور من Google Drive
        ]);
      }
    } catch(e) {
      Logger.log('Error appending to Sheet1: ' + e);
    }
  }

  return { ok: true, message: 'تم تسجيل التركيب بالكامل بنجاح' };
}

/** تركيب إضافي (من العهدة الإضافية) */
function handleExtraInstallation_(data) {
  const delegate = data.delegate;
  const product = data.product;
  const qty = Number(data.quantity);
  const orderNumber = data.orderNumber || '';

  if (!PRODUCTS.includes(product)) return { ok: false, error: 'منتج غير صالح' };
  if (qty <= 0) return { ok: false, error: 'الكمية يجب أن تكون أكبر من صفر' };

  _txCache = null;
  const extraBal = delegateExtraBalance_(delegate, product);
  if (extraBal < qty) {
    return { ok: false, error: 'العهدة الإضافية غير كافية (المتاح: ' + extraBal + ')' };
  }

  addTransaction_(TX_TYPES.EXTRA_INSTALLATION, product, qty, delegate, 'عميل', orderNumber, 'إضافة: ' + (data.reason || ''), data.user || delegate);

  // تسجيل التحصيل إذا كان هناك مبلغ
  if (Number(data.paymentAmount) > 0) {
    try {
      const dbSS = SpreadsheetApp.openById(DB_SHEET_ID);
      const cSheet = dbSS.getSheetByName(SHEET_NAMES.COLLECTIONS);
      if (cSheet) {
        cSheet.appendRow([
          new Date(),
          orderNumber || 'إضافي',
          delegate,
          Number(data.paymentAmount),
          data.paymentStatus || 'كاش',
          'بيع عهدة إضافية: ' + product
        ]);
      }
    } catch(e) {
      Logger.log('Error logging extra collection: ' + e);
    }
  }

  // إضافة سجل مراجعة مالية للعهدة الإضافية
  if (Number(data.paymentAmount) > 0 && data.paymentStatus !== 'آجل') {
    addFinancialReview_(orderNumber ? orderNumber + '-E' : 'إضافي-E', delegate, Number(data.paymentAmount), data.paymentStatus || 'كاش', 'عهدة إضافية');
  }

  // تسجيل الطلب الإضافي في شيت الطلبات الخارجي (ORDERS_SHEET_ID)
  try {
    const orderSS = SpreadsheetApp.openById(ORDERS_SHEET_ID);
    let extraSheet = orderSS.getSheetByName('الطلبات_الإضافية');
    if (!extraSheet) {
      extraSheet = orderSS.insertSheet('الطلبات_الإضافية');
      extraSheet.appendRow(['التاريخ', 'اسم المندوب', 'رقم الطلب', 'نوع المنشر', 'الكمية', 'حالة الدفع', 'المبلغ المحصل', 'السبب']);
    }
    extraSheet.appendRow([
      new Date(), delegate, orderNumber, product, qty, 
      data.paymentStatus || '', data.paymentAmount || '', data.reason || ''
    ]);
  } catch(e) {
    Logger.log('Error logging extra installation to orders sheet: ' + e);
  }

  return { ok: true, message: 'تم تسجيل التركيب الإضافي' };
}

/** مرتجع */
function handleReturn_(data) {
  const delegate = data.delegate;
  const product = data.product;
  const qty = Number(data.quantity);

  if (!PRODUCTS.includes(product)) return { ok: false, error: 'منتج غير صالح' };
  if (qty <= 0) return { ok: false, error: 'الكمية يجب أن تكون أكبر من صفر' };

  _txCache = null;
  const balance = delegateOrderBalance_(delegate, product);
  if (balance < qty) {
    return { ok: false, error: 'المندوب لا يملك كمية كافية (المتاح: ' + balance + ')' };
  }

  addTransaction_(TX_TYPES.RETURN, product, qty, delegate, 'المخزون', '', 'سبب: ' + (data.reason || ''), data.user || 'sorter');
  return { ok: true, message: 'تم تسجيل المرتجع' };
}

/** تالف */
function handleDamaged_(data) {
  const delegate = data.delegate;
  const product = data.product;
  const qty = Number(data.quantity);

  if (!PRODUCTS.includes(product)) return { ok: false, error: 'منتج غير صالح' };
  if (qty <= 0) return { ok: false, error: 'الكمية يجب أن تكون أكبر من صفر' };

  _txCache = null;
  const balance = delegateTotalBalance_(delegate, product);
  if (balance < qty) {
    return { ok: false, error: 'المندوب لا يملك كمية كافية (المتاح: ' + balance + ')' };
  }

  addTransaction_(TX_TYPES.DAMAGED, product, qty, delegate, 'تالف', '', 'سبب: ' + (data.reason || ''), data.user || 'sorter');
  return { ok: true, message: 'تم تسجيل التالف' };
}

/** تصفير التالف (للمصنع) */
function handleClearDamaged_(data) {
  _txCache = null;
  const summaryRes = getDamagedSummary_();
  if (!summaryRes.ok) return summaryRes;
  
  for (const item of summaryRes.data) {
    if (item.quantity > 0) {
      // Create a reverse transaction to nullify the damaged items in the factory view
      addTransaction_('clear_damaged', item.product, -item.quantity, 'تالف', 'مستبدل', '', 'تم استبدال التالف', data.user || 'factory');
      _txCache = null;
    }
  }
  return { ok: true, message: 'تم تصفير التوالف بنجاح' };
}

/** تحصيل */
function handleCollection_(data) {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
  sheet.appendRow([
    new Date(),
    data.orderNumber || '',
    data.delegate || '',
    Number(data.amount) || 0,
    data.method || 'كاش',
    data.notes || ''
  ]);
  return { ok: true, message: 'تم تسجيل التحصيل' };
}

// ==========================================
// عمليات القراءة (GET)
// ==========================================

/** رصيد المخزون الرئيسي */
function getInventoryBalance_() {
  _txCache = null;
  const inventory = PRODUCTS.map(p => {
    const txs = getAllTransactions_().filter(t => t.product === p);
    let totalIn = 0, totalOut = 0, totalReturn = 0, totalDamaged = 0;
    txs.forEach(t => {
      if (t.type === TX_TYPES.INIT_STOCK || t.type === TX_TYPES.FACTORY_DELIVERY) totalIn += t.qty;
      if (t.type === TX_TYPES.RETURN) totalReturn += t.qty;
      if (t.type === TX_TYPES.SORT_TO_DELEGATE || t.type === TX_TYPES.EXTRA_CUSTODY) totalOut += t.qty;
      if (t.type === TX_TYPES.REGION_CUSTODY_ADD || t.type === TX_TYPES.BAR_AUTO_DEDUCTION) totalOut += t.qty;
      if (t.type === TX_TYPES.DAMAGED) totalDamaged += t.qty;
    });
    return {
      product: p,
      totalIn: totalIn,
      totalReturn: totalReturn,
      totalOut: totalOut,
      totalDamaged: totalDamaged,
      balance: totalIn + totalReturn - totalOut
    };
  });
  return { ok: true, data: inventory };
}

/** رصيد مندوب */
function getDelegateBalance_(delegate) {
  if (!delegate) return { ok: false, error: 'لم يتم تحديد المندوب' };
  _txCache = null;

  const balances = PRODUCTS.map(p => ({
    product: p,
    orderReceived: sumTx_(TX_TYPES.SORT_TO_DELEGATE, p, null, delegate),
    extraReceived: sumTx_(TX_TYPES.EXTRA_CUSTODY, p, null, delegate),
    installed: sumTx_(TX_TYPES.INSTALLATION, p, delegate, null),
    extraInstalled: sumTx_(TX_TYPES.EXTRA_INSTALLATION, p, delegate, null),
    returned: sumTx_(TX_TYPES.RETURN, p, delegate, null),
    damaged: sumTx_(TX_TYPES.DAMAGED, p, delegate, null),
    orderBalance: delegateOrderBalance_(delegate, p),
    extraBalance: delegateExtraBalance_(delegate, p)
  }));

  return { ok: true, data: balances, delegate: delegate };
}

/** كل المناديب */
function getAllDelegatesBalance_() {
  _txCache = null;
  const couriers = getCouriers_();
  const result = couriers.map(c => {
    const balances = PRODUCTS.map(p => ({
      product: p,
      orderBalance: delegateOrderBalance_(c.displayName, p),
      extraBalance: delegateExtraBalance_(c.displayName, p),
      totalBalance: delegateTotalBalance_(c.displayName, p)
    }));
    return { delegate: c.displayName, region: c.region, balances: balances };
  });
  return { ok: true, data: result };
}

/** احتياجات المصنع */
function getFactoryNeeds_() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.NEEDS);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const needs = data.map(r => ({
    id: r[0], date: r[1], product: r[2],
    required: r[3], received: r[4], remaining: r[5],
    status: r[6], createdBy: r[7]
  }));
  return { ok: true, data: needs };
}

/** ملخص التالف (للمصنع) */
function getDamagedSummary_() {
  _txCache = null;
  const txs = getAllTransactions_();
  const summary = {};
  PRODUCTS.forEach(p => { summary[p] = 0; });
  txs.forEach(t => { 
    if (t.type === TX_TYPES.DAMAGED) summary[t.product] = (summary[t.product] || 0) + t.qty; 
    if (t.type === TX_TYPES.DAMAGE_RECORD) summary[t.product] = (summary[t.product] || 0) + t.qty;
    if (t.type === 'clear_damaged') summary[t.product] = (summary[t.product] || 0) + t.qty; // negative qty
  });
  const result = PRODUCTS.map(p => ({ product: p, quantity: Math.max(0, summary[p]) }));
  return { ok: true, data: result };
}

/** التحقق من رقم الطلب */
function validateOrder_(orderNumber, delegate) {
  if (!orderNumber) return { ok: false, error: 'رقم الطلب مطلوب' };
  orderNumber = String(orderNumber).trim();

  try {
    const targetSS = SpreadsheetApp.openById(ORDERS_SHEET_ID);
    const targetSheet = targetSS.getSheets()[0];
    const data = targetSheet.getDataRange().getValues();
    if (data.length < 2) return { ok: false, error: 'لا توجد بيانات' };

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const orderCol = headers.indexOf('order_number');
    const payCol = headers.indexOf('pay_method');
    const doneCol = headers.indexOf('done');

    if (orderCol === -1) return { ok: false, error: 'عمود order_number غير موجود' };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][orderCol]).trim() === orderNumber) {
        // تحقق هل اتركب
        if (doneCol !== -1 && data[i][doneCol] === true) {
          return { ok: false, error: 'هذا الطلب تم تركيبه مسبقاً', alreadyDone: true };
        }
        const payMethod = payCol !== -1 ? String(data[i][payCol]).trim() : '';
        const isCOD = payMethod === 'دفع عند الاستلام' || payMethod.toLowerCase() === 'cod' || payMethod.toLowerCase() === 'cash on delivery';
        return { ok: true, found: true, isCOD: isCOD, payMethod: payMethod };
      }
    }

    // الطلب غير موجود - سجل خطأ
    logError_(delegate || '', orderNumber, 'رقم الطلب غير موجود');
    return { ok: false, error: 'رقم الطلب غير موجود في قاعدة البيانات برجاء التأكد من رقم الطلب وإعادة المحاولة', found: false };
  } catch (err) {
    return { ok: false, error: 'خطأ في الاتصال بقاعدة البيانات: ' + err.toString() };
  }
}

/** تحديث الطلب كـ done في الشيت الرئيسي */
function markOrderDone_(orderNumber, delegate, isCollected) {
  try {
    const targetSS = SpreadsheetApp.openById(ORDERS_SHEET_ID);
    const sheet = targetSS.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const orderCol = headers.indexOf('order_number');
    const doneCol = headers.indexOf('done');
    const personCol = headers.indexOf('person');
    const installDateCol = headers.indexOf('data_of_install');
    const doneMoneyCol = headers.indexOf('done_money');

    if (orderCol === -1) return;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][orderCol]).trim() === String(orderNumber).trim()) {
        if (doneCol !== -1) sheet.getRange(i + 1, doneCol + 1).setValue(true);
        if (personCol !== -1) sheet.getRange(i + 1, personCol + 1).setValue(delegate);
        if (installDateCol !== -1) sheet.getRange(i + 1, installDateCol + 1).setValue(new Date());
        if (isCollected && doneMoneyCol !== -1) sheet.getRange(i + 1, doneMoneyCol + 1).setValue(true);
        break;
      }
    }
  } catch (e) {
    Logger.log('markOrderDone error: ' + e.toString());
  }
}

/** سجل الأخطاء */
function logError_(delegate, orderNumber, reason) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.ERROR_LOG);
    sheet.appendRow([new Date(), delegate, orderNumber, reason]);
  } catch (e) { /* silent */ }
}

/** مجموع حركات بنوع معين */
function sumTx_(type, product, from, to) {
  const txs = getAllTransactions_();
  let sum = 0;
  txs.forEach(t => {
    if (t.type !== type || t.product !== product) return;
    if (from && t.from !== from) return;
    if (to && t.to !== to) return;
    sum += t.qty;
  });
  return sum;
}

/** 
 * إعادة هيكلة رأس الجدول في شيت التركيب
 * قم بتشغيل هذه الدالة مرة واحدة من المحرر لتجهيز الأعمدة 
 */
function setupSheet1Headers() {
  try {
    const dbSS = SpreadsheetApp.openById(DB_SHEET_ID);
    let sheet1 = dbSS.getSheetByName('Sheet1');
    
    if (!sheet1) {
      sheet1 = dbSS.insertSheet('Sheet1');
    }
    
    const headers = [
      'التاريخ',
      'اسم المندوب',
      'المنطقة',
      'رقم الطلب',
      'المنتج الأساسي',
      'العدد الأساسي',
      'حالة الدفع',
      'المبلغ المحصل الأساسي',
      'المنتج الإضافي',
      'العدد الإضافي',
      'سبب الإضافي',
      'طريقة دفع الإضافي',
      'مبلغ الإضافي',
      'إجمالي المحصل',
      'روابط الصور'
    ];
    
    // وضع العناوين في الصف الأول
    sheet1.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // تنسيق بسيط للرأس
    sheet1.getRange(1, 1, 1, headers.length)
      .setBackground('#1B3D2F') // لون خلفية داكن
      .setFontColor('#FFFFFF') // نص أبيض
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
      
    // تجميد الصف الأول
    sheet1.setFrozenRows(1);
    
    Logger.log('تم إعداد رأس شيت التركيب بنجاح');
  } catch(e) {
    Logger.log('Error setting up Sheet1 headers: ' + e);
  }
}

/** إحصائيات المندوب */
function getCourierStats_(delegate) {
  if (!delegate) return { ok: false, error: 'المندوب مطلوب' };
  _txCache = null;

  const txs = getAllTransactions_();
  const delegateTx = txs.filter(t => t.from === delegate || t.to === delegate);

  // عدد الطلبات المركبة
  const installedOrders = new Set();
  delegateTx.filter(t => t.type === TX_TYPES.INSTALLATION && t.orderNumber).forEach(t => installedOrders.add(t.orderNumber));

  const balances = PRODUCTS.map(p => ({
    product: p,
    orderReceived: sumTx_(TX_TYPES.SORT_TO_DELEGATE, p, null, delegate),
    extraReceived: sumTx_(TX_TYPES.EXTRA_CUSTODY, p, null, delegate),
    installed: sumTx_(TX_TYPES.INSTALLATION, p, delegate, null),
    extraInstalled: sumTx_(TX_TYPES.EXTRA_INSTALLATION, p, delegate, null),
    returned: sumTx_(TX_TYPES.RETURN, p, delegate, null),
    damaged: sumTx_(TX_TYPES.DAMAGED, p, delegate, null),
    orderBalance: delegateOrderBalance_(delegate, p),
    extraBalance: delegateExtraBalance_(delegate, p)
  }));

  return {
    ok: true,
    data: {
      delegate: delegate,
      totalInstallations: installedOrders.size,
      balances: balances
    }
  };
}

/** لوحة التحكم */
function getDashboard_(role, user) {
  _txCache = null;
  const txs = getAllTransactions_();

  const inventory = {};
  PRODUCTS.forEach(p => { inventory[p] = mainBalance_(p); });

  const totalIn = txs.filter(t => t.type === TX_TYPES.INIT_STOCK || t.type === TX_TYPES.FACTORY_DELIVERY).reduce((s, t) => s + t.qty, 0);
  const totalOut = txs.filter(t => t.type === TX_TYPES.SORT_TO_DELEGATE || t.type === TX_TYPES.EXTRA_CUSTODY).reduce((s, t) => s + t.qty, 0);
  const totalInstalled = txs.filter(t => t.type === TX_TYPES.INSTALLATION || t.type === TX_TYPES.EXTRA_INSTALLATION).reduce((s, t) => s + t.qty, 0);
  const totalReturned = txs.filter(t => t.type === TX_TYPES.RETURN).reduce((s, t) => s + t.qty, 0);
  const totalDamaged = txs.filter(t => t.type === TX_TYPES.DAMAGED).reduce((s, t) => s + t.qty, 0);
  const totalFactoryIn = txs.filter(t => t.type === TX_TYPES.FACTORY_DELIVERY).reduce((s, t) => s + t.qty, 0);

  // تحصيل
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const colSheet = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
  let totalCollected = 0, totalCash = 0, totalTransfer = 0;
  if (colSheet && colSheet.getLastRow() > 1) {
    const cData = colSheet.getRange(2, 1, colSheet.getLastRow() - 1, 6).getValues();
    cData.forEach(r => {
      const amt = Number(r[3]) || 0;
      totalCollected += amt;
      if (r[4] === 'كاش') totalCash += amt; else totalTransfer += amt;
    });
  }

  // احتياجات
  const needs = getFactoryNeeds_();

  return {
    ok: true,
    data: {
      inventory: inventory,
      totalIn: totalIn,
      totalFactoryIn: totalFactoryIn,
      totalOut: totalOut,
      totalInstalled: totalInstalled,
      totalReturned: totalReturned,
      totalDamaged: totalDamaged,
      totalCollected: totalCollected,
      totalCash: totalCash,
      totalTransfer: totalTransfer,
      needs: needs.data || [],
      regionCustody: getRegionCustodyDashboard_().data || []
    }
  };
}

/** حركات المخزون */
function getTransactions_(params) {
  _txCache = null;
  let txs = getAllTransactions_();

  if (params && params.type) txs = txs.filter(t => t.type === params.type);
  if (params && params.delegate) txs = txs.filter(t => t.from === params.delegate || t.to === params.delegate);
  if (params && params.product) txs = txs.filter(t => t.product === params.product);

  const limit = Number(params && params.limit) || 200;
  txs = txs.slice(-limit).reverse();

  return { ok: true, data: txs };
}

/** التحصيلات */
function getCollections_() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const result = data.map(r => ({
    date: r[0], orderNumber: r[1], delegate: r[2],
    amount: r[3], method: r[4], notes: r[5]
  }));
  return { ok: true, data: result };
}

/** تقرير المندوب (الإدارة) */
function getDelegateReport_(delegate) {
  if (!delegate) return { ok: false, error: 'المندوب مطلوب' };
  _txCache = null;
  const txs = getAllTransactions_().filter(t => t.from === delegate || t.to === delegate);
  
  // العهد الإضافية
  const extraReceived = txs.filter(t => t.type === TX_TYPES.EXTRA_CUSTODY && t.to === delegate);
  const extraInstalled = txs.filter(t => t.type === TX_TYPES.EXTRA_INSTALLATION && t.from === delegate);
  
  // التحصيلات (COD & Extra)
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const colSheet = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
  let collections = [];
  if (colSheet && colSheet.getLastRow() > 1) {
    const cData = colSheet.getRange(2, 1, colSheet.getLastRow() - 1, 6).getValues();
    collections = cData.map(r => ({
      date: r[0], orderNumber: r[1], delegate: r[2], amount: Number(r[3]), method: r[4], notes: r[5]
    })).filter(c => c.delegate === delegate);
  }

  // الطلبات المركبة
  const installations = txs.filter(t => t.type === TX_TYPES.INSTALLATION && t.from === delegate);

  const balances = PRODUCTS.map(p => ({
    product: p,
    orderBalance: delegateOrderBalance_(delegate, p),
    extraBalance: delegateExtraBalance_(delegate, p),
    totalBalance: delegateTotalBalance_(delegate, p)
  }));

  return {
    ok: true,
    data: {
      delegate: delegate,
      balances: balances,
      extraReceived: extraReceived,
      extraInstalled: extraInstalled,
      collections: collections,
      installations: installations
    }
  };
}

/** طلب من المصنع (المورد) */
function handleFactoryOrder_(data) {
  const items = data.items || [];
  const notes = data.notes || '';
  
  _txCache = null;
  
  for (const item of items) {
    const qty = Number(item.quantity);
    if (qty <= 0) continue;
    if (!PRODUCTS.includes(item.product)) continue;
    
    addTransaction_(TX_TYPES.FACTORY_ORDER, item.product, qty, 'المصنع', 'المورد', '', notes, data.user || 'factory');
  }
  
  return { ok: true, message: 'تم تسجيل الطلب من المصنع بنجاح' };
}

/** ملخص المصنع التجميعي - 4 قيم لكل منتج */
function getFactoryOverview_() {
  _txCache = null;
  const txs = getAllTransactions_();
  
  // حساب إجمالي الاحتياج من شيت الاحتياجات
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const needsSheet = ss.getSheetByName(SHEET_NAMES.NEEDS);
  const needsTotals = {};
  PRODUCTS.forEach(p => { needsTotals[p] = 0; });
  
  if (needsSheet && needsSheet.getLastRow() > 1) {
    const needsData = needsSheet.getRange(2, 1, needsSheet.getLastRow() - 1, 8).getValues();
    needsData.forEach(r => {
      const product = r[2];
      const required = Number(r[3]) || 0;
      if (PRODUCTS.includes(product)) {
        needsTotals[product] += required;
      }
    });
  }
  
  const overview = PRODUCTS.map(p => {
    const originalNeed = needsTotals[p];
    const totalOrdered = txs.filter(t => t.type === TX_TYPES.FACTORY_ORDER && t.product === p).reduce((s, t) => s + t.qty, 0);
    const totalReceived = txs.filter(t => t.type === TX_TYPES.FACTORY_DELIVERY && t.product === p).reduce((s, t) => s + t.qty, 0);
    
    return {
      product: p,
      originalNeed: originalNeed,
      needToOrder: originalNeed - totalOrdered,
      pendingDelivery: totalOrdered - totalReceived,
      totalReceived: totalReceived,
      remaining: originalNeed - totalReceived
    };
  });
  
  return { ok: true, data: overview };
}

/** جلب سجل طلبات المصنع */
function getFactoryOrders_() {
  _txCache = null;
  const txs = getAllTransactions_().filter(t => t.type === TX_TYPES.FACTORY_ORDER);
  return { ok: true, data: txs.reverse() };
}

/** إسناد طلبية - سحب من المخزون */
function handleOrderAssignment_(data) {
  const orderNumber = String(data.orderNumber || '').trim();
  const items = data.items || [];
  const notes = data.notes || '';
  
  if (!orderNumber) return { ok: false, error: 'رقم الطلبية مطلوب' };
  
  _txCache = null;
  
  let totalBarsNeeded = 0;
  // حساب البارات المطلوبة (السماح بالمخزون السالب)
  for (const item of items) {
    const qty = Number(item.quantity);
    if (qty <= 0) continue;
    if (!PRODUCTS.includes(item.product)) continue;
    // منع إسناد البارات يدوياً - البارات تُخصم تلقائياً
    if (item.product === 'بارات الومنيوم') continue;
    if (BOX_PRODUCTS.includes(item.product)) {
      totalBarsNeeded += qty * BARS_PER_UNIT;
    }
  }
  
  // تنفيذ السحب
  for (const item of items) {
    const qty = Number(item.quantity);
    if (qty <= 0) continue;
    if (!PRODUCTS.includes(item.product)) continue;
    if (item.product === 'بارات الومنيوم') continue;
    addTransaction_(TX_TYPES.ORDER_ASSIGNMENT, item.product, qty, 'المخزون', 'طلبية ' + orderNumber, orderNumber, notes, data.user || 'admin');
    _txCache = null;
  }

  if (totalBarsNeeded > 0) {
    addTransaction_(TX_TYPES.BAR_AUTO_DEDUCTION, 'بارات الومنيوم', totalBarsNeeded, 'المخزون', 'طلبية ' + orderNumber, orderNumber, 'خصم بارات تلقائي مع الطلبية', data.user || 'admin');
    _txCache = null;
  }
  
  return { ok: true, message: 'تم إسناد الطلبية رقم ' + orderNumber + ' بنجاح' + (totalBarsNeeded > 0 ? ' (تم خصم ' + totalBarsNeeded + ' بار تلقائياً)' : '') };
}

/** جلب سجل الطلبيات المسندة */
function getOrderAssignments_() {
  _txCache = null;
  const txs = getAllTransactions_().filter(t => t.type === TX_TYPES.ORDER_ASSIGNMENT);
  return { ok: true, data: txs.reverse() };
}

/** تسجيل تالف بدون مندوب */
function handleDamageRecord_(data) {
  const product = data.product;
  const qty = Number(data.quantity);
  const source = data.source || ''; // من المصنع / من العميل
  const reason = data.reason || '';
  
  if (!PRODUCTS.includes(product)) return { ok: false, error: 'منتج غير صالح' };
  if (qty <= 0) return { ok: false, error: 'الكمية يجب أن تكون أكبر من صفر' };
  if (!source) return { ok: false, error: 'يرجى تحديد مصدر التلف' };
  
  _txCache = null;
  
  addTransaction_(TX_TYPES.DAMAGE_RECORD, product, qty, source, 'تالف', '', 'سبب: ' + reason + ' | مصدر: ' + source, data.user || 'distribution');
  
  return { ok: true, message: 'تم تسجيل التالف بنجاح' };
}

// ==========================================
// إعدادات الحسابات والمناديب
// ==========================================

/** جلب المستخدمين (للإدارة) */
function getUsers_() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const result = data.map(r => ({
    displayName: r[0],
    username: String(r[1]).trim(),
    role: r[2]
  }));
  return { ok: true, data: result };
}

/** تحديث كلمة المرور (للإدارة) */
function handleUpdatePassword_(data) {
  const role = data.role;
  const newPassword = String(data.newPassword).trim();
  if (!role || !newPassword) return { ok: false, error: 'بيانات غير مكتملة' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
    if (!sheet || sheet.getLastRow() < 2) return { ok: false, error: 'جدول المستخدمين غير موجود' };

    const usersData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (let i = 0; i < usersData.length; i++) {
      if (usersData[i][2] === role) {
        sheet.getRange(i + 2, 2).setValue(newPassword);
        return { ok: true, message: 'تم تحديث كلمة المرور بنجاح' };
      }
    }
    return { ok: false, error: 'القسم غير موجود' };
  } catch (err) {
    return { ok: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/** إضافة مندوب (للإدارة) */
function handleAddDelegate_(data) {
  const delegateName = String(data.name).trim();
  const delegateRegion = String(data.region).trim();
  
  if (!delegateName || !delegateRegion) return { ok: false, error: 'يرجى إدخال اسم المندوب والمنطقة' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName('المناديب');
    if (!sheet) return { ok: false, error: 'جدول المناديب غير موجود' };
    
    sheet.appendRow([delegateName, delegateRegion]);
    return { ok: true, message: 'تمت إضافة المندوب بنجاح' };
  } catch (err) {
    return { ok: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/** إضافة سجل مراجعة مالية */
function addFinancialReview_(orderNumber, delegate, amount, method, type) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.FINANCIAL_REVIEW);
    if (!sheet) return;
    const id = 'FR-' + Date.now() + '-' + Math.random().toString(36).substr(2, 4);
    sheet.appendRow([
      id, new Date(), String(orderNumber), delegate, Number(amount),
      method || 'كاش', type || 'أساسي',
      'pending', '', '', 'pending', '', ''
    ]);
  } catch(e) {
    Logger.log('Error adding financial review: ' + e);
  }
}

/** طلبات معلقة عند الفرز */
function getFinancialPendingDist_() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.FINANCIAL_REVIEW);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [], stats: { total: 0, reviewed: 0, pending: 0 } };
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const pending = data.filter(r => String(r[7]).trim() === 'pending').map(r => ({
    id: r[0], date: r[1], orderNumber: String(r[2]), delegate: r[3],
    amount: Number(r[4]), payMethod: r[5], type: r[6]
  }));
  
  const grouped = {};
  pending.forEach(item => {
    if (!grouped[item.delegate]) grouped[item.delegate] = [];
    grouped[item.delegate].push(item);
  });
  
  const result = Object.keys(grouped).map(delegate => ({
    installer: delegate,
    orders: grouped[delegate],
    totalAmount: grouped[delegate].reduce((s, o) => s + o.amount, 0)
  }));
  
  const reviewed = data.filter(r => String(r[7]).trim() === 'confirmed').length;
  
  return { ok: true, data: result, stats: { total: data.length, reviewed: reviewed, pending: pending.length } };
}

/** طلبات معلقة عند المصنع */
function getFinancialPendingFactory_() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.FINANCIAL_REVIEW);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [], stats: { total: 0, reviewed: 0, pending: 0 } };
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const pending = data.filter(r => String(r[7]).trim() === 'confirmed' && String(r[10]).trim() === 'pending').map(r => ({
    id: r[0], date: r[1], orderNumber: String(r[2]), delegate: r[3],
    amount: Number(r[4]), payMethod: r[5], type: r[6],
    distReviewer: r[8], distDate: r[9]
  }));
  
  const grouped = {};
  pending.forEach(item => {
    if (!grouped[item.delegate]) grouped[item.delegate] = [];
    grouped[item.delegate].push(item);
  });
  
  const result = Object.keys(grouped).map(delegate => ({
    installer: delegate,
    orders: grouped[delegate],
    totalAmount: grouped[delegate].reduce((s, o) => s + o.amount, 0)
  }));
  
  const reviewed = data.filter(r => String(r[10]).trim() === 'confirmed').length;
  
  return { ok: true, data: result, stats: { total: data.length, reviewed: reviewed, pending: pending.length } };
}

/** كل السجلات المالية (للإدارة) */
function getFinancialAll_() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.FINANCIAL_REVIEW);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [], stats: {} };
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const records = data.map(r => ({
    id: r[0], date: r[1], orderNumber: String(r[2]), delegate: r[3],
    amount: Number(r[4]), payMethod: r[5], type: r[6],
    distStatus: String(r[7]).trim(), distReviewer: r[8], distDate: r[9],
    factoryStatus: String(r[10]).trim(), factoryReviewer: r[11], factoryDate: r[12]
  })).reverse();
  
  const stats = {
    total: records.length,
    totalAmount: records.reduce((s, r) => s + r.amount, 0),
    pendingDist: records.filter(r => r.distStatus === 'pending').length,
    pendingFactory: records.filter(r => r.distStatus === 'confirmed' && r.factoryStatus === 'pending').length,
    completed: records.filter(r => r.factoryStatus === 'confirmed').length
  };
  
  return { ok: true, data: records, stats: stats };
}

/** تأكيد مراجعة الفرز */
function handleDistFinancialConfirm_(data) {
  const ids = data.ids || [];
  if (ids.length === 0) return { ok: false, error: 'لم يتم تحديد طلبات' };
  
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.FINANCIAL_REVIEW);
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, error: 'لا توجد بيانات' };
  
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  let confirmed = 0;
  
  for (let i = 0; i < allData.length; i++) {
    if (ids.includes(String(allData[i][0])) && String(allData[i][7]).trim() === 'pending') {
      const row = i + 2;
      sheet.getRange(row, 8).setValue('confirmed');
      sheet.getRange(row, 9).setValue('مسؤول الفرز');
      sheet.getRange(row, 10).setValue(new Date());
      confirmed++;
    }
  }
  
  return { ok: true, message: 'تم تأكيد ' + confirmed + ' طلبات بنجاح' };
}

/** تأكيد مراجعة المصنع */
function handleFactoryFinancialConfirm_(data) {
  const ids = data.ids || [];
  if (ids.length === 0) return { ok: false, error: 'لم يتم تحديد طلبات' };
  
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.FINANCIAL_REVIEW);
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, error: 'لا توجد بيانات' };
  
  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  let confirmed = 0;
  
  for (let i = 0; i < allData.length; i++) {
    if (ids.includes(String(allData[i][0])) && String(allData[i][7]).trim() === 'confirmed' && String(allData[i][10]).trim() === 'pending') {
      const row = i + 2;
      sheet.getRange(row, 11).setValue('confirmed');
      sheet.getRange(row, 12).setValue('المصنع');
      sheet.getRange(row, 13).setValue(new Date());
      confirmed++;
    }
  }
  
  return { ok: true, message: 'تم تأكيد ' + confirmed + ' طلبات بنجاح' };
}

/** إحصائيات العهدة الإضافية لكل مندوب */
function getExtraCustodyStats_() {
  _txCache = null;
  const txs = getAllTransactions_().filter(t => t.type === TX_TYPES.EXTRA_INSTALLATION);
  
  const stats = {};
  txs.forEach(t => {
    if (!stats[t.from]) {
      stats[t.from] = {};
      PRODUCTS.forEach(p => { stats[t.from][p] = 0; });
    }
    if (stats[t.from][t.product] !== undefined) {
      stats[t.from][t.product] += t.qty;
    }
  });
  
  const result = Object.keys(stats).map(delegate => ({
    delegate: delegate,
    products: stats[delegate],
    total: Object.values(stats[delegate]).reduce((s, v) => s + v, 0)
  }));
  
  return { ok: true, data: result };
}


// ==========================================
// نظام عُهدة المنطقة
// ==========================================

function isRegionCustodyRegion_(region) {
  const r = normalizeRegion_(region);
  return getCustodyRegions_().includes(r);
}

function normalizeRegion_(region) {
  const r = String(region).trim();
  if (r === 'الغربية' || r === 'الغربيه') return 'الغربيه';
  if (r === 'الشرقية' || r === 'الشرقيه') return 'الشرقية';
  return r;
}

function getDelegateRegion_(delegate) {
  const couriers = getCouriers_();
  const c = couriers.find(c => c.displayName === delegate);
  return c ? normalizeRegion_(c.region) : '';
}

function regionCustodyBalance_(region, product) {
  const normalizedRegion = normalizeRegion_(region);
  const txs = getAllTransactions_();
  let balance = 0;
  txs.forEach(t => {
    if (t.product !== product) return;
    if (t.type === TX_TYPES.REGION_CUSTODY_ADD && normalizeRegion_(t.to) === normalizedRegion) balance += t.qty;
    if (t.type === TX_TYPES.REGION_CUSTODY_USE && normalizeRegion_(t.from) === normalizedRegion) balance -= t.qty;
    if (t.type === TX_TYPES.REGION_CUSTODY_RETURN && normalizeRegion_(t.to) === normalizedRegion) balance += t.qty;
  });
  return balance;
}

function handleRegionCustodyAdd_(data) {
  const region = normalizeRegion_(data.region);
  if (!isRegionCustodyRegion_(region)) {
    return { ok: false, error: 'هذه المنطقة غير مسجلة أو معطلة في النظام' };
  }
  const items = data.items || [];
  if (items.length === 0) return { ok: false, error: 'لم يتم تحديد منتجات' };
  _txCache = null;
  let totalBarsNeeded = 0;
  // حساب البارات المطلوبة (السماح بالمخزون السالب)
  for (const item of items) {
    const qty = Number(item.quantity);
    if (qty <= 0) continue;
    if (!PRODUCTS.includes(item.product)) continue;
    if (item.product === 'بارات الومنيوم') continue;
    if (BOX_PRODUCTS.includes(item.product)) {
      totalBarsNeeded += qty * BARS_PER_UNIT;
    }
  }
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const custodySheet = ss.getSheetByName(SHEET_NAMES.REGION_CUSTODY);
  for (const item of items) {
    const qty = Number(item.quantity);
    if (qty <= 0) continue;
    if (!PRODUCTS.includes(item.product)) continue;
    if (item.product === 'بارات الومنيوم') continue;
    const balanceBefore = regionCustodyBalance_(region, item.product);
    const balanceAfter = balanceBefore + qty;
    addTransaction_(TX_TYPES.REGION_CUSTODY_ADD, item.product, qty, 'المخزون', region, '', 'صرف لعهدة منطقة ' + region, data.user || 'distribution');
    _txCache = null;
    if (custodySheet) {
      const id = 'RC-' + Date.now() + '-' + Math.random().toString(36).substr(2, 4);
      custodySheet.appendRow([id, new Date(), region, 'add', item.product, qty, '', '', balanceBefore, balanceAfter, data.notes || '', data.user || 'distribution']);
    }
  }
  if (totalBarsNeeded > 0) {
    addTransaction_(TX_TYPES.BAR_AUTO_DEDUCTION, 'بارات الومنيوم', totalBarsNeeded, 'المخزون', 'عهدة ' + region, '', 'خصم بارات تلقائي عند إضافة عهدة منطقة ' + region, data.user || 'distribution');
    _txCache = null;
    if (custodySheet) {
      const id = 'RC-BAR-' + Date.now();
      custodySheet.appendRow([id, new Date(), region, 'bar_deduction', 'بارات الومنيوم', totalBarsNeeded, '', '', 0, 0, 'خصم بارات تلقائي مع عهدة المنطقة', data.user || 'distribution']);
    }
  }
  return { ok: true, message: 'تم إضافة عهدة منطقة ' + region + ' بنجاح' + (totalBarsNeeded > 0 ? ' (تم خصم ' + totalBarsNeeded + ' بار تلقائياً)' : '') };
}

function handleRegionCustodyReturn_(data) {
  const region = normalizeRegion_(data.region);
  if (!isRegionCustodyRegion_(region)) return { ok: false, error: 'هذه المنطقة غير مسجلة أو معطلة في النظام' };
  const product = data.product;
  const qty = Number(data.quantity);
  const delegate = data.delegate || '';
  if (!PRODUCTS.includes(product)) return { ok: false, error: 'منتج غير صالح' };
  if (qty <= 0) return { ok: false, error: 'الكمية يجب أن تكون أكبر من صفر' };
  _txCache = null;
  const balanceBefore = regionCustodyBalance_(region, product);
  const balanceAfter = balanceBefore + qty;
  addTransaction_(TX_TYPES.REGION_CUSTODY_RETURN, product, qty, delegate || region, region, '', 'إرجاع عهدة منطقة', data.user || 'distribution');
  _txCache = null;
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const custodySheet = ss.getSheetByName(SHEET_NAMES.REGION_CUSTODY);
  if (custodySheet) {
    const id = 'RC-RET-' + Date.now();
    custodySheet.appendRow([id, new Date(), region, 'return', product, qty, delegate, '', balanceBefore, balanceAfter, data.notes || '', data.user || 'distribution']);
  }
  return { ok: true, message: 'تم إرجاع العهدة بنجاح' };
}

function getRegionCustodyDashboard_() {
  _txCache = null;
  const txs = getAllTransactions_();
  const couriers = getCouriers_();
  const result = getCustodyRegions_().map(region => {
    const nr = normalizeRegion_(region);
    const productsBalance = PRODUCTS.filter(p => p !== 'بارات الومنيوم').map(p => {
      const totalAdded = txs.filter(t => t.type === TX_TYPES.REGION_CUSTODY_ADD && t.product === p && normalizeRegion_(t.to) === nr).reduce((s, t) => s + t.qty, 0);
      const totalUsed = txs.filter(t => t.type === TX_TYPES.REGION_CUSTODY_USE && t.product === p && normalizeRegion_(t.from) === nr).reduce((s, t) => s + t.qty, 0);
      const totalReturned = txs.filter(t => t.type === TX_TYPES.REGION_CUSTODY_RETURN && t.product === p && normalizeRegion_(t.to) === nr).reduce((s, t) => s + t.qty, 0);
      return { product: p, totalAdded, totalUsed, totalReturned, balance: totalAdded - totalUsed + totalReturned };
    });
    const totalAdded = productsBalance.reduce((s, p) => s + p.totalAdded, 0);
    const totalUsed = productsBalance.reduce((s, p) => s + p.totalUsed, 0);
    const totalReturned = productsBalance.reduce((s, p) => s + p.totalReturned, 0);
    const totalBalance = productsBalance.reduce((s, p) => s + p.balance, 0);
    const orderNumbers = new Set();
    txs.filter(t => t.type === TX_TYPES.REGION_CUSTODY_USE && normalizeRegion_(t.from) === nr && t.orderNumber).forEach(t => orderNumbers.add(t.orderNumber));
    const regionCouriers = couriers.filter(c => normalizeRegion_(c.region) === nr);
    return { region: nr, productsBalance, totalAdded, totalUsed, totalReturned, totalBalance, ordersCount: orderNumbers.size, couriersCount: regionCouriers.length, couriers: regionCouriers.map(c => c.displayName) };
  });
  return { ok: true, data: result };
}

function getRegionCustodyDetail_(region) {
  if (!region) return { ok: false, error: 'المنطقة مطلوبة' };
  const nr = normalizeRegion_(region);
  if (!isRegionCustodyRegion_(nr)) return { ok: false, error: 'هذه المنطقة لا تطبق نظام عهدة المنطقة' };
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const custodySheet = ss.getSheetByName(SHEET_NAMES.REGION_CUSTODY);
  let movements = [];
  if (custodySheet && custodySheet.getLastRow() > 1) {
    const data = custodySheet.getRange(2, 1, custodySheet.getLastRow() - 1, 12).getValues();
    movements = data.filter(r => normalizeRegion_(String(r[2]).trim()) === nr)
      .map(r => ({ id: r[0], date: r[1], region: r[2], type: r[3], product: r[4], quantity: Number(r[5]), delegate: r[6], orderNumber: r[7], balanceBefore: Number(r[8]), balanceAfter: Number(r[9]), notes: r[10], user: r[11] })).reverse();
  }
  return { ok: true, data: { region: nr, movements } };
}

function getCourierRegionUsage_() {
  _txCache = null;
  const txs = getAllTransactions_().filter(t => t.type === TX_TYPES.REGION_CUSTODY_USE);
  const couriers = getCouriers_();
  const usage = {};
  txs.forEach(t => {
    const delegate = t.to;
    if (!usage[delegate]) {
      usage[delegate] = { delegate, region: '', products: {}, total: 0 };
      PRODUCTS.filter(p => p !== 'بارات الومنيوم').forEach(p => { usage[delegate].products[p] = 0; });
    }
    if (usage[delegate].products[t.product] !== undefined) {
      usage[delegate].products[t.product] += t.qty;
      usage[delegate].total += t.qty;
    }
  });
  Object.values(usage).forEach(u => {
    const c = couriers.find(c => c.displayName === u.delegate);
    if (c) u.region = normalizeRegion_(c.region);
  });
  return { ok: true, data: Object.values(usage).sort((a, b) => b.total - a.total) };
}

function getFactoryInventoryDashboard_() {
  _txCache = null;
  const txs = getAllTransactions_();
  const inventory = PRODUCTS.map(p => {
    const balance = mainBalance_(p);
    const totalReceived = txs.filter(t => (t.type === TX_TYPES.FACTORY_DELIVERY || t.type === TX_TYPES.INIT_STOCK) && t.product === p).reduce((s, t) => s + t.qty, 0);
    const totalSorted = txs.filter(t => t.type === TX_TYPES.SORT_TO_DELEGATE && t.product === p).reduce((s, t) => s + t.qty, 0);
    const totalExtraCustody = txs.filter(t => t.type === TX_TYPES.EXTRA_CUSTODY && t.product === p).reduce((s, t) => s + t.qty, 0);
    const totalRegionCustody = txs.filter(t => t.type === TX_TYPES.REGION_CUSTODY_ADD && t.product === p).reduce((s, t) => s + t.qty, 0);
    const totalDamaged = txs.filter(t => (t.type === TX_TYPES.DAMAGED || t.type === TX_TYPES.DAMAGE_RECORD) && t.product === p).reduce((s, t) => s + t.qty, 0);
    const totalReturned = txs.filter(t => t.type === TX_TYPES.RETURN && t.product === p).reduce((s, t) => s + t.qty, 0);
    const totalInstalled = txs.filter(t => (t.type === TX_TYPES.INSTALLATION || t.type === TX_TYPES.EXTRA_INSTALLATION) && t.product === p).reduce((s, t) => s + t.qty, 0);
    const totalBarDeduction = txs.filter(t => t.type === TX_TYPES.BAR_AUTO_DEDUCTION && t.product === p).reduce((s, t) => s + t.qty, 0);
    return { product: p, isBar: p === 'بارات الومنيوم', balance, totalReceived, totalSorted, totalExtraCustody, totalRegionCustody, totalDamaged, totalReturned, totalInstalled, totalBarDeduction };
  });
  const needs = getFactoryNeeds_();
  const totalRequired = (needs.data || []).reduce((s, n) => s + (Number(n.required) || 0), 0);
  const totalNeedReceived = (needs.data || []).reduce((s, n) => s + (Number(n.received) || 0), 0);
  const totalRemaining = (needs.data || []).reduce((s, n) => s + (Number(n.remaining) || 0), 0);
  return { ok: true, data: { inventory, totalStock: inventory.reduce((s, i) => s + i.balance, 0), totalReceived: inventory.reduce((s, i) => s + i.totalReceived, 0), totalDamaged: inventory.reduce((s, i) => s + i.totalDamaged, 0), totalRequired, totalNeedReceived, totalNeedRemaining: totalRemaining } };
}

/** جلب المناطق */
function getRegions_() {
  const ss = SpreadsheetApp.openById(DB_SHEET_ID);
  const sheet = ss.getSheetByName('المناطق');
  if (!sheet || sheet.getLastRow() < 2) {
    return { ok: true, data: DEFAULT_REGIONS.map(r => ({ name: r, status: 'نشطة', date: '' })) };
  }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const result = data.map((r, i) => ({
    index: i,
    name: String(r[0]).trim(),
    status: String(r[1]).trim() || 'نشطة',
    date: r[2]
  }));
  return { ok: true, data: result };
}

/** إضافة منطقة جديدة */
function handleAddRegion_(data) {
  const name = String(data.name || '').trim();
  if (!name) return { ok: false, error: 'يرجى إدخال اسم المنطقة' };
  
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    let sheet = ss.getSheetByName('المناطق');
    if (!sheet) {
      sheet = ss.insertSheet('المناطق');
      sheet.getRange(1, 1, 1, 3).setValues([['الاسم', 'الحالة', 'تاريخ_الإضافة']]).setFontWeight('bold').setBackground('#3D271D').setFontColor('white');
    }
    // التحقق من عدم تكرار الاسم
    if (sheet.getLastRow() > 1) {
      const existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim());
      if (existing.includes(name)) return { ok: false, error: 'المنطقة موجودة بالفعل' };
    }
    sheet.appendRow([name, 'نشطة', new Date()]);
    return { ok: true, message: 'تمت إضافة المنطقة بنجاح' };
  } catch(err) {
    return { ok: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/** تعديل منطقة */
function handleUpdateRegion_(data) {
  const index = Number(data.index);
  const newName = String(data.name || '').trim();
  if (!newName) return { ok: false, error: 'يرجى إدخال اسم المنطقة' };
  
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName('المناطق');
    if (!sheet) return { ok: false, error: 'جدول المناطق غير موجود' };
    sheet.getRange(index + 2, 1).setValue(newName);
    return { ok: true, message: 'تم تعديل المنطقة بنجاح' };
  } catch(err) {
    return { ok: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/** تفعيل/تعطيل منطقة */
function handleToggleRegion_(data) {
  const index = Number(data.index);
  const newStatus = data.status; // 'نشطة' or 'معطلة'
  
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName('المناطق');
    if (!sheet) return { ok: false, error: 'جدول المناطق غير موجود' };
    sheet.getRange(index + 2, 2).setValue(newStatus);
    return { ok: true, message: 'تم تحديث حالة المنطقة' };
  } catch(err) {
    return { ok: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/** حذف منطقة */
function handleDeleteRegion_(data) {
  const index = Number(data.index);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const sheet = ss.getSheetByName('المناطق');
    if (!sheet) return { ok: false, error: 'جدول المناطق غير موجود' };
    // التحقق من عدم وجود عمليات مرتبطة
    const regionName = sheet.getRange(index + 2, 1).getValue();
    const txs = getAllTransactions_();
    const hasTransactions = txs.some(t => 
      normalizeRegion_(t.from) === normalizeRegion_(regionName) || 
      normalizeRegion_(t.to) === normalizeRegion_(regionName)
    );
    if (hasTransactions) {
      return { ok: false, error: 'لا يمكن حذف المنطقة لوجود عمليات مرتبطة بها. يمكنك تعطيلها بدلاً من حذفها.' };
    }
    sheet.deleteRow(index + 2);
    return { ok: true, message: 'تم حذف المنطقة بنجاح' };
  } catch(err) {
    return { ok: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deductRegionCustody_(region, product, qty, delegate, orderNumber, userName) {
  const nr = normalizeRegion_(region);
  if (!isRegionCustodyRegion_(nr)) return { ok: true };
  if (product === 'بارات الومنيوم') return { ok: true };
  const balance = regionCustodyBalance_(nr, product);
  const balanceBefore = balance;
  const balanceAfter = balance - qty;
  addTransaction_(TX_TYPES.REGION_CUSTODY_USE, product, qty, nr, delegate, orderNumber, 'استخدام من عهدة منطقة ' + nr, userName);
  _txCache = null;
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    const custodySheet = ss.getSheetByName(SHEET_NAMES.REGION_CUSTODY);
    if (custodySheet) {
      const id = 'RC-USE-' + Date.now();
      custodySheet.appendRow([id, new Date(), nr, 'use', product, qty, delegate, orderNumber, balanceBefore, balanceAfter, '', userName]);
    }
  } catch(e) { Logger.log('Error logging region custody usage: ' + e); }
  return { ok: true };
}