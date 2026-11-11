/**
 * Store Online — السيرفر (Google Apps Script)
 * يستقبل بيانات العملاء والطلبات، ويرسل المنتجات للمتجر لحظياً.
 */

// ===== الإعدادات (غيّرها قبل النشر) =====
const ADMIN_KEY = 'غيّر-هذه-الكلمة-لكلمة-سر-قوية'; // كلمة سر البائع
const OWNER_EMAIL = '';                              // فارغ = بريد حسابك في جوجل
const TELEGRAM_TOKEN = '';                           // اختياري: إشعار فوري على تيليجرام
const TELEGRAM_CHAT_ID = '';                         // اختياري

const HEADERS = {
  Products:  ['id', 'الاسم', 'السعر', 'القسم', 'الوصف', 'الصورة', 'متوفر'],
  Customers: ['الهاتف', 'الاسم', 'المحافظة', 'المدينة', 'العنوان', 'أول تسجيل', 'آخر تحديث'],
  Orders:    ['رقم الطلب', 'التاريخ', 'الاسم', 'الموبايل', 'المحافظة', 'المدينة', 'العنوان',
              'ملاحظات', 'المنتجات', 'المجموع', 'الشحن', 'الإجمالي', 'طريقة الدفع', 'الحالة',
              'ملاحظة للعميل', 'آخر تحديث', 'سجل التتبع'],
  Returns:   ['رقم الاسترجاع', 'التاريخ', 'رقم الطلب', 'الاسم', 'الموبايل', 'المنتج', 'السبب',
              'التفاصيل', 'الحالة', 'ملاحظة للعميل', 'آخر تحديث']
};
const ORDER_STATUSES  = ['جديد', 'قيد التجهيز', 'تم الشحن', 'خرج للتوصيل', 'تم التسليم', 'ملغي'];
const RETURN_STATUSES = ['قيد المراجعة', 'تمت الموافقة', 'مرفوض', 'تم استلام المنتج', 'تم رد المبلغ'];

// شغّل هذه الدالة مرة واحدة من المحرر (Run) لإنشاء الجداول ومنح الصلاحيات
// (وشغّلها مرة تانية بعد أي تحديث للكود)
function setup() {
  ['Products', 'Customers', 'Orders', 'Returns'].forEach(function (n) {
    const sh = sheet_(n);
    sh.getRange(1, 1, 1, HEADERS[n].length).setValues([HEADERS[n]]);
  });
  const rule = function (list) {
    return SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build();
  };
  sheet_('Orders').getRange('N2:N1000').setDataValidation(rule(ORDER_STATUSES));
  sheet_('Returns').getRange('I2:I1000').setDataValidation(rule(RETURN_STATUSES));
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEADERS[name]);
    sh.setFrozenRows(1);
    if (name === 'Customers') sh.getRange('A:A').setNumberFormat('@');
    if (name === 'Orders') { sh.getRange('A:A').setNumberFormat('@'); sh.getRange('D:D').setNumberFormat('@'); }
    if (name === 'Returns') { ['A:A', 'C:C', 'E:E'].forEach(function (r) { sh.getRange(r).setNumberFormat('@'); }); }
  }
  return sh;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ===== قراءة المنتجات (للعملاء) =====
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'products';
    if (action !== 'products') return out_({ ok: false, error: 'bad action' });
    const sh = sheet_('Products');
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try { ensureIds_(sh); } finally { lock.releaseLock(); }
    const hidden = ['لا', 'no', 'نفد', 'نفذ', 'false', '0'];
    const products = sh.getDataRange().getValues().slice(1)
      .filter(r => r[0] !== '' && String(r[1]).trim() && hidden.indexOf(String(r[6]).trim().toLowerCase()) < 0)
      .map(r => ({ id: r[0], name: r[1], price: r[2], cat: r[3], desc: r[4], img: r[5] }))
      .reverse(); // الأحدث أولاً
    return out_({ ok: true, products: products });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

// لو أضفت منتجاً مباشرة في الجدول بدون id، نعطيه id تلقائياً
function ensureIds_(sh) {
  const lr = sh.getLastRow();
  if (lr < 2) return;
  const v = sh.getRange(2, 1, lr - 1, 2).getValues();
  let changed = false;
  v.forEach((r, i) => { if (r[0] === '' && String(r[1]).trim()) { r[0] = Date.now() + i; changed = true; } });
  if (changed) sh.getRange(2, 1, lr - 1, 1).setValues(v.map(r => [r[0]]));
}

// ===== الكتابة (عملاء / طلبات / إدارة المنتجات) =====
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const m = JSON.parse(e.postData.contents);
    switch (m.action) {
      case 'register':      return out_(register_(m));
      case 'order':         return out_(order_(m.order));
      case 'auth':          needKey_(m); return out_({ ok: true });
      case 'saveProduct':   needKey_(m); return out_(saveProduct_(m.product));
      case 'deleteProduct': needKey_(m); return out_(deleteProduct_(m.id));
      case 'track':         return out_(track_(m.orders));
      case 'return':        return out_(return_(m));
      case 'adminOrders':   needKey_(m); return out_({ ok: true, rows: rows_('Orders', 100) });
      case 'adminReturns':  needKey_(m); return out_({ ok: true, rows: rows_('Returns', 100) });
      case 'setStatus':     needKey_(m); return out_(setStatus_(m));
    }
    return out_({ ok: false, error: 'bad action' });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function needKey_(m) {
  if (m.key !== ADMIN_KEY) throw new Error('unauthorized');
}

function s_(v, n) { return String(v == null ? '' : v).slice(0, n || 200); }

function notify_(subject, body) {
  try { MailApp.sendEmail(OWNER_EMAIL || Session.getEffectiveUser().getEmail(), subject, body); } catch (e) {}
  if (TELEGRAM_TOKEN && TELEGRAM_CHAT_ID) {
    try {
      UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_TOKEN + '/sendMessage',
        { method: 'post', payload: { chat_id: TELEGRAM_CHAT_ID, text: subject + '\n\n' + body } });
    } catch (e) {}
  }
}

function register_(m) {
  const phone = s_(m.phone, 11);
  if (!/^01\d{9}$/.test(phone)) throw new Error('bad phone');
  const sh = sheet_('Customers');
  const now = new Date();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === phone) { // عميل موجود: نحدّث بياناته
      sh.getRange(i + 1, 2, 1, 3).setValues([[s_(m.name), s_(m.gov, 60) || data[i][2], s_(m.city, 80) || data[i][3]]]);
      if (m.addr) sh.getRange(i + 1, 5).setValue(s_(m.addr, 300));
      sh.getRange(i + 1, 7).setValue(now);
      return { ok: true, updated: true };
    }
  }
  sh.appendRow([phone, s_(m.name), s_(m.gov, 60), s_(m.city, 80), s_(m.addr, 300), now, now]);
  notify_('👤 عميل جديد: ' + s_(m.name), 'الاسم: ' + s_(m.name) + '\nالموبايل: ' + phone);
  return { ok: true };
}

function order_(o) {
  if (!o || !o.id || !o.items || !o.items.length) throw new Error('bad order');
  const sh = sheet_('Orders');
  const last = sh.getLastRow();
  if (last > 1) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
    if (ids.indexOf(String(o.id)) > -1) return { ok: true, duplicate: true }; // لا نكرر نفس الطلب
  }
  // نعيد حساب الأسعار من جدول المنتجات حتى لا يتلاعب أحد بها
  const prices = {};
  sheet_('Products').getDataRange().getValues().slice(1).forEach(r => { prices[r[0]] = Number(r[2]); });
  let sub = 0;
  const lines = o.items.slice(0, 50).map(i => {
    const price = prices[i.id] !== undefined ? prices[i.id] : Number(i.price) || 0;
    const qty = Math.max(1, Math.min(99, Number(i.qty) || 1));
    sub += price * qty;
    return '• ' + s_(i.name, 150) + ' × ' + qty + ' = ' + (price * qty) + ' جنيه';
  });
  const ship = Math.min(Math.max(Number(o.ship) || 0, 0), 200);
  const total = sub + ship;
  const pay = { cod: 'الدفع عند الاستلام', card: 'بطاقة بنكية', wallet: 'محفظة إلكترونية / إنستاباي' }[o.pay] || s_(o.pay, 40);
  const items = lines.join('\n');
  sh.appendRow([s_(o.id, 30), new Date(), s_(o.name), s_(o.phone, 11), s_(o.gov, 60), s_(o.city, 80),
    s_(o.addr, 300), s_(o.note, 500), items, sub, ship, total, pay, 'جديد', '', new Date(),
    JSON.stringify([{ s: 'جديد', t: new Date().toISOString(), n: '' }])]);
  notify_('🛒 طلب جديد ' + o.id + ' — ' + total + ' جنيه',
    'الاسم: ' + s_(o.name) + '\nالموبايل: ' + s_(o.phone, 11) +
    '\nالعنوان: ' + s_(o.gov, 60) + ' - ' + s_(o.city, 80) + ' - ' + s_(o.addr, 300) +
    (o.note ? '\nملاحظات: ' + s_(o.note, 500) : '') +
    '\n\nالمنتجات:\n' + items +
    '\n\nالمجموع: ' + sub + '\nالشحن: ' + ship + '\nالإجمالي: ' + total + ' جنيه\nالدفع: ' + pay);
  return { ok: true };
}

// ===== إدارة المنتجات (للبائع فقط) =====
function saveProduct_(p) {
  if (!p || !String(p.name || '').trim()) throw new Error('bad product');
  const sh = sheet_('Products');
  const id = Number(p.id) || Date.now();
  let img = String(p.img || '');
  if (img.indexOf('data:image') === 0) img = uploadImage_(img, id);
  const vals = [id, s_(p.name, 150), Number(p.price) || 0, s_(p.cat || 'عام', 60), s_(p.desc, 1000), img];
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === id) { // تعديل (لا نغيّر عمود "متوفر")
      sh.getRange(i + 1, 1, 1, 6).setValues([vals]);
      return { ok: true, product: { id: id, img: img } };
    }
  }
  sh.appendRow(vals.concat(['نعم']));
  return { ok: true, product: { id: id, img: img } };
}

function deleteProduct_(id) {
  const sh = sheet_('Products');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(id)) { sh.deleteRow(i + 1); break; }
  }
  return { ok: true };
}

function uploadImage_(dataUrl, id) {
  const m = dataUrl.match(/^data:(image\/[a-z+]+);base64,([A-Za-z0-9+\/=]+)$/);
  if (!m) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'product-' + id);
  const it = DriveApp.getFoldersByName('StoreOnline-Images');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('StoreOnline-Images');
  const f = folder.createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w800';
}

// ===== تتبع الطلبات والاسترجاع =====
// مقارنة مرنة: نتجاهل المسافات والأصفار الأولى (جوجل شيت قد يحذف الصفر من الموبايل)
function samePhone_(a, b) {
  return String(a).replace(/\D/g, '').slice(-10) === String(b).replace(/\D/g, '').slice(-10);
}
function sameId_(a, b) {
  return String(a).replace(/\s/g, '').toUpperCase() === String(b).replace(/\s/g, '').toUpperCase();
}
// التتبع يحتاج رقم الطلب + رقم الموبايل معاً
function track_(list) {
  list = (list || []).slice(0, 30);
  const orders = sheet_('Orders').getDataRange().getValues().slice(1);
  const rets = sheet_('Returns').getDataRange().getValues().slice(1);
  const results = list.map(function (q) {
    const id = String(q.id || '').trim().toUpperCase();
    const phone = String(q.phone || '').trim();
    const r = orders.filter(function (x) { return sameId_(x[0], id) && samePhone_(x[3], phone); })[0];
    if (!r) return { id: q.id, found: false };
    return {
      id: String(r[0]), found: true, date: r[1], status: r[13] || 'جديد', note: r[14] || '',
      updated: r[15] || r[1], total: r[11], items: r[8], history: hist_(r),
      returns: rets.filter(function (x) { return String(x[2]) === String(r[0]) && samePhone_(x[4], phone); })
        .map(function (x) { return { id: x[0], item: x[5], status: x[8] || 'قيد المراجعة', note: x[9] || '' }; })
    };
  });
  return { ok: true, results: results };
}

function return_(m) {
  const id = String(m.orderId || '').trim().toUpperCase();
  const phone = String(m.phone || '').trim();
  const order = sheet_('Orders').getDataRange().getValues().slice(1)
    .filter(function (x) { return sameId_(x[0], id) && samePhone_(x[3], phone); })[0];
  if (!order) throw new Error('order_not_found');
  const item = s_(m.item, 150), reason = s_(m.reason, 100);
  if (!item || !reason) throw new Error('bad return');
  const rid = 'RT-' + Date.now().toString().slice(-6);
  const now = new Date();
  sheet_('Returns').appendRow([rid, now, String(order[0]), s_(order[2]), phone, item, reason, s_(m.details, 800), 'قيد المراجعة', '', now]);
  notify_('↩️ طلب استرجاع ' + rid + ' للطلب ' + order[0],
    'الاسم: ' + s_(order[2]) + '\nالموبايل: ' + phone + '\nالمنتج: ' + item + '\nالسبب: ' + reason +
    (m.details ? '\nتفاصيل: ' + s_(m.details, 800) : ''));
  return { ok: true, id: rid };
}

// آخر n صفوف (للوحة البائع)
function rows_(name, n) {
  return sheet_(name).getDataRange().getValues().slice(1).reverse().slice(0, n)
    .map(function (r) { return r.map(function (v) { return v instanceof Date ? v.toISOString() : v; }); });
}

function setStatus_(m) {
  const isOrder = m.kind === 'orders';
  const sh = sheet_(isOrder ? 'Orders' : 'Returns');
  const col = isOrder ? 14 : 9; // عمود الحالة، ثم الملاحظة، ثم آخر تحديث
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(m.id)) {
      sh.getRange(i + 1, col).setValue(s_(m.status, 40));
      sh.getRange(i + 1, col + 1).setValue(s_(m.note, 300));
      sh.getRange(i + 1, col + 2).setValue(new Date());
      if (isOrder) pushHistory_(sh, i + 1, s_(m.status, 40), s_(m.note, 300));
      return { ok: true };
    }
  }
  throw new Error('not found');
}

// ===== سجل التتبع (يظهر للعميل كخط زمني مثل أمازون ونون) =====
function hist_(r) {
  let h = [];
  try { h = JSON.parse(r[16] || '[]'); } catch (e) {}
  if (!h.length) h = [{ s: 'جديد', t: r[1], n: '' }];
  const cur = r[13] || 'جديد';
  if (h[h.length - 1].s !== cur) h.push({ s: cur, t: r[15] || new Date(), n: r[14] || '' });
  return h;
}

function pushHistory_(sh, row, status, note) {
  const cell = sh.getRange(row, 17);
  let h = [];
  try { h = JSON.parse(cell.getValue() || '[]'); } catch (e) {}
  const last = h[h.length - 1];
  if (last && last.s === status && (last.n || '') === (note || '')) return;
  h.push({ s: status, t: new Date().toISOString(), n: note || '' });
  cell.setValue(JSON.stringify(h));
}

// لو غيّرت "الحالة" يدوياً من الجدول، يُسجَّل التحديث تلقائياً (يعمل تلقائياً بدون إعداد)
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== 'Orders' || e.range.getRow() < 2 || e.range.getColumn() !== 14) return;
    const row = e.range.getRow();
    sh.getRange(row, 16).setValue(new Date());
    pushHistory_(sh, row, String(e.range.getValue()), String(sh.getRange(row, 15).getValue() || ''));
  } catch (err) {}
}
