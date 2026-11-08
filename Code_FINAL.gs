/*************************************************************************************
 *  A'S STORE — Shopify → Google Sheets Auto Sync + Daily Cash Closing
 *  ---------------------------------------------------------------------------------
 *  يسحب المبيعات والمدفوعات والمرتجعات والمنتجات من Shopify تلقائياً إلى الشيت،
 *  ويحسب التقفيل اليومي (الكاش المتوقع / الفرق). المصاريف تبقى إدخال يدوي بموافقة الأدمن.
 *
 *  SECURITY: التوكن يُخزَّن في Script Properties فقط — لا يُكتب أبداً داخل خلايا الشيت.
 *  Setup steps are at the bottom of this file (اقرأ قسم SETUP).
 *************************************************************************************/

// ================================ CONFIG (ثابت) ====================================
const API_VERSION = '2026-10';                 // Shopify Admin API version
const TIMEZONE    = 'Asia/Amman';              // لاشتقاق BusinessDate
const CURRENCY    = 'JOD';
const DEFAULT_LOCATION = 'MAIN';               // LocationID الافتراضي
const FIRST_RUN_LOOKBACK_DAYS = 90;            // أول مزامنة: كم يوم للخلف
const COUNT_MANUAL_AS_CASH = false;            // هل "manual" يُحتسب كاش بالدرج؟ (غيّرها true لو بدك)

// تصنيف البند من اسمه (بدون read_products) بالكلمات المفتاحية:
// 1) إذا فيه كلمة إكسسوار (كفر/شاحن/كيبل/حماية...) → إكسسوار (حتى لو فيه "iPhone" مثل "Screen Protector for iPhone").
// 2) غير هيك، إذا فيه موديل جهاز (iPhone/iPad/MacBook...) → جهاز (مثل "iPhone 14 Black").
// 3) غير هيك → إكسسوار (الافتراضي).
const ACCESSORY_WORDS = ['case','cover','protector','screen','glass','tempered','cable','charger','adapter',
  'silicone','soundcore','anker','powerbank','power bank','bank','airpod','earphone','headphone','buds',
  'strap','band','holder','mount','stand','dock','microfon','usb','type-c','type c','lightning',
  'gorilla','uag','baseus','spigen','otterbox','wireless','magsafe'];
const DEVICE_WORDS = ['iphone','ipad','macbook','imac','mac mini','mac studio','apple watch'];
function classifyCategory(title) {
  const t = String(title || '').toLowerCase();
  for (var i = 0; i < ACCESSORY_WORDS.length; i++) if (t.indexOf(ACCESSORY_WORDS[i]) >= 0) return 'Accessory';
  for (var j = 0; j < DEVICE_WORDS.length; j++) if (t.indexOf(DEVICE_WORDS[j]) >= 0) return 'Device';
  return 'Accessory';
}

// خريطة بوابة الدفع → طريقة موحّدة (عدّلها لو ظهرت بوابة جديدة)
const GATEWAY_MAP = {
  'cash': 'Cash',
  'cliq': 'CliQ', 'CLIQ': 'CliQ',
  'visa': 'Card', 'VISA': 'Card', 'mastercard': 'Card', 'credit card': 'Card',
  'shopify payments': 'Card', 'bank card': 'Card', 'card': 'Card',
  'manual': 'Manual'
};

// أسماء التبويبات
const T = {
  orders: 'Orders', items: 'OrderItems', payments: 'Payments', refunds: 'Refunds',
  products: 'Products', closings: 'DailyClosings', expenses: 'Expenses',
  movements: 'CashMovements', sync: 'SyncEvents', config: 'Config'
};

// ============================== MENU + TRIGGERS ====================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("A'S STORE Sync")
    .addItem('🔄 مزامنة الآن (Sync now)', 'syncAll')
    .addItem('📅 إعادة حساب التقفيل (Recompute closings)', 'recomputeClosings')
    .addSeparator()
    .addItem('⏱️ تفعيل المزامنة التلقائية (كل 15 دقيقة)', 'installTrigger')
    .addItem('⏹️ إيقاف المزامنة التلقائية', 'removeTriggers')
    .addToUi();
}

function installTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('syncAll').timeBased().everyMinutes(15).create();
  SpreadsheetApp.getActive().toast('تم تفعيل المزامنة التلقائية كل 15 دقيقة ✅');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncAll') ScriptApp.deleteTrigger(t);
  });
}

// ================================ MAIN ENTRY ======================================
function syncAll() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { logSync('LOCK', '', 'SKIPPED', 'another sync running'); return; }
  try {
    syncOrders();      // Orders + OrderItems + Payments + Refunds (read_orders)
    try { syncProducts(); }  // Products (يحتاج read_products — لو مش متوفر نتخطاه)
    catch (pe) { logSync('products/sync', '', 'SKIPPED', 'need read_products scope? ' + pe); }
    recomputeClosings(); // DailyClosings computed fields
    SpreadsheetApp.getActive().toast('تمت المزامنة ✅');
  } catch (e) {
    logSync('FATAL', '', 'ERROR', String(e && e.stack || e));
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// ============================ SHOPIFY AUTH + GRAPHQL ===============================
// المصادقة بطريقة client_credentials (نفس طريقة testShopifyConnection الموجودة عندك):
// SHOPIFY_SHOP + SHOPIFY_CLIENT_ID + SHOPIFY_CLIENT_SECRET في Script Properties.
function shopDomain() {
  const d = PropertiesService.getScriptProperties().getProperty('SHOPIFY_SHOP');
  if (!d) throw new Error('SHOPIFY_SHOP غير موجود في Script Properties');
  return String(d).trim().toLowerCase().replace(/^https?:\/\//, '').replace(/\/.*$/, '');
}

// تطبيقات Dev Dashboard: نجيب توكن عبر client_credentials (CLIENT_ID + CLIENT_SECRET).
// مهم: الطلب لازم يكون form-urlencoded (payload كـ object) — مش JSON — وإلا 400 invalid_request.
var _ACCESS_TOKEN = null; // cache داخل نفس التشغيل
function getAccessToken() {
  if (_ACCESS_TOKEN) return _ACCESS_TOKEN;
  const cache = CacheService.getScriptCache();
  const cached = cache.get('SHOPIFY_ACCESS_TOKEN');
  if (cached) { _ACCESS_TOKEN = cached; return cached; }

  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHOPIFY_CLIENT_ID');
  const secret = props.getProperty('SHOPIFY_CLIENT_SECRET');
  if (!id || !secret) throw new Error('SHOPIFY_CLIENT_ID / SHOPIFY_CLIENT_SECRET ناقصين في Script Properties');

  const res = UrlFetchApp.fetch('https://' + shopDomain() + '/admin/oauth/access_token', {
    method: 'post',
    // payload كـ object → UrlFetchApp يرسله application/x-www-form-urlencoded تلقائياً
    payload: { client_id: id, client_secret: secret, grant_type: 'client_credentials' },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Auth failed ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
  }
  const body = JSON.parse(res.getContentText());
  if (!body.access_token) throw new Error('No access_token: ' + res.getContentText().slice(0, 300));
  _ACCESS_TOKEN = body.access_token;
  const ttl = Math.max(60, Math.min(21000, (body.expires_in || 86400) - 300)); // كاش آبس سكربت أقصاه 6 ساعات
  cache.put('SHOPIFY_ACCESS_TOKEN', _ACCESS_TOKEN, ttl);
  return _ACCESS_TOKEN;
}

function gql(query, variables) {
  const url = 'https://' + shopDomain() + '/admin/api/' + API_VERSION + '/graphql.json';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Shopify-Access-Token': getAccessToken() },
    payload: JSON.stringify({ query: query, variables: variables || {} }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText() || '{}');
  if (code !== 200 || body.errors) {
    throw new Error('Shopify API ' + code + ': ' + JSON.stringify(body.errors || body).slice(0, 500));
  }
  return body.data;
}

// اختبار سريع (زي الموجود عندك) — بيطبع عدد الأوردرات بالـ Execution log
function testConnection() {
  const data = gql('{ orders(first:3){ edges{ node{ name totalPriceSet{ shopMoney{ amount currencyCode } } } } } }', {});
  const n = data.orders.edges.length;
  Logger.log('SHOPIFY CONNECTION SUCCESS · orders sample: ' + n);
  if (n) Logger.log('first: ' + data.orders.edges[0].node.name);
  SpreadsheetApp.getActive().toast('اتصال Shopify ناجح ✅ (' + n + ' عينة)');
}

// ============================== SYNC ORDERS ========================================
const ORDERS_QUERY =
'query($cursor:String,$q:String){' +
'  orders(first:40,after:$cursor,query:$q,sortKey:UPDATED_AT){' +
'    pageInfo{hasNextPage endCursor}' +
'    edges{node{' +
'      id name createdAt processedAt updatedAt cancelledAt displayFinancialStatus currencyCode' +
'      totalPriceSet{shopMoney{amount}}' +
'      lineItems(first:50){edges{node{id title quantity sku ' +
'        discountedTotalSet{shopMoney{amount}}}}}' +
'      transactions{id kind status gateway processedAt ' +
'        parentTransaction{id} amountSet{shopMoney{amount}}}' +
'      refunds{id createdAt note totalRefundedSet{shopMoney{amount}}}' +
'    }}' +
'  }' +
'}';

function syncOrders() {
  const props = PropertiesService.getScriptProperties();
  const since = props.getProperty('ORDERS_SINCE') ||
    Utilities.formatDate(new Date(Date.now() - FIRST_RUN_LOOKBACK_DAYS * 864e5), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");

  const ordersRows = [], itemRows = [], payRows = [], refundRows = [];
  let cursor = null, hasNext = true, maxUpdated = since, count = 0;

  while (hasNext) {
    const data = gql(ORDERS_QUERY, { cursor: cursor, q: 'updated_at:>=' + since });
    const conn = data.orders;
    conn.edges.forEach(function (e) {
      const o = e.node; count++;
      const oid = numId(o.id);
      if (o.updatedAt > maxUpdated) maxUpdated = o.updatedAt;

      ordersRows.push({
        OrderID: oid, OrderNumber: o.name, LocationID: DEFAULT_LOCATION,
        CreatedAt: o.createdAt, ProcessedAt: o.processedAt, Currency: o.currencyCode || CURRENCY,
        TotalAmount: num(o.totalPriceSet), FinancialStatus: o.displayFinancialStatus,
        PaymentMethod: orderPaymentMethod(o.transactions),
        CancelledAt: o.cancelledAt || '', UpdatedAt: o.updatedAt
      });

      o.lineItems.edges.forEach(function (li) {
        const n = li.node;
        itemRows.push({
          LineItemID: numId(n.id), OrderID: oid,
          VariantID: '',
          ProductTitle: n.title,
          FinancialCategory: classifyCategory(n.title),  // Device/Accessory (من الاسم)
          Quantity: n.quantity, NetSalesAmount: num(n.discountedTotalSet),
          Currency: o.currencyCode || CURRENCY
        });
      });

      (o.transactions || []).forEach(function (tx) {
        payRows.push({
          TransactionID: numId(tx.id), OrderID: oid, LocationID: DEFAULT_LOCATION,
          ProcessedAt: tx.processedAt, Gateway: tx.gateway || '',
          NormalizedMethod: normalizeMethod(tx.gateway),
          Kind: tx.kind, Status: tx.status, Amount: num(tx.amountSet),
          Currency: o.currencyCode || CURRENCY,
          ParentTransactionID: tx.parentTransaction ? numId(tx.parentTransaction.id) : ''
        });
      });

      (o.refunds || []).forEach(function (r) {
        refundRows.push({
          RefundID: numId(r.id), OrderID: oid, ProcessedAt: r.createdAt,
          Amount: num(r.totalRefundedSet), Currency: o.currencyCode || CURRENCY,
          Reason: r.note || ''
        });
      });
    });
    hasNext = conn.pageInfo.hasNextPage;
    cursor = conn.pageInfo.endCursor;
  }

  upsert(T.orders,   'OrderID',       ordersRows);
upsert(T.items, 'LineItemID', itemRows);
  upsert(T.payments, 'TransactionID', payRows);
  upsert(T.refunds,  'RefundID',      refundRows);

  props.setProperty('ORDERS_SINCE', maxUpdated);
  logSync('orders/sync', '', 'OK', count + ' orders · ' + payRows.length + ' tx · ' + refundRows.length + ' refunds');
}

// ============================= SYNC PRODUCTS =======================================
const PRODUCTS_QUERY =
'query($cursor:String){' +
'  products(first:40,after:$cursor,sortKey:UPDATED_AT){' +
'    pageInfo{hasNextPage endCursor}' +
'    edges{node{id title productType updatedAt ' +
'      variants(first:50){edges{node{id sku}}}}}' +
'  }' +
'}';

function syncProducts() {
  const rows = [];
  let cursor = null, hasNext = true, count = 0;
  while (hasNext) {
    const data = gql(PRODUCTS_QUERY, { cursor: cursor });
    const conn = data.products;
    conn.edges.forEach(function (e) {
      const p = e.node; count++;
      const vs = p.variants.edges;
      if (vs.length === 0) {
        rows.push({ ProductID: numId(p.id), VariantID: '', SKU: '', Title: p.title,
                    FinancialCategory: p.productType || '', UpdatedAt: p.updatedAt });
      } else {
        vs.forEach(function (v) {
          rows.push({ ProductID: numId(p.id), VariantID: numId(v.node.id), SKU: v.node.sku || '',
                      Title: p.title, FinancialCategory: p.productType || '', UpdatedAt: p.updatedAt });
        });
      }
    });
    hasNext = conn.pageInfo.hasNextPage;
    cursor = conn.pageInfo.endCursor;
  }
  upsert(T.products, 'VariantID', rows.filter(function (r) { return r.VariantID; }));
  logSync('products/sync', '', 'OK', count + ' products');
}

// ======================= DAILY CLOSINGS (compute) =================================
function recomputeClosings() {
  const payments  = readAll(T.payments);
  const expenses  = readAll(T.expenses);
  const movements = readAll(T.movements);
  const orders    = readAll(T.orders);
  const items     = readAll(T.items);

  // خريطة الأوردر → (تاريخ العمل، الفرع) لربط بنود البيع بتاريخها
  const orderInfo = {};
  orders.forEach(function (o) {
    orderInfo[String(o.OrderID)] = { date: businessDate(o.ProcessedAt), loc: o.LocationID || DEFAULT_LOCATION };
  });

  // aggregate per (BusinessDate|Location)
  const agg = {};
  function bucket(date, loc) {
    const k = date + '|' + (loc || DEFAULT_LOCATION);
    if (!agg[k]) agg[k] = { date: date, loc: loc || DEFAULT_LOCATION,
      cashSales: 0, cliqSales: 0, cardSales: 0, manualSales: 0,
      deviceSales: 0, accessorySales: 0,
      cashRefunds: 0, expenses: 0, cashIn: 0, cashOut: 0 };
    return agg[k];
  }

  // مبيعات حسب النوع (جهاز/إكسسوار) من بنود الأوردرات
  items.forEach(function (it) {
    const info = orderInfo[String(it.OrderID)];
    if (!info || !info.date) return;
    const b = bucket(info.date, info.loc);
    const amt = Number(it.NetSalesAmount) || 0;
    if (String(it.FinancialCategory) === 'Device') b.deviceSales += amt;
    else b.accessorySales += amt;
  });

   // Net sales by payment method: sales minus successful refunds
  payments.forEach(function (p) {
    if (String(p.Status).toUpperCase() !== 'SUCCESS') return;

    const d = businessDate(p.ProcessedAt);
    if (!d) return;

    const b = bucket(d, p.LocationID);
    const amt = Math.abs(Number(p.Amount) || 0);
    const kind = String(p.Kind).toUpperCase();
    const m = String(p.NormalizedMethod);

    const isSale = (kind === 'SALE' || kind === 'CAPTURE');
    const isRefund = (kind === 'REFUND');

    if (isSale) {
      if (m === 'Cash') b.cashSales += amt;
      else if (m === 'CliQ') b.cliqSales += amt;
      else if (m === 'Card') b.cardSales += amt;
      else {
        b.manualSales += amt;
        if (COUNT_MANUAL_AS_CASH) b.cashSales += amt;
      }
    } else if (isRefund) {
      // Subtract each refund from its matching payment method
      if (m === 'Cash') {
        b.cashSales -= amt;
        b.cashRefunds += amt;
      } else if (m === 'CliQ') {
        b.cliqSales -= amt;
      } else if (m === 'Card') {
        b.cardSales -= amt;
      } else {
        b.manualSales -= amt;
        if (COUNT_MANUAL_AS_CASH) b.cashSales -= amt;
      }
    }
  });

  // المصاريف: تُحتسب المعلّقة + الموافَق عليها، وتُستثنى المرفوضة فقط (REJECTED)
  expenses.forEach(function (x) {
    if (String(x.ApprovalStatus).toUpperCase() === 'REJECTED') return;
    const d = fmtDate(x.BusinessDate); if (!d) return;
    bucket(d, x.LocationID).expenses += Number(x.AmountJOD) || 0;
  });

  // cash movements
  movements.forEach(function (mv) {
    const d = fmtDate(mv.BusinessDate); if (!d) return;
    const b = bucket(d, mv.LocationID);
    const amt = Number(mv.AmountJOD) || 0;
    if (String(mv.MovementType).toUpperCase().indexOf('OUT') >= 0) b.cashOut += amt;
    else b.cashIn += amt;
  });

  // write into DailyClosings, preserving staff-entered OpeningCash & CountedCash
  const sh = sheet(T.closings);
  const H = headers(sh);
  const existing = readAll(T.closings);
  const index = {};
  existing.forEach(function (r, i) { index[fmtDate(r.BusinessDate) + '|' + r.LocationID] = i; });

  const data = sh.getDataRange().getValues(); // includes header row at 0
  Object.keys(agg).forEach(function (k) {
    const b = agg[k];
    const opening = findNum(existing, index[k], 'OpeningCashJOD');
    const counted = findVal(existing, index[k], 'CountedCashJOD');
    const benefit = findNum(existing, index[k], 'BenefitJOD'); // يدوي: دخل التصليح (كاش)
    // الكاش المتوقع بالدرج = افتتاحي + مبيعات كاش + بنفت(تصليح كاش) - مرتجعات كاش - مصاريف + داخل - خارج
    const expected = round3(opening + b.cashSales + benefit - b.cashRefunds - b.expenses + b.cashIn - b.cashOut);
    const variance = (counted === '' || counted === null) ? '' : round3(Number(counted) - expected);

    const vals = {
      LocationID: b.loc, BusinessDate: b.date,
      CashSalesJOD: round3(b.cashSales), CliqSalesJOD: round3(b.cliqSales),
      CardSalesJOD: round3(b.cardSales), ManualSalesJOD: round3(b.manualSales),
      DeviceSalesJOD: round3(b.deviceSales), AccessorySalesJOD: round3(b.accessorySales),
      CashRefundsJOD: round3(b.cashRefunds), ExpensesJOD: round3(b.expenses),
      CashInJOD: round3(b.cashIn), CashOutJOD: round3(b.cashOut),
      ExpectedCashJOD: expected, VarianceJOD: variance
    };

    if (index[k] === undefined) {
      // new row → حالة ابتدائية "Submitted" (تنتظر موافقة الأدمن)
      const row = new Array(data[0].length).fill('');
      setCell(row, H, 'ClosingID', b.date + '_' + b.loc);
      setCell(row, H, 'Status', 'Submitted');
      Object.keys(vals).forEach(function (c) { setCell(row, H, c, vals[c]); });
      data.push(row);
      index[k] = data.length - 2; // relative to existing (minus header)
    } else {
      const rowIdx = index[k] + 1; // +1 for header
      Object.keys(vals).forEach(function (c) { setCell(data[rowIdx], H, c, vals[c]); });
      // اضبط الحالة "Submitted" فقط إذا فاضية — لا تلمس "Approved" أو غيرها
      const curStatus = findVal(existing, index[k], 'Status');
      if (curStatus === '' || curStatus === null || curStatus === undefined) {
        setCell(data[rowIdx], H, 'Status', 'Submitted');
      }
    }
  });

  sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  logSync('closings/compute', '', 'OK', Object.keys(agg).length + ' day/location buckets');
}

// ============================== HELPERS ===========================================
function numId(gid) { return gid ? String(gid).split('/').pop() : ''; }
function num(moneySet) { return moneySet && moneySet.shopMoney ? Number(moneySet.shopMoney.amount) : 0; }
function round3(n) { return Math.round((Number(n) || 0) * 1000) / 1000; }

/** يحدد طريقة دفع الأوردر من حركاته: طريقة وحدة، أو "Mixed" لو أكثر من وحدة */
function orderPaymentMethod(transactions) {
  const set = {};
  (transactions || []).forEach(function (tx) {
    const kind = String(tx.kind).toUpperCase();
    if ((kind === 'SALE' || kind === 'CAPTURE') && String(tx.status).toUpperCase() === 'SUCCESS') {
      set[normalizeMethod(tx.gateway)] = true;
    }
  });
  const methods = Object.keys(set);
  if (methods.length === 0) return '';
  if (methods.length === 1) return methods[0];
  return 'Mixed';
}

function normalizeMethod(gateway) {
  if (!gateway) return 'Other';
  const key = String(gateway).toLowerCase();
  return GATEWAY_MAP[key] || GATEWAY_MAP[gateway] || 'Other';
}

// اليوم التجاري يبدأ 10 الصبح ويمتد لـ 2 بعد منتصف الليل.
// أي أوردر قبل الساعة 2 بالليل بينحسب على اليوم اللي قبل (مش اليوم الجديد).
var DAY_CUTOFF_HOUR = 2;
function businessDate(iso) {
  if (!iso) return '';
  var dt = new Date(iso);
  if (isNaN(dt.getTime())) return '';
  dt = new Date(dt.getTime() - DAY_CUTOFF_HOUR * 3600 * 1000);  // اطرح ساعتين قبل اشتقاق التاريخ
  return Utilities.formatDate(dt, TIMEZONE, 'yyyy-MM-dd');
}
function fmtDate(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd');
  const s = String(v);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd');
}

function sheet(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('التبويب غير موجود: ' + name);
  return sh;
}
function headers(sh) {
  const row = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  row.forEach(function (h, i) { if (h !== '' && h !== null) map[String(h).trim()] = i; });
  return map;
}
function readAll(name) {
  const sh = sheet(name);
  const rng = sh.getDataRange().getValues();
  if (rng.length < 2) return [];
  const head = rng[0];
  return rng.slice(1).map(function (r) {
    const o = {}; head.forEach(function (h, i) { if (h) o[String(h).trim()] = r[i]; }); return o;
  });
}
function setCell(rowArr, H, colName, val) { if (H[colName] !== undefined) rowArr[H[colName]] = val; }
function findNum(rows, i, col) { return (i === undefined || !rows[i]) ? 0 : (Number(rows[i][col]) || 0); }
function findVal(rows, i, col) { return (i === undefined || !rows[i]) ? '' : rows[i][col]; }

/**
 * Upsert rows (array of objects) into a sheet, keyed by keyCol. Header-driven.
 * Existing keys are updated in place; new keys appended. Columns not present are ignored.
 */
function upsert(name, keyCol, rows) {
  if (!rows || rows.length === 0) return;
  const sh = sheet(name);
  const H = headers(sh);
  if (H[keyCol] === undefined) throw new Error(name + ': لا يوجد عمود المفتاح ' + keyCol);
  const data = sh.getDataRange().getValues();
  const width = data[0].length;
  const keyIdx = H[keyCol];
  const pos = {};
  for (var i = 1; i < data.length; i++) pos[String(data[i][keyIdx])] = i;

  rows.forEach(function (obj) {
    const key = String(obj[keyCol]);
    var rowArr;
    if (pos[key] !== undefined) rowArr = data[pos[key]];
    else { rowArr = new Array(width).fill(''); data.push(rowArr); pos[key] = data.length - 1; }
    Object.keys(obj).forEach(function (c) { if (H[c] !== undefined) rowArr[H[c]] = obj[c]; });
  });

  sh.getRange(1, 1, data.length, width).setValues(data);
}

function logSync(topic, resourceId, status, msg) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(T.sync);
    if (!sh) return;
    const H = headers(sh);
    const row = new Array(sh.getLastColumn()).fill('');
    setCell(row, H, 'EventID', Utilities.getUuid().slice(0, 8));
    setCell(row, H, 'Topic', topic);
    setCell(row, H, 'ResourceID', resourceId);
    setCell(row, H, 'ReceivedAt', new Date());
    setCell(row, H, 'ProcessedAt', new Date());
    setCell(row, H, 'Status', status);
    setCell(row, H, 'ErrorMessage', String(msg || '').slice(0, 500));
    sh.appendRow(row);
  } catch (e) { /* no-op */ }
}
const EXPENSE_OWNER_EMAIL = 'asabualenan@gmail.com';
const EXPENSE_SHEET_NAME = 'Expenses';

function notifyOwnerOnExpenseEdit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== EXPENSE_SHEET_NAME) return;

  const row = e.range.getRow();
  if (row < 2) return;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(h => String(h).trim());

  const col = name => headers.indexOf(name) + 1;

  const statusCol = col('ApprovalStatus');
  const amountCol = col('AmountJOD');
  const categoryCol = col('Category');
  const dateCol = col('BusinessDate');
  const descriptionCol = col('Description');
  const notifiedCol = col('OwnerEmailSentAt');

  if (!statusCol || !amountCol || !notifiedCol) {
    throw new Error(
      'تأكد من وجود أعمدة ApprovalStatus و AmountJOD، وأضف عمود OwnerEmailSentAt'
    );
  }

  const status = String(sh.getRange(row, statusCol).getDisplayValue()).toUpperCase();
  const alreadySent = sh.getRange(row, notifiedCol).getValue();

  if (status !== 'PENDING' || alreadySent) return;

  const getValue = c => c ? sh.getRange(row, c).getDisplayValue() : '';

  const subject = 'New Expense Request - Approval Required';

  const body =
    'A new expense request requires your approval.\n\n' +
    'Date: ' + getValue(dateCol) + '\n' +
    'Category: ' + getValue(categoryCol) + '\n' +
    'Amount: ' + getValue(amountCol) + ' JOD\n' +
    'Description: ' + getValue(descriptionCol) + '\n' +
    'Status: ' + status + '\n\n' +
    'Please open AS STORE DAILY CLOSE and review the request.';

  MailApp.sendEmail(EXPENSE_OWNER_EMAIL, subject, body);

  sh.getRange(row, notifiedCol).setValue(new Date());
}

function installExpenseEmailTrigger() {
  const ss = SpreadsheetApp.getActive();

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'notifyOwnerOnExpenseEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('notifyOwnerOnExpenseEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert('تم تفعيل إشعار الإيميل للمصاريف.');
}

function notifyPendingExpensesByTimer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Expenses');

  if (!sh || sh.getLastRow() < 2) return;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(h => String(h).trim());

  const col = name => headers.indexOf(name) + 1;

  const statusCol = col('ApprovalStatus');
  const amountCol = col('AmountJOD');
  const dateCol = col('BusinessDate');
  const categoryCol = col('Category');
  const descriptionCol = col('Description');
  const notifiedCol = col('OwnerEmailSentAt');

  if (!statusCol || !amountCol || !notifiedCol) {
    throw new Error(
      'Missing column: ApprovalStatus, AmountJOD, or OwnerEmailSentAt'
    );
  }

  const lastRow = sh.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    const status = String(
      sh.getRange(row, statusCol).getDisplayValue()
    ).trim().toUpperCase();

    const alreadySent = sh.getRange(row, notifiedCol).getValue();

    if (status !== 'PENDING' || alreadySent) continue;

    const getValue = c =>
      c ? sh.getRange(row, c).getDisplayValue() : '';

    const expenseId = getValue(col('ExpenseID'));

    const subject = 'New Expense Request - Approval Required';

    const body =
      'A new expense request requires your approval.\n\n' +
      'Expense ID: ' + expenseId + '\n' +
      'Date: ' + getValue(dateCol) + '\n' +
      'Category: ' + getValue(categoryCol) + '\n' +
      'Amount: ' + getValue(amountCol) + ' JOD\n' +
      'Description: ' + getValue(descriptionCol) + '\n' +
      'Status: ' + status + '\n\n' +
      'Please review the expense in AS STORE DAILY CLOSE.';

    MailApp.sendEmail({
      to: 'asabualenan@gmail.com',
      subject: subject,
      body: body
    });

    sh.getRange(row, notifiedCol).setValue(new Date());
    SpreadsheetApp.flush();

    Logger.log('EMAIL SENT | Row ' + row + ' | ExpenseID=' + expenseId);
  }
}


function installExpenseEmailTimer() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Expenses');
  if (!sh) throw new Error('Expenses tab not found');

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(h => String(h).trim());

  const notifiedCol = headers.indexOf('OwnerEmailSentAt') + 1;
  if (!notifiedCol) {
    throw new Error('أضف عمود OwnerEmailSentAt أولاً');
  }

  // علّم الصفوف القديمة حتى لا يرسل إشعارات عنها
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const range = sh.getRange(2, notifiedCol, lastRow - 1, 1);
    const values = range.getValues();

    for (let i = 0; i < values.length; i++) {
      if (!values[i][0]) values[i][0] = 'EXISTING';
    }

    range.setValues(values);
  }

  // احذف التريغرات القديمة لنفس وظيفة الإشعار
  ScriptApp.getProjectTriggers().forEach(trigger => {
    const fn = trigger.getHandlerFunction();
    if (fn === 'notifyOwnerOnExpenseEdit' ||
        fn === 'notifyPendingExpensesByTimer') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('notifyPendingExpensesByTimer')
    .timeBased()
    .everyMinutes(1)
    .create();

  SpreadsheetApp.getUi().alert('تم تفعيل مراقبة المصاريف. الإيميل يصل تلقائياً.');
}

// ===== جرد المخزون: صباحي/مسائي + موافقة المسؤول =====
// استبدل فيه كامل قسم الجرد القديم (من "// ===== جرد المخزون ..." لين نهاية fixSummary).
// ما بيلمس المزامنة ولا المصاريف.

const CFG = {
  MISSING_TAB: 'Missing',    // تابة المخزون المرجعية (قراءة فقط)
  SERIAL_COL: 2,             // B = Serial number داخل Missing
  SCANS_TAB: 'Scans',        // اللي انمسح
  OUT_TAB: 'Unscanned',      // الناقص (AppSheet بيقرا منها)
  SUM_TAB: 'Summary',        // عداد الشاشة
  CHECK_TAB: 'InventoryChecks' // 🆕 سجل جلسات الجرد (صباحي/مسائي) + الموافقة
};
const OUT_HEAD  = ['Serial','Model','Color','Storage','Battery','Notes','ExtraNote','Price','Reason','Note','InvDate'];
const SUM_HEAD  = ['ID','Total','Scanned','Missing','UpdatedAt','InvDate','Approved','ApprovedBy','ApprovedAt'];
const CHECK_HEAD= ['CheckID','BusinessDate','Session','Total','Scanned','Missing','Status','Reason','ApprovedBy','ApprovedAt','CreatedAt'];

const norm = v => String(v == null ? '' : v).trim().toUpperCase();
// سيريال حقيقي: ما فيه مسافات وفيه رقم واحد على الأقل (يستبعد "windows" و"not work")
const validSerial = s => /^(?=.*\d)[A-Z0-9]{6,}$/.test(s);

// اليوم التجاري (نفس منطق businessDate: قطع الساعة 2 بالليل)
function jardBusinessDate_() {
  var cut = (typeof DAY_CUTOFF_HOUR !== 'undefined') ? DAY_CUTOFF_HOUR : 2;
  var tz  = (typeof TIMEZONE !== 'undefined') ? TIMEZONE : 'Asia/Amman';
  var dt  = new Date(Date.now() - cut * 3600 * 1000);
  return Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
}
function normSession_(s) {
  s = String(s || '').trim().toLowerCase();
  if (s.indexOf('even') >= 0 || s.indexOf('مسائ') >= 0 || s === 'pm' || s === 'e') return 'Evening';
  return 'Morning';
}
function arSession_(s) { return s === 'Evening' ? 'المسائي' : 'الصباحي'; }

function out_(e, obj) {
  const cb = e && e.parameter && e.parameter.callback;
  return ContentService
    .createTextOutput(cb ? cb + '(' + JSON.stringify(obj) + ')' : JSON.stringify(obj))
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function tab_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]);
    sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@'); // السيريال نص
  }
  return sh;
}

// كل سيريالات Missing: { NORM: row[] }
function missingMap_() {
  const m = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.MISSING_TAB);
  const map = {};
  if (!m || m.getLastRow() < 2) return map;
  m.getRange(2, 1, m.getLastRow() - 1, 8).getValues().forEach(r => {
    const s = norm(r[CFG.SERIAL_COL - 1]);
    if (validSerial(s)) map[s] = r;
  });
  return map;
}

function scannedSet_() {
  const sh = tab_(CFG.SCANS_TAB, ['Serial', 'ScannedAt']);
  const set = {};
  if (sh.getLastRow() > 1)
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => { set[norm(r[0])] = true; });
  return set;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'reset')  return out_(e, resetInventory());
  if (p.action === 'finish') return out_(e, finishInventory(p.session, p.reason));
  return out_(e, scanSerial(p.serial));
}

function scanSerial(raw) {
  const s = norm(raw);
  if (!s) return { text: 'سيريال فاضي', cls: 'bad' };
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const row = missingMap_()[s];
    if (!row) return { text: '❌ ' + s + ' — مش موجود بقائمة Missing', cls: 'bad' };
    if (scannedSet_()[s]) return { text: '⚠️ ' + s + ' — مَسحته قبل', cls: 'dup' };
    const sh = tab_(CFG.SCANS_TAB, ['Serial', 'ScannedAt']);
    sh.getRange(sh.getLastRow() + 1, 1, 1, 2).setValues([[s, new Date()]]);
    return { text: '✅ ' + s + ' — ' + row[0], cls: 'ok' };
  } catch (err) {
    return { text: 'خطأ: ' + err, cls: 'bad' };
  } finally { lock.releaseLock(); }
}

function resetInventory() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    [CFG.SCANS_TAB, CFG.OUT_TAB].forEach((name, i) => {
      const sh = i ? tab_(name, OUT_HEAD) : tab_(name, ['Serial', 'ScannedAt']);
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    });
    writeSummary_(0, 0, 0);
    return { text: '♻️ تم تصفير الجرد', cls: 'ok' };
  } catch (err) {
    return { text: 'خطأ: ' + err, cls: 'bad' };
  } finally { lock.releaseLock(); }
}

function finishInventory(session, reason) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sess  = normSession_(session);
    const bdate = jardBusinessDate_();
    const miss  = missingMap_();
    const scanned = scannedSet_();
    const total = Object.keys(miss).length;
    const nScanned = Object.keys(scanned).filter(s => miss[s]).length;
    if (!nScanned) return { text: 'ما في ولا جهاز ممسوح — ما بنفع أنهي', cls: 'bad' };

    // ابنِ قائمة الناقص (نحافظ على الأسباب المدخلة سابقاً بـ Unscanned)
    const sh = tab_(CFG.OUT_TAB, OUT_HEAD);
    const old = {};
    if (sh.getLastRow() > 1)
      sh.getRange(2, 1, sh.getLastRow() - 1, OUT_HEAD.length).getValues().forEach(r => { old[norm(r[0])] = r; });

    const today = new Date();
    const rows = [];
    Object.keys(miss).forEach(s => {
      if (scanned[s]) return;
      const r = miss[s], o = old[s];
      rows.push([s, r[0], r[2], r[3], r[4], r[5], r[6], r[7], o ? o[8] : '', o ? o[9] : '', o ? o[10] : today]);
    });
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, OUT_HEAD.length).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, OUT_HEAD.length).setValues(rows);

    const missing = rows.length;
    const status  = (missing === 0) ? 'Approved' : 'PendingApproval';
    upsertCheck_(bdate, sess, total, nScanned, missing, status, reason || '');

    writeSummary_(total, nScanned, missing, true);

    const msg = (missing === 0)
      ? ('✅ جرد ' + arSession_(sess) + ' مطابق — تمت الموافقة تلقائياً (' + nScanned + '/' + total + ')')
      : ('⚠️ جرد ' + arSession_(sess) + ' ناقص ' + missing + ' — بانتظار موافقة المسؤول');
    return { text: msg, cls: (missing === 0 ? 'ok' : 'dup'), missing: missing, total: total, scanned: nScanned, status: status, session: sess };
  } catch (err) {
    return { text: 'خطأ: ' + err, cls: 'bad' };
  } finally { lock.releaseLock(); }
}

// سجل/حدّث جلسة الجرد. ما يدُس موافقة/رفض اتخذها المسؤول يدوياً.
function upsertCheck_(bdate, sess, total, scanned, missing, status, reason) {
  const sh = tab_(CFG.CHECK_TAB, CHECK_HEAD);
  const id = bdate + '_' + sess;
  const data = sh.getDataRange().getValues();
  let ri = -1;
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === id) { ri = i; break; } }

  let outStatus = status, by = '', at = '', rsn = reason;
  if (missing === 0) { by = 'AUTO'; at = new Date(); }     // مطابق = موافقة تلقائية
  if (ri >= 0) {
    const cur = String(data[ri][6] || '');
    if ((cur === 'Approved' || cur === 'Rejected') && missing > 0) { // احترم قرار المسؤول اليدوي
      outStatus = cur; by = data[ri][8]; at = data[ri][9];
    }
    if (!reason && data[ri][7]) rsn = data[ri][7];           // احفظ السبب القديم
  }
  const rowVals = [id, bdate, sess, total, scanned, missing, outStatus, rsn, by, at, new Date()];
  if (ri >= 0) sh.getRange(ri + 1, 1, 1, CHECK_HEAD.length).setValues([rowVals]);
  else sh.appendRow(rowVals);
}

const SUM_HEAD_ = SUM_HEAD; // توافق
function writeSummary_(total, scanned, missing, finished) {
  const sh = tab_(CFG.SUM_TAB, SUM_HEAD);
  sh.getRange(1, 1, 1, SUM_HEAD.length).setValues([SUM_HEAD]);
  sh.getRange(2, 1, 1, 9).setValues([['1', total, scanned, missing, new Date(), finished ? new Date() : '', false, '', '']]);
  sh.getRange(2, 6).setNumberFormat('yyyy-mm-dd');
}

// تنظيف تلقائي: يشيل من Unscanned أي جهاز انباع/انحذف من الأصلي (مش موجود بـ Missing)
function pruneUnscanned() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.OUT_TAB);
  if (!sh || sh.getLastRow() < 2) return;
  const miss = missingMap_();
  if (!Object.keys(miss).length) return;
  const ser = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = ser.length - 1; i >= 0; i--) if (!miss[norm(ser[i][0])]) sh.deleteRow(i + 2);
  const left = Math.max(sh.getLastRow() - 1, 0);
  const sum = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SUM_TAB);
  if (sum && sum.getLastRow() > 1) sum.getRange(2, 4).setValue(left);
}

function installTrigger() { // شغّلها مرة وحدة (نفس اسمها القديم)
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'pruneUnscanned') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pruneUnscanned').timeBased().everyMinutes(5).create();
}
function setupTabs() {
  tab_(CFG.SCANS_TAB, ['Serial', 'ScannedAt']);
  tab_(CFG.OUT_TAB, OUT_HEAD);
  tab_(CFG.CHECK_TAB, CHECK_HEAD);
  writeSummary_(0, 0, 0);
}
function fixSummary() {
  const sh = tab_(CFG.SUM_TAB, SUM_HEAD);
  sh.getRange(1, 1, 1, SUM_HEAD.length).setValues([SUM_HEAD]);
  sh.getRange(2, 7).setValue(false);
  sh.getRange(2, 6).setNumberFormat('yyyy-mm-dd');
}

/*************************************************************************************
 *  SETUP — خطوات التفعيل (مرة وحدة)
 *  المصادقة: client_credentials لتطبيق Dev Dashboard (CLIENT_ID + CLIENT_SECRET).
 *  1) من الشيت: Extensions → Apps Script. الصق هذا الملف كاملاً فوق الكود القديم واحفظ.
 *  2) Script Properties (Project Settings ⚙️ → Script properties) — نفس اللي عندك:
 *        SHOPIFY_SHOP          =  as-store-97300.myshopify.com
 *        SHOPIFY_CLIENT_ID     =  <Client ID من Dev Dashboard → Settings>
 *        SHOPIFY_CLIENT_SECRET =  <Client secret>
 *     (مش لازم SHOPIFY_TOKEN — هالنوع من التطبيقات ما بيعطي توكن ثابت.)
 *     تأكد التطبيق عنده صلاحيات: read_orders, read_products, read_locations.
 *  3) اختبر: اختر فنكشن "testConnection" واضغط Run (لازم يطبع SUCCESS).
 *  4) ارجع للشيت، حدّث الصفحة → منيو "A'S STORE Sync" → "مزامنة الآن".
 *  5) بعد نجاح المزامنة → "تفعيل المزامنة التلقائية" (كل 15 دقيقة).
 *
 *  ملاحظات:
 *   - جداول Orders/OrderItems/Payments/Refunds/Products = قراءة فقط (يملأها السكربت).
 *   - Expenses يدوي 100% مع موافقة الأدمن (ما بلمسه السكربت).
 *   - DailyClosings: السكربت يحسب المبيعات والمتوقع والفرق، وأنت/الموظف تدخل
 *     OpeningCashJOD و CountedCashJOD فقط، والباقي تلقائي.
 *   - "manual" غير محتسب كاش افتراضياً. لو بدك تحتسبه: COUNT_MANUAL_AS_CASH = true.
 *************************************************************************************/