/**
 * Newspaper Customer Management - Final (Annual-prepaid logic)
 *
 * Run:
 *   npm install
 *   npm start
 *
 * Notes:
 * - Monthly subscribers: billed using subscription.monthly_rate (default 295)
 * - Sunday-only subscribers: billed ₹10 × number_of_sundays_in_month
 * - Annual (prepaid) subscribers: NOT billed monthly (flag is_annual = 1)
 * - Carry-forward handled by payments table (balance)
 */

const express = require('express');
const session = require('express-session');
const bodyParser = require('body-parser');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({ secret: 'secretKey', resave: false, saveUninitialized: true }));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// ---------- DB & schema ----------
const DB_FILE = path.join(__dirname, 'newspaper.db');
const db = new sqlite3.Database(DB_FILE);

db.serialize(() => {
  // Core tables
  db.run(`CREATE TABLE IF NOT EXISTS customers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    address TEXT,
    mobile TEXT,
    area_name TEXT,
    route_name TEXT,
    active INTEGER DEFAULT 1
  )`);

  db.run(`CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE,
    item_type TEXT,      -- 'newspaper'|'magazine' (informational)
    monthly_rate REAL,
    active INTEGER DEFAULT 1
  )`);

  // subscriptions: one per customer item (supports multi-items)
  db.run(`CREATE TABLE IF NOT EXISTS subscriptions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_id INTEGER,
    item_id INTEGER,
    item_name TEXT,
    monthly_rate REAL,
    start_date TEXT,
    end_date TEXT,
    active INTEGER DEFAULT 1
  )`);

  db.run(`CREATE TABLE IF NOT EXISTS payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    subscription_id INTEGER,
    customer_id INTEGER,
    month TEXT,            -- 'YYYY-MM' (billing month)
    amount_due REAL,
    amount_paid REAL DEFAULT 0,
    balance REAL DEFAULT 0,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
  )`);

  db.run(`CREATE TABLE IF NOT EXISTS delivery_boys (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    mobile TEXT
  )`);

  db.run(`CREATE TABLE IF NOT EXISTS delivery_assignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    delivery_boy_id INTEGER,
    customer_id INTEGER,
    assigned_at TEXT DEFAULT CURRENT_TIMESTAMP
  )`);

  // settings: store global per-paper rate for deliveries (not used for sunday-only billing; preserved for other uses)
  db.run(`CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)`);

  // ensure default item (optional) and default setting for per_paper_rate
  db.get(`SELECT value FROM settings WHERE key='per_paper_rate'`, [], (err, row) => {
    if (!row) db.run(`INSERT INTO settings (key, value) VALUES ('per_paper_rate','5')`);
  });

  // migration: if older DB didn't have columns, ensure columns exist safely
  function addColumnIfNotExists(table, column, definition, callback) {
    db.all(`PRAGMA table_info(${table})`, [], (err, rows) => {
      if (err) return callback(err);
      const exists = rows.some(r => r.name === column);
      if (exists) return callback();
      db.run(`ALTER TABLE ${table} ADD COLUMN ${column} ${definition}`, callback);
    });
  }

  addColumnIfNotExists('items', 'active', 'INTEGER DEFAULT 1', (errItems) => {
    if (errItems) console.error('Migration error adding items.active:', errItems.message);
    db.run(`UPDATE items SET active = 1 WHERE active IS NULL`, () => {
      addColumnIfNotExists('subscriptions', 'is_annual', 'INTEGER DEFAULT 0', (err) => {
        if (err) console.error('Migration error adding is_annual:', err.message);
        addColumnIfNotExists('subscriptions', 'is_sunday_only', 'INTEGER DEFAULT 0', (err2) => {
          if (err2) console.error('Migration error adding is_sunday_only:', err2.message);
          console.log('Database ready:', DB_FILE);
        });
      });
    });
  });
});

// ---------- Helpers ----------
function pad2(n) { return n.toString().padStart(2, '0'); }

function countSundaysInMonth(year, month) {
  // month: 1-12
  const daysInMonth = new Date(year, month, 0).getDate();
  let count = 0;
  for (let d = 1; d <= daysInMonth; d++) {
    const dt = new Date(year, month - 1, d);
    if (dt.getDay() === 0) count++; // 0 = Sunday
  }
  return count;
}

/** Same fee rules as processOneMonthBilling for this payment month (YYYY-MM). */
function subscriptionFeeForPaymentMonth(row) {
  if (!row || row.subscription_id == null) return null;
  if (Number(row.subscription_is_annual) === 1) return 0;
  const parts = String(row.month || '').split('-').map(Number);
  const year = parts[0];
  const monthNum = parts[1];
  if (!year || !monthNum || monthNum < 1 || monthNum > 12) return null;
  if (Number(row.subscription_is_sunday_only) === 1) {
    return countSundaysInMonth(year, monthNum) * 10;
  }
  const rate = row.subscription_monthly_rate;
  return Number(rate != null && rate !== '' ? rate : 295);
}

function getMonthKeyFromDate(d) {
  const dt = d ? new Date(d) : new Date();
  return `${dt.getFullYear()}-${pad2(dt.getMonth()+1)}`;
}
function getCurrentMonthKey() {
  const dt = new Date();
  return `${dt.getFullYear()}-${pad2(dt.getMonth()+1)}`;
}

// find last payment balance for a subscription (carry forward)
function getLastBalance(subscriptionId) {
  return new Promise((resolve, reject) => {
    db.get(`SELECT balance FROM payments WHERE subscription_id = ? ORDER BY month DESC LIMIT 1`, [subscriptionId], (err, row) => {
      if (err) return resolve(0);
      resolve(row ? Number(row.balance || 0) : 0);
    });
  });
}

/** First day of calendar month after adding `delta` months to YYYY-MM. */
function addCalendarMonthsToKey(monthKey, delta) {
  const [y, m] = monthKey.split('-').map(Number);
  if (!y || !m || m < 1 || m > 12) return null;
  const d = new Date(y, m - 1, 1);
  d.setMonth(d.getMonth() + delta);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
}

/** Inclusive list of YYYY-MM keys: monthly = 1, quarterly = 3, yearly = 12. */
function periodMonthKeys(startKey, frequency) {
  const count =
    frequency === 'quarterly' ? 3 : frequency === 'yearly' ? 12 : 1;
  const keys = [];
  for (let i = 0; i < count; i++) {
    const k = addCalendarMonthsToKey(startKey, i);
    if (!k) return [];
    keys.push(k);
  }
  return keys;
}

/** Oldest → newest, `n` calendar months ending in `endMonthKey` (inclusive). */
function lastNMonthKeysEndingAt(endMonthKey, n) {
  const keys = [];
  for (let i = n - 1; i >= 0; i--) {
    const k = addCalendarMonthsToKey(endMonthKey, -i);
    if (!k) return [];
    keys.push(k);
  }
  return keys;
}

function csvEscapeCell(val) {
  const s = val == null ? '' : String(val);
  if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

/** Escape `%`, `_`, `\` for SQL LIKE patterns. */
function escapeLikePattern(s) {
  return String(s).replace(/\\/g, '\\\\').replace(/%/g, '\\%').replace(/_/g, '\\_');
}

function maxMonthKey(a, b) {
  return a >= b ? a : b;
}

/** First day of month as YYYY-MM from SQLite date 'YYYY-MM-DD'. */
function monthKeyFromStartDateString(dateStr) {
  if (dateStr == null || String(dateStr).trim() === '') return null;
  const s = String(dateStr).trim().slice(0, 10);
  const m = /^(\d{4})-(\d{2})/.exec(s);
  if (!m) return null;
  return `${m[1]}-${m[2]}`;
}

function getMinStartMonthForBillingScope(customerId) {
  return new Promise((resolve, reject) => {
    let sql = `SELECT MIN(s.start_date) AS d FROM subscriptions s
               WHERE s.active = 1 AND COALESCE(s.is_annual, 0) != 1
                 AND s.start_date IS NOT NULL AND TRIM(s.start_date) != ''`;
    const params = [];
    if (customerId != null && Number.isFinite(customerId)) {
      sql += ' AND s.customer_id = ?';
      params.push(customerId);
    }
    db.get(sql, params, (err, row) => {
      if (err) return reject(err);
      resolve(monthKeyFromStartDateString(row && row.d));
    });
  });
}

/**
 * Create/update payment rows for each calendar month from a safe start through today,
 * using the same rules as Generate Billing. Fixes empty payment tables when billing was never run.
 * @param {number|null} customerId null = all customers (dashboard); number = one customer page.
 */
async function ensureBillingThroughCurrentForScope(customerId) {
  const cur = getCurrentMonthKey();
  const floor = addCalendarMonthsToKey(cur, -36);
  if (!floor) return;

  const hasBillableSubs = await new Promise((resolve) => {
    if (customerId != null && Number.isFinite(customerId)) {
      db.get(
        'SELECT 1 AS o FROM subscriptions WHERE customer_id = ? AND active = 1 AND COALESCE(is_annual, 0) = 0 LIMIT 1',
        [customerId],
        (err, row) => resolve(!!row)
      );
    } else {
      db.get(
        'SELECT 1 AS o FROM subscriptions WHERE active = 1 AND COALESCE(is_annual, 0) = 0 LIMIT 1',
        [],
        (err, row) => resolve(!!row)
      );
    }
  });
  if (!hasBillableSubs) return;

  let minStart = null;
  try {
    minStart = await getMinStartMonthForBillingScope(customerId);
  } catch {
    return;
  }

  const fromKey = minStart ? maxMonthKey(floor, minStart) : floor;
  if (fromKey > cur) return;

  let m = fromKey;
  for (;;) {
    await processOneMonthBilling(m, customerId);
    if (m >= cur) break;
    const next = addCalendarMonthsToKey(m, 1);
    if (!next || next === m) break;
    m = next;
  }
}

/**
 * Create/update payment rows for one calendar month. `customerId` null = all customers.
 * Resolves with count of subscriptions that received billing (annual prepaid excluded).
 */
function processOneMonthBilling(monthKey, customerId) {
  return new Promise((resolve, reject) => {
    const [yearStr, monthStr] = monthKey.split('-').map((x) => Number(x));
    const year = yearStr;
    const month = monthStr;
    let sql = `SELECT s.*, c.name as customer_name, c.mobile as customer_mobile,
          COALESCE(i.monthly_rate, s.monthly_rate, 295) AS effective_monthly_rate
          FROM subscriptions s
          JOIN customers c ON c.id = s.customer_id
          LEFT JOIN items i ON i.id = s.item_id
          WHERE s.active = 1
            AND (s.end_date IS NULL OR date(s.end_date) >= date(?,'start of month'))
            AND (s.start_date IS NULL OR date(s.start_date) <= date(?,'start of month','+1 month','-1 day'))`;
    const params = [monthKey + '-01', monthKey + '-01'];
    if (customerId != null && Number.isFinite(customerId)) {
      sql += ' AND c.id = ?';
      params.push(customerId);
    }
    db.all(sql, params, async (err, subs) => {
      if (err) return reject(err);
      const billedCount = subs.filter((s) => Number(s.is_annual) !== 1).length;
      const tasks = subs.map(
        (s) =>
          new Promise(async (resolveTask) => {
            if (Number(s.is_annual) === 1) return resolveTask();

            let amountForThisSub = 0;
            if (Number(s.is_sunday_only) === 1) {
              const sundays = countSundaysInMonth(year, month);
              amountForThisSub = sundays * 10;
            } else {
              amountForThisSub = Number(
                s.effective_monthly_rate != null && s.effective_monthly_rate !== ''
                  ? s.effective_monthly_rate
                  : 295
              );
            }

            const prevBalance = await getLastBalance(s.id);
            const totalDue = (prevBalance || 0) + amountForThisSub;

            db.get(
              'SELECT * FROM payments WHERE subscription_id = ? AND month = ?',
              [s.id, monthKey],
              (e2, existing) => {
                if (existing) {
                  const newBalance = totalDue - (existing.amount_paid || 0);
                  db.run(
                    'UPDATE payments SET amount_due=?, balance=? WHERE id=?',
                    [totalDue, newBalance, existing.id],
                    () => resolveTask()
                  );
                } else {
                  db.run(
                    'INSERT INTO payments (subscription_id,customer_id,month,amount_due,amount_paid,balance) VALUES(?,?,?,?,?,?)',
                    [s.id, s.customer_id, monthKey, totalDue, 0, totalDue],
                    () => resolveTask()
                  );
                }
              }
            );
          })
      );

      Promise.all(tasks).then(() => resolve(billedCount));
    });
  });
}

// ---------- Auth ----------
/** Form `active` / `active` select: urlencoded may send 0 as number or string */
function parseActiveFromBody(body) {
  const v = body.active;
  if (v === 0 || v === '0' || v === false || v === 'false') return 0;
  return 1;
}

/** Parse `item_ids` / legacy `item_id` from form body → unique positive item ids. */
function parseItemIdsFromBody(body) {
  const raw = body.item_ids;
  let ids = [];
  if (Array.isArray(raw)) {
    ids = raw.map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0);
  } else if (raw != null && String(raw).trim() !== '') {
    const one = Number(raw);
    if (Number.isFinite(one) && one > 0) ids = [one];
  }
  if (!ids.length && body.item_id != null && String(body.item_id).trim() !== '') {
    const legacy = Number(body.item_id);
    if (Number.isFinite(legacy) && legacy > 0) ids = [legacy];
  }
  return [...new Set(ids)];
}

/**
 * Add subscriptions for newly checked items; deactivate subscriptions for items unchecked (edit flow).
 * New rows use subscription_start_date / subscription_end_date from the form.
 */
function syncSubscriptionsFromCustomerForm(customerId, body, done) {
  const desiredIds = parseItemIdsFromBody(body);
  const startRaw = body.subscription_start_date;
  const endRaw = body.subscription_end_date;
  const startDate =
    startRaw != null && String(startRaw).trim() !== '' ? String(startRaw).trim() : null;
  const endDate = endRaw != null && String(endRaw).trim() !== '' ? String(endRaw).trim() : null;

  db.all(
    'SELECT id, item_id FROM subscriptions WHERE customer_id = ? AND active = 1 AND item_id IS NOT NULL',
    [customerId],
    (e, activeSubs) => {
      if (e) return done(e);
      const rows = activeSubs || [];
      const toDeactivate = [];
      rows.forEach((s) => {
        if (s.item_id != null && !desiredIds.some((d) => Number(d) === Number(s.item_id))) {
          toDeactivate.push(s.id);
        }
      });
      const activeItemSet = new Set(
        rows.map((s) => s.item_id).filter((id) => id != null).map((id) => Number(id))
      );
      const toCreate = desiredIds.filter((id) => !activeItemSet.has(Number(id)));

      let di = 0;
      function deactivateNext() {
        if (di >= toDeactivate.length) return createNext();
        const subId = toDeactivate[di];
        di += 1;
        db.run(
          'UPDATE subscriptions SET active=0, end_date=date(\'now\') WHERE id=?',
          [subId],
          (err) => {
            if (err) return done(err);
            deactivateNext();
          }
        );
      }

      let ci = 0;
      function createNext() {
        if (ci >= toCreate.length) return done();
        const itemId = toCreate[ci];
        ci += 1;
        db.get(
          'SELECT * FROM items WHERE id = ? AND COALESCE(active, 1) != 0',
          [itemId],
          (e2, item) => {
            if (e2) return done(e2);
            if (!item) return createNext();
            const rate = item.monthly_rate != null ? Number(item.monthly_rate) : 295;
            db.run(
              `INSERT INTO subscriptions (customer_id,item_id,item_name,monthly_rate,is_annual,is_sunday_only,start_date,end_date,active)
               VALUES(?,?,?,?,0,0,?,?,1)`,
              [customerId, itemId, item.name, rate, startDate, endDate],
              (e3) => {
                if (e3) return done(e3);
                createNext();
              }
            );
          }
        );
      }

      deactivateNext();
    }
  );
}

function requireAuth(req, res, next) {
  if (req.session.loggedIn) return next();
  res.redirect('/login');
}

// ---------- Routes: general ----------
app.get('/', (req, res) => res.redirect('/login'));

app.get('/login', (req, res) => res.render('login', { error: null }));
app.post('/login', (req, res) => {
  const { username, password } = req.body;
  // static simple login
  if (username === 'admin' && password === 'admin') {
    req.session.loggedIn = true;
    return res.redirect('/home');
  }
  res.render('login', { error: 'Invalid credentials' });
});
app.get('/logout', (req, res) => { req.session.destroy(()=>res.redirect('/login')); });

// ---------- Dashboard ----------
app.get('/home', requireAuth, async (req, res) => {
  try {
    await ensureBillingThroughCurrentForScope(null);
  } catch (e) {
    console.error('ensureBilling home:', e);
  }
  const monthKey = getCurrentMonthKey();
  const stats = {};
  db.get('SELECT COUNT(*) AS c FROM customers', [], (e, r1) => {
    stats.customers = r1 ? r1.c : 0;
    db.get('SELECT COUNT(*) AS c FROM subscriptions WHERE active=1', [], (e2, r2) => {
      stats.active_subscriptions = r2 ? r2.c : 0;
      db.get('SELECT IFNULL(SUM(balance),0) as pending FROM payments WHERE month = ?', [monthKey], (e3, r3) => {
        stats.current_month_pending = r3 ? r3.pending : 0;
        db.get('SELECT IFNULL(SUM(balance),0) as older_pending FROM payments WHERE month <> ?', [monthKey], (e4, r4) => {
          stats.older_pending = r4 ? r4.older_pending : 0;
          res.render('home', { stats, monthKey });
        });
      });
    });
  });
});

app.get('/home/pending-details', requireAuth, (req, res) => {
  const monthKey = req.query.month || getCurrentMonthKey();
  db.all(
    `SELECT p.id AS payment_id, p.balance, p.amount_due, p.amount_paid, p.month,
            c.id AS customer_id, c.name AS customer_name, c.mobile AS customer_mobile,
            COALESCE(s.item_name, '(subscription)') AS subscription_name
     FROM payments p
     JOIN customers c ON c.id = p.customer_id
     LEFT JOIN subscriptions s ON s.id = p.subscription_id
     WHERE p.month = ? AND CAST(COALESCE(p.balance, 0) AS REAL) > 0
     ORDER BY c.name COLLATE NOCASE, subscription_name`,
    [monthKey],
    (err, rows) => {
      if (err) return res.status(500).json({ error: err.message });
      res.json({ month: monthKey, rows: rows || [] });
    }
  );
});

// ---------- Customers ----------
app.get('/customer/add', requireAuth, (req, res) => {
  const editId = req.query.edit;
  const renderForm = (customer, editing, activeItemIds) => {
    db.all(
      `SELECT * FROM items WHERE COALESCE(active, 1) != 0 ORDER BY item_type, name`,
      [],
      (e, subscriptionItems) => {
        if (e) return res.status(500).send('Error loading items: ' + e.message);
        db.all('SELECT * FROM items ORDER BY item_type,name', [], (e2, items) => {
          if (e2) return res.status(500).send('Error loading items: ' + e2.message);
          res.render('customer_add', {
            items: items || [],
            subscriptionItems: subscriptionItems || [],
            activeItemIds: activeItemIds || [],
            customer: customer == null ? null : customer,
            editing: !!editing,
          });
        });
      }
    );
  };
  if (editId != null && String(editId).trim() !== '') {
    db.get('SELECT * FROM customers WHERE id = ?', [editId], (e, customer) => {
      if (!customer) return res.redirect('/customers');
      db.all(
        'SELECT item_id FROM subscriptions WHERE customer_id = ? AND active = 1 AND item_id IS NOT NULL',
        [editId],
        (e2, rows) => {
          if (e2) return res.status(500).send('Error loading subscriptions: ' + e2.message);
          const activeItemIds = (rows || []).map((r) => r.item_id).filter((id) => id != null);
          renderForm(customer, true, activeItemIds);
        }
      );
    });
  } else {
    renderForm(null, false, []);
  }
});

app.get('/customer/edit/:id', requireAuth, (req, res) => {
  res.redirect(302, '/customer/add?edit=' + encodeURIComponent(req.params.id));
});

app.post('/customer/add', requireAuth, (req, res) => {
  const { name, address, mobile, area_name, route_name } = req.body;
  const active = parseActiveFromBody(req.body);
  db.run(`INSERT INTO customers (name,address,mobile,area_name,route_name,active) VALUES(?,?,?,?,?,?)`,
    [name, address, mobile, area_name || null, route_name || null, active], function(err) {
      if (err) return res.send('Error creating customer');
      const newId = this.lastID;
      syncSubscriptionsFromCustomerForm(newId, req.body, (err) => {
        if (err) return res.status(500).send('Error syncing subscriptions: ' + err.message);
        res.redirect('/customers');
      });
    });
});

app.post('/customer/edit/:id', requireAuth, (req, res) => {
  const id = req.params.id;
  const { name, address, mobile, area_name, route_name } = req.body;
  const active = parseActiveFromBody(req.body);
  db.run(
    `UPDATE customers SET name=?, address=?, mobile=?, area_name=?, route_name=?, active=? WHERE id=?`,
    [name, address, mobile, area_name || null, route_name || null, active, id],
    (err) => {
      if (err) return res.send('Error updating customer');
      syncSubscriptionsFromCustomerForm(Number(id), req.body, (err) => {
        if (err) return res.status(500).send('Error syncing subscriptions: ' + err.message);
        res.redirect('/customers');
      });
    }
  );
});

app.get('/customers', requireAuth, (req, res) => {
  const q = String(req.query.q || '').trim();
  if (q) {
    const pat = `%${escapeLikePattern(q)}%`;
    db.all(
      `SELECT * FROM customers
       WHERE name LIKE ? ESCAPE '\\' OR IFNULL(mobile, '') LIKE ? ESCAPE '\\'
       ORDER BY name COLLATE NOCASE`,
      [pat, pat],
      (err, rows) => {
        if (err) return res.status(500).send('Error loading customers: ' + err.message);
        res.render('customer_list', { customers: rows || [], searchQuery: q });
      }
    );
  } else {
    db.all('SELECT * FROM customers ORDER BY name COLLATE NOCASE', [], (err, rows) => {
      if (err) return res.status(500).send('Error loading customers: ' + err.message);
      res.render('customer_list', { customers: rows || [], searchQuery: '' });
    });
  }
});

// view single customer and their subscriptions
app.get('/customer/view/:id', requireAuth, async (req, res) => {
  const id = req.params.id;
  const cid = Number(id);
  if (!Number.isFinite(cid)) return res.status(400).send('Invalid customer id');
  try {
    await ensureBillingThroughCurrentForScope(cid);
  } catch (e) {
    console.error('ensureBilling customer view:', e);
  }
  db.get('SELECT * FROM customers WHERE id = ?', [id], (e, customer) => {
    if (!customer) return res.send('Customer not found');
    db.all('SELECT s.*, i.name as item_name FROM subscriptions s LEFT JOIN items i ON i.id = s.item_id WHERE s.customer_id = ?', [id], (err, subs) => {
      db.all(
        `SELECT p.*, COALESCE(s.item_name, '') AS sub_item_name
         FROM payments p
         LEFT JOIN subscriptions s ON s.id = p.subscription_id
         WHERE p.customer_id = ?
         ORDER BY p.month DESC
         LIMIT 48`,
        [id],
        (err2, payments) => {
          res.render('customer_view', { customer, subscriptions: subs, payments });
        }
      );
    });
  });
});

// Legacy URL → query form (avoids "Cannot GET" if old links/bookmarks exist)
app.get('/items/edit/:id', requireAuth, (req, res) => {
  res.redirect(302, '/items?edit=' + encodeURIComponent(req.params.id));
});

// ---------- Items (news/magazines rates) ----------
app.get('/items', requireAuth, (req, res) => {
  const editId = req.query.edit;
  const loadList = (editItem, activeSubCount) => {
    db.all('SELECT * FROM items ORDER BY item_type,name', [], (err, rows) => {
      res.render('items', {
        items: rows,
        editItem: editItem == null ? null : editItem,
        activeSubCount: activeSubCount == null ? 0 : activeSubCount,
      });
    });
  };
  if (editId != null && String(editId).trim() !== '') {
    db.get('SELECT * FROM items WHERE id = ?', [editId], (e, editItem) => {
      if (!editItem) return res.redirect('/items');
      db.get(
        'SELECT COUNT(*) AS c FROM subscriptions WHERE item_id = ? AND active = 1',
        [editItem.id],
        (e2, row) => {
          const cnt = row && row.c != null ? Number(row.c) : 0;
          loadList(editItem, cnt);
        }
      );
    });
  } else {
    loadList(null, 0);
  }
});

app.post('/items', requireAuth, (req, res) => {
  const { name, item_type, monthly_rate } = req.body;
  const active = parseActiveFromBody(req.body);
  db.run(`INSERT OR REPLACE INTO items (name,item_type,monthly_rate,active) VALUES(?,?,?,?)`,
    [name, item_type || 'newspaper', monthly_rate || 295, active], () => res.redirect('/items'));
});

app.post('/items/edit/:id', requireAuth, (req, res) => {
  const id = req.params.id;
  const { name, item_type, monthly_rate } = req.body;
  const active = parseActiveFromBody(req.body);
  const removeLinked =
    req.body.remove_linked_subscriptions === '1' || req.body.remove_linked_subscriptions === 1;
  const rate = monthly_rate === '' || monthly_rate == null ? 295 : Number(monthly_rate);

  db.get('SELECT * FROM items WHERE id = ?', [id], (eItem, existing) => {
    if (!existing) return res.send('Item not found');
    const wasActive = Number(existing.active) !== 0;

    db.get(
      'SELECT COUNT(*) AS c FROM subscriptions WHERE item_id = ? AND active = 1',
      [id],
      (eCnt, row) => {
        const activeSubCount = row && row.c != null ? Number(row.c) : 0;

        if (active === 0 && activeSubCount > 0 && wasActive && !removeLinked) {
          return res.status(400).send(
            'This item has active subscriptions. Use the form (confirm when prompted) to make it inactive and stop those subscriptions.'
          );
        }

        const runItemUpdate = () => {
          db.run(
            `UPDATE items SET name=?, item_type=?, monthly_rate=?, active=? WHERE id=?`,
            [name, item_type || 'newspaper', rate, active, id],
            function (err) {
              if (err) return res.send('Error updating item: ' + err.message);
              db.run(`UPDATE subscriptions SET item_name=? WHERE item_id=?`, [name, id], () => res.redirect('/items'));
            }
          );
        };

        if (active === 0 && removeLinked && activeSubCount > 0) {
          db.run(
            `UPDATE subscriptions SET active = 0, end_date = COALESCE(end_date, date('now')) WHERE item_id = ? AND active = 1`,
            [id],
            (eSub) => {
              if (eSub) return res.send('Error stopping subscriptions: ' + eSub.message);
              runItemUpdate();
            }
          );
        } else {
          runItemUpdate();
        }
      }
    );
  });
});

// ---------- Subscriptions ----------
app.get('/subscription/add/:customerId', requireAuth, (req, res) => {
  const cid = req.params.customerId;
  db.get('SELECT * FROM customers WHERE id = ?', [cid], (e, customer) => {
    if (!customer) return res.send('Customer not found');
    db.all(
      'SELECT item_id FROM subscriptions WHERE customer_id = ? AND active = 1 AND item_id IS NOT NULL',
      [cid],
      (e2, subRows) => {
        const activeItemIds = (subRows || []).map((r) => r.item_id).filter((id) => id != null);
        db.all(
          `SELECT * FROM items WHERE COALESCE(active, 1) != 0 ORDER BY item_type,name`,
          [],
          (err, items) => {
            if (err) return res.status(500).send('Error loading items: ' + err.message);
            res.render('subscription_add', { customer, items: items || [], activeItemIds });
          }
        );
      }
    );
  });
});
app.post('/subscription/add', requireAuth, (req, res) => {
  const customer_id = req.body.customer_id;
  const rawIds = req.body.item_ids;
  let itemIds = [];
  if (Array.isArray(rawIds)) {
    itemIds = rawIds.map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0);
  } else if (rawIds != null && String(rawIds).trim() !== '') {
    const one = Number(rawIds);
    if (Number.isFinite(one) && one > 0) itemIds = [one];
  }
  if (!itemIds.length && req.body.item_id != null && String(req.body.item_id).trim() !== '') {
    const legacy = Number(req.body.item_id);
    if (Number.isFinite(legacy) && legacy > 0) itemIds = [legacy];
  }
  if (!itemIds.length) {
    return res.status(400).send('Select at least one publication (item).');
  }

  const startRaw = req.body.start_date;
  const start_date = startRaw != null && String(startRaw).trim() !== '' ? String(startRaw).trim() : null;
  if (!start_date) {
    return res.status(400).send('Start date is required.');
  }
  const end_date =
    req.body.end_date != null && String(req.body.end_date).trim() !== '' ? String(req.body.end_date).trim() : null;

  const itemRates = req.body.item_rates && typeof req.body.item_rates === 'object' ? req.body.item_rates : {};
  function rateFromFormForItem(itemId, catalogFallback) {
    const raw = itemRates[String(itemId)] ?? itemRates[itemId];
    if (raw != null && String(raw).trim() !== '') {
      const n = Number(raw);
      if (Number.isFinite(n) && n >= 0) return n;
    }
    return Number(catalogFallback != null ? catalogFallback : 295);
  }

  const s_annual = req.body.is_annual ? 1 : 0;
  const s_sunday = req.body.is_sunday_only ? 1 : 0;

  function processIndex(i) {
    if (i >= itemIds.length) {
      return res.redirect('/customer/view/' + customer_id);
    }
    const item_id = itemIds[i];
    db.get(
      'SELECT 1 AS x FROM subscriptions WHERE customer_id = ? AND item_id = ? AND active = 1',
      [customer_id, item_id],
      (eDup, dup) => {
        if (eDup) return res.send('Error checking subscriptions: ' + eDup.message);
        if (dup) return processIndex(i + 1);
        db.get('SELECT name, monthly_rate FROM items WHERE id = ?', [item_id], (err, item) => {
          if (err) return res.send('Error loading item: ' + err.message);
          if (!item) return processIndex(i + 1);
          const itemName = item.name || 'Custom';
          const catalog = item.monthly_rate != null ? item.monthly_rate : 295;
          const rate = rateFromFormForItem(item_id, catalog);
          db.run(
            `INSERT INTO subscriptions (customer_id,item_id,item_name,monthly_rate,is_annual,is_sunday_only,start_date,end_date,active)
            VALUES(?,?,?,?,?,?,?,?,1)`,
            [customer_id, item_id, itemName, rate, s_annual, s_sunday, start_date, end_date],
            (err2) => {
              if (err2) return res.send('Error adding subscription: ' + err2.message);
              processIndex(i + 1);
            }
          );
        });
      }
    );
  }

  processIndex(0);
});

app.get('/subscription/edit/:id', requireAuth, (req, res) => {
  const subId = req.params.id;
  db.get(
    `SELECT s.*, c.name AS customer_name
     FROM subscriptions s
     JOIN customers c ON c.id = s.customer_id
     WHERE s.id = ?`,
    [subId],
    (e, row) => {
      if (e) return res.status(500).send('Error loading subscription: ' + e.message);
      if (!row) return res.status(404).send('Subscription not found');
      if (Number(row.active) === 0) {
        return res.status(400).send('This subscription is stopped. Start a new subscription from the customer page if needed.');
      }
      res.render('subscription_edit', { subscription: row });
    }
  );
});

app.post('/subscription/edit/:id', requireAuth, (req, res) => {
  const subId = req.params.id;
  const startRaw = req.body.start_date;
  const start_date = startRaw != null && String(startRaw).trim() !== '' ? String(startRaw).trim() : null;
  if (!start_date) {
    return res.status(400).send('Start date is required.');
  }
  const end_date =
    req.body.end_date != null && String(req.body.end_date).trim() !== '' ? String(req.body.end_date).trim() : null;
  const rateRaw = req.body.monthly_rate;
  const monthly_rate =
    rateRaw != null && String(rateRaw).trim() !== '' ? Number(rateRaw) : NaN;
  if (!Number.isFinite(monthly_rate) || monthly_rate < 0) {
    return res.status(400).send('Enter a valid monthly rate (0 or greater).');
  }
  const s_annual = req.body.is_annual ? 1 : 0;
  const s_sunday = req.body.is_sunday_only ? 1 : 0;

  db.get('SELECT id, customer_id, active FROM subscriptions WHERE id = ?', [subId], (e, sub) => {
    if (e) return res.status(500).send(e.message);
    if (!sub) return res.status(404).send('Subscription not found');
    if (Number(sub.active) === 0) {
      return res.status(400).send('Cannot edit a stopped subscription.');
    }
    db.run(
      `UPDATE subscriptions SET monthly_rate=?, start_date=?, end_date=?, is_annual=?, is_sunday_only=? WHERE id=?`,
      [monthly_rate, start_date, end_date, s_annual, s_sunday, subId],
      (err2) => {
        if (err2) return res.status(500).send('Error updating subscription: ' + err2.message);
        res.redirect('/customer/view/' + sub.customer_id);
      }
    );
  });
});

// stop subscription (deactivate)
app.post('/subscription/stop/:id', requireAuth, (req, res) => {
  const subId = req.params.id;
  const customerIdCheck = req.body.customer_id;
  db.get('SELECT id, customer_id FROM subscriptions WHERE id = ?', [subId], (e, sub) => {
    if (e) return res.status(500).send(e.message);
    if (!sub) return res.status(404).send('Subscription not found');
    if (customerIdCheck != null && String(customerIdCheck).trim() !== '' && String(customerIdCheck) !== String(sub.customer_id)) {
      return res.status(403).send('This subscription does not belong to that customer.');
    }
    const ed =
      req.body.end_date != null && String(req.body.end_date).trim() !== '' ? String(req.body.end_date).trim() : null;
    db.run(
      'UPDATE subscriptions SET active=0, end_date=COALESCE(?, date(\'now\')) WHERE id=?',
      [ed, subId],
      (err2) => {
        if (err2) return res.status(500).send(err2.message);
        res.redirect('/customer/view/' + sub.customer_id);
      }
    );
  });
});

// list subscriptions
app.get('/subscriptions', requireAuth, (req, res) => {
  db.all(`SELECT s.*, c.name as customer_name FROM subscriptions s LEFT JOIN customers c ON c.id = s.customer_id ORDER BY c.name`, [], (err, rows) => {
    res.render('subscriptions', { subscriptions: rows });
  });
});

// ---------- Billing: generate monthly billing ----------
app.get('/billing/form', requireAuth, (req, res) => {
  const m = req.query.month || getCurrentMonthKey();
  const error = req.query.error ? String(req.query.error) : null;
  db.all(
    'SELECT id, name, mobile FROM customers ORDER BY name COLLATE NOCASE',
    [],
    (e, customers) => {
      if (e) return res.send('Error loading customers: ' + e.message);
      res.render('billing_form', { month: m, customers: customers || [], error });
    }
  );
});

/** After POST /billing/generate: show payment lines for the months that were run. */
function enrichBillingSummaryRows(list, done) {
  const subIds = [...new Set(list.map((r) => r.subscription_id).filter((id) => id != null))];
  if (!subIds.length) {
    return done(
      null,
      list.map((r) => {
        const fromSub = subscriptionFeeForPaymentMonth(r);
        return {
          ...r,
          prior_closing_balance: null,
          period_charge: fromSub != null ? fromSub : Number(r.amount_due || 0),
        };
      })
    );
  }
  const rowMonths = [...new Set(list.map((r) => r.month))].filter(Boolean).sort();
  const maxDisplayed = rowMonths[rowMonths.length - 1];
  const ph = subIds.map(() => '?').join(',');
  const sql = `SELECT subscription_id, month, balance FROM payments WHERE subscription_id IN (${ph}) AND month <= ?`;
  db.all(sql, [...subIds, maxDisplayed], (e2, hist) => {
    if (e2) return done(e2);
    const bySub = new Map();
    for (const pr of hist || []) {
      const id = pr.subscription_id;
      if (!bySub.has(id)) bySub.set(id, []);
      bySub.get(id).push({ month: pr.month, bal: Number(pr.balance || 0) });
    }
    for (const arr of bySub.values()) {
      arr.sort((a, b) => (a.month < b.month ? -1 : a.month > b.month ? 1 : 0));
    }

    const enriched = list.map((r) => {
      const sid = r.subscription_id;
      if (sid == null) {
        return { ...r, prior_closing_balance: null, period_charge: Number(r.amount_due || 0) };
      }
      const arr = bySub.get(sid) || [];
      let pick = null;
      for (const entry of arr) {
        if (entry.month >= r.month) continue;
        if (!pick || entry.month > pick.month) pick = entry;
      }
      const priorClosing = pick != null ? pick.bal : 0;
      const fromSub = subscriptionFeeForPaymentMonth(r);
      const periodCharge =
        fromSub != null ? fromSub : Math.max(0, Number(r.amount_due || 0) - priorClosing);
      return {
        ...r,
        prior_closing_balance: pick != null ? pick.bal : null,
        period_charge: periodCharge,
      };
    });
    done(null, enriched);
  });
}

app.get('/billing/summary', requireAuth, (req, res) => {
  const months = String(req.query.months || '')
    .split(',')
    .map((s) => s.trim())
    .filter((m) => /^\d{4}-\d{2}$/.test(m));
  if (!months.length) return res.redirect('/billing/form');

  const frequency = ['monthly', 'quarterly', 'yearly'].includes(String(req.query.frequency || ''))
    ? req.query.frequency
    : 'monthly';
  const scope = req.query.scope === 'customer' ? 'customer' : 'all';
  const customerIdRaw = req.query.customer_id;
  const customerId =
    customerIdRaw != null && String(customerIdRaw).trim() !== '' && Number.isFinite(Number(customerIdRaw))
      ? Number(customerIdRaw)
      : null;

  const placeholders = months.map(() => '?').join(',');
  let sql = `SELECT p.*, c.name AS customer_name, c.mobile AS customer_mobile, s.item_name,
       COALESCE(i.monthly_rate, s.monthly_rate, 295) AS subscription_monthly_rate,
       COALESCE(s.is_sunday_only, 0) AS subscription_is_sunday_only,
       COALESCE(s.is_annual, 0) AS subscription_is_annual
     FROM payments p
     LEFT JOIN customers c ON c.id = p.customer_id
     LEFT JOIN subscriptions s ON s.id = p.subscription_id
     LEFT JOIN items i ON i.id = s.item_id
     WHERE p.month IN (${placeholders})`;
  const params = [...months];
  if (customerId != null) {
    sql += ' AND p.customer_id = ?';
    params.push(customerId);
  }
  sql += ' ORDER BY p.month, COALESCE(c.name, \'\') COLLATE NOCASE, COALESCE(s.item_name, \'\')';

  db.all(sql, params, (err, rows) => {
    if (err) return res.status(500).send('Could not load billing summary: ' + err.message);
    const list = rows || [];
    enrichBillingSummaryRows(list, (e3, enriched) => {
      if (e3) return res.status(500).send('Could not load billing summary: ' + e3.message);
      const totals = enriched.reduce(
        (acc, r) => {
          acc.amountDue += Number(r.amount_due || 0);
          acc.amountPaid += Number(r.amount_paid || 0);
          acc.balance += Number(r.balance || 0);
          const pc = Number(r.period_charge);
          acc.periodCharge += Number.isFinite(pc) ? pc : 0;
          return acc;
        },
        { amountDue: 0, amountPaid: 0, balance: 0, periodCharge: 0 }
      );
      const subscriptionMonths = Number(req.query.billed || '') || enriched.length;

      const finish = (customerRow) => {
        res.render('billing_summary', {
          months,
          frequency,
          scope,
          customer: customerRow,
          rows: enriched,
          totals,
          subscriptionMonths,
        });
      };

      if (customerId != null) {
        db.get('SELECT id, name, mobile FROM customers WHERE id = ?', [customerId], (e2, c) => finish(c || null));
      } else {
        finish(null);
      }
    });
  });
});

app.post('/billing/generate', requireAuth, async (req, res) => {
  const monthKey = String(req.body.month || '').trim() || getCurrentMonthKey();
  if (!/^\d{4}-\d{2}$/.test(monthKey)) {
    return res.redirect('/billing/form?error=' + encodeURIComponent('Invalid period start. Use YYYY-MM.'));
  }
  const frequency = ['monthly', 'quarterly', 'yearly'].includes(req.body.frequency)
    ? req.body.frequency
    : 'monthly';
  const scope = req.body.billing_scope === 'customer' ? 'customer' : 'all';
  let customerId = null;
  if (scope === 'customer') {
    customerId = Number(req.body.customer_id);
    if (!Number.isFinite(customerId) || customerId <= 0) {
      return res.redirect(
        '/billing/form?error=' + encodeURIComponent('Select a customer for people-based billing.')
      );
    }
  }

  const monthKeys = periodMonthKeys(monthKey, frequency);
  if (monthKeys.length === 0) {
    return res.redirect(
      '/billing/form?error=' + encodeURIComponent('Could not build billing period from that month.')
    );
  }

  try {
    let totalBilledSubs = 0;
    for (const mk of monthKeys) {
      totalBilledSubs += await processOneMonthBilling(mk, customerId);
    }
    const q = new URLSearchParams({
      months: monthKeys.join(','),
      frequency,
      scope,
      billed: String(totalBilledSubs),
    });
    if (customerId != null) q.set('customer_id', String(customerId));
    return res.redirect(302, '/billing/summary?' + q.toString());
  } catch (err) {
    return res.redirect(
      '/billing/form?error=' + encodeURIComponent('Error generating billing: ' + err.message)
    );
  }
});

// ---------- Payments ----------
app.get('/payment/customer/:customerId', requireAuth, async (req, res) => {
  const cidRaw = req.params.customerId;
  const cid = Number(cidRaw);
  if (!Number.isFinite(cid)) return res.status(400).send('Invalid customer id');
  try {
    await ensureBillingThroughCurrentForScope(cid);
  } catch (e) {
    console.error('ensureBilling payment customer:', e);
  }
  db.get('SELECT * FROM customers WHERE id = ?', [cid], (e, cust) => {
    if (!cust) return res.send('Customer not found');
    db.all(`SELECT p.*, s.item_name FROM payments p LEFT JOIN subscriptions s ON s.id = p.subscription_id WHERE p.customer_id = ? ORDER BY p.month DESC`, [cid], (err, rows) => {
      res.render('payment_customer', { customer: cust, payments: rows });
    });
  });
});

// payment page for one payment record
app.get('/payment/:paymentId', requireAuth, (req, res) => {
  const pid = req.params.paymentId;
  db.get('SELECT p.*, s.item_name, c.name AS customer_name FROM payments p LEFT JOIN subscriptions s ON s.id=p.subscription_id LEFT JOIN customers c ON c.id=p.customer_id WHERE p.id = ?', [pid], (e, p) => {
    if (!p) return res.send('Payment record not found');
    res.render('payment_single', { payment: p });
  });
});

app.post('/payment/pay/:paymentId', requireAuth, (req, res) => {
  const pid = req.params.paymentId;
  const amount = Number(req.body.amount || 0);
  db.get('SELECT * FROM payments WHERE id = ?', [pid], (e, p) => {
    if (!p) return res.send('Payment not found');
    const newPaid = Number(p.amount_paid || 0) + amount;
    const newBalance = Number(p.amount_due || 0) - newPaid;
    db.run('UPDATE payments SET amount_paid=?, balance=? WHERE id=?', [newPaid, newBalance, pid], (err) => {
      if (err) return res.send('Error updating payment');
      res.redirect('/payment/' + pid);
    });
  });
});

// export payments: ?month=YYYY-MM OR ?from=&to= ; optional customer_id= ; ?format=csv|xlsx (default xlsx)
app.get('/export/payments', requireAuth, (req, res) => {
  const format = String(req.query.format || 'xlsx').toLowerCase();
  const month = String(req.query.month || '').trim();
  const from = String(req.query.from || '').trim();
  const to = String(req.query.to || '').trim();
  const custRaw = req.query.customer_id;
  const exportCustomerId =
    custRaw != null && String(custRaw).trim() !== '' && Number.isFinite(Number(custRaw))
      ? Number(custRaw)
      : null;

  let sql =
    'SELECT p.id, p.month, p.amount_due, p.amount_paid, p.balance, s.item_name, c.name as customer_name, c.mobile as customer_mobile ' +
    'FROM payments p LEFT JOIN subscriptions s ON s.id=p.subscription_id LEFT JOIN customers c ON c.id=p.customer_id';
  const params = [];
  if (month && /^\d{4}-\d{2}$/.test(month)) {
    sql += ' WHERE p.month = ?';
    params.push(month);
  } else if (from && to && /^\d{4}-\d{2}$/.test(from) && /^\d{4}-\d{2}$/.test(to)) {
    if (from > to) {
      return res
        .status(400)
        .type('text')
        .send('from= must be on or before to= (both YYYY-MM).');
    }
    sql += ' WHERE p.month >= ? AND p.month <= ?';
    params.push(from, to);
  } else {
    return res
      .status(400)
      .type('text')
      .send(
        'Provide month=YYYY-MM or both from=YYYY-MM and to=YYYY-MM (inclusive). Optional customer_id=, format=csv|xlsx (default).'
      );
  }
  if (exportCustomerId != null) {
    sql += ' AND p.customer_id = ?';
    params.push(exportCustomerId);
  }
  sql += ' ORDER BY p.month DESC, c.name COLLATE NOCASE';

  db.all(sql, params, (err, rows) => {
    if (err) return res.status(500).type('text').send(err.message);
    const headers = [
      'payment_id',
      'month',
      'customer_name',
      'customer_mobile',
      'item_name',
      'amount_due',
      'amount_paid',
      'balance',
    ];
    const records = (rows || []).map((r) => ({
      payment_id: r.id,
      month: r.month,
      customer_name: r.customer_name,
      customer_mobile: r.customer_mobile,
      item_name: r.item_name,
      amount_due: r.amount_due,
      amount_paid: r.amount_paid,
      balance: r.balance,
    }));

    const custSuffix = exportCustomerId != null ? `_customer_${exportCustomerId}` : '';

    if (format === 'csv') {
      const lines = [headers.join(',')];
      records.forEach((rec) => {
        lines.push(headers.map((h) => csvEscapeCell(rec[h])).join(','));
      });
      const body = '\uFEFF' + lines.join('\r\n');
      const fname = month
        ? `payments_${month}${custSuffix}.csv`
        : `payments_${from}_to_${to}${custSuffix}.csv`;
      res.setHeader('Content-Type', 'text/csv; charset=utf-8');
      res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
      return res.send(body);
    }

    if (format !== 'xlsx') {
      return res.status(400).type('text').send('format must be csv or xlsx');
    }

    const sheetRows = records.map((r) => ({
      payment_id: r.payment_id,
      month: r.month,
      customer: r.customer_name,
      mobile: r.customer_mobile,
      item: r.item_name,
      amount_due: r.amount_due,
      amount_paid: r.amount_paid,
      balance: r.balance,
    }));
    const ws = XLSX.utils.json_to_sheet(sheetRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Payments');
    const fn = month
      ? `payments_${month}${custSuffix}.xlsx`
      : `payments_${from}_to_${to}${custSuffix}.xlsx`;
    const filepath = path.join(__dirname, fn);
    try {
      XLSX.writeFile(wb, filepath);
      res.download(filepath, fn, () => {
        fs.unlink(filepath, () => {});
      });
    } catch (e) {
      res.status(500).type('text').send(String(e.message || e));
    }
  });
});

// ---------- Settings ----------
app.get('/settings', requireAuth, (req, res) => {
  db.get(`SELECT value FROM settings WHERE key='per_paper_rate'`, [], (e, row) => {
    res.render('settings', { per_paper_rate: row ? row.value : '5' });
  });
});
app.post('/settings', requireAuth, (req, res) => {
  const { per_paper_rate } = req.body;
  db.run(`INSERT OR REPLACE INTO settings (key,value) VALUES('per_paper_rate',?)`, [String(per_paper_rate)], () => res.redirect('/settings'));
});

// ---------- Minimal extra pages ----------
app.get('/reports', requireAuth, (req, res) => {
  const currentMonth = getCurrentMonthKey();
  const trendMonths = lastNMonthKeysEndingAt(currentMonth, 12);
  const oldestMonth = trendMonths[0];

  db.get(
    'SELECT COUNT(*) AS c FROM customers WHERE COALESCE(active, 1) != 0',
    [],
    (e1, rowCustomers) => {
      if (e1) return res.status(500).send('Database error: ' + e1.message);
      db.get(
        'SELECT COUNT(*) AS c FROM subscriptions WHERE active = 1',
        [],
        (e2, rowSubs) => {
          if (e2) return res.status(500).send('Database error: ' + e2.message);
          db.all(
            `SELECT month,
                    SUM(COALESCE(amount_due, 0)) AS amount_due,
                    SUM(COALESCE(amount_paid, 0)) AS amount_paid,
                    SUM(COALESCE(balance, 0)) AS balance
             FROM payments WHERE month >= ? GROUP BY month ORDER BY month`,
            [oldestMonth],
            (e3, trendRows) => {
              if (e3) return res.status(500).send('Database error: ' + e3.message);
              const trendMap = {};
              (trendRows || []).forEach((r) => {
                trendMap[r.month] = r;
              });
              const trend = trendMonths.map((m) => ({
                month: m,
                amount_due: trendMap[m] ? Number(trendMap[m].amount_due) : 0,
                amount_paid: trendMap[m] ? Number(trendMap[m].amount_paid) : 0,
                balance: trendMap[m] ? Number(trendMap[m].balance) : 0,
              }));

              db.all(
                `SELECT COALESCE(s.item_name, '(no item)') AS item_name,
                        SUM(COALESCE(p.balance, 0)) AS total_balance
                 FROM payments p
                 LEFT JOIN subscriptions s ON s.id = p.subscription_id
                 WHERE p.month = ?
                 GROUP BY COALESCE(s.item_name, '(no item)')
                 ORDER BY total_balance DESC`,
                [currentMonth],
                (e4, byItemRows) => {
                  if (e4) return res.status(500).send('Database error: ' + e4.message);
                  const byItem = (byItemRows || []).map((r) => ({
                    label: r.item_name,
                    value: Number(r.total_balance) || 0,
                  }));

                  db.all(
                    `SELECT c.name AS customer_name, SUM(COALESCE(p.balance, 0)) AS total_balance
                     FROM payments p
                     JOIN customers c ON c.id = p.customer_id
                     WHERE CAST(COALESCE(p.balance, 0) AS REAL) > 0
                     GROUP BY c.id
                     ORDER BY total_balance DESC
                     LIMIT 10`,
                    [],
                    (e5, debtRows) => {
                      if (e5) return res.status(500).send('Database error: ' + e5.message);
                      const topDebtors = (debtRows || []).map((r) => ({
                        label: r.customer_name,
                        value: Number(r.total_balance) || 0,
                      }));

                      const cur = trend[trend.length - 1] || {
                        amount_due: 0,
                        amount_paid: 0,
                        balance: 0,
                      };
                      const chartPayload = {
                        trend,
                        byItem,
                        topDebtors,
                        currentMonth,
                      };
                      res.render('reports', {
                        currentMonth,
                        exportOldest: oldestMonth,
                        kpi: {
                          activeCustomers: rowCustomers ? rowCustomers.c : 0,
                          activeSubscriptions: rowSubs ? rowSubs.c : 0,
                          monthDue: cur.amount_due,
                          monthPaid: cur.amount_paid,
                          monthBalance: cur.balance,
                        },
                        chartPayloadJson: JSON.stringify(chartPayload).replace(/</g, '\\u003c'),
                      });
                    }
                  );
                }
              );
            }
          );
        }
      );
    }
  );
});
app.get('/delivery', requireAuth, (req, res) => {
  db.all('SELECT * FROM delivery_boys', [], (e, rows) => res.render('delivery', { boys: rows }));
});

// ---------- Start ----------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));