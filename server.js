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
    monthly_rate REAL
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

  addColumnIfNotExists('subscriptions', 'is_annual', 'INTEGER DEFAULT 0', (err) => {
    if (err) console.error('Migration error adding is_annual:', err.message);
    addColumnIfNotExists('subscriptions', 'is_sunday_only', 'INTEGER DEFAULT 0', (err2) => {
      if (err2) console.error('Migration error adding is_sunday_only:', err2.message);
      console.log('Database ready:', DB_FILE);
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

// ---------- Auth ----------
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
app.get('/home', requireAuth, (req, res) => {
  const monthKey = getCurrentMonthKey();
  const stats = {};
  db.get('SELECT COUNT(*) AS c FROM customers', [], (e, r1) => {
    stats.customers = r1 ? r1.c : 0;
    db.get('SELECT COUNT(*) AS c FROM subscriptions WHERE active=1', [], (e2, r2) => {
      stats.active_subscriptions = r2 ? r2.c : 0;
      db.get('SELECT IFNULL(SUM(balance),0) as pending FROM payments WHERE month = ?', [monthKey], (e3, r3) => {
        stats.current_month_pending = r3 ? r3.pending : 0;
        db.get('SELECT IFNULL(SUM(balance),0) as older_pending FROM payments WHERE month <> ?', [monthKey], (e4, r4) => {
          stats.older_pending = r4 ? r4.older_pending : r4 ? r4.pending : 0;
          res.render('home', { stats, monthKey });
        });
      });
    });
  });
});

// ---------- Customers ----------
app.get('/customer/add', requireAuth, (req, res) => {
  db.all('SELECT * FROM items ORDER BY item_type,name', [], (err, items) => {
    res.render('customer_add', { items });
  });
});
app.post('/customer/add', requireAuth, (req, res) => {
  const { name, address, mobile, area_name, route_name } = req.body;
  db.run(`INSERT INTO customers (name,address,mobile,area_name,route_name,active) VALUES(?,?,?,?,?,1)`,
    [name, address, mobile, area_name || null, route_name || null], function(err) {
      if (err) return res.send('Error creating customer');
      res.redirect('/customers');
  });
});

app.get('/customers', requireAuth, (req, res) => {
  db.all('SELECT * FROM customers ORDER BY name', [], (err, rows) => {
    res.render('customer_list', { customers: rows });
  });
});

// view single customer and their subscriptions
app.get('/customer/view/:id', requireAuth, (req, res) => {
  const id = req.params.id;
  db.get('SELECT * FROM customers WHERE id = ?', [id], (e, customer) => {
    if (!customer) return res.send('Customer not found');
    db.all('SELECT s.*, i.name as item_name FROM subscriptions s LEFT JOIN items i ON i.id = s.item_id WHERE s.customer_id = ?', [id], (err, subs) => {
      db.all('SELECT * FROM payments WHERE customer_id = ? ORDER BY month DESC LIMIT 24', [id], (err2, payments) => {
        res.render('customer_view', { customer, subscriptions: subs, payments });
      });
    });
  });
});

// ---------- Items (news/magazines rates) ----------
app.get('/items', requireAuth, (req, res) => {
  db.all('SELECT * FROM items ORDER BY item_type,name', [], (err, rows) => res.render('items', { items: rows }));
});
app.post('/items', requireAuth, (req, res) => {
  const { name, item_type, monthly_rate } = req.body;
  db.run(`INSERT OR REPLACE INTO items (name,item_type,monthly_rate) VALUES(?,?,?)`,
    [name, item_type || 'newspaper', monthly_rate || 295], () => res.redirect('/items'));
});

// ---------- Subscriptions ----------
app.get('/subscription/add/:customerId', requireAuth, (req, res) => {
  const cid = req.params.customerId;
  db.get('SELECT * FROM customers WHERE id = ?', [cid], (e, customer) => {
    if (!customer) return res.send('Customer not found');
    db.all('SELECT * FROM items ORDER BY item_type,name', [], (err, items) => res.render('subscription_add', { customer, items }));
  });
});
app.post('/subscription/add', requireAuth, (req, res) => {
  const { customer_id, item_id, monthly_rate, is_annual, is_sunday_only, start_date, end_date } = req.body;
  // fetch item name if item_id given
  db.get('SELECT name,monthly_rate FROM items WHERE id = ?', [item_id || null], (err, item) => {
    const itemName = item ? item.name : req.body.item_name || 'Custom';
    const rate = monthly_rate || (item ? item.monthly_rate : 295);
    const s_annual = is_annual ? 1 : 0;
    const s_sunday = is_sunday_only ? 1 : 0;
    db.run(`INSERT INTO subscriptions (customer_id,item_id,item_name,monthly_rate,is_annual,is_sunday_only,start_date,end_date,active)
            VALUES(?,?,?,?,?,?,?,?,1)`, [customer_id, item_id || null, itemName, rate, s_annual, s_sunday, start_date || null, end_date || null], function(err2) {
      if (err2) return res.send('Error adding subscription: ' + err2.message);
      // If annual prepaid, optionally one could create a payment record immediately when paid (not implemented automatically)
      res.redirect('/customer/view/' + customer_id);
    });
  });
});

// stop subscription (deactivate)
app.post('/subscription/stop/:id', requireAuth, (req, res) => {
  const id = req.params.id;
  const { end_date } = req.body;
  db.run('UPDATE subscriptions SET active=0, end_date=? WHERE id=?', [end_date || null, id], () => res.send('Subscription stopped'));
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
  res.render('billing_form', { month: m });
});

app.post('/billing/generate', requireAuth, async (req, res) => {
  // month param 'YYYY-MM' or default current
  const monthKey = req.body.month || getCurrentMonthKey();
  const [yearStr, monthStr] = monthKey.split('-').map(x=>Number(x));
  const year = yearStr, month = monthStr;
  // select active subscriptions overlapping the month and which are not annual
  db.all(`SELECT s.*, c.name as customer_name, c.mobile as customer_mobile
          FROM subscriptions s
          JOIN customers c ON c.id = s.customer_id
          WHERE s.active = 1
            AND (s.end_date IS NULL OR date(s.end_date) >= date(?,'start of month'))
            AND (s.start_date IS NULL OR date(s.start_date) <= date(?,'start of month','+1 month','-1 day'))`,
    [monthKey + '-01', monthKey + '-01'], async (err, subs) => {
    if (err) return res.send('Error selecting subscriptions: '+err.message);

    const tasks = subs.map(s => new Promise(async (resolve) => {
      // skip annual prepaid subscribers
      if (s.is_annual == 1) return resolve();

      // determine month-specific amount
      let amountForThisSub = 0;
      if (s.is_sunday_only == 1) {
        // sunday-only: ₹10 per sunday
        const sundays = countSundaysInMonth(year, month);
        amountForThisSub = sundays * 10;
      } else {
        // monthly subscriber
        amountForThisSub = Number(s.monthly_rate || 295);
      }

      // add carry-forward (last balance)
      const prevBalance = await getLastBalance(s.id); // returns number
      const totalDue = (prevBalance || 0) + amountForThisSub;

      // check if payment row exists for this sub + month
      db.get('SELECT * FROM payments WHERE subscription_id = ? AND month = ?', [s.id, monthKey], (e2, existing) => {
        if (existing) {
          // update amount_due and balance but keep amount_paid unchanged
          const newBalance = totalDue - (existing.amount_paid || 0);
          db.run('UPDATE payments SET amount_due=?, balance=? WHERE id=?', [totalDue, newBalance, existing.id], () => resolve());
        } else {
          db.run('INSERT INTO payments (subscription_id,customer_id,month,amount_due,amount_paid,balance) VALUES(?,?,?,?,?,?)',
            [s.id, s.customer_id, monthKey, totalDue, 0, totalDue], () => resolve());
        }
      });
    }));

    Promise.all(tasks).then(()=> res.send(`Billing generated for ${monthKey}. Processed ${subs.length} subscriptions (annual skipped).`));
  });
});

// ---------- Payments ----------
app.get('/payment/customer/:customerId', requireAuth, (req, res) => {
  const cid = req.params.customerId;
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

// export payments by month
app.get('/export/payments', requireAuth, (req, res) => {
  const month = req.query.month;
  if (!month) return res.send('Provide ?month=YYYY-MM');
  db.all('SELECT p.*, s.item_name, c.name as customer_name, c.mobile as customer_mobile FROM payments p LEFT JOIN subscriptions s ON s.id=p.subscription_id LEFT JOIN customers c ON c.id=p.customer_id WHERE p.month = ?', [month], (err, rows) => {
    const data = rows.map(r => ({ customer: r.customer_name, mobile: r.customer_mobile, item: r.item_name, month: r.month, amount_due: r.amount_due, amount_paid: r.amount_paid, balance: r.balance }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Payments');
    const fn = `payments_${month}.xlsx`;
    XLSX.writeFile(wb, fn);
    res.download(path.join(__dirname, fn));
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
app.get('/reports', requireAuth, (req, res) => res.render('reports'));
app.get('/delivery', requireAuth, (req, res) => {
  db.all('SELECT * FROM delivery_boys', [], (e, rows) => res.render('delivery', { boys: rows }));
});

// ---------- Start ----------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));