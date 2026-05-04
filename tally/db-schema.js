/**
 * Tally integration DB schema — run once on startup via initTallySchema(db)
 * All tables are company-scoped for multi-tenant SaaS support.
 */

function initTallySchema(db) {
    const tables = [

        // Tenants / clients
        `CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            gstin TEXT DEFAULT "",
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`,

        // Per-company Tally connection + preferences
        `CREATE TABLE IF NOT EXISTS tally_config (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL UNIQUE,
            tally_url TEXT DEFAULT "http://localhost:9000",
            tally_company_name TEXT DEFAULT "",
            narration_template TEXT DEFAULT "Being purchase vide Invoice No. {invoice_number} dated {invoice_date} from {vendor_name}",
            voucher_numbering TEXT DEFAULT "invoice_number",
            roundoff_ledger TEXT DEFAULT "Round Off",
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (company_id) REFERENCES companies(id)
        )`,

        // GST ledger mapping: tax_type (cgst/sgst/igst/cess) + rate → Tally ledger name
        `CREATE TABLE IF NOT EXISTS tally_gst_ledger_map (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            tax_type TEXT NOT NULL,
            tax_rate REAL NOT NULL,
            ledger_name TEXT NOT NULL,
            UNIQUE(company_id, tax_type, tax_rate),
            FOREIGN KEY (company_id) REFERENCES companies(id)
        )`,

        // Purchase/expense ledger mapping: gst_type (igst/local/exempt) + optional rate/category → ledger
        // tax_rate uses -1 sentinel and category uses "" to allow UNIQUE without COALESCE (unsupported in SQLite constraints)
        `CREATE TABLE IF NOT EXISTS tally_purchase_ledger_map (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            gst_type TEXT NOT NULL,
            tax_rate REAL NOT NULL DEFAULT -1,
            category TEXT NOT NULL DEFAULT "",
            ledger_name TEXT NOT NULL,
            UNIQUE(company_id, gst_type, tax_rate, category),
            FOREIGN KEY (company_id) REFERENCES companies(id)
        )`,

        // Vendor mapping: GSTIN → Tally ledger name + group
        `CREATE TABLE IF NOT EXISTS tally_vendor_map (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            vendor_gstin TEXT NOT NULL,
            vendor_name TEXT DEFAULT "",
            ledger_name TEXT NOT NULL,
            tally_group TEXT DEFAULT "Sundry Creditors",
            UNIQUE(company_id, vendor_gstin),
            FOREIGN KEY (company_id) REFERENCES companies(id)
        )`,

        // Stock item mapping: invoice description → Tally stock item name
        `CREATE TABLE IF NOT EXISTS tally_stock_item_map (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            invoice_description TEXT NOT NULL,
            tally_item_name TEXT NOT NULL,
            UNIQUE(company_id, invoice_description),
            FOREIGN KEY (company_id) REFERENCES companies(id)
        )`,

        // RCM config: tax_type + rate → liability ledger + input credit ledger
        `CREATE TABLE IF NOT EXISTS tally_rcm_config (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            tax_type TEXT NOT NULL,
            tax_rate REAL NOT NULL,
            liability_ledger TEXT NOT NULL,
            input_credit_ledger TEXT NOT NULL,
            UNIQUE(company_id, tax_type, tax_rate),
            FOREIGN KEY (company_id) REFERENCES companies(id)
        )`,

        // Default company seed so single-tenant setups work without setup
        `INSERT OR IGNORE INTO companies (id, name) VALUES (1, "Default Company")`
    ];

    // Columns to add to expenses if not present
    const expenseAlters = [
        `ALTER TABLE expenses ADD COLUMN company_id INTEGER DEFAULT 1`,
        `ALTER TABLE expenses ADD COLUMN tally_pushed INTEGER DEFAULT 0`,
        `ALTER TABLE expenses ADD COLUMN tally_pushed_at DATETIME DEFAULT NULL`,
        `ALTER TABLE expenses ADD COLUMN tally_voucher_type TEXT DEFAULT ""`
    ];

    return new Promise((resolve, reject) => {
        db.serialize(() => {
            tables.forEach(sql => {
                db.run(sql, err => {
                    if (err) console.error('Tally schema error:', err.message);
                });
            });

            // ALTER TABLE ignores "duplicate column" errors intentionally
            expenseAlters.forEach(sql => {
                db.run(sql, err => {
                    if (err && !err.message.includes('duplicate column')) {
                        console.error('Expenses alter error:', err.message);
                    }
                });
            });

            db.run('SELECT 1', err => {
                if (err) reject(err);
                else {
                    console.log('✅ Tally schema initialized');
                    resolve();
                }
            });
        });
    });
}

module.exports = { initTallySchema };
