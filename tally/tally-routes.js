/**
 * Tally integration routes — mounted at /api by server.js
 * Covers: companies, tally config, live Tally data fetch, all mapping tables
 */

const express = require('express');
const axios = require('axios');
const router = express.Router();

// ─── helpers ────────────────────────────────────────────────────────────────

function dbAll(db, sql, params = []) {
    return new Promise((resolve, reject) =>
        db.all(sql, params, (err, rows) => (err ? reject(err) : resolve(rows)))
    );
}
function dbGet(db, sql, params = []) {
    return new Promise((resolve, reject) =>
        db.get(sql, params, (err, row) => (err ? reject(err) : resolve(row)))
    );
}
function dbRun(db, sql, params = []) {
    return new Promise((resolve, reject) =>
        db.run(sql, params, function (err) {
            if (err) reject(err);
            else resolve({ lastID: this.lastID, changes: this.changes });
        })
    );
}

async function getTallyUrl(db, companyId) {
    const cfg = await dbGet(db, 'SELECT tally_url FROM tally_config WHERE company_id = ?', [companyId]);
    return cfg ? cfg.tally_url : (process.env.TALLY_URL || 'http://localhost:9000');
}

async function postToTally(url, xml) {
    const res = await axios.post(url, xml, {
        headers: { 'Content-Type': 'application/xml' },
        timeout: 15000
    });
    return res.data;
}

// ─── companies ───────────────────────────────────────────────────────────────

router.get('/companies', async (req, res) => {
    try {
        const db = req.app.get('db');
        const rows = await dbAll(db, 'SELECT * FROM companies ORDER BY id');
        res.json(rows);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

router.post('/companies', async (req, res) => {
    try {
        const db = req.app.get('db');
        const { name, gstin = '' } = req.body;
        if (!name) return res.status(400).json({ error: 'name is required' });
        const result = await dbRun(db, 'INSERT INTO companies (name, gstin) VALUES (?, ?)', [name, gstin]);
        res.json({ id: result.lastID, name, gstin });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

router.put('/companies/:id', async (req, res) => {
    try {
        const db = req.app.get('db');
        const { name, gstin } = req.body;
        await dbRun(db, 'UPDATE companies SET name = ?, gstin = ? WHERE id = ?', [name, gstin, req.params.id]);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// ─── tally config (per company) ──────────────────────────────────────────────

router.get('/tally/config/:companyId', async (req, res) => {
    try {
        const db = req.app.get('db');
        const row = await dbGet(db, 'SELECT * FROM tally_config WHERE company_id = ?', [req.params.companyId]);
        res.json(row || {});
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

router.post('/tally/config', async (req, res) => {
    try {
        const db = req.app.get('db');
        const {
            company_id, tally_url, tally_company_name,
            narration_template, voucher_numbering, roundoff_ledger
        } = req.body;
        await dbRun(db, `
            INSERT INTO tally_config (company_id, tally_url, tally_company_name, narration_template, voucher_numbering, roundoff_ledger)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(company_id) DO UPDATE SET
                tally_url = excluded.tally_url,
                tally_company_name = excluded.tally_company_name,
                narration_template = excluded.narration_template,
                voucher_numbering = excluded.voucher_numbering,
                roundoff_ledger = excluded.roundoff_ledger
        `, [company_id, tally_url, tally_company_name, narration_template, voucher_numbering, roundoff_ledger]);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// ─── live Tally data ──────────────────────────────────────────────────────────

router.get('/tally/test-connection', async (req, res) => {
    try {
        const db = req.app.get('db');
        const companyId = req.query.company_id || 1;
        const url = await getTallyUrl(db, companyId);
        const xml = `<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER><BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Companies</REPORTNAME></REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;
        await postToTally(url, xml);
        res.json({ connected: true, url });
    } catch (e) {
        res.json({ connected: false, error: e.message });
    }
});

router.get('/tally/ledgers', async (req, res) => {
    try {
        const db = req.app.get('db');
        const companyId = req.query.company_id || 1;
        const url = await getTallyUrl(db, companyId);
        const cfg = await dbGet(db, 'SELECT tally_company_name FROM tally_config WHERE company_id = ?', [companyId]);
        const company = cfg ? cfg.tally_company_name : '';

        const xml = `<ENVELOPE>
<HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
<BODY><EXPORTDATA><REQUESTDESC>
<REPORTNAME>List of Accounts</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<ACCOUNTTYPE>Ledgers</ACCOUNTTYPE>
</STATICVARIABLES>
</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;

        const data = await postToTally(url, xml);
        const ledgers = parseLedgersFromXml(data);
        res.json({ ledgers });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

router.get('/tally/voucher-types', async (req, res) => {
    try {
        const db = req.app.get('db');
        const companyId = req.query.company_id || 1;
        const url = await getTallyUrl(db, companyId);
        const cfg = await dbGet(db, 'SELECT tally_company_name FROM tally_config WHERE company_id = ?', [companyId]);
        const company = cfg ? cfg.tally_company_name : '';

        const xml = `<ENVELOPE>
<HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
<BODY><EXPORTDATA><REQUESTDESC>
<REPORTNAME>List of Voucher Types</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
</STATICVARIABLES>
</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;

        const data = await postToTally(url, xml);
        const types = parseNamesFromXml(data, 'VOUCHERTYPENAME');
        res.json({ voucher_types: types.length ? types : ['Purchase', 'Sales', 'Journal', 'Receipt', 'Payment', 'Contra', 'Debit Note', 'Credit Note'] });
    } catch (e) {
        res.status(500).json({ error: e.message, voucher_types: ['Purchase', 'Sales', 'Journal', 'Receipt', 'Payment', 'Contra', 'Debit Note', 'Credit Note'] });
    }
});

router.get('/tally/groups', async (req, res) => {
    try {
        const db = req.app.get('db');
        const companyId = req.query.company_id || 1;
        const url = await getTallyUrl(db, companyId);

        const xml = `<ENVELOPE>
<HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
<BODY><EXPORTDATA><REQUESTDESC>
<REPORTNAME>List of Accounts</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<ACCOUNTTYPE>Groups</ACCOUNTTYPE>
</STATICVARIABLES>
</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;

        const data = await postToTally(url, xml);
        const groups = parseNamesFromXml(data, 'GROUPNAME');
        res.json({ groups: groups.length ? groups : ['Sundry Creditors', 'Sundry Debtors', 'Bank Accounts', 'Cash-in-Hand', 'Duties & Taxes', 'Indirect Expenses'] });
    } catch (e) {
        res.status(500).json({ error: e.message, groups: ['Sundry Creditors', 'Sundry Debtors', 'Bank Accounts', 'Cash-in-Hand', 'Duties & Taxes', 'Indirect Expenses'] });
    }
});

router.get('/tally/stock-items', async (req, res) => {
    try {
        const db = req.app.get('db');
        const companyId = req.query.company_id || 1;
        const url = await getTallyUrl(db, companyId);

        const xml = `<ENVELOPE>
<HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
<BODY><EXPORTDATA><REQUESTDESC>
<REPORTNAME>List of Accounts</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<ACCOUNTTYPE>Stock Items</ACCOUNTTYPE>
</STATICVARIABLES>
</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;

        const data = await postToTally(url, xml);
        const items = parseNamesFromXml(data, 'STOCKITEMNAME');
        res.json({ stock_items: items });
    } catch (e) {
        res.status(500).json({ error: e.message, stock_items: [] });
    }
});

// ─── XML parse helpers ────────────────────────────────────────────────────────

function parseLedgersFromXml(xmlStr) {
    const ledgers = [];
    const seen = new Set();

    // TallyPrime format: <LEDGER NAME="Cash"><PARENT>Cash-in-Hand</PARENT>...</LEDGER>
    const blockRegex = /<LEDGER\s[^>]*NAME="([^"]+)"[^>]*>([\s\S]*?)<\/LEDGER>/gi;
    let block;
    while ((block = blockRegex.exec(xmlStr)) !== null) {
        const name = block[1].trim();
        const inner = block[2];
        const parentMatch = inner.match(/<PARENT>(.*?)<\/PARENT>/i);
        const group = parentMatch ? parentMatch[1].trim() : '';
        if (name && !seen.has(name)) {
            seen.add(name);
            ledgers.push({ name, group });
        }
    }

    // Fallback: <NAME> child tags (older Tally format)
    if (ledgers.length === 0) {
        const nameRegex = /<NAME>(.*?)<\/NAME>/g;
        let nm;
        while ((nm = nameRegex.exec(xmlStr)) !== null) {
            const n = nm[1].trim();
            if (n && !seen.has(n)) { seen.add(n); ledgers.push({ name: n, group: '' }); }
        }
    }

    return ledgers;
}

function parseNamesFromXml(xmlStr, tag) {
    const results = [];
    const seen = new Set();

    // Attribute-based: <VOUCHERTYPENAME NAME="Purchase"> or <VOUCHERTYPE NAME="Purchase">
    const attrRegex = new RegExp(`<(?:${tag}|[A-Z]+)\\s[^>]*NAME="([^"]+)"`, 'gi');
    let m;
    while ((m = attrRegex.exec(xmlStr)) !== null) {
        const v = m[1].trim();
        if (v && !seen.has(v)) { seen.add(v); results.push(v); }
    }

    // Child tag: <VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>
    const childRegex = new RegExp(`<${tag}>(.*?)<\/${tag}>`, 'g');
    while ((m = childRegex.exec(xmlStr)) !== null) {
        const v = m[1].trim();
        if (v && !seen.has(v)) { seen.add(v); results.push(v); }
    }

    return results;
}

// ─── Debug: raw Tally response ───────────────────────────────────────────────
router.get('/tally/debug-raw', async (req, res) => {
    try {
        const db = req.app.get('db');
        const companyId = req.query.company_id || 1;
        const url = await getTallyUrl(db, companyId);
        const cfg = await dbGet(db, 'SELECT tally_company_name FROM tally_config WHERE company_id = ?', [companyId]);
        const company = cfg ? cfg.tally_company_name : '';

        const xml = `<ENVELOPE>
<HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
<BODY><EXPORTDATA><REQUESTDESC>
<REPORTNAME>List of Accounts</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<ACCOUNTTYPE>Ledgers</ACCOUNTTYPE>
</STATICVARIABLES>
</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>`;

        const raw = await postToTally(url, xml);
        const parsed = parseLedgersFromXml(raw);
        res.json({ raw: raw.substring(0, 3000), parsed_count: parsed.length, first_5: parsed.slice(0, 5) });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// ─── GST ledger map ───────────────────────────────────────────────────────────

router.get('/tally/gst-map/:companyId', async (req, res) => {
    try {
        const db = req.app.get('db');
        const rows = await dbAll(db, 'SELECT * FROM tally_gst_ledger_map WHERE company_id = ?', [req.params.companyId]);
        res.json(rows);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

router.post('/tally/gst-map', async (req, res) => {
    try {
        const db = req.app.get('db');
        const { company_id, mappings } = req.body; // mappings: [{tax_type, tax_rate, ledger_name}]
        for (const m of mappings) {
            await dbRun(db, `
                INSERT INTO tally_gst_ledger_map (company_id, tax_type, tax_rate, ledger_name)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(company_id, tax_type, tax_rate) DO UPDATE SET ledger_name = excluded.ledger_name
            `, [company_id, m.tax_type, m.tax_rate, m.ledger_name]);
        }
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// ─── Purchase / expense ledger map ────────────────────────────────────────────

router.get('/tally/purchase-map/:companyId', async (req, res) => {
    try {
        const db = req.app.get('db');
        const rows = await dbAll(db, 'SELECT * FROM tally_purchase_ledger_map WHERE company_id = ?', [req.params.companyId]);
        res.json(rows);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

router.post('/tally/purchase-map', async (req, res) => {
    try {
        const db = req.app.get('db');
        const { company_id, gst_type, tax_rate, category, ledger_name } = req.body;
        await dbRun(db, `
            INSERT INTO tally_purchase_ledger_map (company_id, gst_type, tax_rate, category, ledger_name)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(company_id, gst_type, tax_rate, category)
            DO UPDATE SET ledger_name = excluded.ledger_name
        `, [company_id, gst_type, tax_rate ?? -1, category ?? '', ledger_name]);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// ─── Vendor map ───────────────────────────────────────────────────────────────

router.get('/tally/vendor-map/:companyId', async (req, res) => {
    try {
        const db = req.app.get('db');
        const rows = await dbAll(db, 'SELECT * FROM tally_vendor_map WHERE company_id = ?', [req.params.companyId]);
        res.json(rows);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

router.post('/tally/vendor-map', async (req, res) => {
    try {
        const db = req.app.get('db');
        const { company_id, vendor_gstin, vendor_name, ledger_name, tally_group } = req.body;
        await dbRun(db, `
            INSERT INTO tally_vendor_map (company_id, vendor_gstin, vendor_name, ledger_name, tally_group)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(company_id, vendor_gstin) DO UPDATE SET
                vendor_name = excluded.vendor_name,
                ledger_name = excluded.ledger_name,
                tally_group = excluded.tally_group
        `, [company_id, vendor_gstin, vendor_name || '', ledger_name, tally_group || 'Sundry Creditors']);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// ─── Stock item map ───────────────────────────────────────────────────────────

router.get('/tally/stock-map/:companyId', async (req, res) => {
    try {
        const db = req.app.get('db');
        const rows = await dbAll(db, 'SELECT * FROM tally_stock_item_map WHERE company_id = ?', [req.params.companyId]);
        res.json(rows);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

router.post('/tally/stock-map', async (req, res) => {
    try {
        const db = req.app.get('db');
        const { company_id, invoice_description, tally_item_name } = req.body;
        await dbRun(db, `
            INSERT INTO tally_stock_item_map (company_id, invoice_description, tally_item_name)
            VALUES (?, ?, ?)
            ON CONFLICT(company_id, invoice_description) DO UPDATE SET tally_item_name = excluded.tally_item_name
        `, [company_id, invoice_description, tally_item_name]);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// ─── RCM config ───────────────────────────────────────────────────────────────

router.get('/tally/rcm-config/:companyId', async (req, res) => {
    try {
        const db = req.app.get('db');
        const rows = await dbAll(db, 'SELECT * FROM tally_rcm_config WHERE company_id = ?', [req.params.companyId]);
        res.json(rows);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

router.post('/tally/rcm-config', async (req, res) => {
    try {
        const db = req.app.get('db');
        const { company_id, mappings } = req.body; // [{tax_type, tax_rate, liability_ledger, input_credit_ledger}]
        for (const m of mappings) {
            await dbRun(db, `
                INSERT INTO tally_rcm_config (company_id, tax_type, tax_rate, liability_ledger, input_credit_ledger)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(company_id, tax_type, tax_rate) DO UPDATE SET
                    liability_ledger = excluded.liability_ledger,
                    input_credit_ledger = excluded.input_credit_ledger
            `, [company_id, m.tax_type, m.tax_rate, m.liability_ledger, m.input_credit_ledger]);
        }
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

module.exports = router;
