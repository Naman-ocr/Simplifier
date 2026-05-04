/**
 * Tally push engine — DB record → XML → Tally
 * Handles: multi-rate invoices, RCM, round-off, vendor auto-create, narration templates
 */

const express = require('express');
const axios = require('axios');
const router = express.Router();

// ─── helpers ─────────────────────────────────────────────────────────────────

function dbGet(db, sql, params = []) {
    return new Promise((resolve, reject) =>
        db.get(sql, params, (err, row) => (err ? reject(err) : resolve(row)))
    );
}
function dbAll(db, sql, params = []) {
    return new Promise((resolve, reject) =>
        db.all(sql, params, (err, rows) => (err ? reject(err) : resolve(rows)))
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

function tallyDate(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    if (isNaN(d)) return dateStr.replace(/-/g, '');
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${y}${m}${dd}`;
}

function escapeXml(str) {
    return String(str || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

function buildNarration(template, invoice) {
    return template
        .replace('{invoice_number}', invoice.invoice_number || '')
        .replace('{invoice_date}', invoice.invoice_date || '')
        .replace('{vendor_name}', invoice.vendor_name || '')
        .replace('{vendor_gstn}', invoice.vendor_gstn || '')
        .replace('{invoice_value}', invoice.invoice_value || '')
        .replace('{customer_name}', invoice.customer_name || '');
}

function round2(n) { return Math.round((n || 0) * 100) / 100; }

async function postToTally(url, xml) {
    const res = await axios.post(url, xml, {
        headers: { 'Content-Type': 'application/xml' },
        timeout: 30000
    });
    return res.data;
}

function parseTallyResponse(xmlStr) {
    const created = parseInt((xmlStr.match(/<CREATED>(\d+)<\/CREATED>/) || [])[1] || 0);
    const altered = parseInt((xmlStr.match(/<ALTERED>(\d+)<\/ALTERED>/) || [])[1] || 0);
    const exceptions = parseInt((xmlStr.match(/<EXCEPTIONS>(\d+)<\/EXCEPTIONS>/) || [])[1] || 0);
    const errorMsg = (xmlStr.match(/<LINEERROR>(.*?)<\/LINEERROR>/) || [])[1] || '';
    return { created, altered, exceptions, errorMsg };
}

// ─── vendor auto-create XML ───────────────────────────────────────────────────

function buildVendorCreateXml(companyName, vendorLedgerName, vendorGroup, vendorGstin, vendorAddress) {
    return `<ENVELOPE>
<HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
<BODY><IMPORTDATA><REQUESTDESC>
<REPORTNAME>All Masters</REPORTNAME>
<STATICVARIABLES><SVCURRENTCOMPANY>${escapeXml(companyName)}</SVCURRENTCOMPANY></STATICVARIABLES>
</REQUESTDESC>
<REQUESTDATA><TALLYMESSAGE xmlns:UDF="TallyUDF">
<LEDGER NAME="${escapeXml(vendorLedgerName)}" ACTION="Create">
<NAME>${escapeXml(vendorLedgerName)}</NAME>
<PARENT>${escapeXml(vendorGroup)}</PARENT>
<GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>
<PARTYGSTIN>${escapeXml(vendorGstin)}</PARTYGSTIN>
<ADDRESS>${escapeXml(vendorAddress || '')}</ADDRESS>
<MAILINGNAME>${escapeXml(vendorLedgerName)}</MAILINGNAME>
<ISBILLWISEON>Yes</ISBILLWISEON>
<AFFECTSSTOCK>No</AFFECTSSTOCK>
</LEDGER>
</TALLYMESSAGE></REQUESTDATA></IMPORTDATA></BODY></ENVELOPE>`;
}

// ─── main voucher XML builder ────────────────────────────────────────────────

function buildVoucherXml(companyName, voucherData) {
    const {
        voucher_type, voucher_number, voucher_date, narration,
        due_date, bill_ref, is_rcm,
        vendor_ledger, ledger_lines
    } = voucherData;

    const allLedgers = ledger_lines
        .map(l => `
    <ALLLEDGERENTRIES.LIST>
        <LEDGERNAME>${escapeXml(l.ledger)}</LEDGERNAME>
        <ISDEEMEDPOSITIVE>${l.amount < 0 ? 'Yes' : 'No'}</ISDEEMEDPOSITIVE>
        <AMOUNT>${round2(l.amount)}</AMOUNT>
        ${l.bill_ref ? `<BILLALLOCATIONS.LIST>
            <NAME>${escapeXml(l.bill_ref)}</NAME>
            <BILLTYPE>New Ref</BILLTYPE>
            <AMOUNT>${round2(l.amount)}</AMOUNT>
            ${due_date ? `<DUEDATE>${tallyDate(due_date)}</DUEDATE>` : ''}
        </BILLALLOCATIONS.LIST>` : ''}
    </ALLLEDGERENTRIES.LIST>`).join('');

    return `<ENVELOPE>
<HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
<BODY><IMPORTDATA><REQUESTDESC>
<REPORTNAME>Vouchers</REPORTNAME>
<STATICVARIABLES><SVCURRENTCOMPANY>${escapeXml(companyName)}</SVCURRENTCOMPANY></STATICVARIABLES>
</REQUESTDESC>
<REQUESTDATA><TALLYMESSAGE xmlns:UDF="TallyUDF">
<VOUCHER REMOTEID="${escapeXml(voucher_number)}" VCHTYPE="${escapeXml(voucher_type)}" ACTION="Create">
<DATE>${tallyDate(voucher_date)}</DATE>
<VOUCHERTYPENAME>${escapeXml(voucher_type)}</VOUCHERTYPENAME>
<VOUCHERNUMBER>${escapeXml(voucher_number)}</VOUCHERNUMBER>
<PARTYLEDGERNAME>${escapeXml(vendor_ledger)}</PARTYLEDGERNAME>
<NARRATION>${escapeXml(narration)}</NARRATION>
<ISOPTIONAL>No</ISOPTIONAL>
<ISINVOICE>Yes</ISINVOICE>
${is_rcm ? '<ISRCMAPPLICABLE>Yes</ISRCMAPPLICABLE>' : ''}
${allLedgers}
</VOUCHER>
</TALLYMESSAGE></REQUESTDATA></IMPORTDATA></BODY></ENVELOPE>`;
}

// ─── push endpoint ────────────────────────────────────────────────────────────

router.post('/expenses/:id/push-to-tally', async (req, res) => {
    const db = req.app.get('db');
    const expenseId = req.params.id;

    try {
        // 1. Load invoice
        const invoice = await dbGet(db, 'SELECT * FROM expenses WHERE id = ?', [expenseId]);
        if (!invoice) return res.status(404).json({ error: 'Invoice not found' });
        if (invoice.status !== 'approved') return res.status(400).json({ error: 'Only approved invoices can be pushed to Tally' });
        if (invoice.tally_pushed) return res.status(400).json({ error: 'Invoice already pushed to Tally', pushed_at: invoice.tally_pushed_at });

        const {
            company_id = 1,
            voucher_type,
            narration,
            due_date,
            is_rcm = false,
            vendor_ledger_name,
            vendor_group,
            purchase_ledger_overrides = {},   // { "rate_gsttype_key": "Ledger Name" }
            stock_item_overrides = {},          // { "invoice_description": "Tally Item Name" }
        } = req.body;

        // 2. Load config
        const cfg = await dbGet(db, 'SELECT * FROM tally_config WHERE company_id = ?', [company_id]);
        const tallyUrl = cfg ? cfg.tally_url : (process.env.TALLY_URL || 'http://localhost:9000');
        const companyName = cfg ? cfg.tally_company_name : (process.env.TALLY_COMPANY_NAME || '');
        const narrationTemplate = cfg ? cfg.narration_template : 'Being purchase vide Invoice No. {invoice_number} dated {invoice_date} from {vendor_name}';
        const roundoffLedger = cfg ? cfg.roundoff_ledger : 'Round Off';
        const voucherNumbering = cfg ? cfg.voucher_numbering : 'invoice_number';

        // 3. Resolve narration
        const finalNarration = narration || buildNarration(narrationTemplate, invoice);

        // 4. Resolve voucher number
        const voucherNumber = voucherNumbering === 'invoice_number' ? (invoice.invoice_number || String(expenseId)) : '';

        // 5. Resolve vendor ledger
        let vendorLedger = vendor_ledger_name;
        if (!vendorLedger && invoice.vendor_gstn) {
            const vm = await dbGet(db, 'SELECT ledger_name FROM tally_vendor_map WHERE company_id = ? AND vendor_gstin = ?', [company_id, invoice.vendor_gstn]);
            if (vm) vendorLedger = vm.ledger_name;
        }
        if (!vendorLedger) vendorLedger = invoice.vendor_name || 'Unknown Vendor';

        // 6. Auto-create vendor in Tally if requested
        if (req.body.create_vendor) {
            const group = vendor_group || 'Sundry Creditors';
            const createXml = buildVendorCreateXml(companyName, vendorLedger, group, invoice.vendor_gstn, invoice.vendor_address);
            await postToTally(tallyUrl, createXml);
            // Save to vendor map
            await dbRun(db, `
                INSERT INTO tally_vendor_map (company_id, vendor_gstin, vendor_name, ledger_name, tally_group)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(company_id, vendor_gstin) DO UPDATE SET ledger_name = excluded.ledger_name, tally_group = excluded.tally_group
            `, [company_id, invoice.vendor_gstn || '', invoice.vendor_name || '', vendorLedger, group]);
        }

        // 7. Build ledger lines
        const lineItems = typeof invoice.line_items === 'string' ? JSON.parse(invoice.line_items || '[]') : (invoice.line_items || []);
        const gstMap = await dbAll(db, 'SELECT * FROM tally_gst_ledger_map WHERE company_id = ?', [company_id]);
        const purchaseMap = await dbAll(db, 'SELECT * FROM tally_purchase_ledger_map WHERE company_id = ?', [company_id]);
        const rcmMap = is_rcm ? await dbAll(db, 'SELECT * FROM tally_rcm_config WHERE company_id = ?', [company_id]) : [];

        const gstLookup = (type, rate) => {
            const m = gstMap.find(g => g.tax_type === type && Math.abs(g.tax_rate - rate) < 0.01);
            return m ? m.ledger_name : `${type.toUpperCase()} @ ${rate}%`;
        };

        const purchaseLookup = (gstType, rate) => {
            // Check override from request first
            const key = `${gstType}_${rate}`;
            if (purchase_ledger_overrides[key]) return purchase_ledger_overrides[key];
            // Then saved map: exact rate match, fallback to wildcard (-1 sentinel = any rate)
            let m = purchaseMap.find(p => p.gst_type === gstType && p.tax_rate >= 0 && Math.abs(p.tax_rate - rate) < 0.01);
            if (!m) m = purchaseMap.find(p => p.gst_type === gstType && (p.tax_rate === -1 || p.tax_rate === null));
            return m ? m.ledger_name : null;
        };

        const rcmLookup = (type, rate) => {
            const m = rcmMap.find(r => r.tax_type === type && Math.abs(r.tax_rate - rate) < 0.01);
            return m || null;
        };

        // Determine invoice GST type: igst or local
        const isIgst = (invoice.igst_amount || 0) > 0 && (invoice.cgst_amount || 0) === 0;
        const gstType = isIgst ? 'igst' : 'local';

        // Group line items by tax rate for multi-rate support
        const rateGroups = {};
        let totalTaxable = 0;
        let totalCgst = 0, totalSgst = 0, totalIgst = 0, totalCess = 0;

        if (lineItems.length > 0) {
            lineItems.forEach(li => {
                const rate = li.tax_rate || 0;
                const key = String(rate);
                if (!rateGroups[key]) rateGroups[key] = { taxable: 0, cgst: 0, sgst: 0, igst: 0, cess: 0, rate };
                rateGroups[key].taxable += round2(li.line_total || 0);
                rateGroups[key].cgst += round2(li.cgst_amount || 0);
                rateGroups[key].sgst += round2(li.sgst_amount || 0);
                rateGroups[key].igst += round2(li.igst_amount || 0);
            });
        } else {
            // No line items — use invoice-level totals as single rate group
            const rate = invoice.igst_rate || invoice.cgst_rate ? (invoice.igst_rate || invoice.cgst_rate * 2) : 0;
            rateGroups[String(rate)] = {
                rate,
                taxable: round2(invoice.taxable_amount || 0),
                cgst: round2(invoice.cgst_amount || 0),
                sgst: round2(invoice.sgst_amount || 0),
                igst: round2(invoice.igst_amount || 0),
                cess: round2(invoice.cess_amount || 0)
            };
        }

        // Build debit ledger lines (purchase/expense + GST)
        const ledgerLines = [];
        let debitTotal = 0;

        for (const [, grp] of Object.entries(rateGroups)) {
            // Purchase/expense line
            const purchaseLedger = purchaseLookup(gstType, grp.rate);
            if (!purchaseLedger) {
                return res.status(400).json({
                    error: `No purchase ledger mapped for ${gstType} @ ${grp.rate}%. Please configure it in Tally Setup.`,
                    missing_mapping: { gst_type: gstType, tax_rate: grp.rate }
                });
            }

            if (is_rcm) {
                // Under RCM: purchase DR at full value (taxable + tax), vendor CR at full value
                const rcmEntry = rcmLookup(isIgst ? 'igst' : 'cgst', grp.rate);
                if (rcmEntry) {
                    ledgerLines.push({ ledger: rcmEntry.liability_ledger, amount: -(grp.igst || grp.cgst + grp.sgst) });
                    ledgerLines.push({ ledger: rcmEntry.input_credit_ledger, amount: grp.igst || grp.cgst + grp.sgst });
                }
                ledgerLines.push({ ledger: purchaseLedger, amount: grp.taxable });
                debitTotal += grp.taxable;
            } else {
                ledgerLines.push({ ledger: purchaseLedger, amount: grp.taxable });
                debitTotal += grp.taxable;

                // GST lines
                if (isIgst && grp.igst > 0) {
                    const igstLedger = gstLookup('igst', grp.rate);
                    ledgerLines.push({ ledger: igstLedger, amount: grp.igst });
                    debitTotal += grp.igst;
                } else {
                    const halfRate = round2(grp.rate / 2);
                    if (grp.cgst > 0) {
                        const cgstLedger = gstLookup('cgst', halfRate);
                        ledgerLines.push({ ledger: cgstLedger, amount: grp.cgst });
                        debitTotal += grp.cgst;
                    }
                    if (grp.sgst > 0) {
                        const sgstLedger = gstLookup('sgst', halfRate);
                        ledgerLines.push({ ledger: sgstLedger, amount: grp.sgst });
                        debitTotal += grp.sgst;
                    }
                }
                if (grp.cess > 0) {
                    const cessLedger = gstLookup('cess', grp.rate);
                    ledgerLines.push({ ledger: cessLedger, amount: grp.cess });
                    debitTotal += grp.cess;
                }
            }
        }

        // Vendor credit line (negative = credit in Tally)
        const invoiceValue = round2(invoice.invoice_value || 0);
        const diff = round2(debitTotal - invoiceValue);

        if (Math.abs(diff) > 0.02 && roundoffLedger) {
            ledgerLines.push({ ledger: roundoffLedger, amount: round2(-diff) });
        }

        // Vendor CR — with bill ref for outstanding tracking
        ledgerLines.push({
            ledger: vendorLedger,
            amount: -invoiceValue,
            bill_ref: invoice.invoice_number || String(expenseId)
        });

        // 8. Build and send XML
        const voucherData = {
            voucher_type: voucher_type || invoice.tally_voucher_type || 'Purchase',
            voucher_number: voucherNumber,
            voucher_date: invoice.invoice_date,
            narration: finalNarration,
            due_date,
            is_rcm,
            vendor_ledger: vendorLedger,
            ledger_lines: ledgerLines
        };

        const voucherXml = buildVoucherXml(companyName, voucherData);
        const tallyResponse = await postToTally(tallyUrl, voucherXml);
        const result = parseTallyResponse(tallyResponse);

        if (result.exceptions > 0 && result.created === 0) {
            return res.status(400).json({
                error: `Tally rejected the voucher: ${result.errorMsg || 'Check Tally for details'}`,
                tally_response: tallyResponse
            });
        }

        // 9. Mark as pushed
        await dbRun(db,
            'UPDATE expenses SET tally_pushed = 1, tally_pushed_at = CURRENT_TIMESTAMP, tally_voucher_type = ? WHERE id = ?',
            [voucherData.voucher_type, expenseId]
        );

        // 10. Save vendor map for future
        if (invoice.vendor_gstn && vendorLedger) {
            await dbRun(db, `
                INSERT INTO tally_vendor_map (company_id, vendor_gstin, vendor_name, ledger_name, tally_group)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(company_id, vendor_gstin) DO NOTHING
            `, [company_id, invoice.vendor_gstn, invoice.vendor_name || '', vendorLedger, vendor_group || 'Sundry Creditors']);
        }

        res.json({
            success: true,
            created: result.created,
            voucher_type: voucherData.voucher_type,
            voucher_number: voucherNumber,
            pushed_at: new Date().toISOString()
        });

    } catch (e) {
        console.error('Push to Tally error:', e.message);
        if (e.code === 'ECONNREFUSED' || e.code === 'ETIMEDOUT') {
            return res.status(503).json({
                error: 'Cannot connect to Tally. Please ensure Tally is open with the correct company selected, then try again.'
            });
        }
        res.status(500).json({ error: e.message });
    }
});

// Preflight: check what's missing before the push modal opens
router.get('/expenses/:id/tally-preflight', async (req, res) => {
    const db = req.app.get('db');
    const companyId = req.query.company_id || 1;

    try {
        const invoice = await dbGet(db, 'SELECT * FROM expenses WHERE id = ?', [req.params.id]);
        if (!invoice) return res.status(404).json({ error: 'Invoice not found' });

        const isIgst = (invoice.igst_amount || 0) > 0 && (invoice.cgst_amount || 0) === 0;
        const gstType = isIgst ? 'igst' : 'local';

        const lineItems = typeof invoice.line_items === 'string' ? JSON.parse(invoice.line_items || '[]') : (invoice.line_items || []);
        const rates = lineItems.length > 0
            ? [...new Set(lineItems.map(li => li.tax_rate || 0))]
            : [(invoice.igst_rate || invoice.cgst_rate ? (invoice.igst_rate || invoice.cgst_rate * 2) : 0)];

        const purchaseMap = await dbAll(db, 'SELECT * FROM tally_purchase_ledger_map WHERE company_id = ?', [companyId]);
        const vendorMap = invoice.vendor_gstn
            ? await dbGet(db, 'SELECT * FROM tally_vendor_map WHERE company_id = ? AND vendor_gstin = ?', [companyId, invoice.vendor_gstn])
            : null;

        const missingPurchaseLedgers = rates.filter(rate => {
            const exact = purchaseMap.find(p => p.gst_type === gstType && p.tax_rate >= 0 && Math.abs(p.tax_rate - rate) < 0.01);
            const fallback = purchaseMap.find(p => p.gst_type === gstType && (p.tax_rate === -1 || p.tax_rate === null));
            return !exact && !fallback;
        });

        res.json({
            invoice_id: invoice.id,
            already_pushed: !!invoice.tally_pushed,
            pushed_at: invoice.tally_pushed_at,
            vendor: {
                name: invoice.vendor_name,
                gstin: invoice.vendor_gstn,
                mapped_ledger: vendorMap ? vendorMap.ledger_name : null,
                needs_mapping: !vendorMap
            },
            gst_type: gstType,
            rates,
            missing_purchase_ledgers: missingPurchaseLedgers,
            has_line_items: lineItems.length > 0,
            rcm_detected: !!(invoice.extracted_text && /reverse.charge.*yes/i.test(invoice.extracted_text))
        });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

module.exports = router;
