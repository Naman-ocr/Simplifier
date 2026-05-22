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

// ─── unit resolution ──────────────────────────────────────────────────────────

// Global auto-map: normalize common invoice unit spellings to Tally unit names
const UNIT_AUTOMAP = {
    'NOS': 'Nos', 'NO': 'Nos', 'NO.': 'Nos', 'NUM': 'Nos', 'NUMBER': 'Nos', 'UNIT': 'Nos', 'UNITS': 'Nos',
    'PCS': 'Pcs', 'PC': 'Pcs', 'PIECES': 'Pcs', 'PIECE': 'Pcs',
    'KGS': 'Kgs', 'KG': 'Kgs', 'KILO': 'Kgs', 'KILOGRAM': 'Kgs', 'KILOGRAMS': 'Kgs',
    'GMS': 'Gms', 'GM': 'Gms', 'GR': 'Gms', 'GRAM': 'Gms', 'GRAMS': 'Gms',
    'MTR': 'Mtr', 'METER': 'Mtr', 'METRE': 'Mtr', 'METERS': 'Mtr', 'METRES': 'Mtr',
    'LTR': 'Ltr', 'LT': 'Ltr', 'LITRE': 'Ltr', 'LITER': 'Ltr', 'LITRES': 'Ltr', 'LITERS': 'Ltr',
    'BOX': 'Box', 'BOXES': 'Box',
    'SET': 'Set', 'SETS': 'Set',
    'RMT': 'Rmt', 'RUNNING MTR': 'Rmt', 'RUNNING METER': 'Rmt',
    'SQM': 'Sqm', 'SQMTR': 'Sqm', 'SQ.M': 'Sqm', 'SQ MTR': 'Sqm',
    'SQF': 'Sqf', 'SQFT': 'Sqf', 'SQ.FT': 'Sqf', 'SQ FT': 'Sqf',
    'PKT': 'Pkt', 'PACKET': 'Pkt', 'PACKETS': 'Pkt',
    'BTL': 'Btl', 'BOTTLE': 'Btl', 'BOTTLES': 'Btl',
    'PAIR': 'Pair', 'PR': 'Pair', 'PAIRS': 'Pair',
    'DOZ': 'Doz', 'DOZEN': 'Doz', 'DZ': 'Doz',
    'TON': 'Ton', 'TONNE': 'Ton', 'TONNES': 'Ton', 'TONS': 'Ton',
    'MTS': 'Ton', 'MT': 'Mtr'  // MT is ambiguous; default to Mtr (metric ton is MTS)
};

/**
 * Resolve invoice unit → Tally unit + optional quantity conversion factor.
 * @param {string} invoiceUnit - raw unit from invoice (e.g. "KGS", "mtr", "NOS")
 * @param {Array}  customUnitMap - rows from tally_unit_map for this company
 * @param {boolean} enableConversion - whether cross-unit conversion is enabled
 * @returns {{ tally_unit: string, factor: number }}
 */
function resolveUnit(invoiceUnit, customUnitMap = [], enableConversion = false) {
    const normalized = (invoiceUnit || '').trim().toUpperCase().replace(/\.+$/, '');

    // 1. Company-specific override (only when conversion feature is enabled)
    if (enableConversion && customUnitMap.length) {
        const custom = customUnitMap.find(u => u.invoice_unit.toUpperCase() === normalized);
        if (custom) return { tally_unit: custom.tally_unit, factor: parseFloat(custom.conversion_factor) || 1.0 };
    }

    // 2. Global auto-normalization (always active)
    const autoMapped = UNIT_AUTOMAP[normalized];
    if (autoMapped) return { tally_unit: autoMapped, factor: 1.0 };

    // 3. Fallback — use as-is (let Tally handle unknown units)
    return { tally_unit: invoiceUnit || 'Nos', factor: 1.0 };
}

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
    const errorMsg = (xmlStr.match(/<LINEERROR>(.*?)<\/LINEERROR>/i) || [])[1] || '';
    // Tally sometimes returns <RESPONSE><LINEERROR>...</LINEERROR></RESPONSE> with no CREATED tag
    const failed = created === 0 && altered === 0 && errorMsg !== '';
    return { created, altered, exceptions, errorMsg, failed };
}

// ─── stock item auto-create XML ──────────────────────────────────────────────

function buildStockItemCreateXml(companyName, itemName, unit) {
    return `<ENVELOPE>
<HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
<BODY><IMPORTDATA><REQUESTDESC>
<REPORTNAME>All Masters</REPORTNAME>
<STATICVARIABLES><SVCURRENTCOMPANY>${escapeXml(companyName)}</SVCURRENTCOMPANY></STATICVARIABLES>
</REQUESTDESC>
<REQUESTDATA><TALLYMESSAGE xmlns:UDF="TallyUDF">
<STOCKITEM NAME="${escapeXml(itemName)}" ACTION="Create">
<NAME>${escapeXml(itemName)}</NAME>
<PARENT>Primary</PARENT>
<BASEUNITS>${escapeXml(unit || 'Nos')}</BASEUNITS>
</STOCKITEM>
</TALLYMESSAGE></REQUESTDATA></IMPORTDATA></BODY></ENVELOPE>`;
}

// ─── inventory entries XML ────────────────────────────────────────────────────

/**
 * Build ALLINVENTORYENTRIES.LIST XML for each stock line item.
 * @param {Array} stockItems - [{tally_item_name, tally_unit, tally_qty, unit_rate, line_total, purchase_ledger}]
 */
function buildInventoryEntriesXml(stockItems) {
    return stockItems.map(item => {
        const amt = round2(-item.line_total);                    // negative = DR in Tally
        const qtyStr = `${item.tally_qty} ${item.tally_unit}`;
        const rateStr = item.unit_rate > 0 ? `${item.unit_rate}/${item.tally_unit}` : `0/${item.tally_unit}`;
        return `
    <ALLINVENTORYENTRIES.LIST>
        <STOCKITEMNAME>${escapeXml(item.tally_item_name)}</STOCKITEMNAME>
        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
        <RATE>${rateStr}</RATE>
        <AMOUNT>${amt}</AMOUNT>
        <ACTUALQTY>${qtyStr}</ACTUALQTY>
        <BILLEDQTY>${qtyStr}</BILLEDQTY>
        <ACCOUNTINGALLOCATIONS.LIST>
            <LEDGERNAME>${escapeXml(item.purchase_ledger)}</LEDGERNAME>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <AMOUNT>${amt}</AMOUNT>
        </ACCOUNTINGALLOCATIONS.LIST>
    </ALLINVENTORYENTRIES.LIST>`;
    }).join('');
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
        vendor_ledger, ledger_lines,
        inventory_entries = []   // stock items for ISINVOICE=Yes mode
    } = voucherData;
    const hasInventory = inventory_entries.length > 0;

    // Tally sign convention:
    //   DR entries (purchase, gst input)  → ISDEEMEDPOSITIVE=Yes,  AMOUNT = negative
    //   CR entries (vendor/party credit)  → ISDEEMEDPOSITIVE=No,   AMOUNT = positive
    // Our internal convention: positive amounts = DR, negative amounts = CR
    const allLedgers = ledger_lines
        .map(l => {
            const isDR = l.amount >= 0;
            const tallyAmt = round2(-l.amount); // flip: our +DR → Tally negative, our -CR → Tally positive
            const billingAmt = round2(-l.amount);
            return `
    <ALLLEDGERENTRIES.LIST>
        <LEDGERNAME>${escapeXml(l.ledger)}</LEDGERNAME>
        <ISDEEMEDPOSITIVE>${isDR ? 'Yes' : 'No'}</ISDEEMEDPOSITIVE>
        <AMOUNT>${tallyAmt}</AMOUNT>
        ${l.bill_ref ? `<BILLALLOCATIONS.LIST>
            <NAME>${escapeXml(l.bill_ref)}</NAME>
            <BILLTYPE>New Ref</BILLTYPE>
            <AMOUNT>${billingAmt}</AMOUNT>
            ${due_date ? `<DUEDATE>${tallyDate(due_date)}</DUEDATE>` : ''}
        </BILLALLOCATIONS.LIST>` : ''}
    </ALLLEDGERENTRIES.LIST>`;
        }).join('');

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
<ISINVOICE>${hasInventory ? 'Yes' : 'No'}</ISINVOICE>
${is_rcm ? '<ISRCMAPPLICABLE>Yes</ISRCMAPPLICABLE>' : ''}
${allLedgers}
${hasInventory ? buildInventoryEntriesXml(inventory_entries) : ''}
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
            narration,
            due_date,
            is_rcm = false,
            vendor_ledger_name,
            vendor_group,
            create_vendor = false,
            purchase_ledger_overrides = {},     // { "gsttype_rate": "Ledger Name" } — purchase invoices
            expense_ledger_overrides = {},       // { "description": "Ledger Name" } — expense/journal invoices
            stock_item_overrides = {},           // { "invoice_description": "Tally Item Name" }
            unit_overrides = {},                 // { "description": {unit, qty} }
            tds_section_code = null,             // e.g. "194C" — from push modal TDS toggle
            tds_amount_override = 0,             // manually entered TDS amount
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

        // Determine invoice type
        const isExpense = invoice.tally_voucher_type === 'Journal';

        // Build debit ledger lines
        const ledgerLines = [];
        let debitTotal = 0;

        if (isExpense && lineItems.length > 0) {
            // ── Expense / Journal invoice: one DR line per description ──────────
            const expenseMap = await dbAll(db, 'SELECT * FROM tally_expense_ledger_map WHERE company_id = ?', [company_id]);
            const expenseLedgerLookup = (desc) => {
                if (expense_ledger_overrides[desc]) return expense_ledger_overrides[desc];
                const m = expenseMap.find(e => e.description === desc);
                return m ? m.ledger_name : null;
            };

            for (const li of lineItems) {
                const desc = (li.description || li.item_description || li.item_name || '').trim();
                if (!desc) continue;
                const expLedger = expenseLedgerLookup(desc);
                if (!expLedger) {
                    return res.status(400).json({
                        error: `No expense ledger mapped for "${desc}". Please select it in the push modal.`,
                        missing_description: desc
                    });
                }
                ledgerLines.push({ ledger: expLedger, amount: round2(li.line_total || 0) });
                debitTotal += round2(li.line_total || 0);
            }

            // GST lines grouped by rate (same logic for both invoice types)
            for (const [, grp] of Object.entries(rateGroups)) {
                if (isIgst && grp.igst > 0) {
                    ledgerLines.push({ ledger: gstLookup('igst', grp.rate), amount: grp.igst });
                    debitTotal += grp.igst;
                } else {
                    const halfRate = round2(grp.rate / 2);
                    if (grp.cgst > 0) { ledgerLines.push({ ledger: gstLookup('cgst', halfRate), amount: grp.cgst }); debitTotal += grp.cgst; }
                    if (grp.sgst > 0) { ledgerLines.push({ ledger: gstLookup('sgst', halfRate), amount: grp.sgst }); debitTotal += grp.sgst; }
                }
                if (grp.cess > 0) { ledgerLines.push({ ledger: gstLookup('cess', grp.rate), amount: grp.cess }); debitTotal += grp.cess; }
            }

        } else {
            // ── Purchase invoice: one DR line per GST rate group ─────────────────
            for (const [, grp] of Object.entries(rateGroups)) {
                const purchaseLedger = purchaseLookup(gstType, grp.rate);
                if (!purchaseLedger) {
                    return res.status(400).json({
                        error: `No purchase ledger mapped for ${gstType} @ ${grp.rate}%. Please configure it in Tally Setup.`,
                        missing_mapping: { gst_type: gstType, tax_rate: grp.rate }
                    });
                }

                if (is_rcm) {
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

                    if (isIgst && grp.igst > 0) {
                        ledgerLines.push({ ledger: gstLookup('igst', grp.rate), amount: grp.igst });
                        debitTotal += grp.igst;
                    } else {
                        const halfRate = round2(grp.rate / 2);
                        if (grp.cgst > 0) { ledgerLines.push({ ledger: gstLookup('cgst', halfRate), amount: grp.cgst }); debitTotal += grp.cgst; }
                        if (grp.sgst > 0) { ledgerLines.push({ ledger: gstLookup('sgst', halfRate), amount: grp.sgst }); debitTotal += grp.sgst; }
                    }
                    if (grp.cess > 0) { ledgerLines.push({ ledger: gstLookup('cess', grp.rate), amount: grp.cess }); debitTotal += grp.cess; }
                }
            }
        }

        // ── TDS credit line ─────────────────────────────────────────────────────
        // Resolve TDS section → ledger + amount
        let tdsLedger = null, tdsAmount = 0;
        if (tds_section_code) {
            const tdsSection = await dbGet(db, 'SELECT * FROM tally_tds_section_map WHERE company_id = ? AND section_code = ?', [company_id, tds_section_code]);
            if (tdsSection) {
                tdsLedger = tdsSection.ledger_name;
                // Use manually entered amount if provided, otherwise calculate from taxable
                tdsAmount = tds_amount_override > 0
                    ? round2(tds_amount_override)
                    : round2((invoice.taxable_amount || 0) * tdsSection.default_rate / 100);
            }
        } else if (invoice.tds_amount > 0 && cfg && cfg.tds_ledger) {
            // Fallback: invoice had TDS extracted and config has a default TDS ledger
            tdsLedger = cfg.tds_ledger;
            tdsAmount = round2(invoice.tds_amount);
        }
        if (tdsAmount > 0 && tdsLedger) {
            ledgerLines.push({ ledger: tdsLedger, amount: -tdsAmount }); // CR: TDS payable
        }

        // Vendor credit line (negative = credit in Tally)
        const invoiceValue = round2(invoice.invoice_value || 0);
        // After TDS, remaining diff is true rounding
        const diff = round2(debitTotal - tdsAmount - invoiceValue);

        if (Math.abs(diff) > 0.02 && roundoffLedger) {
            ledgerLines.push({ ledger: roundoffLedger, amount: round2(-diff) });
        }

        // Vendor CR — with bill ref for outstanding tracking
        ledgerLines.push({
            ledger: vendorLedger,
            amount: -invoiceValue,
            bill_ref: invoice.invoice_number || String(expenseId)
        });

        // 8. Build inventory entries (when company has records_inventory = 1 and line items exist)
        const recordsInventory = !!(cfg && cfg.records_inventory);
        const enableUnitConversion = !!(cfg && cfg.enable_unit_conversion);
        const customUnitMap = enableUnitConversion
            ? await dbAll(db, 'SELECT * FROM tally_unit_map WHERE company_id = ?', [company_id])
            : [];

        const inventoryEntries = [];
        if (recordsInventory && lineItems.length > 0) {
            for (const li of lineItems) {
                const desc = (li.description || li.item_description || li.item_name || '').trim();
                if (!desc) continue;

                // Tally stock item name: user override → saved mapping → invoice description
                let tallyItemName = (stock_item_overrides && stock_item_overrides[desc]) || null;
                if (!tallyItemName) {
                    const savedMap = await dbGet(db, 'SELECT tally_item_name FROM tally_stock_item_map WHERE company_id = ? AND invoice_description = ?', [company_id, desc]);
                    tallyItemName = savedMap ? savedMap.tally_item_name : desc;
                }

                // Unit resolution
                let tallyUnit, tallyQty;
                if (unit_overrides && unit_overrides[desc]) {
                    // User explicitly overrode unit in push modal
                    tallyUnit = unit_overrides[desc].unit || 'Nos';
                    tallyQty = round2(parseFloat(unit_overrides[desc].qty) || li.quantity || 1);
                } else {
                    const resolved = resolveUnit(li.unit || li.uom || '', customUnitMap, enableUnitConversion);
                    tallyUnit = resolved.tally_unit;
                    tallyQty = round2((parseFloat(li.quantity) || 1) * resolved.factor);
                }

                // Purchase ledger for this line item's rate
                const liRate = li.tax_rate || 0;
                const liPurchaseLedger = purchaseLookup(gstType, liRate) || purchaseLookup(gstType, 0) || 'Purchase';

                inventoryEntries.push({
                    tally_item_name: tallyItemName,
                    tally_unit: tallyUnit,
                    tally_qty: tallyQty,
                    unit_rate: round2(parseFloat(li.unit_rate) || 0),
                    line_total: round2(parseFloat(li.line_total) || 0),
                    purchase_ledger: liPurchaseLedger
                });

                // Auto-create stock item in Tally if it doesn't exist yet (best-effort)
                if (req.body.auto_create_stock_items && tallyItemName) {
                    const createXml = buildStockItemCreateXml(companyName, tallyItemName, tallyUnit);
                    await postToTally(tallyUrl, createXml).catch(() => {}); // ignore "already exists" errors
                }

                // Persist stock item mapping for future pushes
                if (desc && tallyItemName !== desc) {
                    await dbRun(db, `
                        INSERT INTO tally_stock_item_map (company_id, invoice_description, tally_item_name)
                        VALUES (?, ?, ?)
                        ON CONFLICT(company_id, invoice_description) DO UPDATE SET tally_item_name = excluded.tally_item_name
                    `, [company_id, desc, tallyItemName]).catch(() => {});
                }
            }
        }

        // 9. Build and send XML
        const voucherData = {
            voucher_type: invoice.tally_voucher_type || 'Purchase',
            voucher_number: voucherNumber,
            voucher_date: invoice.invoice_date,
            narration: finalNarration,
            due_date,
            is_rcm,
            vendor_ledger: vendorLedger,
            ledger_lines: ledgerLines,
            inventory_entries: inventoryEntries
        };

        const voucherXml = buildVoucherXml(companyName, voucherData);
        console.log('📤 Tally XML sent:\n', voucherXml);
        const tallyResponse = await postToTally(tallyUrl, voucherXml);
        console.log('📥 Tally response:\n', tallyResponse);
        const result = parseTallyResponse(tallyResponse);
        console.log('📊 Parsed result:', result);

        if (result.failed || (result.exceptions > 0 && result.created === 0)) {
            return res.status(400).json({
                error: `Tally rejected the voucher: ${result.errorMsg || 'Check Tally for details'}`,
                tally_response: tallyResponse
            });
        }

        // 10. Mark as pushed
        await dbRun(db,
            'UPDATE expenses SET tally_pushed = 1, tally_pushed_at = CURRENT_TIMESTAMP, tally_voucher_type = ? WHERE id = ?',
            [voucherData.voucher_type, expenseId]
        );

        // 11. Save vendor map for future
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
            altered: result.altered,
            exceptions: result.exceptions,
            tally_message: result.errorMsg || null,
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

        const cfg = await dbGet(db, 'SELECT * FROM tally_config WHERE company_id = ?', [companyId]);
        const recordsInventory = !!(cfg && cfg.records_inventory);
        const enableUnitConversion = !!(cfg && cfg.enable_unit_conversion);
        const customUnitMap = enableUnitConversion
            ? await dbAll(db, 'SELECT * FROM tally_unit_map WHERE company_id = ?', [companyId])
            : [];

        // Invoice type
        const isExpense = invoice.tally_voucher_type === 'Journal';

        // Expense items (per-description ledger needed for Journal invoices)
        const expenseItems = [];
        if (isExpense && lineItems.length > 0) {
            const expenseMap = await dbAll(db, 'SELECT * FROM tally_expense_ledger_map WHERE company_id = ?', [companyId]);
            for (const li of lineItems) {
                const desc = (li.description || li.item_description || li.item_name || '').trim();
                if (!desc) continue;
                const saved = expenseMap.find(e => e.description === desc);
                expenseItems.push({ description: desc, saved_ledger: saved ? saved.ledger_name : null });
            }
        }

        // Stock items (purchase invoices with inventory)
        const stockItemsNeeded = [];
        if (!isExpense && recordsInventory && lineItems.length > 0) {
            const stockMap = await dbAll(db, 'SELECT * FROM tally_stock_item_map WHERE company_id = ?', [companyId]);
            for (const li of lineItems) {
                const desc = (li.description || li.item_description || li.item_name || '').trim();
                if (!desc) continue;
                const saved = stockMap.find(s => s.invoice_description === desc);
                const resolved = resolveUnit(li.unit || li.uom || '', customUnitMap, enableUnitConversion);
                const invQty = round2((parseFloat(li.quantity) || 1) * resolved.factor);
                stockItemsNeeded.push({
                    description: desc,
                    invoice_qty: parseFloat(li.quantity) || 1,
                    invoice_unit: li.unit || li.uom || '',
                    tally_unit: resolved.tally_unit,
                    tally_qty: invQty,
                    conversion_factor: resolved.factor,
                    saved_tally_item: saved ? saved.tally_item_name : null
                });
            }
        }

        // TDS sections configured for this company
        const tdsSections = await dbAll(db, 'SELECT * FROM tally_tds_section_map WHERE company_id = ? ORDER BY section_code', [companyId]);

        res.json({
            invoice_id: invoice.id,
            invoice_number: invoice.invoice_number,
            already_pushed: !!invoice.tally_pushed,
            pushed_at: invoice.tally_pushed_at,
            voucher_type: invoice.tally_voucher_type || 'Purchase',
            is_expense: isExpense,
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
            expense_items: expenseItems,
            rcm_detected: !!(invoice.extracted_text && /reverse.charge.*yes/i.test(invoice.extracted_text)),
            tds_amount: round2(invoice.tds_amount || 0),
            tds_rate: invoice.tds_rate || 0,
            taxable_amount: round2(invoice.taxable_amount || 0),
            tds_sections: tdsSections,
            records_inventory: recordsInventory,
            stock_items_needed: stockItemsNeeded
        });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

module.exports = router;
