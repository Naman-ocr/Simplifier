const axios = require('axios');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const TALLY_URL = 'http://localhost:9000';

// All your existing XML requests remain the same
const xmlRequestEnhancedLedgers = `<ENVELOPE>
<HEADER>
<TALLYREQUEST>Export Data</TALLYREQUEST>
</HEADER>
<BODY>
<EXPORTDATA>
<REQUESTDESC>
<REPORTNAME>List of Accounts</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<ACCOUNTTYPE>Ledgers</ACCOUNTTYPE>
</STATICVARIABLES>
</REQUESTDESC>
</EXPORTDATA>
</BODY>
</ENVELOPE>`;

const xmlRequestAccountGroups = `<ENVELOPE>
<HEADER>
<TALLYREQUEST>Export Data</TALLYREQUEST>
</HEADER>
<BODY>
<EXPORTDATA>
<REQUESTDESC>
<REPORTNAME>List of Accounts</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<ACCOUNTTYPE>Groups</ACCOUNTTYPE>
</STATICVARIABLES>
</REQUESTDESC>
</EXPORTDATA>
</BODY>
</ENVELOPE>`;

const xmlRequestStockItems = `<ENVELOPE>
<HEADER>
<TALLYREQUEST>Export Data</TALLYREQUEST>
</HEADER>
<BODY>
<EXPORTDATA>
<REQUESTDESC>
<REPORTNAME>List of Accounts</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<ACCOUNTTYPE>Stock Items</ACCOUNTTYPE>
</STATICVARIABLES>
</REQUESTDESC>
</EXPORTDATA>
</BODY>
</ENVELOPE>`;

const xmlRequestStockGroups = `<ENVELOPE>
<HEADER>
<TALLYREQUEST>Export Data</TALLYREQUEST>
</HEADER>
<BODY>
<EXPORTDATA>
<REQUESTDESC>
<REPORTNAME>List of Accounts</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<ACCOUNTTYPE>Stock Groups</ACCOUNTTYPE>
</STATICVARIABLES>
</REQUESTDESC>
</EXPORTDATA>
</BODY>
</ENVELOPE>`;

const xmlRequestUnits = `<ENVELOPE>
<HEADER>
<TALLYREQUEST>Export Data</TALLYREQUEST>
</HEADER>
<BODY>
<EXPORTDATA>
<REQUESTDESC>
<REPORTNAME>List of Accounts</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<ACCOUNTTYPE>Units</ACCOUNTTYPE>
</STATICVARIABLES>
</REQUESTDESC>
</EXPORTDATA>
</BODY>
</ENVELOPE>`;

const xmlRequestTrialBalance = `<ENVELOPE>
<HEADER>
<TALLYREQUEST>Export Data</TALLYREQUEST>
</HEADER>
<BODY>
<EXPORTDATA>
<REQUESTDESC>
<REPORTNAME>Trial Balance</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<SVFROMDATE>##DTYEARFROM##</SVFROMDATE>
<SVTODATE>##DTYEARTO##</SVTODATE>
</STATICVARIABLES>
</REQUESTDESC>
</EXPORTDATA>
</BODY>
</ENVELOPE>`;

async function extractCompleteData() {
  try {
    console.log('🚀 Starting comprehensive Tally data extraction...');
    
    const requests = [
      { name: 'Ledgers', xml: xmlRequestEnhancedLedgers },
      { name: 'Account Groups', xml: xmlRequestAccountGroups },
      { name: 'Stock Items', xml: xmlRequestStockItems },
      { name: 'Stock Groups', xml: xmlRequestStockGroups },
      { name: 'Units', xml: xmlRequestUnits },
      { name: 'Trial Balance', xml: xmlRequestTrialBalance }
    ];

    const responses = {};
    
    console.log('📊 Fetching data from Tally...');
    
    for (const request of requests) {
      try {
        console.log(`🔄 Fetching ${request.name}...`);
        const response = await axios.post(TALLY_URL, request.xml.trim(), {
          headers: { 'Content-Type': 'text/xml' },
          timeout: 30000,
        });
        
        responses[request.name] = response.data;
        console.log(`✅ ${request.name} fetched successfully`);
        
        // Add small delay between requests
        await new Promise(resolve => setTimeout(resolve, 500));
        
      } catch (error) {
        console.log(`❌ Failed to fetch ${request.name}: ${error.message}`);
        responses[request.name] = null;
      }
    }
    
    console.log('✅ All data fetch attempts completed!');
    
    // Process the extracted data
    const processedData = processAllData(responses);
    
    // Create single Excel file with multiple sheets
    await createSingleExcelFile(processedData);
    
    // Display summary
    displaySummary(processedData);
    
  } catch (error) {
    console.error('❌ Error:', error.message);
  }
}

function processAllData(responses) {
  const result = {
    ledgers: [],
    accountGroups: [],
    stockItems: [],
    stockGroups: [],
    units: [],
    accountBalances: []
  };

  // Process Ledgers
  if (responses['Ledgers']) {
    result.ledgers = extractLedgers(responses['Ledgers']);
    console.log(`📋 Processed ${result.ledgers.length} ledgers`);
  }

  // Process Account Groups
  if (responses['Account Groups']) {
    result.accountGroups = extractAccountGroups(responses['Account Groups']);
    console.log(`🏛️ Processed ${result.accountGroups.length} account groups`);
  }

  // Process Stock Items
  if (responses['Stock Items']) {
    result.stockItems = extractStockItems(responses['Stock Items']);
    console.log(`📦 Processed ${result.stockItems.length} stock items`);
  }

  // Process Stock Groups
  if (responses['Stock Groups']) {
    result.stockGroups = extractStockGroups(responses['Stock Groups']);
    console.log(`📊 Processed ${result.stockGroups.length} stock groups`);
  }

  // Process Units
  if (responses['Units']) {
    result.units = extractUnits(responses['Units']);
    console.log(`📏 Processed ${result.units.length} units`);
  }

  // Process Trial Balance
  if (responses['Trial Balance']) {
    result.accountBalances = extractAccountBalances(responses['Trial Balance']);
    console.log(`💰 Processed ${result.accountBalances.length} account balances`);
  }

  return result;
}

function extractLedgers(xmlData) {
  const ledgers = [];
  const ledgerBlockPattern = /<TALLYMESSAGE[^>]*>[\s\S]*?<LEDGER[^>]*>([\s\S]*?)<\/LEDGER>[\s\S]*?<\/TALLYMESSAGE>/gi;
  let blockMatch;
  
  while ((blockMatch = ledgerBlockPattern.exec(xmlData)) !== null) {
    const ledgerBlock = blockMatch[1];
    
    const nameMatch = ledgerBlock.match(/<NAME>([^<]+)<\/NAME>/);
    if (!nameMatch) continue;
    
    const ledgerName = decodeXmlEntities(nameMatch[1]);
    
    // Extract all the details
    const parentMatch = ledgerBlock.match(/<PARENT>([^<]*)<\/PARENT>/);
    const gstnPatterns = [
      /<GSTIN>([^<]*)<\/GSTIN>/i,
      /<GSTREGISTRATIONNUMBER>([^<]*)<\/GSTREGISTRATIONNUMBER>/i,
      /<PARTYGSTIN>([^<]*)<\/PARTYGSTIN>/i
    ];
    
    let gstn = '';
    for (const pattern of gstnPatterns) {
      const gstnMatch = ledgerBlock.match(pattern);
      if (gstnMatch && gstnMatch[1].trim()) {
        gstn = decodeXmlEntities(gstnMatch[1].trim());
        break;
      }
    }
    
    const msmeRegMatch = ledgerBlock.match(/<MSMEREGNUMBER>([^<]*)<\/MSMEREGNUMBER>/);
    const msmeStatusMatch = ledgerBlock.match(/<MSMETYPEID>([^<]*)<\/MSMETYPEID>/);
    const mobileMatch = ledgerBlock.match(/<LEDGERMOBILE>([^<]*)<\/LEDGERMOBILE>/);
    const emailMatch = ledgerBlock.match(/<EMAIL>([^<]*)<\/EMAIL>/);
    const panMatch = ledgerBlock.match(/<INCOMETAXNUMBER>([^<]*)<\/INCOMETAXNUMBER>/);
    
    // Extract address
    const addressMatch = ledgerBlock.match(/<ADDRESS\.LIST[^>]*>([\s\S]*?)<\/ADDRESS\.LIST>/);
    let address = '';
    if (addressMatch) {
      const addressLines = addressMatch[1].match(/<ADDRESS>([^<]*)<\/ADDRESS>/g);
      if (addressLines) {
        address = addressLines.map(line => {
          const match = line.match(/<ADDRESS>([^<]*)<\/ADDRESS>/);
          return match ? decodeXmlEntities(match[1]) : '';
        }).filter(line => line.trim()).join(', ');
      }
    }
    
    ledgers.push({
      'Ledger Name': ledgerName,
      'Aliases': '',      
      'Parent Group': parentMatch ? decodeXmlEntities(parentMatch[1]) : '',
      'GSTN': gstn,
      'MSME Registration': msmeRegMatch ? decodeXmlEntities(msmeRegMatch[1]) : '',
      'MSME Status': msmeStatusMatch ? decodeXmlEntities(msmeStatusMatch[1]) : '',
      'Mobile': mobileMatch ? decodeXmlEntities(mobileMatch[1]) : '',
      'Email': emailMatch ? decodeXmlEntities(emailMatch[1]) : '',
      'PAN': panMatch ? decodeXmlEntities(panMatch[1]) : '',
      'Address': address
    });
  }
  
  return ledgers;
}

function extractAccountGroups(xmlData) {
  const groups = [];
  const groupBlockPattern = /<TALLYMESSAGE[^>]*>[\s\S]*?<GROUP[^>]*>([\s\S]*?)<\/GROUP>[\s\S]*?<\/TALLYMESSAGE>/gi;
  let blockMatch;
  
  while ((blockMatch = groupBlockPattern.exec(xmlData)) !== null) {
    const groupBlock = blockMatch[1];
    
    const nameMatch = groupBlock.match(/<NAME>([^<]+)<\/NAME>/);
    if (!nameMatch) continue;
    
    const groupName = decodeXmlEntities(nameMatch[1]);
    const parentMatch = groupBlock.match(/<PARENT>([^<]*)<\/PARENT>/);
    const natureMatch = groupBlock.match(/<NATURE>([^<]*)<\/NATURE>/);
    
    groups.push({
      'Group Name': groupName,
      'Parent Group': parentMatch ? decodeXmlEntities(parentMatch[1]) : '',
      'Nature': natureMatch ? decodeXmlEntities(natureMatch[1]) : ''
    });
  }
  
  return groups;
}

function extractStockItems(xmlData) {
  const stockItems = [];
  const stockBlockPattern = /<TALLYMESSAGE[^>]*>[\s\S]*?<STOCKITEM[^>]*>([\s\S]*?)<\/STOCKITEM>[\s\S]*?<\/TALLYMESSAGE>/gi;
  let blockMatch;
  
  while ((blockMatch = stockBlockPattern.exec(xmlData)) !== null) {
    const stockBlock = blockMatch[1];
    
    const nameMatch = stockBlock.match(/<NAME>([^<]+)<\/NAME>/);
    if (!nameMatch) continue;
    
    const stockName = decodeXmlEntities(nameMatch[1]);
    const parentMatch = stockBlock.match(/<PARENT>([^<]*)<\/PARENT>/);
    const baseUnitsMatch = stockBlock.match(/<BASEUNITS>([^<]*)<\/BASEUNITS>/);
    const additionalUnitsMatch = stockBlock.match(/<ADDITIONALUNITS>([^<]*)<\/ADDITIONALUNITS>/);
    
    // Extract HSN code
    let hsnCode = '';
    const gstDetailsMatch = stockBlock.match(/<GSTDETAILS\.LIST[^>]*>([\s\S]*?)<\/GSTDETAILS\.LIST>/);
    if (gstDetailsMatch) {
      const gstBlock = gstDetailsMatch[1];
      const hsnMatch = gstBlock.match(/<HSN>([^<]*)<\/HSN>/);
      if (hsnMatch) {
        hsnCode = decodeXmlEntities(hsnMatch[1]);
      }
    }
    
    const partNumberMatch = stockBlock.match(/<PARTNUMBER>([^<]*)<\/PARTNUMBER>/);
    const descriptionMatch = stockBlock.match(/<DESCRIPTION>([^<]*)<\/DESCRIPTION>/);
    
    stockItems.push({
      'Stock Item Name': stockName,
      'Parent Stock Group': parentMatch ? decodeXmlEntities(parentMatch[1]) : '',
      'Base Units': baseUnitsMatch ? decodeXmlEntities(baseUnitsMatch[1]) : '',
      'Additional Units': additionalUnitsMatch ? decodeXmlEntities(additionalUnitsMatch[1]) : '',
      'HSN/SAC Code': hsnCode,
      'Part Number': partNumberMatch ? decodeXmlEntities(partNumberMatch[1]) : '',
      'Description': descriptionMatch ? decodeXmlEntities(descriptionMatch[1]) : ''
    });
  }
  
  return stockItems;
}

function extractStockGroups(xmlData) {
  const stockGroups = [];
  const groupBlockPattern = /<TALLYMESSAGE[^>]*>[\s\S]*?<STOCKGROUP[^>]*>([\s\S]*?)<\/STOCKGROUP>[\s\S]*?<\/TALLYMESSAGE>/gi;
  let blockMatch;
  
  while ((blockMatch = groupBlockPattern.exec(xmlData)) !== null) {
    const groupBlock = blockMatch[1];
    
    const nameMatch = groupBlock.match(/<NAME>([^<]+)<\/NAME>/);
    if (!nameMatch) continue;
    
    const groupName = decodeXmlEntities(nameMatch[1]);
    const parentMatch = groupBlock.match(/<PARENT>([^<]*)<\/PARENT>/);
    
    // Extract HSN code for stock group
    let hsnCode = '';
    const gstDetailsMatch = groupBlock.match(/<GSTDETAILS\.LIST[^>]*>([\s\S]*?)<\/GSTDETAILS\.LIST>/);
    if (gstDetailsMatch) {
      const gstBlock = gstDetailsMatch[1];
      const hsnMatch = gstBlock.match(/<HSN>([^<]*)<\/HSN>/);
      if (hsnMatch) {
        hsnCode = decodeXmlEntities(hsnMatch[1]);
      }
    }
    
    stockGroups.push({
      'Stock Group Name': groupName,
      'Parent Stock Group': parentMatch ? decodeXmlEntities(parentMatch[1]) : '',
      'HSN/SAC Code': hsnCode
    });
  }
  
  return stockGroups;
}

function extractUnits(xmlData) {
  const units = [];
  const unitBlockPattern = /<TALLYMESSAGE[^>]*>[\s\S]*?<UNIT[^>]*>([\s\S]*?)<\/UNIT>[\s\S]*?<\/TALLYMESSAGE>/gi;
  let blockMatch;
  
  while ((blockMatch = unitBlockPattern.exec(xmlData)) !== null) {
    const unitBlock = blockMatch[1];
    
    const nameMatch = unitBlock.match(/<NAME>([^<]+)<\/NAME>/);
    if (!nameMatch) continue;
    
    const unitName = decodeXmlEntities(nameMatch[1]);
    const symbolMatch = unitBlock.match(/<FORMALNAME>([^<]*)<\/FORMALNAME>/);
    const decimalPlacesMatch = unitBlock.match(/<DECIMALPLACES>([^<]*)<\/DECIMALPLACES>/);
    const isSimpleUnitMatch = unitBlock.match(/<ISSIMPLEUNIT>([^<]*)<\/ISSIMPLEUNIT>/);
    
    units.push({
      'Unit Name': unitName,
      'Symbol': symbolMatch ? decodeXmlEntities(symbolMatch[1]) : '',
      'Decimal Places': decimalPlacesMatch ? decodeXmlEntities(decimalPlacesMatch[1]) : '',
      'Is Simple Unit': isSimpleUnitMatch ? decodeXmlEntities(isSimpleUnitMatch[1]) : ''
    });
  }
  
  return units;
}

function extractAccountBalances(xmlData) {
  const balances = [];
  const lines = xmlData.split('\n');
  let currentAccount = '';
  let currentDebit = '0.00';
  let currentCredit = '0.00';
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    const accountMatch = line.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
    if (accountMatch) {
      if (currentAccount && (currentDebit !== '0.00' || currentCredit !== '0.00')) {
        const netBalance = (parseFloat(currentCredit) + parseFloat(currentDebit)).toFixed(2);
        balances.push({
          'Account Name': decodeXmlEntities(currentAccount),
          'Debit Amount': currentDebit,
          'Credit Amount': currentCredit,
          'Net Balance': netBalance
        });
      }
      currentAccount = accountMatch[1];
      currentDebit = '0.00';
      currentCredit = '0.00';
    }
    
    const debitMatch = line.match(/<DSPCLDRAMTA>([^<]*)<\/DSPCLDRAMTA>/);
    if (debitMatch) {
      currentDebit = debitMatch[1] || '0.00';
    }
    
    const creditMatch = line.match(/<DSPCLCRAMTA>([^<]*)<\/DSPCLCRAMTA>/);
    if (creditMatch) {
      currentCredit = creditMatch[1] || '0.00';
    }
  }
  
  if (currentAccount && (currentDebit !== '0.00' || currentCredit !== '0.00')) {
    const netBalance = (parseFloat(currentCredit) + parseFloat(currentDebit)).toFixed(2);
    balances.push({
      'Account Name': decodeXmlEntities(currentAccount),
      'Debit Amount': currentDebit,
      'Credit Amount': currentCredit,
      'Net Balance': netBalance
    });
  }
  
  return balances;
}

async function createSingleExcelFile(data) {
  console.log('📊 Creating single Excel file with multiple sheets...');
  
  // Create a new workbook
  const workbook = XLSX.utils.book_new();
  
  // Define the sheets with their data and custom names
  const sheets = [
    {
      name: '📋 Ledgers',
      data: data.ledgers,
      description: 'All ledger accounts with groups, GSTN, MSME details'
    },
    {
      name: '🏛️ Account Groups',
      data: data.accountGroups,
      description: 'Account group hierarchy and nature'
    },
    {
      name: '📦 Stock Items',
      data: data.stockItems,
      description: 'Stock items with groups, HSN codes, units'
    },
    {
      name: '📊 Stock Groups',
      data: data.stockGroups,
      description: 'Stock group hierarchy with HSN codes'
    },
    {
      name: '📏 Units',
      data: data.units,
      description: 'Units of measurement definitions'
    },
    {
      name: '💰 Account Balances',
      data: data.accountBalances,
      description: 'Account balances from Trial Balance'
    }
  ];
  
  // Add each sheet to the workbook
  sheets.forEach(sheet => {
    if (sheet.data.length > 0) {
      console.log(`📄 Adding sheet: ${sheet.name} (${sheet.data.length} entries)`);
      
      // Create worksheet from JSON data
      const worksheet = XLSX.utils.json_to_sheet(sheet.data);
      
      // Auto-fit column widths
      const colWidths = [];
      if (sheet.data.length > 0) {
        Object.keys(sheet.data[0]).forEach((key, index) => {
          const maxLength = Math.max(
            key.length,
            ...sheet.data.map(row => String(row[key] || '').length)
          );
          colWidths[index] = { wch: Math.min(maxLength + 2, 50) }; // Max width 50
        });
        worksheet['!cols'] = colWidths;
      }
      
      // Add the worksheet to the workbook
      XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
    } else {
      console.log(`⚠️ Skipping empty sheet: ${sheet.name}`);
    }
  });
  
  // Create summary sheet
  const summaryData = [
    { 'Data Type': 'Ledgers', 'Count': data.ledgers.length, 'Description': 'All ledger accounts with groups, GSTN, MSME details' },
    { 'Data Type': 'Account Groups', 'Count': data.accountGroups.length, 'Description': 'Account group hierarchy and nature' },
    { 'Data Type': 'Stock Items', 'Count': data.stockItems.length, 'Description': 'Stock items with groups, HSN codes, units' },
    { 'Data Type': 'Stock Groups', 'Count': data.stockGroups.length, 'Description': 'Stock group hierarchy with HSN codes' },
    { 'Data Type': 'Units', 'Count': data.units.length, 'Description': 'Units of measurement definitions' },
    { 'Data Type': 'Account Balances', 'Count': data.accountBalances.length, 'Description': 'Account balances from Trial Balance' }
  ];
  
  const summarySheet = XLSX.utils.json_to_sheet(summaryData);
  summarySheet['!cols'] = [
    { wch: 20 }, // Data Type
    { wch: 10 }, // Count
    { wch: 60 }  // Description
  ];
  
  // Insert summary sheet as the first sheet
  XLSX.utils.book_append_sheet(workbook, summarySheet, '📊 Summary');
  
  // Generate timestamp for filename
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const filename = `TallyData_Complete_${timestamp}.xlsx`;
  
  // Write the Excel file
  XLSX.writeFile(workbook, filename);
  
  console.log(`✅ Excel file created successfully: ${filename}`);
  console.log(`📁 File location: ${path.resolve(filename)}`);
  
  // Clean up any old debug files (optional)
  try {
    const files = fs.readdirSync('.');
    files.forEach(file => {
      if (file.startsWith('debug_') && file.endsWith('.xml')) {
        fs.unlinkSync(file);
      }
    });
    console.log('🧹 Cleaned up debug XML files');
  } catch (error) {
    // Ignore cleanup errors
  }
  
  return filename;
}

function displaySummary(data) {
  console.log('\n📋 Comprehensive Tally Data Extraction Summary');
  console.log('='.repeat(80));
  
  console.log(`\n📊 Data Overview:`);
  console.log(`   📋 Ledgers: ${data.ledgers.length}`);
  console.log(`   🏛️ Account Groups: ${data.accountGroups.length}`);
  console.log(`   📦 Stock Items: ${data.stockItems.length}`);
  console.log(`   📊 Stock Groups: ${data.stockGroups.length}`);
  console.log(`   📏 Units: ${data.units.length}`);
  console.log(`   💰 Account Balances: ${data.accountBalances.length}`);
  
  // Show statistics
  const ledgersWithMSME = data.ledgers.filter(l => l['MSME Registration']).length;
  const stockItemsWithHSN = data.stockItems.filter(s => s['HSN/SAC Code']).length;
  const ledgersWithGroups = data.ledgers.filter(l => l['Parent Group']).length;
  const stockItemsWithGroups = data.stockItems.filter(s => s['Parent Stock Group']).length;
  const ledgersWithGSTN = data.ledgers.filter(l => l['GSTN']).length;
  
  console.log(`\n📈 Data Quality Statistics:`);
  console.log(`   Ledgers with MSME details: ${ledgersWithMSME}`);
  console.log(`   Ledgers with GSTN: ${ledgersWithGSTN}`);
  console.log(`   Ledgers with parent groups: ${ledgersWithGroups}`);
  console.log(`   Stock items with HSN codes: ${stockItemsWithHSN}`);
  console.log(`   Stock items with parent groups: ${stockItemsWithGroups}`);
  
  // Show samples
  if (data.ledgers.length > 0) {
    console.log(`\n💼 Sample Ledgers:`);
    data.ledgers.slice(0, 3).forEach((item, index) => {
      console.log(`   ${index + 1}. ${item['Ledger Name']} (Group: ${item['Parent Group'] || 'N/A'})`);
    });
  }
  
  if (data.stockItems.length > 0) {
    console.log(`\n📦 Sample Stock Items:`);
    data.stockItems.slice(0, 3).forEach((item, index) => {
      console.log(`   ${index + 1}. ${item['Stock Item Name']} (Group: ${item['Parent Stock Group'] || 'N/A'}) [${item['Base Units'] || 'N/A'}]`);
    });
  }
  
  console.log(`\n🎯 Result: Single Excel file created with ${Object.keys(data).length + 1} sheets (including Summary)`);
}

function decodeXmlEntities(text) {
  if (!text) return text;
  
  return text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#13;&#10;/g, ' ')
    .replace(/&#10;/g, ' ')
    .replace(/&#13;/g, ' ')
    .replace(/&#9;/g, ' ')
    .replace(/&#32;/g, ' ')
    .replace(/&#(\d+);/g, (match, num) => {
      return String.fromCharCode(parseInt(num, 10));
    })
    .replace(/&#x([0-9A-Fa-f]+);/g, (match, hex) => {
      return String.fromCharCode(parseInt(hex, 16));
    })
    .replace(/^\s+|\s+$/g, '');
}

// Execute the extraction
console.log('🎬 Starting enhanced Tally data extraction...');
console.log('🎯 Output: Single Excel file with multiple sheets');
extractCompleteData();