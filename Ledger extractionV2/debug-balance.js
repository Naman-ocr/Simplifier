const axios = require('axios');
const fs = require('fs');
const path = require('path');

const TALLY_URL = 'http://localhost:9000';

// Let's try different XML requests to find the right one for balances
const xmlRequests = [
  {
    name: 'Trial Balance',
    xml: `<ENVELOPE>
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
</ENVELOPE>`
  },
  {
    name: 'Balance Sheet',
    xml: `<ENVELOPE>
<HEADER>
<TALLYREQUEST>Export Data</TALLYREQUEST>
</HEADER>
<BODY>
<EXPORTDATA>
<REQUESTDESC>
<REPORTNAME>Balance Sheet</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<SVFROMDATE>##DTYEARFROM##</SVFROMDATE>
<SVTODATE>##DTYEARTO##</SVTODATE>
</STATICVARIABLES>
</REQUESTDESC>
</EXPORTDATA>
</BODY>
</ENVELOPE>`
  },
  {
    name: 'Ledger Monthly Summary',
    xml: `<ENVELOPE>
<HEADER>
<TALLYREQUEST>Export Data</TALLYREQUEST>
</HEADER>
<BODY>
<EXPORTDATA>
<REQUESTDESC>
<REPORTNAME>Ledger Monthly Summary</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<SVFROMDATE>##DTYEARFROM##</SVFROMDATE>
<SVTODATE>##DTYEARTO##</SVTODATE>
</STATICVARIABLES>
</REQUESTDESC>
</EXPORTDATA>
</BODY>
</ENVELOPE>`
  }
];

async function debugBalanceStructure() {
  console.log('🔍 Debugging balance data structure...');
  
  for (let i = 0; i < xmlRequests.length; i++) {
    const request = xmlRequests[i];
    console.log(`\n📝 Testing: ${request.name}`);
    
    try {
      const response = await axios.post(TALLY_URL, request.xml.trim(), {
        headers: { 'Content-Type': 'text/xml' },
        timeout: 15000,
      });

      console.log('✅ Response received');
      console.log('📏 Size:', response.data.length, 'characters');
      
      // Save the response
      const fileName = `debug_${request.name.replace(/[^a-zA-Z0-9]/g, '_')}.xml`;
      fs.writeFileSync(fileName, response.data);
      console.log('💾 Saved to:', fileName);
      
      // Show structure preview
      const preview = response.data.substring(0, 1000);
      console.log('\n👀 Structure preview:');
      console.log('---START---');
      console.log(preview);
      console.log('---END---');
      
      // Look for balance-related tags
      const balanceTags = findBalanceTags(response.data);
      if (balanceTags.length > 0) {
        console.log('\n💰 Found potential balance tags:');
        balanceTags.forEach(tag => console.log(`   - ${tag}`));
      }
      
      // Look for amount patterns
      const amountPatterns = findAmountPatterns(response.data);
      if (amountPatterns.length > 0) {
        console.log('\n💵 Found amount patterns:');
        amountPatterns.slice(0, 5).forEach(pattern => console.log(`   - ${pattern}`));
      }
      
    } catch (error) {
      console.error('❌ Error:', error.message);
    }
  }
}

function findBalanceTags(xmlData) {
  const balanceTagPatterns = [
    /<[^>]*balance[^>]*>/gi,
    /<[^>]*amount[^>]*>/gi,
    /<[^>]*closing[^>]*>/gi,
    /<[^>]*opening[^>]*>/gi,
    /<[^>]*debit[^>]*>/gi,
    /<[^>]*credit[^>]*>/gi,
    /<[^>]*dr[^>]*>/gi,
    /<[^>]*cr[^>]*>/gi
  ];
  
  const foundTags = new Set();
  
  balanceTagPatterns.forEach(pattern => {
    let match;
    while ((match = pattern.exec(xmlData)) !== null) {
      foundTags.add(match[0]);
    }
  });
  
  return Array.from(foundTags).slice(0, 20); // Limit to first 20
}

function findAmountPatterns(xmlData) {
  // Look for patterns that might contain amounts
  const amountPattern = /<([^>]+)>([^<]*(?:\d+\.?\d*)[^<]*)<\/\1>/gi;
  const foundPatterns = [];
  let match;
  
  while ((match = amountPattern.exec(xmlData)) !== null && foundPatterns.length < 10) {
    if (match[2].trim() && /\d/.test(match[2])) {
      foundPatterns.push(`<${match[1]}>${match[2]}</${match[1]}>`);
    }
  }
  
  return foundPatterns;
}

// Also create a more specific ledger balance request
async function trySpecificLedgerBalance() {
  console.log('\n🎯 Trying specific ledger balance request...');
  
  const specificXML = `<ENVELOPE>
<HEADER>
<TALLYREQUEST>Export Data</TALLYREQUEST>
</HEADER>
<BODY>
<EXPORTDATA>
<REQUESTDESC>
<REPORTNAME>Ledger</REPORTNAME>
<STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<SVFROMDATE>##DTYEARFROM##</SVFROMDATE>
<SVTODATE>##DTYEARTO##</SVTODATE>
<LEDGERNAME>Cash</LEDGERNAME>
</STATICVARIABLES>
</REQUESTDESC>
</EXPORTDATA>
</BODY>
</ENVELOPE>`;

  try {
    const response = await axios.post(TALLY_URL, specificXML, {
      headers: { 'Content-Type': 'text/xml' },
      timeout: 10000,
    });
    
    console.log('✅ Specific ledger response received');
    console.log('📏 Size:', response.data.length);
    
    fs.writeFileSync('debug_specific_ledger.xml', response.data);
    console.log('💾 Saved to: debug_specific_ledger.xml');
    
    const preview = response.data.substring(0, 800);
    console.log('\n👀 Specific ledger preview:');
    console.log(preview);
    
  } catch (error) {
    console.log('❌ Specific ledger request failed:', error.message);
  }
}

// Execute debugging
console.log('🎬 Starting balance structure debugging...');
debugBalanceStructure()
  .then(() => trySpecificLedgerBalance())
  .then(() => {
    console.log('\n🎉 Debugging complete!');
    console.log('📁 Check the generated debug_*.xml files to see the structure');
    console.log('💡 Look for tags containing balance/amount information');
    console.log('📧 Share the structure and I\'ll fix the parsing logic');
  });