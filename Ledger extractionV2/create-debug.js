const axios = require('axios');
const fs = require('fs');

const TALLY_URL = 'http://localhost:9000';

const xmlRequest = `<ENVELOPE>
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

async function createDebugFile() {
  try {
    console.log('🚀 Fetching Trial Balance XML...');
    
    const response = await axios.post(TALLY_URL, xmlRequest.trim(), {
      headers: { 'Content-Type': 'text/xml' },
      timeout: 15000,
    });

    console.log('✅ Response received!');
    console.log('📏 Size:', response.data.length, 'characters');
    
    // Save the XML file
    fs.writeFileSync('debug_balance_structure.xml', response.data);
    console.log('💾 Saved debug_balance_structure.xml');
    
    // Immediately analyze it
    console.log('\n🔍 Analyzing the XML structure...\n');
    
    const xmlData = response.data;
    
    // Show first 1000 characters
    console.log('📄 First 1000 characters:');
    console.log('='.repeat(60));
    console.log(xmlData.substring(0, 1000));
    console.log('='.repeat(60));
    
    // Find account names and amounts
    console.log('\n👤 Looking for account names:');
    const accountPattern = /<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/g;
    let accountMatch;
    let accountCount = 0;
    
    while ((accountMatch = accountPattern.exec(xmlData)) && accountCount < 10) {
      console.log(`   ${accountCount + 1}. ${accountMatch[1]}`);
      accountCount++;
    }
    
    console.log('\n💰 Looking for debit amounts:');
    const debitPattern = /<DSPCLDRAMTA>([^<]*)<\/DSPCLDRAMTA>/g;
    let debitMatch;
    let debitCount = 0;
    
    while ((debitMatch = debitPattern.exec(xmlData)) && debitCount < 10) {
      console.log(`   ${debitCount + 1}. ${debitMatch[1]}`);
      debitCount++;
    }
    
    console.log('\n💳 Looking for credit amounts:');
    const creditPattern = /<DSPCLCRAMTA>([^<]*)<\/DSPCLCRAMTA>/g;
    let creditMatch;
    let creditCount = 0;
    
    while ((creditMatch = creditPattern.exec(xmlData)) && creditCount < 10) {
      console.log(`   ${creditCount + 1}. ${creditMatch[1]}`);
      creditCount++;
    }
    
    // Try to match accounts with amounts
    console.log('\n🔗 Trying to match accounts with amounts:');
    const balances = extractBalancesStep(xmlData);
    
    if (balances.length > 0) {
      console.log('✅ Successfully extracted balances:');
      balances.forEach((balance, index) => {
        console.log(`${index + 1}. ${balance.name}: Dr=${balance.debit}, Cr=${balance.credit}`);
      });
    } else {
      console.log('❌ Could not extract balances');
      console.log('\n🔍 Let me try a different approach...');
      
      // Show the raw structure with line numbers
      const lines = xmlData.split('\n');
      console.log('\n📝 Raw XML lines (first 30):');
      lines.slice(0, 30).forEach((line, index) => {
        if (line.trim()) {
          console.log(`${(index + 1).toString().padStart(3)}: ${line.trim()}`);
        }
      });
    }
    
  } catch (error) {
    console.error('❌ Error:', error.message);
  }
}

function extractBalancesStep(xmlData) {
  const balances = [];
  
  // Use your debug output structure exactly
  const lines = xmlData.split('\n');
  let currentAccount = '';
  let currentDebit = '0.00';
  let currentCredit = '0.00';
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    // Account name
    const accountMatch = line.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
    if (accountMatch) {
      // Save previous account if it had amounts
      if (currentAccount && (currentDebit !== '0.00' || currentCredit !== '0.00')) {
        balances.push({
          name: currentAccount,
          debit: currentDebit,
          credit: currentCredit
        });
      }
      
      // Start new account
      currentAccount = accountMatch[1];
      currentDebit = '0.00';
      currentCredit = '0.00';
    }
    
    // Debit amount
    const debitMatch = line.match(/<DSPCLDRAMTA>([^<]*)<\/DSPCLDRAMTA>/);
    if (debitMatch) {
      currentDebit = debitMatch[1] || '0.00';
    }
    
    // Credit amount
    const creditMatch = line.match(/<DSPCLCRAMTA>([^<]*)<\/DSPCLCRAMTA>/);
    if (creditMatch) {
      currentCredit = creditMatch[1] || '0.00';
    }
  }
  
  // Don't forget the last account
  if (currentAccount && (currentDebit !== '0.00' || currentCredit !== '0.00')) {
    balances.push({
      name: currentAccount,
      debit: currentDebit,
      credit: currentCredit
    });
  }
  
  return balances;
}

// Run it
console.log('🎬 Creating debug file and analyzing...');
createDebugFile();