const fs = require('fs');

function analyzeBalanceXML() {
  try {
    // Read the debug file that was created
    const xmlData = fs.readFileSync('debug_balance_structure.xml', 'utf8');
    
    console.log('🔍 Analyzing Trial Balance XML Structure...\n');
    
    // 1. Show first 1000 characters
    console.log('📄 First 1000 characters of XML:');
    console.log('='.repeat(50));
    console.log(xmlData.substring(0, 1000));
    console.log('='.repeat(50));
    
    // 2. Find all unique XML tags
    console.log('\n🏷️ All XML tags found:');
    const tagPattern = /<\/?([^>\s]+)[^>]*>/g;
    const tags = new Set();
    let match;
    
    while ((match = tagPattern.exec(xmlData)) !== null) {
      tags.add(match[1]);
    }
    
    Array.from(tags).sort().forEach(tag => {
      console.log(`   <${tag}>`);
    });
    
    // 3. Find all content with numbers
    console.log('\n💰 All XML elements containing numbers:');
    const numberPattern = /<([^>]+)>([^<]*\d[^<]*)<\/([^>]+)>/g;
    let numberMatch;
    let count = 0;
    
    while ((numberMatch = numberPattern.exec(xmlData)) && count < 20) {
      console.log(`   <${numberMatch[1]}>${numberMatch[2]}</${numberMatch[3]}>`);
      count++;
    }
    
    // 4. Look specifically for account name patterns
    console.log('\n👤 Account name patterns:');
    const namePatterns = [
      /<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/g,
      /<NAME>([^<]+)<\/NAME>/g,
      /<LEDGERNAME>([^<]+)<\/LEDGERNAME>/g
    ];
    
    namePatterns.forEach((pattern, index) => {
      console.log(`\nPattern ${index + 1}:`);
      let nameMatch;
      let nameCount = 0;
      while ((nameMatch = pattern.exec(xmlData)) && nameCount < 5) {
        console.log(`   ${nameMatch[1]}`);
        nameCount++;
      }
    });
    
    // 5. Look for amount patterns specifically
    console.log('\n💵 Amount-specific patterns:');
    const amountPatterns = [
      /<DSPCLDRAMTA>([^<]*)<\/DSPCLDRAMTA>/g,
      /<DSPCLCRAMTA>([^<]*)<\/DSPCLCRAMTA>/g,
      /<AMOUNT>([^<]*)<\/AMOUNT>/g,
      /<BALANCE>([^<]*)<\/BALANCE>/g
    ];
    
    amountPatterns.forEach((pattern, index) => {
      console.log(`\nAmount Pattern ${index + 1}:`);
      let amountMatch;
      let amountCount = 0;
      while ((amountMatch = pattern.exec(xmlData)) && amountCount < 5) {
        console.log(`   ${amountMatch[1]}`);
        amountCount++;
      }
    });
    
    // 6. Show the full structure around the first amount
    console.log('\n🎯 Context around first amount found:');
    const firstAmountMatch = xmlData.match(/.*<DSPCLDRAMTA>([^<]*)<\/DSPCLDRAMTA>.*/);
    if (firstAmountMatch) {
      const position = xmlData.indexOf(firstAmountMatch[0]);
      const contextStart = Math.max(0, position - 200);
      const contextEnd = Math.min(xmlData.length, position + 400);
      const context = xmlData.substring(contextStart, contextEnd);
      console.log(context);
    } else {
      console.log('No DSPCLDRAMTA tags found!');
    }
    
    // 7. Manual parsing attempt
    console.log('\n🔧 Manual parsing attempt:');
    parseBalanceManually(xmlData);
    
  } catch (error) {
    console.error('❌ Error reading debug file:', error.message);
    console.log('💡 Make sure you ran the previous script first to create debug_balance_structure.xml');
  }
}

function parseBalanceManually(xmlData) {
  // Split into lines and look for patterns
  const lines = xmlData.split('\n');
  
  let currentAccount = '';
  let foundAccounts = [];
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    // Look for account names
    if (line.includes('<DSPDISPNAME>')) {
      const nameMatch = line.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
      if (nameMatch) {
        currentAccount = nameMatch[1];
        console.log(`📝 Found account: ${currentAccount}`);
        
        // Look for amounts in the next few lines
        for (let j = i + 1; j < Math.min(i + 10, lines.length); j++) {
          const nextLine = lines[j].trim();
          
          const debitMatch = nextLine.match(/<DSPCLDRAMTA>([^<]*)<\/DSPCLDRAMTA>/);
          if (debitMatch) {
            console.log(`   💸 Debit: ${debitMatch[1]}`);
          }
          
          const creditMatch = nextLine.match(/<DSPCLCRAMTA>([^<]*)<\/DSPCLCRAMTA>/);
          if (creditMatch) {
            console.log(`   💰 Credit: ${creditMatch[1]}`);
          }
          
          // If we find amounts, add to found accounts
          if (debitMatch || creditMatch) {
            foundAccounts.push({
              name: currentAccount,
              debit: debitMatch ? debitMatch[1] : '0',
              credit: creditMatch ? creditMatch[1] : '0'
            });
            break;
          }
        }
      }
    }
  }
  
  console.log(`\n✅ Successfully parsed ${foundAccounts.length} accounts with amounts:`);
  foundAccounts.forEach((account, index) => {
    console.log(`${index + 1}. ${account.name}: Dr=${account.debit}, Cr=${account.credit}`);
  });
  
  return foundAccounts;
}

// Run the analysis
console.log('🎬 Starting detailed XML analysis...');
analyzeBalanceXML();