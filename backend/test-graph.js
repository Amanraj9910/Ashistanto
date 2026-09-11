require('dotenv').config();
const graphTools = require('./graph-tools');

console.log('\n╔═══════════════════════════════════════════╗');
console.log('║   Microsoft Graph API Test                ║');
console.log('╚═══════════════════════════════════════════╝\n');

// Test 1: Check Environment Variables
console.log('1️⃣  Checking Environment Variables...');
console.log('  MICROSOFT_CLIENT_ID:', process.env.MICROSOFT_CLIENT_ID ? '✓ Set' : '✗ Not set');
console.log('  MICROSOFT_CLIENT_SECRET:', process.env.MICROSOFT_CLIENT_SECRET ? '✓ Set' : '✗ Not set');
console.log('  MICROSOFT_TENANT_ID:', process.env.MICROSOFT_TENANT_ID ? '✓ Set' : '✗ Not set');
console.log('  MICROSOFT_ACCESS_TOKEN:', process.env.MICROSOFT_ACCESS_TOKEN ? '✓ Set (using manual token)' : '⚠ Not set (will use app auth)');

if (!process.env.MICROSOFT_CLIENT_ID || !process.env.MICROSOFT_CLIENT_SECRET || !process.env.MICROSOFT_TENANT_ID) {
  console.log('\n❌ Microsoft Graph credentials missing!');
  console.log('Please follow MICROSOFT_GRAPH_SETUP.md to configure Azure app registration.\n');
  process.exit(1);
}

// Run all tests
async function runTests() {
  console.log('\n2️⃣  Testing Graph API Connections...\n');
  
  const tests = [
    {
      name: 'Get User Profile',
      func: graphTools.getUserProfile,
      args: []
    },
    {
      name: 'Get Recent Emails',
      func: graphTools.getRecentEmails,
      args: [3]
    },
    {
      name: 'Get Calendar Events',
      func: graphTools.getCalendarEvents,
      args: [7]
    },
    {
      name: 'Get Recent Files',
      func: graphTools.getRecentFiles,
      args: [5]
    },
    {
      name: 'Get Teams',
      func: graphTools.getTeams,
      args: []
    }
  ];
  
  let passedTests = 0;
  let failedTests = 0;
  
  for (const test of tests) {
    try {
      console.log(`📋 Testing: ${test.name}...`);
      const result = await test.func(...test.args);
      console.log(`   ✓ Success!`);
      
      // Show sample of results
      if (Array.isArray(result) && result.length > 0) {
        console.log(`   📊 Retrieved ${result.length} item(s)`);
        console.log(`   Sample:`, JSON.stringify(result[0], null, 2).substring(0, 200) + '...');
      } else if (typeof result === 'object') {
        console.log(`   📊 Result:`, JSON.stringify(result, null, 2).substring(0, 200) + '...');
      }
      console.log('');
      passedTests++;
    } catch (error) {
      console.log(`   ✗ Failed: ${error.message}`);
      console.log('');
      failedTests++;
    }
  }
  
  console.log('╔═══════════════════════════════════════════╗');
  console.log('║   Test Summary                            ║');
  console.log('╚═══════════════════════════════════════════╝');
  console.log(`✓ Passed: ${passedTests}`);
  console.log(`✗ Failed: ${failedTests}`);
  
  if (failedTests > 0) {
    console.log('\n⚠️  Some tests failed. Common issues:');
    console.log('   1. Missing API permissions in Azure Portal');
    console.log('   2. Admin consent not granted');
    console.log('   3. Access token expired (if using manual token)');
    console.log('   4. Wrong authentication flow for your setup');
    console.log('\nRefer to MICROSOFT_GRAPH_SETUP.md for detailed instructions.\n');
  } else {
    console.log('\n🎉 All tests passed! Your Microsoft Graph integration is working!\n');
    console.log('You can now use voice commands like:');
    console.log('   • "Check my recent emails"');
    console.log('   • "What\'s on my calendar today?"');
    console.log('   • "Show my recent files"');
    console.log('   • "What teams am I in?"\n');
  }
}

runTests().catch(error => {
  console.error('\n❌ Test execution error:', error.message);
  process.exit(1);
});