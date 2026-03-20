require('dotenv').config();
const graphTools = require('./graph-tools');

async function testResolution() {
  const nameToSearch = process.argv[2] || 'Aman';
  console.log(`\n🔍 Testing searchContactEmail parameter: "${nameToSearch}"\n`);
  
  try {
    const result = await graphTools.searchContactEmail(nameToSearch);
    if (result.found) {
      console.log('✅ Search SUCCESS');
      console.log('Selected Email:', result.results[0].email);
      console.log('Source:', result.results[0].source);
      console.log('\nFull Result Array:');
      console.log(JSON.stringify(result.results, null, 2));
    } else {
      console.log('❌ Search Failed or No Match');
      console.log(result.message);
    }
  } catch (err) {
    console.error('Error in test:', err.message);
  }
  
  process.exit(0);
}

testResolution();
