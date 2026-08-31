/**
 * Unit tests for contact-resolver.js. No Azure credentials or network needed.
 * Run: node test-contact-resolver.js
 */
const { resolveRecipientMatch, buildDisambiguationResult } = require('./contact-resolver');

let pass = 0;
let fail = 0;

function check(label, actual, expected) {
  const ok = actual === expected;
  console.log((ok ? '  PASS  ' : '  FAIL  ') + label + (ok ? '' : `\n          expected ${expected}, got ${actual}`));
  ok ? pass++ : fail++;
}

const amanRaj = { name: 'Aman Raj', email: 'amanr@hoshodigital.com', source: 'people_api' };
const amanGupta = { name: 'Aman Gupta', email: 'amang@hoshodigital.com', source: 'people_api' };
const sarah = { name: 'Sarah Wihbow', email: 'sarah@hoshodigital.com', source: 'personal_contacts' };
const dupName = { name: 'Aman Raj', email: 'aman.raj@contractor.com', source: 'org' };

console.log('\n--- exact: proceed without asking ---');
check('full name matches one candidate', resolveRecipientMatch('Aman Raj', [amanRaj]).status, 'exact');
check('full name, case/space insensitive', resolveRecipientMatch('  aman   raj ', [amanRaj]).status, 'exact');
check('full name wins over a second person', resolveRecipientMatch('Aman Raj', [amanRaj, amanGupta]).status, 'exact');
check('email local-part matches', resolveRecipientMatch('amanr', [amanRaj]).status, 'exact');
check('user supplied an email address', resolveRecipientMatch('amanr@hoshodigital.com', [amanRaj, amanGupta]).status, 'exact');
check('name with a trailing period', resolveRecipientMatch('Aman Raj.', [amanRaj]).status, 'exact');
check('picks the right person of two', resolveRecipientMatch('Aman Gupta', [amanRaj, amanGupta]).match.email, 'amang@hoshodigital.com');
check('email input picks the matching address', resolveRecipientMatch('amang@hoshodigital.com', [amanRaj, amanGupta]).match.email, 'amang@hoshodigital.com');

console.log('\n--- ambiguous: must ask ---');
check('partial first name, one hit', resolveRecipientMatch('aman', [amanRaj]).status, 'ambiguous');
check('partial first name, two hits', resolveRecipientMatch('aman', [amanRaj, amanGupta]).status, 'ambiguous');
check('same display name, two addresses', resolveRecipientMatch('Aman Raj', [amanRaj, dupName]).status, 'ambiguous');
check('no candidates at all', resolveRecipientMatch('nobody', []).status, 'ambiguous');
check('candidates without an email are dropped', resolveRecipientMatch('x', [{ name: 'No Email' }]).status, 'ambiguous');
check('duplicate-name case narrows to those two', resolveRecipientMatch('Aman Raj', [amanRaj, dupName]).candidates.length, 2);

console.log('\n--- the question the user actually sees ---');
const one = buildDisambiguationResult('send the email', 'aman', [amanRaj]);
check('single near-match is not a success', one.success, false);
check('single near-match flags disambiguation', one.needsDisambiguation, true);
console.log('\n' + one.message + '\n');

const many = buildDisambiguationResult('send the email', 'aman', [amanRaj, amanGupta]);
check('multi-match lists both candidates', many.candidates.length, 2);
console.log(many.message + '\n');

const none = buildDisambiguationResult('send the message', 'zzz', []);
console.log(none.message + '\n');

console.log('--- caps the list ---');
const lots = Array.from({ length: 9 }, (_, i) => ({ name: 'Aman ' + i, email: 'a' + i + '@x.com' }));
check('shows at most 5', buildDisambiguationResult('send the email', 'aman', lots).candidates.length, 5);
check('but reports the true total in the text', buildDisambiguationResult('send the email', 'aman', lots).message.includes('There are 9 people'), true);

console.log(`\n${pass} passed, ${fail} failed\n`);
process.exitCode = fail ? 1 : 0;
