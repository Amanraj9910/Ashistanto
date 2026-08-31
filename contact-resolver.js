/**
 * Recipient disambiguation for actions that contact a person.
 *
 * graph-tools.searchContactEmail() returns every candidate it found, ranked (it prefers
 * @hoshodigital.com addresses). Callers used to just take results[0], so "send it to aman"
 * silently picked whichever Aman sorted first and mailed the wrong person with nothing in the
 * UI to reveal it. This module decides when a match is certain enough to act on, and builds the
 * question to ask when it isn't.
 */

const MAX_CANDIDATES_SHOWN = 5;

function normalizeName(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[.,]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function looksLikeEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

/**
 * Decides whether a contact search result is safe to act on without asking.
 *
 * Certain ("exact"):
 *   - the user typed an email address, or
 *   - exactly one candidate's display name equals the searched name, or
 *   - exactly one candidate's email local-part equals the searched name.
 *
 * Everything else is ambiguous, including a single loose hit such as "aman" -> "Aman Raj".
 * Confirming a near-match costs one reply; sending to the wrong colleague cannot be undone.
 *
 * @param {string} searchedName
 * @param {Array<{name:string,email:string,source?:string}>} results
 * @returns {{status:'exact'|'ambiguous', match?:object, candidates?:object[]}}
 */
function resolveRecipientMatch(searchedName, results) {
  const candidates = (results || []).filter((entry) => entry && entry.email);
  if (candidates.length === 0) return { status: 'ambiguous', candidates: [] };

  const searched = normalizeName(searchedName);

  if (looksLikeEmail(searchedName)) {
    const byEmail = candidates.find((entry) => normalizeName(entry.email) === searched);
    return { status: 'exact', match: byEmail || candidates[0] };
  }

  const exact = candidates.filter((entry) =>
    normalizeName(entry.name) === searched ||
    normalizeName(String(entry.email).split('@')[0]) === searched);

  if (exact.length === 1) return { status: 'exact', match: exact[0] };
  if (exact.length > 1) return { status: 'ambiguous', candidates: exact };

  return { status: 'ambiguous', candidates };
}

/**
 * Builds the tool result that makes the assistant ask the user which person they meant.
 *
 * Shaped like the existing notFound result so it travels the same path: the assistant relays
 * `message`, the user answers with a full name or address, and the tool is called again -- this
 * time resolving to an exact match.
 *
 * @param {string} actionLabel e.g. 'send the email'
 * @param {string} searchedName
 * @param {Array<{name:string,email:string}>} candidates
 */
function buildDisambiguationResult(actionLabel, searchedName, candidates) {
  const shown = (candidates || []).slice(0, MAX_CANDIDATES_SHOWN);
  const list = shown.map((entry, i) => (i + 1) + '. ' + entry.name + ' <' + entry.email + '>').join('\n');

  let message;
  if (shown.length === 0) {
    message = 'I could not confirm who "' + searchedName + '" is. Please give me their full name or email address.';
  } else if (shown.length === 1) {
    message = 'I found one close match for "' + searchedName + '":\n\n' + list +
      '\n\nDid you mean ' + shown[0].name + '? Reply with the full name or email to confirm and I will ' +
      actionLabel + '.';
  } else {
    message = 'There are ' + candidates.length + ' people matching "' + searchedName + '":\n\n' + list +
      '\n\nWhich one did you mean? Reply with the full name or the email address.';
  }

  return {
    success: false,
    needsDisambiguation: true,
    searchedName: searchedName,
    candidates: shown.map((entry) => ({ name: entry.name, email: entry.email })),
    // The assistant must relay this and must not choose on the user's behalf.
    message: message
  };
}

module.exports = {
  resolveRecipientMatch,
  buildDisambiguationResult,
  normalizeName,
  looksLikeEmail
};
