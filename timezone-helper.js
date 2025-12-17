/**
 * ============================================================
 * ⏰ TIMEZONE HELPER MODULE
 * ============================================================
 * 
 * Handles user timezone detection and all date/time formatting
 * - Detects timezone from user's Graph profile
 * - Converts all times to user's timezone
 * - Stores timezone in session for consistency
 * - No hardcoding - timezone detected per user
 * 
 * ============================================================
 */

// 🔄 Lazy-load graphTools to prevent circular dependency
// This is required because graph-tools.js also imports timezone-helper.js
let graphTools;
function getGraphTools() {
  if (!graphTools) {
    graphTools = require('./graph-tools');
  }
  return graphTools;
}

// Map to store user timezone by sessionId
// In production, store this in database
const userTimeZoneStore = new Map();

/**
 * Get or detect user's timezone
 * First checks cache, then fetches from Graph API, then defaults
 * 
 * @param {String} sessionId - User session ID
 * @param {String} userToken - User's access token
 * @returns {Promise<String>} Timezone string (e.g., 'Asia/Singapore')
 */
async function getUserTimeZone(sessionId, userToken) {
  // Check if already cached
  if (userTimeZoneStore.has(sessionId)) {
    console.log(`✓ Timezone from cache: ${userTimeZoneStore.get(sessionId)}`);
    return userTimeZoneStore.get(sessionId);
  }

  try {
    // Try to detect from user's mailbox settings
    const timeZone = await detectTimeZoneFromGraph(userToken);
    
    // Cache for future use
    userTimeZoneStore.set(sessionId, timeZone);
    console.log(`✓ Timezone detected: ${timeZone}`);
    
    return timeZone;
  } catch (err) {
    console.warn(`⚠ Could not detect timezone:`, err.message);
    
    // Fallback to UTC
    const defaultTimeZone = 'UTC';
    userTimeZoneStore.set(sessionId, defaultTimeZone);
    console.log(`✓ Using default timezone: ${defaultTimeZone}`);
    
    return defaultTimeZone;
  }
}

/**
 * Detect timezone from Microsoft Graph mailbox settings
 * @param {String} userToken 
 * @returns {Promise<String>} Timezone
 */
async function detectTimeZoneFromGraph(userToken) {
  if (!userToken) throw new Error('User token required');

  try {
    // Get user's mailbox settings which includes timezone
    const graphTools = getGraphTools();
    const client = await graphTools.getGraphClient(userToken);
    
    const settings = await client
      .api('/me/mailboxSettings')
      .get();

    if (settings && settings.timeZone) {
      console.log(`   📍 User timezone from Graph: ${settings.timeZone}`);
      return settings.timeZone;
    }

    // Fallback: try to get from user profile
    const profile = await client
      .api('/me')
      .select('preferredLanguage')
      .get();

    // If all else fails, return UTC
    return 'UTC';
  } catch (err) {
    console.error('   ❌ Error detecting timezone:', err.message);
    throw new Error('Could not detect timezone: ' + err.message);
  }
}

/**
 * Clear cached timezone for a session (e.g., on logout)
 * @param {String} sessionId 
 */
function clearCachedTimeZone(sessionId) {
  if (userTimeZoneStore.has(sessionId)) {
    userTimeZoneStore.delete(sessionId);
    console.log(`✓ Timezone cleared for session: ${sessionId}`);
  }
}

/**
 * Format a date/time in user's timezone
 * Used throughout the application for consistent display
 * 
 * @param {Date|String} dateTime - Date to format
 * @param {String} userTimeZone - User's timezone (e.g., 'Asia/Singapore')
 * @param {String} format - Format type: 'full' (default), 'date-only', 'time-only'
 * @returns {String} Formatted datetime
 */
function formatForUser(dateTime, userTimeZone = 'UTC', format = 'full') {
  if (!dateTime) return 'N/A';

  try {
    const date = new Date(dateTime);
    
    if (isNaN(date)) {
      return 'Invalid date';
    }

    const options = getFormatOptions(format);

    return date.toLocaleString('en-US', {
      ...options,
      timeZone: userTimeZone
    });
  } catch (err) {
    console.error('Error formatting date:', err);
    return 'Invalid date';
  }
}

/**
 * Get Intl.DateTimeFormat options based on format type
 * @param {String} format 
 * @returns {Object} Options for toLocaleString
 */
function getFormatOptions(format) {
  const baseOptions = {
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    hour12: true
  };

  switch (format) {
    case 'date-only':
      return {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
      };

    case 'time-only':
      return {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: true
      };

    case 'full':
    default:
      return baseOptions;
  }
}

/**
 * Get list of valid timezones
 * Used for timezone selection in UI
 * @returns {Array} Array of timezone strings
 */
function getValidTimeZones() {
  // Common timezones organized by region
  return [
    // Asia
    'Asia/Shanghai',
    'Asia/Hong_Kong',
    'Asia/Singapore',
    'Asia/Bangkok',
    'Asia/Jakarta',
    'Asia/Manila',
    'Asia/Tokyo',
    'Asia/Seoul',
    'Asia/Kolkata',
    'Asia/Dubai',
    'Asia/Bangkok',
    
    // America
    'America/New_York',
    'America/Chicago',
    'America/Denver',
    'America/Los_Angeles',
    'America/Toronto',
    'America/Mexico_City',
    'America/Sao_Paulo',
    'America/Buenos_Aires',
    
    // Europe
    'Europe/London',
    'Europe/Paris',
    'Europe/Berlin',
    'Europe/Amsterdam',
    'Europe/Rome',
    'Europe/Madrid',
    'Europe/Moscow',
    'Europe/Dublin',
    
    // Australia & Pacific
    'Australia/Sydney',
    'Australia/Melbourne',
    'Australia/Brisbane',
    'Pacific/Auckland',
    'Pacific/Fiji',
    
    // Africa
    'Africa/Cairo',
    'Africa/Johannesburg',
    'Africa/Lagos',
    'Africa/Nairobi',
    
    // UTC
    'UTC'
  ];
}

/**
 * Check if a timezone string is valid
 * @param {String} timeZone 
 * @returns {Boolean}
 */
function isValidTimeZone(timeZone) {
  try {
    Intl.DateTimeFormat(undefined, { timeZone });
    return true;
  } catch (ex) {
    return false;
  }
}

/**
 * Get user's current time in their timezone
 * @param {String} userTimeZone 
 * @returns {Object} {date, time, dateTime}
 */
function getUserCurrentTime(userTimeZone = 'UTC') {
  const now = new Date();
  
  try {
    const fullTime = now.toLocaleString('en-US', {
      timeZone: userTimeZone,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    });

    const [datePart, timePart] = fullTime.split(', ');

    return {
      dateTime: `${datePart} ${timePart}`,
      date: datePart,
      time: timePart,
      timezone: userTimeZone
    };
  } catch (err) {
    console.error('Error getting current time:', err);
    return {
      dateTime: now.toISOString(),
      date: now.toDateString(),
      time: now.toTimeString(),
      timezone: 'UTC'
    };
  }
}

/**
 * Convert a UTC time to user's timezone
 * Useful for calendar event conversions
 * 
 * @param {String} utcTime - ISO string
 * @param {String} userTimeZone 
 * @returns {Object} Converted time details
 */
function convertToUserTimeZone(utcTime, userTimeZone = 'UTC') {
  try {
    const date = new Date(utcTime);
    
    const formatted = date.toLocaleString('en-US', {
      timeZone: userTimeZone,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    });

    return {
      original: utcTime,
      converted: formatted,
      timezone: userTimeZone,
      date: date
    };
  } catch (err) {
    console.error('Error converting timezone:', err);
    return {
      original: utcTime,
      converted: utcTime,
      timezone: userTimeZone,
      error: err.message
    };
  }
}

// ============================================================
// EXPORTS
// ============================================================
module.exports = {
  getUserTimeZone,
  detectTimeZoneFromGraph,
  clearCachedTimeZone,
  formatForUser,
  getValidTimeZones,
  isValidTimeZone,
  getUserCurrentTime,
  convertToUserTimeZone,
  userTimeZoneStore // For testing/debugging only
};
