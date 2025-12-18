const { userTokenStore } = require('./auth');

/**
 * Token Refresh Middleware
 * Handles automatic token refresh when access tokens expire
 */

/**
 * Check if token needs refresh (expires in less than 5 minutes)
 * @param {string} sessionId - User session ID
 * @returns {boolean} - True if token needs refresh
 */
function needsRefresh(sessionId) {
    const tokenData = userTokenStore.get(sessionId);
    if (!tokenData || !tokenData.expiresAt) {
        return false;
    }

    const timeUntilExpiry = tokenData.expiresAt - Date.now();
    const REFRESH_THRESHOLD = 5 * 60 * 1000; // 5 minutes

    return timeUntilExpiry < REFRESH_THRESHOLD;
}

/**
 * Refresh access token using refresh token
 * @param {string} sessionId - User session ID
 * @returns {Promise<string>} - New access token
 */
async function refreshTokenIfNeeded(sessionId) {
    const tokenData = userTokenStore.get(sessionId);

    if (!tokenData) {
        throw new Error('Session not found');
    }

    // Check if refresh is needed
    if (!needsRefresh(sessionId)) {
        console.log(`✓ Token still valid for session: ${sessionId}`);
        return tokenData.accessToken;
    }

    console.log(`🔄 Refreshing token for session: ${sessionId}`);

    try {
        // Lazy import to avoid circular dependency
        const { getAccessTokenByRefreshToken } = require('./graph-tools');
        
        // Get new access token using refresh token
        const newTokenResponse = await getAccessTokenByRefreshToken(tokenData.refreshToken);

        // Update token store with new tokens
        const updatedTokenData = {
            accessToken: newTokenResponse.accessToken || newTokenResponse,
            refreshToken: newTokenResponse.refreshToken || tokenData.refreshToken, // Keep old refresh token if new one not provided
            expiresAt: Date.now() + ((newTokenResponse.expiresIn || 3600) * 1000),
            email: tokenData.email
        };

        userTokenStore.set(sessionId, updatedTokenData);

        console.log(`✅ Token refreshed successfully for session: ${sessionId}`);
        return updatedTokenData.accessToken;

    } catch (error) {
        console.error(`❌ Token refresh failed for session ${sessionId}:`, error.message);

        // Remove invalid session
        userTokenStore.delete(sessionId);

        throw new Error('Token refresh failed. Please log in again.');
    }
}

/**
 * Wrapper for Graph API calls that handles token refresh on 401 errors
 * @param {Function} apiCall - Async function that makes the Graph API call
 * @param {string} sessionId - User session ID
 * @param {number} retryCount - Current retry attempt (internal use)
 * @returns {Promise<any>} - Result from the API call
 */
async function handleGraphApiCall(apiCall, sessionId, retryCount = 0) {
    const MAX_RETRIES = 1;

    try {
        // Proactively refresh token if needed
        const tokenData = userTokenStore.get(sessionId);
        if (!tokenData) {
            throw new Error('Session not found. Please log in again.');
        }

        if (needsRefresh(sessionId)) {
            await refreshTokenIfNeeded(sessionId);
            // Get updated token
            const updatedTokenData = userTokenStore.get(sessionId);
            // Execute API call with fresh token
            return await apiCall(updatedTokenData.accessToken);
        }

        // Execute API call with current token
        return await apiCall(tokenData.accessToken);

    } catch (error) {
        // Check if it's a 401 error (token expired)
        const is401Error =
            error.statusCode === 401 ||
            error.code === 'InvalidAuthenticationToken' ||
            (error.message && error.message.includes('token is expired'));

        if (is401Error && retryCount < MAX_RETRIES) {
            console.log(`⚠️ 401 error detected, attempting token refresh (retry ${retryCount + 1}/${MAX_RETRIES})`);

            try {
                // Refresh token
                await refreshTokenIfNeeded(sessionId);

                // Retry the API call with new token
                const updatedTokenData = userTokenStore.get(sessionId);
                return await apiCall(updatedTokenData.accessToken);

            } catch (refreshError) {
                console.error(`❌ Token refresh failed:`, refreshError.message);
                throw new Error('Authentication failed. Please log in again.');
            }
        }

        // If not a 401 error or max retries exceeded, throw original error
        throw error;
    }
}

/**
 * Get current access token for a session (with automatic refresh)
 * @param {string} sessionId - User session ID
 * @returns {Promise<string>} - Current valid access token
 */
async function getAccessToken(sessionId) {
    return await refreshTokenIfNeeded(sessionId);
}

module.exports = {
    refreshTokenIfNeeded,
    handleGraphApiCall,
    getAccessToken,
    needsRefresh
};
