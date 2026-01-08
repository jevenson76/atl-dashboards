/**
 * spApi.js - SharePoint REST API Helper for ATL Integration Hub
 * Version: 1.0.0
 *
 * Provides unified SharePoint REST access for all dashboard pages.
 * Works in both embedded iframe and direct page access contexts.
 *
 * HARD CONSTRAINTS:
 * - NO Power Automate flow triggers
 * - NO email/Teams/notification triggers
 * - Read-only GET operations only
 */

(function(global) {
    'use strict';

    // ============================================================
    // Configuration
    // ============================================================

    var CONFIG = {
        // Fallback site URL if detection fails
        fallbackSiteUrl: 'https://chamberlaingroup.sharepoint.com/sites/PrincipalGTMStrategy-InternalUseOnly-ATLIntegrationProject',

        // Power Automate HTTP Trigger URL for external hosting (bypasses CORS)
        // SET THIS AFTER CREATING THE FLOW - get URL from Flow trigger
        flowProxyUrl: null,  // e.g., 'https://prod-XX.westus.logic.azure.com:443/workflows/...'

        // Auto-detect if running externally (not on SharePoint)
        // When true and flowProxyUrl is set, uses Flow proxy for data
        useFlowProxy: null,  // null = auto-detect, true = force proxy, false = force direct

        // Canonical list names (ONLY use these - verified 2026-01-07)
        lists: {
            tasks: 'ATL_Project_Plan.v21',      // CANONICAL - 172 items
            salesData: 'ATL-SalesData',
            activityLog: 'ActivityLog',
            peopleMap: 'PeopleMap'
        },

        // Default OData settings
        defaultTop: 500,

        // Debug mode - set true to enable console logging
        debug: true
    };

    // ============================================================
    // State
    // ============================================================

    var state = {
        webAbsoluteUrl: null,
        lastError: null,
        lastFetch: null
    };

    // ============================================================
    // Debug Panel
    // ============================================================

    /**
     * Log message to console and optional debug panel
     */
    function log(level, message, data) {
        if (!CONFIG.debug && level === 'debug') return;

        var timestamp = new Date().toISOString().substr(11, 12);
        var prefix = '[spApi ' + timestamp + '] ';

        if (level === 'error') {
            console.error(prefix + message, data || '');
        } else if (level === 'warn') {
            console.warn(prefix + message, data || '');
        } else {
            console.log(prefix + message, data || '');
        }

        // Update debug panel if it exists
        updateDebugPanel(level, message, data);
    }

    /**
     * Update on-page debug panel (if present)
     */
    function updateDebugPanel(level, message, data) {
        var panel = document.getElementById('sp-api-debug');
        if (!panel) return;

        var entry = document.createElement('div');
        entry.className = 'sp-debug-entry sp-debug-' + level;
        entry.textContent = '[' + level.toUpperCase() + '] ' + message;

        if (data) {
            var details = document.createElement('pre');
            details.textContent = typeof data === 'string' ? data : JSON.stringify(data, null, 2);
            entry.appendChild(details);
        }

        panel.insertBefore(entry, panel.firstChild);

        // Keep only last 20 entries
        while (panel.children.length > 20) {
            panel.removeChild(panel.lastChild);
        }
    }

    /**
     * Create debug panel HTML (call once on page load)
     */
    function createDebugPanel() {
        if (document.getElementById('sp-api-debug')) return;

        // Create container
        var container = document.createElement('div');
        container.id = 'sp-api-debug-container';

        // Create and append style element
        var style = document.createElement('style');
        style.textContent = [
            '#sp-api-debug-container { position: fixed; bottom: 0; left: 0; right: 0; z-index: 9999; font-family: monospace; font-size: 11px; }',
            '#sp-api-debug-toggle { background: #003087; color: white; border: none; padding: 4px 12px; cursor: pointer; }',
            '#sp-api-debug { display: none; background: #1a1a2e; color: #eee; max-height: 200px; overflow-y: auto; padding: 8px; }',
            '#sp-api-debug.visible { display: block; }',
            '.sp-debug-entry { padding: 2px 0; border-bottom: 1px solid #333; }',
            '.sp-debug-error { color: #ff6b6b; }',
            '.sp-debug-warn { color: #ffd93d; }',
            '.sp-debug-info { color: #6bcb77; }',
            '.sp-debug-entry pre { margin: 4px 0 4px 16px; font-size: 10px; color: #888; white-space: pre-wrap; word-break: break-all; }'
        ].join('\n');
        container.appendChild(style);

        // Create toggle button
        var toggleBtn = document.createElement('button');
        toggleBtn.id = 'sp-api-debug-toggle';
        toggleBtn.textContent = 'Debug Panel';
        toggleBtn.addEventListener('click', function() {
            document.getElementById('sp-api-debug').classList.toggle('visible');
        });
        container.appendChild(toggleBtn);

        // Create debug output panel
        var debugPanel = document.createElement('div');
        debugPanel.id = 'sp-api-debug';
        container.appendChild(debugPanel);

        document.body.appendChild(container);
    }

    // ============================================================
    // Proxy Detection
    // ============================================================

    /**
     * Detect if we should use the Flow proxy for data requests.
     * Returns true if:
     * - CONFIG.useFlowProxy is explicitly true, OR
     * - CONFIG.useFlowProxy is null (auto) AND we're not on SharePoint domain
     */
    function shouldUseFlowProxy() {
        // Explicit override
        if (CONFIG.useFlowProxy === true) return true;
        if (CONFIG.useFlowProxy === false) return false;

        // Auto-detect: check if we're on SharePoint
        var isSharePoint = location.hostname.indexOf('sharepoint.com') !== -1;
        var hasFlowUrl = CONFIG.flowProxyUrl && CONFIG.flowProxyUrl.length > 0;

        if (!isSharePoint && hasFlowUrl) {
            log('info', 'External hosting detected - using Flow proxy');
            return true;
        }

        return false;
    }

    /**
     * Call the Power Automate HTTP trigger to get list data
     *
     * @param {string} listName - SharePoint list name
     * @param {Object} options - Query options (select, filter, top, orderby)
     * @returns {Promise<Array>} Array of list items
     */
    function callFlowProxy(listName, options) {
        if (!CONFIG.flowProxyUrl) {
            return Promise.reject(new Error('Flow proxy URL not configured. Set spApi.config.flowProxyUrl'));
        }

        options = options || {};

        // Build request body for the Flow
        var requestBody = {
            listName: listName,
            select: options.select || null,
            filter: options.filter || null,
            top: options.top || CONFIG.defaultTop,
            orderby: options.orderby || null
        };

        log('debug', 'Calling Flow proxy for: ' + listName, requestBody);

        return fetch(CONFIG.flowProxyUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(requestBody)
        })
        .then(function(response) {
            if (!response.ok) {
                return response.text().then(function(text) {
                    throw new Error('Flow proxy error ' + response.status + ': ' + text.slice(0, 200));
                });
            }
            return response.json();
        })
        .then(function(data) {
            // Flow returns { items: [...], count: N }
            var items = data.items || data.value || [];
            log('info', 'Flow proxy returned ' + items.length + ' items from ' + listName);
            return items;
        })
        .catch(function(error) {
            log('error', 'Flow proxy failed: ' + error.message);
            state.lastError = { url: CONFIG.flowProxyUrl, error: error.message, time: new Date() };
            throw error;
        });
    }

    // ============================================================
    // URL Detection
    // ============================================================

    /**
     * Get the absolute URL of the current SharePoint web/site.
     * Uses multiple detection methods with fallbacks.
     *
     * @returns {string} The web absolute URL (e.g., https://tenant.sharepoint.com/sites/SiteName)
     */
    function getWebAbsoluteUrl() {
        // Return cached value if available
        if (state.webAbsoluteUrl) {
            return state.webAbsoluteUrl;
        }

        var url = null;
        var method = 'unknown';

        // Method 1: _spPageContextInfo (most reliable when available)
        if (typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo.webAbsoluteUrl) {
            url = _spPageContextInfo.webAbsoluteUrl;
            method = '_spPageContextInfo';
        }

        // Method 2: Parse from current URL pathname
        if (!url) {
            var pathMatch = location.pathname.match(/^(\/sites\/[^\/]+)/i);
            if (pathMatch) {
                url = location.origin + pathMatch[1];
                method = 'pathname';
            }
        }

        // Method 3: Check for SP context in parent frame (for embeds)
        if (!url && window.parent !== window) {
            try {
                if (window.parent._spPageContextInfo && window.parent._spPageContextInfo.webAbsoluteUrl) {
                    url = window.parent._spPageContextInfo.webAbsoluteUrl;
                    method = 'parent._spPageContextInfo';
                }
            } catch (e) {
                // Cross-origin access denied - expected for security
            }
        }

        // Method 4: Fallback to configured URL
        if (!url) {
            url = CONFIG.fallbackSiteUrl;
            method = 'fallback';
            log('warn', 'Using fallback site URL - detection failed');
        }

        // Remove trailing slash if present
        url = url.replace(/\/$/, '');

        state.webAbsoluteUrl = url;
        log('info', 'Site URL detected via ' + method + ': ' + url);

        return url;
    }

    /**
     * Build full REST API URL from relative path
     *
     * @param {string} path - Relative path (e.g., "/_api/web/lists")
     * @returns {string} Full URL
     */
    function buildUrl(path) {
        var baseUrl = getWebAbsoluteUrl();

        // If path is already absolute URL, return as-is
        if (path.indexOf('http') === 0) {
            return path;
        }

        // Ensure path starts with /
        if (path.charAt(0) !== '/') {
            path = '/' + path;
        }

        return baseUrl + path;
    }

    // ============================================================
    // REST API Methods
    // ============================================================

    /**
     * Execute a GET request against SharePoint REST API
     *
     * @param {string} pathOrUrl - Relative path or full URL
     * @param {Object} options - Optional fetch options override
     * @returns {Promise<Object>} Parsed JSON response
     */
    function spGet(pathOrUrl, options) {
        var url = buildUrl(pathOrUrl);

        var fetchOptions = {
            method: 'GET',
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-Type': 'application/json'
            },
            credentials: 'include'
        };

        // Merge custom options
        if (options && options.headers) {
            Object.assign(fetchOptions.headers, options.headers);
        }

        log('debug', 'GET ' + url);
        state.lastFetch = { url: url, time: new Date() };

        return fetch(url, fetchOptions)
            .then(function(response) {
                if (!response.ok) {
                    // Extract error details from response body
                    return response.text().then(function(text) {
                        var errorDetail = text.slice(0, 300);
                        var errMsg = 'REST ' + response.status + ' ' + response.statusText + ': ' + errorDetail;
                        log('error', 'REST FAILED: ' + url, errMsg);
                        throw new Error(errMsg);
                    });
                }
                return response.json();
            })
            .then(function(data) {
                var itemCount = data.value ? data.value.length : 'N/A';
                log('info', 'REST PASS: ' + itemCount + ' items from ' + url);
                state.lastError = null;
                return data;
            })
            .catch(function(error) {
                state.lastError = { url: url, error: error.message, time: new Date() };
                log('error', 'REST FAIL: ' + url, error.message);
                // Surface error in UI if debug panel exists
                var debugEl = document.getElementById('sp-api-debug');
                if (debugEl) {
                    debugEl.classList.add('visible');
                }
                throw error;
            });
    }

    /**
     * Get items from a SharePoint list with OData query options
     *
     * @param {string} listTitle - Display name of the list
     * @param {Object} options - Query options
     * @param {string} options.select - Comma-separated field names (default: all)
     * @param {string} options.expand - Expand lookup fields
     * @param {string} options.filter - OData filter expression
     * @param {number} options.top - Max items to return (default: 500)
     * @param {string} options.orderby - Sort expression (e.g., "Modified desc")
     * @returns {Promise<Array>} Array of list items
     */
    function listItems(listTitle, options) {
        options = options || {};

        // Route through Flow proxy if external hosting detected
        if (shouldUseFlowProxy()) {
            return callFlowProxy(listTitle, options);
        }

        // Direct SharePoint REST API call
        // Build query string
        var queryParts = [];

        if (options.select) {
            queryParts.push('$select=' + encodeURIComponent(options.select));
        }

        if (options.expand) {
            queryParts.push('$expand=' + encodeURIComponent(options.expand));
        }

        if (options.filter) {
            queryParts.push('$filter=' + encodeURIComponent(options.filter));
        }

        queryParts.push('$top=' + (options.top || CONFIG.defaultTop));

        if (options.orderby) {
            queryParts.push('$orderby=' + encodeURIComponent(options.orderby));
        }

        var path = "/_api/web/lists/getbytitle('" + encodeURIComponent(listTitle) + "')/items";
        if (queryParts.length > 0) {
            path += '?' + queryParts.join('&');
        }

        log('info', 'Fetching list: ' + listTitle);

        return spGet(path).then(function(data) {
            var items = data.value || [];
            log('info', 'Loaded ' + items.length + ' items from ' + listTitle);
            return items;
        });
    }

    /**
     * Get a single list item by ID
     *
     * @param {string} listTitle - Display name of the list
     * @param {number} itemId - Item ID
     * @param {Object} options - Query options (select, expand)
     * @returns {Promise<Object>} List item
     */
    function getItem(listTitle, itemId, options) {
        options = options || {};

        var queryParts = [];

        if (options.select) {
            queryParts.push('$select=' + encodeURIComponent(options.select));
        }

        if (options.expand) {
            queryParts.push('$expand=' + encodeURIComponent(options.expand));
        }

        var path = "/_api/web/lists/getbytitle('" + encodeURIComponent(listTitle) + "')/items(" + itemId + ")";
        if (queryParts.length > 0) {
            path += '?' + queryParts.join('&');
        }

        return spGet(path);
    }

    /**
     * Get list metadata (field definitions, item count, etc.)
     *
     * @param {string} listTitle - Display name of the list
     * @returns {Promise<Object>} List metadata
     */
    function getListInfo(listTitle) {
        var path = "/_api/web/lists/getbytitle('" + encodeURIComponent(listTitle) + "')?$select=Title,ItemCount,LastItemModifiedDate,Description";
        return spGet(path);
    }

    /**
     * Get current web/site information
     *
     * @returns {Promise<Object>} Web info including Title, Url, etc.
     */
    function getWebInfo() {
        return spGet('/_api/web?$select=Title,Url,Description,Created');
    }

    // ============================================================
    // Convenience Methods for ATL Lists
    // ============================================================

    /**
     * Get project tasks from ATL-CG Project Tasks list
     *
     * @param {Object} options - Filter/select options
     * @returns {Promise<Array>} Task items
     */
    function getTasks(options) {
        options = options || {};

        // Default select for common task fields
        if (!options.select) {
            options.select = 'Id,Title,TaskID,Phase,Workstream,Status,PercentComplete,Priority,DueDate,Owner,Modified,TaskType,Description,IsBlocked,BlockerReason';
        }

        // Default sort by due date
        if (!options.orderby) {
            options.orderby = 'DueDate asc';
        }

        return listItems(CONFIG.lists.tasks, options);
    }

    /**
     * Get sales data from ATL-SalesData list
     *
     * @param {Object} options - Filter/select options
     * @returns {Promise<Array>} Sales data items
     */
    function getSalesData(options) {
        options = options || {};

        if (!options.select) {
            options.select = 'Id,Title,SalesDate,DailySalesActual,DailyBudgetTarget,DailyVariance,DailyVariancePercent,VolumeLevel,MTDSalesActual,MTDBudgetTarget,YTDSalesActual,YTDBudgetTarget,ImportTimestamp,DataQuality';
        }

        if (!options.orderby) {
            options.orderby = 'SalesDate desc';
        }

        return listItems(CONFIG.lists.salesData, options);
    }

    /**
     * Get data freshness timestamp for a list
     *
     * @param {string} listTitle - List to check
     * @returns {Promise<Date>} Last modified date
     */
    function getDataFreshness(listTitle) {
        return getListInfo(listTitle).then(function(info) {
            return new Date(info.LastItemModifiedDate);
        });
    }

    // ============================================================
    // UI Helpers
    // ============================================================

    /**
     * Format date for display
     *
     * @param {Date|string} date - Date to format
     * @param {boolean} includeTime - Include time portion
     * @returns {string} Formatted date string
     */
    function formatDate(date, includeTime) {
        if (!date) return 'N/A';

        var d = date instanceof Date ? date : new Date(date);

        if (isNaN(d.getTime())) return 'Invalid Date';

        var options = {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        };

        if (includeTime) {
            options.hour = '2-digit';
            options.minute = '2-digit';
        }

        return d.toLocaleDateString('en-US', options);
    }

    /**
     * Create/update a "Data Updated As Of" timestamp element
     *
     * @param {string} containerId - ID of container element
     * @param {Date} timestamp - Last update time
     */
    function showDataTimestamp(containerId, timestamp) {
        var container = document.getElementById(containerId);
        if (!container) {
            // Create default container if not exists
            container = document.createElement('div');
            container.id = containerId;
            container.className = 'data-timestamp';
            container.style.cssText = 'font-size: 11px; color: #666; text-align: right; padding: 4px 8px; background: #f5f5f5; border-radius: 4px; margin-top: 8px;';
            document.body.appendChild(container);
        }

        container.textContent = 'Data Updated: ' + formatDate(timestamp, true);
    }

    /**
     * Show loading indicator
     *
     * @param {string} containerId - Container to show loading in
     * @param {string} message - Loading message
     */
    function showLoading(containerId, message) {
        var container = document.getElementById(containerId);
        if (!container) return;

        // Clear existing content
        while (container.firstChild) {
            container.removeChild(container.firstChild);
        }

        // Create loading element
        var loadingDiv = document.createElement('div');
        loadingDiv.className = 'sp-loading';
        loadingDiv.style.cssText = 'text-align: center; padding: 40px; color: #666;';

        var iconDiv = document.createElement('div');
        iconDiv.style.cssText = 'font-size: 24px; margin-bottom: 8px;';
        iconDiv.textContent = '\u23F3'; // hourglass

        var msgDiv = document.createElement('div');
        msgDiv.textContent = message || 'Loading...';

        loadingDiv.appendChild(iconDiv);
        loadingDiv.appendChild(msgDiv);
        container.appendChild(loadingDiv);
    }

    /**
     * Show error message
     *
     * @param {string} containerId - Container to show error in
     * @param {string} message - Error message
     * @param {boolean} showRetry - Show retry button
     */
    function showError(containerId, message, showRetry) {
        var container = document.getElementById(containerId);
        if (!container) return;

        // Clear existing content
        while (container.firstChild) {
            container.removeChild(container.firstChild);
        }

        // Create error element
        var errorDiv = document.createElement('div');
        errorDiv.className = 'sp-error';
        errorDiv.style.cssText = 'text-align: center; padding: 40px; color: #dc3545; background: #fff5f5; border-radius: 8px;';

        var iconDiv = document.createElement('div');
        iconDiv.style.cssText = 'font-size: 24px; margin-bottom: 8px;';
        iconDiv.textContent = '\u26A0'; // warning sign

        var msgDiv = document.createElement('div');
        msgDiv.style.cssText = 'margin-bottom: 16px;';
        msgDiv.textContent = message;

        errorDiv.appendChild(iconDiv);
        errorDiv.appendChild(msgDiv);

        if (showRetry) {
            var retryBtn = document.createElement('button');
            retryBtn.textContent = 'Retry';
            retryBtn.style.cssText = 'background: #003087; color: white; border: none; padding: 8px 24px; border-radius: 4px; cursor: pointer;';
            retryBtn.addEventListener('click', function() {
                location.reload();
            });
            errorDiv.appendChild(retryBtn);
        }

        container.appendChild(errorDiv);
    }

    // ============================================================
    // Status Helpers
    // ============================================================

    /**
     * Get status information (for debugging/validation)
     */
    function getStatus() {
        return {
            webAbsoluteUrl: state.webAbsoluteUrl,
            lastFetch: state.lastFetch,
            lastError: state.lastError,
            config: CONFIG,
            usingFlowProxy: shouldUseFlowProxy(),
            isSharePointHost: location.hostname.indexOf('sharepoint.com') !== -1
        };
    }

    /**
     * Set the Flow proxy URL (call this before fetching data)
     * @param {string} url - The HTTP trigger URL from Power Automate
     */
    function setFlowProxyUrl(url) {
        CONFIG.flowProxyUrl = url;
        log('info', 'Flow proxy URL configured');
    }

    /**
     * Test connectivity to SharePoint
     *
     * @returns {Promise<Object>} Test results
     */
    function testConnection() {
        var results = {
            siteUrl: getWebAbsoluteUrl(),
            webInfo: null,
            tasksListInfo: null,
            errors: []
        };

        return getWebInfo()
            .then(function(info) {
                results.webInfo = info;
                return getListInfo(CONFIG.lists.tasks);
            })
            .then(function(info) {
                results.tasksListInfo = info;
                return results;
            })
            .catch(function(error) {
                results.errors.push(error.message);
                return results;
            });
    }

    // ============================================================
    // Export Public API
    // ============================================================

    global.spApi = {
        // Configuration
        config: CONFIG,

        // URL helpers
        getWebAbsoluteUrl: getWebAbsoluteUrl,
        buildUrl: buildUrl,

        // Core REST methods
        spGet: spGet,
        listItems: listItems,
        getItem: getItem,
        getListInfo: getListInfo,
        getWebInfo: getWebInfo,

        // ATL convenience methods
        getTasks: getTasks,
        getSalesData: getSalesData,
        getDataFreshness: getDataFreshness,

        // UI helpers
        formatDate: formatDate,
        showDataTimestamp: showDataTimestamp,
        showLoading: showLoading,
        showError: showError,
        createDebugPanel: createDebugPanel,

        // Status/debugging
        getStatus: getStatus,
        testConnection: testConnection,
        log: log,

        // Flow proxy configuration
        setFlowProxyUrl: setFlowProxyUrl,
        shouldUseFlowProxy: shouldUseFlowProxy
    };

    // Auto-initialize on DOM ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function() {
            log('info', 'spApi initialized');
            if (CONFIG.debug) {
                createDebugPanel();
            }
        });
    } else {
        log('info', 'spApi initialized');
        if (CONFIG.debug) {
            createDebugPanel();
        }
    }

})(typeof window !== 'undefined' ? window : this);
