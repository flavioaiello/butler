/**
 * Butler Extension - Popup Script
 * Archives emails that have been replied to
 */

"use strict";

const DOM_ELEMENT_IDS = Object.freeze({
  DRY_RUN_BTN: "dryRunBtn",
  ARCHIVE_BTN: "archiveBtn",
  CLEAR_BTN: "clearBtn",
  TOKEN_STATUS: "tokenStatus",
  STATUS: "status",
  RESULT_SECTION: "resultSection",
  RESULT_SUMMARY: "resultSummary",
  LOG_OUTPUT: "logOutput"
});

const CSS_CLASSES = Object.freeze({
  HIDDEN: "hidden",
  SUCCESS: "success",
  ERROR: "error",
  WARNING: "warning",
  PROCESSING: "processing"
});

const MESSAGES = Object.freeze({
  NO_TOKEN: "No valid token found. Please open Outlook Web first.",
  TOKEN_NOT_FOUND: "Token not found. Please open Outlook Web first.",
  SCANNING: "Scanning emails...",
  ARCHIVING: "Archiving emails...",
  NO_EMAILS: "No replied-to emails found to archive.",
  FOUND: (count) => `🔍 Found ${count} email(s) to archive`,
  ARCHIVED: (count) => `✅ Archived ${count} email(s)`,
  ERROR: (msg) => `Error: ${msg}`,
  CLEARED: "Tokens cleared"
});

const TIMEOUTS = Object.freeze({
  STATUS_CLEAR_MS: 3000,
  ARCHIVE_TIMEOUT_MS: 120000
});

// Optional behavior toggle (no UI change): include Inbox subfolders in scan/archive.
const INCLUDE_SUBFOLDERS = true;

/**
 * Safe DOM element retrieval with error handling
 * @param {string} id - Element ID
 * @returns {HTMLElement|null}
 */
function getElement(id) {
  const element = document.getElementById(id);
  if (!element) {
    console.warn(`Element not found: ${id}`);
  }
  return element;
}

/**
 * Safely set element text content
 * @param {string} id - Element ID
 * @param {string} text - Text to set
 */
function setText(id, text) {
  const element = getElement(id);
  if (element) {
    element.textContent = text;
  }
}

/**
 * Show status message with optional class
 * @param {string} message - Status message
 * @param {string} [className] - CSS class for styling
 */
function showStatus(message, className = "") {
  const statusElement = getElement(DOM_ELEMENT_IDS.STATUS);
  if (statusElement) {
    statusElement.textContent = message;
    statusElement.className = `status ${className}`.trim();
    statusElement.classList.remove(CSS_CLASSES.HIDDEN);
  }
}

/**
 * Clear status message
 */
function clearStatus() {
  const statusElement = getElement(DOM_ELEMENT_IDS.STATUS);
  if (statusElement) {
    statusElement.textContent = "";
    statusElement.classList.add(CSS_CLASSES.HIDDEN);
  }
}

/**
 * Set button loading state
 * @param {boolean} loading - Whether button is loading
 * @param {string} [activeButton] - Which button triggered the action ('dryRun' or 'archive')
 */
function setButtonLoading(loading, activeButton = 'archive') {
  const archiveBtn = getElement(DOM_ELEMENT_IDS.ARCHIVE_BTN);
  const dryRunBtn = getElement(DOM_ELEMENT_IDS.DRY_RUN_BTN);
  
  if (archiveBtn) {
    archiveBtn.disabled = loading;
    if (activeButton === 'archive') {
      archiveBtn.textContent = loading ? "⏳ Archiving..." : "📦 Archive All";
    } else {
      archiveBtn.textContent = "📦 Archive All";
    }
  }
  
  if (dryRunBtn) {
    dryRunBtn.disabled = loading;
    if (activeButton === 'dryRun') {
      dryRunBtn.textContent = loading ? "⏳ Scanning..." : "🔍 Dry Run";
    } else {
      dryRunBtn.textContent = "🔍 Dry Run";
    }
  }
}

/**
 * Append to log output
 * @param {string} message - Log message
 */
function appendLog(message) {
  const logOutput = getElement(DOM_ELEMENT_IDS.LOG_OUTPUT);
  if (logOutput) {
    logOutput.textContent += message + "\n";
    logOutput.scrollTop = logOutput.scrollHeight;
  }
}

function appendFolderCounts(title, rows) {
  if (!Array.isArray(rows) || rows.length === 0) return;
  appendLog(`\n${title}`);
  rows.forEach((row) => {
    if (!row || typeof row.folder !== 'string' || typeof row.count !== 'number') return;
    appendLog(`  - ${row.folder}: ${row.count}`);
  });
}

function appendFolderScanStats(title, rows) {
  if (!Array.isArray(rows) || rows.length === 0) return;
  appendLog(`\n${title}`);
  rows.forEach((row) => {
    if (!row || typeof row.folder !== 'string') return;
    const fetched = typeof row.fetched === 'number' ? row.fetched : 0;
    const included = typeof row.included === 'number' ? row.included : 0;
    const suffix = row.error
      ? ` (error: ${row.error})`
      : (row.truncated ? ' (truncated)' : '');
    appendLog(`  - ${row.folder}: fetched ${fetched}, included ${included}${suffix}`);
  });
}

/**
 * Clear log output
 */
function clearLog() {
  const logOutput = getElement(DOM_ELEMENT_IDS.LOG_OUTPUT);
  if (logOutput) {
    logOutput.textContent = "";
  }
}

/**
 * Show result section
 * @param {boolean} show - Whether to show
 */
function showResultSection(show) {
  const resultSection = getElement(DOM_ELEMENT_IDS.RESULT_SECTION);
  if (resultSection) {
    resultSection.classList.toggle(CSS_CLASSES.HIDDEN, !show);
  }
}

// Token expiry time (must match background.js)
const TOKEN_EXPIRY_MS = 24 * 60 * 60 * 1000;

/**
 * Check if a token is still valid based on timestamp
 */
function isTokenValid(token) {
  if (!token || !token.timestamp) return false;
  return (Date.now() - token.timestamp) < TOKEN_EXPIRY_MS;
}

/**
 * Update token status display
 */
async function updateTokenStatus() {
  try {
    const data = await chrome.storage.local.get(["capturedTokens"]);
    const tokens = data.capturedTokens || [];
    const validTokens = tokens.filter(isTokenValid);
    
    const statusElement = getElement(DOM_ELEMENT_IDS.TOKEN_STATUS);
    if (statusElement) {
      if (validTokens.length > 0) {
        // Token exists - hide the status section entirely
        statusElement.style.display = "none";
      } else {
        // No token - show warning message
        statusElement.style.display = "block";
        statusElement.textContent = MESSAGES.TOKEN_NOT_FOUND;
        statusElement.className = "token-status empty";
      }
    }
  } catch (error) {
    console.error("Failed to update token status:", error);
  }
}

/**
 * Get most recent valid token
 * @returns {Promise<string|null>}
 */
async function getToken() {
  try {
    const data = await chrome.storage.local.get(["capturedTokens"]);
    const tokens = data.capturedTokens || [];
    const validTokens = tokens
      .filter(isTokenValid)
      .sort((a, b) => b.timestamp - a.timestamp);
    
    return validTokens.length > 0 ? validTokens[0].token : null;
  } catch (error) {
    console.error("Failed to get token:", error);
    return null;
  }
}

/**
 * Perform dry run - scan and show what would be archived
 */
async function performDryRun() {
  const token = await getToken();
  
  if (!token) {
    showStatus(MESSAGES.NO_TOKEN, CSS_CLASSES.ERROR);
    return;
  }
  
  setButtonLoading(true, 'dryRun');
  clearLog();
  showResultSection(true);
  showStatus(MESSAGES.SCANNING, CSS_CLASSES.PROCESSING);
  
  appendLog("🔍 DRY RUN - Scanning inbox...");
  appendLog("Fetching inbox messages (up to 2000)...");
  
  try {
    const response = await chrome.runtime.sendMessage({
      action: "archiveRepliedEmails",
      dryRun: true,
      includeSubfolders: INCLUDE_SUBFOLDERS
    });
    
    if (!response) {
      throw new Error("No response from background script");
    }
    
    if (!response.success) {
      throw new Error(response.error || "Unknown error occurred");
    }
    
    const { foundCount, totalScanned, foundSubjects, duplicateCount, duplicateGroups, folderStats, toArchiveByFolder } = response;
    
    appendLog(`\\nScanned ${totalScanned} emails`);

    appendFolderScanStats('Scan breakdown (per folder):', folderStats);
    appendFolderCounts('Replied-to emails to archive (per folder):', toArchiveByFolder);
    appendLog(`Found ${foundCount} replied-to emails to archive`);
    appendLog(`Found ${duplicateCount || 0} duplicate emails (same Message-ID)\n`);
    
    if (foundCount > 0) {
      appendLog("Would archive (replied-to):");
      foundSubjects.forEach((subject, i) => {
        const displaySubject = subject.length > 50 
          ? subject.substring(0, 47) + "..." 
          : subject;
        appendLog(`  ${i + 1}. ${displaySubject}`);
      });
    }
    
    if (duplicateCount > 0 && duplicateGroups) {
      appendLog("\nDuplicate groups (same Message-ID):");
      duplicateGroups.forEach((group, i) => {
        const displaySubject = group.subject.length > 40 
          ? group.subject.substring(0, 37) + "..." 
          : group.subject;
        appendLog(`  ${i + 1}. "${displaySubject}" (${group.count} copies)`);
      });
    }
    
    if (foundCount > 0 || duplicateCount > 0) {
      const statusMsg = [];
      if (foundCount > 0) statusMsg.push(`${foundCount} to archive`);
      if (duplicateCount > 0) statusMsg.push(`${duplicateCount} duplicates`);
      showStatus(`🔍 Found: ${statusMsg.join(', ')}`, CSS_CLASSES.SUCCESS);
      setText(DOM_ELEMENT_IDS.RESULT_SUMMARY, 
        `🔍 ${foundCount} to archive, ${duplicateCount || 0} duplicates (of ${totalScanned} scanned)`);
    } else {
      appendLog("\nNo emails to archive and no duplicates found.");
      showStatus(MESSAGES.NO_EMAILS, CSS_CLASSES.WARNING);
      setText(DOM_ELEMENT_IDS.RESULT_SUMMARY, 
        `Scanned ${totalScanned} emails - none need archiving, no duplicates`);
    }
    
  } catch (error) {
    console.error("Dry run error:", error);
    appendLog(`\n❌ Error: ${error.message}`);
    showStatus(MESSAGES.ERROR(error.message), CSS_CLASSES.ERROR);
    setText(DOM_ELEMENT_IDS.RESULT_SUMMARY, "Scan failed");
  } finally {
    setButtonLoading(false, 'dryRun');
  }
}

/**
 * Archive emails that have been replied to
 */
async function archiveRepliedEmails() {
  const token = await getToken();
  
  if (!token) {
    showStatus(MESSAGES.NO_TOKEN, CSS_CLASSES.ERROR);
    return;
  }
  
  setButtonLoading(true, 'archive');
  clearLog();
  showResultSection(true);
  showStatus(MESSAGES.ARCHIVING, CSS_CLASSES.PROCESSING);
  
  appendLog("📦 ARCHIVE - Starting archive process...");
  appendLog("Fetching inbox messages (up to 2000)...");
  
  try {
    const response = await chrome.runtime.sendMessage({
      action: "archiveRepliedEmails",
      dryRun: false,
      includeSubfolders: INCLUDE_SUBFOLDERS
    });
    
    if (!response) {
      throw new Error("No response from background script");
    }
    
    if (!response.success) {
      throw new Error(response.error || "Unknown error occurred");
    }
    
    const { archivedCount, totalScanned, archivedSubjects, duplicatesMovedCount, folderStats, toArchiveByFolder, archivedByFolder } = response;

    appendLog(`\\nScanned ${totalScanned} emails`);

    appendFolderScanStats('Scan breakdown (per folder):', folderStats);
    appendFolderCounts('Emails to archive (per folder):', toArchiveByFolder);

    if (duplicatesMovedCount > 0) {
      appendLog(`Moved ${duplicatesMovedCount} duplicates to "Duplicates" folder`);
    }
    appendLog(`Archived ${archivedCount} replied-to emails\\n`);
    
    if (archivedCount > 0) {
      appendLog("Archived:");
      archivedSubjects.forEach((subject, i) => {
        const displaySubject = subject.length > 50 
          ? subject.substring(0, 47) + "..." 
          : subject;
        appendLog(`  ${i + 1}. ${displaySubject}`);
      });
    }
    
    if (archivedCount > 0 || duplicatesMovedCount > 0) {
      const parts = [];
      if (duplicatesMovedCount > 0) parts.push(`${duplicatesMovedCount} duplicates moved`);
      if (archivedCount > 0) parts.push(`${archivedCount} archived`);
      showStatus(`✅ ${parts.join(', ')}`, CSS_CLASSES.SUCCESS);
      setText(DOM_ELEMENT_IDS.RESULT_SUMMARY, 
        `✅ ${duplicatesMovedCount || 0} duplicates moved, ${archivedCount} archived`);
    } else {
      appendLog("No emails to archive and no duplicates found.");
      showStatus(MESSAGES.NO_EMAILS, CSS_CLASSES.WARNING);
      setText(DOM_ELEMENT_IDS.RESULT_SUMMARY, 
        `Scanned ${totalScanned} emails - nothing to process`);
    }

    appendFolderCounts('Archived emails (per folder):', archivedByFolder);
    
  } catch (error) {
    console.error("Archive error:", error);
    appendLog(`\n❌ Error: ${error.message}`);
    showStatus(MESSAGES.ERROR(error.message), CSS_CLASSES.ERROR);
    setText(DOM_ELEMENT_IDS.RESULT_SUMMARY, "Archive failed");
  } finally {
    setButtonLoading(false, 'archive');
  }
}

/**
 * Clear all stored tokens
 */
async function clearTokens() {
  try {
    await chrome.storage.local.remove(["capturedTokens"]);
    showStatus(MESSAGES.CLEARED, CSS_CLASSES.SUCCESS);
    await updateTokenStatus();
    showResultSection(false);
    clearLog();
    
    setTimeout(clearStatus, TIMEOUTS.STATUS_CLEAR_MS);
  } catch (error) {
    console.error("Failed to clear tokens:", error);
    showStatus(MESSAGES.ERROR(error.message), CSS_CLASSES.ERROR);
  }
}

/**
 * Initialize popup
 */
function init() {
  // Update token status on load
  updateTokenStatus();
  
  // Hide result section initially
  showResultSection(false);
  
  // Set up event listeners
  const dryRunBtn = getElement(DOM_ELEMENT_IDS.DRY_RUN_BTN);
  if (dryRunBtn) {
    dryRunBtn.addEventListener("click", performDryRun);
  }
  
  const archiveBtn = getElement(DOM_ELEMENT_IDS.ARCHIVE_BTN);
  if (archiveBtn) {
    archiveBtn.addEventListener("click", archiveRepliedEmails);
  }
  
  const clearBtn = getElement(DOM_ELEMENT_IDS.CLEAR_BTN);
  if (clearBtn) {
    clearBtn.addEventListener("click", clearTokens);
  }
  
  // Listen for storage changes to update token status
  chrome.storage.onChanged.addListener((changes, areaName) => {
    if (areaName === "local" && changes.capturedTokens) {
      updateTokenStatus();
    }
  });
}

// Initialize when DOM is ready
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", init);
} else {
  init();
}
