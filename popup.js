/**
 * Butler Extension - Popup Script
 * Archives emails that have been replied to
 * Integrates with local Ollama for AI-powered email management
 */

"use strict";

const DOM_ELEMENT_IDS = Object.freeze({
  DRY_RUN_BTN: "dryRunBtn",
  ARCHIVE_BTN: "archiveBtn",
  CLEAR_BTN: "clearBtn",
  ASK_BTN: "askBtn",
  PROMPT_INPUT: "promptInput",
  ITERATIONS_INPUT: "iterationsInput",
  PROGRESS_BAR: "progressBar",
  PROGRESS_FILL: "progressFill",
  PROGRESS_TEXT: "progressText",
  MODEL_SELECT: "modelSelect",
  TOKEN_STATUS: "tokenStatus",
  STATUS: "status",
  RESULT_SECTION: "resultSection",
  RESULT_SUMMARY: "resultSummary",
  PLAN_SECTION: "planSection",
  PLAN_LIST: "planList",
  LOG_OUTPUT: "logOutput"
});

const CSS_CLASSES = Object.freeze({
  HIDDEN: "hidden",
  SUCCESS: "success",
  ERROR: "error",
  WARNING: "warning",
  PROCESSING: "processing",
  CONNECTED: "connected",
  DISCONNECTED: "disconnected"
});

const MESSAGES = Object.freeze({
  NO_TOKEN: "No valid token found. Please open Outlook Web first.",
  TOKEN_NOT_FOUND: "Token not found. Please open Outlook Web first.",
  SCANNING: "Scanning emails...",
  ARCHIVING: "Archiving emails...",
  NO_EMAILS: "No replied-to emails found to archive.",
  FOUND: (count) => `Found ${count} email(s) to archive`,
  ARCHIVED: (count) => `Archived ${count} email(s)`,
  ERROR: (msg) => `Error: ${msg}`,
  CLEARED: "Tokens cleared",
  ASKING_OLLAMA: "Analyzing with Ollama...",
  EXECUTING: "Executing plan...",
  OLLAMA_CONNECTED: "Ollama",
  OLLAMA_DISCONNECTED: "Ollama"
});

const TIMEOUTS = Object.freeze({
  STATUS_CLEAR_MS: 3000,
  ARCHIVE_TIMEOUT_MS: 120000,
  OLLAMA_CHECK_INTERVAL_MS: 30000
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

/**
 * Show/hide plan section
 * @param {boolean} show - Whether to show
 */
function showPlanSection(show) {
  const planSection = getElement(DOM_ELEMENT_IDS.PLAN_SECTION);
  if (planSection) {
    planSection.classList.toggle(CSS_CLASSES.HIDDEN, !show);
  }
}

/**
 * Populate plan list with items
 * @param {Array} plan - Array of plan items
 */
function populatePlanList(plan) {
  const planList = getElement(DOM_ELEMENT_IDS.PLAN_LIST);
  if (!planList || !Array.isArray(plan)) return;
  
  planList.innerHTML = '';
  
  plan.forEach((item, index) => {
    const itemDiv = document.createElement('div');
    itemDiv.className = 'plan-item';
    
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = true;
    checkbox.className = 'plan-item-checkbox';
    checkbox.dataset.index = index;
    
    const textDiv = document.createElement('div');
    textDiv.className = 'plan-item-text';
    
    const subjectSpan = document.createElement('div');
    subjectSpan.className = 'plan-item-subject';
    subjectSpan.textContent = (item.subject || '(No Subject)').substring(0, 60);
    
    const metaSpan = document.createElement('div');
    metaSpan.className = 'plan-item-meta';
    let metaText = `From: ${(item.from || 'unknown').substring(0, 40)}`;
    if (item.reasoning) {
      metaText += ` | ${item.reasoning}`;
    }
    metaSpan.textContent = metaText;
    
    textDiv.appendChild(subjectSpan);
    textDiv.appendChild(metaSpan);
    
    itemDiv.appendChild(checkbox);
    itemDiv.appendChild(textDiv);
    planList.appendChild(itemDiv);
  });
}

/**
 * Set button states for AI workflow
 * @param {string} state - 'idle', 'asking', 'executing'
 */
function setAIButtonState(state) {
  const askBtn = getElement(DOM_ELEMENT_IDS.ASK_BTN);
  
  switch (state) {
    case 'asking':
      if (askBtn) {
        askBtn.textContent = 'Abort';
        askBtn.className = 'btn btn-abort btn-full';
        askBtn.disabled = false;
      }
      break;
    case 'executing':
      if (askBtn) {
        askBtn.textContent = 'Executing...';
        askBtn.className = 'btn btn-execute btn-full';
        askBtn.disabled = true;
      }
      break;
    case 'idle':
    default:
      if (askBtn) {
        askBtn.textContent = 'Triage Inbox';
        askBtn.className = 'btn btn-primary btn-full';
        askBtn.disabled = false;
      }
      break;
  }
}

/**
 * Populate the model selector dropdown with available Ollama models
 */
async function populateModelSelector() {
  const selectElement = getElement(DOM_ELEMENT_IDS.MODEL_SELECT);
  if (!selectElement) {
    console.error('[Butler Popup] Model select element not found');
    return;
  }
  
  try {
    console.log('[Butler Popup] Requesting Ollama models...');
    const response = await chrome.runtime.sendMessage({ action: 'getOllamaModels' });
    console.log('[Butler Popup] Response:', response);
    
    if (!response || !response.available) {
      selectElement.innerHTML = '<option value="">Ollama offline</option>';
      selectElement.className = 'model-select disconnected';
      selectElement.title = response?.error || 'Ollama not available';
      return;
    }
    
    const models = response.models || [];
    const activeModel = response.activeModel || '';
    
    if (models.length === 0) {
      selectElement.innerHTML = '<option value="">No models</option>';
      selectElement.className = 'model-select disconnected';
      selectElement.title = 'No models installed';
      return;
    }
    
    // Build options HTML
    selectElement.innerHTML = models.map(model => {
      const selected = model === activeModel ? ' selected' : '';
      const displayName = model.length > 20 ? model.substring(0, 18) + '...' : model;
      return `<option value="${model}"${selected}>${displayName}</option>`;
    }).join('');
    
    selectElement.className = 'model-select connected';
    selectElement.title = `Active: ${activeModel}`;
  } catch (error) {
    selectElement.innerHTML = '<option value="">Error</option>';
    selectElement.className = 'model-select disconnected';
    selectElement.title = 'Connection error';
  }
}

/**
 * Handle model selection change
 */
async function handleModelChange(event) {
  const selectedModel = event.target.value;
  if (!selectedModel) return;
  
  try {
    await chrome.runtime.sendMessage({ action: 'setOllamaModel', model: selectedModel });
    event.target.title = `Active: ${selectedModel}`;
  } catch (error) {
    console.error('Failed to set model:', error);
  }
}

/**
 * Update Ollama connection status (refresh model list)
 */
async function updateOllamaStatus() {
  await populateModelSelector();
}

/**
 * Show or hide progress bar
 * @param {boolean} show - Whether to show progress bar
 */
function showProgressBar(show) {
  const progressBar = getElement(DOM_ELEMENT_IDS.PROGRESS_BAR);
  if (progressBar) {
    progressBar.classList.toggle(CSS_CLASSES.HIDDEN, !show);
  }
}

/**
 * Update progress bar
 * @param {number} current - Current progress
 * @param {number} total - Total items
 * @param {number} moved - Number moved so far
 */
function updateProgressBar(current, total, moved) {
  const progressFill = getElement(DOM_ELEMENT_IDS.PROGRESS_FILL);
  const progressText = getElement(DOM_ELEMENT_IDS.PROGRESS_TEXT);
  
  if (progressFill && progressText) {
    const percent = total > 0 ? (current / total) * 100 : 0;
    progressFill.style.width = `${percent}%`;
    const movedText = typeof moved === 'number' && moved > 0 ? ` (${moved} moved)` : '';
    progressText.textContent = `${current}/${total}${movedText}`;
  }
}

// Track last seen result index to avoid duplicates
let lastSeenResultIndex = 0;

/**
 * Poll classification progress and append results
 * @param {AbortSignal} signal - Abort signal to stop polling
 */
async function pollProgress(signal) {
  lastSeenResultIndex = 0;
  
  while (!signal.aborted) {
    try {
      const progress = await chrome.runtime.sendMessage({ action: 'getClassificationProgress' });
      if (progress && progress.total > 0) {
        updateProgressBar(progress.current, progress.total, progress.moved);
        
        // If there's a new result, append it to the log
        if (progress.lastResult && progress.current > lastSeenResultIndex) {
          const r = progress.lastResult;
          const from = (r.from || 'unknown').substring(0, 30);
          const subject = (r.subject || '').substring(0, 40);
          const folder = r.folder || '';
          
          if (r.moved && folder) {
            appendLog(`[${progress.current}] ${folder}: ${from} | ${subject}`);
          } else if (r.moved) {
            appendLog(`[${progress.current}] MOVED: ${from} | ${subject}`);
          } else if (r.match && r.error) {
            appendLog(`[${progress.current}] ERROR: ${from} | ${subject} - ${r.error}`);
          } else if (r.error) {
            appendLog(`[${progress.current}] ERROR: ${from} | ${subject}`);
          }
          // Silently skip non-matching emails (no log entry)
          
          lastSeenResultIndex = progress.current;
        }
      }
    } catch {
      // Ignore polling errors
    }
    await new Promise(resolve => setTimeout(resolve, 300));
  }
}

// Track if we're currently processing
let isProcessing = false;

/**
 * Handle Ask button click - toggles between Ask and Abort
 */
async function handleAskButtonClick() {
  if (isProcessing) {
    // Currently processing - abort
    await handleAbort();
  } else {
    // Not processing - start
    await handleAskOllama();
  }
}

/**
 * Handle abort action
 */
async function handleAbort() {
  const askBtn = getElement(DOM_ELEMENT_IDS.ASK_BTN);
  if (askBtn) {
    askBtn.disabled = true;
    askBtn.textContent = 'Aborting...';
  }
  
  appendLog('\nAborting...');
  
  try {
    await chrome.runtime.sendMessage({ action: 'abortClassification' });
  } catch (error) {
    appendLog(`Abort error: ${error.message}`);
  }
}

/**
 * Handle Triage Inbox - classify and move emails to P1/P2/P3/FYI
 */
async function handleAskOllama() {
  const iterationsInput = getElement(DOM_ELEMENT_IDS.ITERATIONS_INPUT);
  const maxIterations = parseInt(iterationsInput?.value, 10) || 20;
  
  const token = await getToken();
  if (!token) {
    showStatus(MESSAGES.NO_TOKEN, CSS_CLASSES.ERROR);
    return;
  }
  
  isProcessing = true;
  setAIButtonState('asking');
  clearLog();
  showResultSection(true);
  showPlanSection(false);
  showProgressBar(true);
  updateProgressBar(0, maxIterations);
  showStatus('Triaging inbox...', CSS_CLASSES.PROCESSING);
  
  appendLog(`Triaging up to ${maxIterations} emails...`);
  
  // Start progress polling
  const abortController = new AbortController();
  pollProgress(abortController.signal);
  
  try {
    const response = await chrome.runtime.sendMessage({
      action: 'askOllama',
      prompt: 'triage',
      maxIterations: maxIterations
    });
    
    // Stop progress polling
    abortController.abort();
    showProgressBar(false);
    
    if (!response) {
      throw new Error('No response from background script');
    }
    
    if (response.log && Array.isArray(response.log)) {
      response.log.forEach(line => appendLog(line));
    }
    
    console.log('[Butler Popup] Response:', response);
    
    if (!response.success && !response.aborted) {
      throw new Error(response.error || 'Unknown error');
    }
    
    // Triage is the only action - handle the response
    const moved = response.moved || 0;
    const processed = response.processed || 0;
    const results = response.results || [];
    const folderCounts = response.folderCounts || { '1-Urgent': 0, '2-Action': 0, '3-Attention': 0, '4-FYI': 0, '5-CORP': 0, '6-Zero': 0 };
    
    // Show results in plan list
    const movedResults = results.filter(r => r.moved);
    if (movedResults.length > 0) {
      populatePlanList(movedResults.map(r => ({
        subject: r.subject,
        from: r.from,
        reasoning: `${r.folder}: ${r.reasoning}`
      })));
      showPlanSection(true);
    }
    
    if (response.aborted) {
      setText(DOM_ELEMENT_IDS.RESULT_SUMMARY, `Aborted: ${moved} triaged of ${processed} processed`);
      showStatus(`Aborted after triaging ${moved} emails`, CSS_CLASSES.WARNING);
    } else if (moved === 0) {
      setText(DOM_ELEMENT_IDS.RESULT_SUMMARY, `No emails triaged (${processed} checked)`);
      showStatus('No emails triaged', CSS_CLASSES.WARNING);
    } else {
      const dist = `Urgent:${folderCounts['1-Urgent']} Action:${folderCounts['2-Action']} Attention:${folderCounts['3-Attention']} FYI:${folderCounts['4-FYI']} CORP:${folderCounts['5-CORP']} Zero:${folderCounts['6-Zero']}`;
      setText(DOM_ELEMENT_IDS.RESULT_SUMMARY, `Triaged ${moved} emails (${dist})`);
      showStatus(`Done: triaged ${moved} of ${processed} emails`, CSS_CLASSES.SUCCESS);
    }
    
  } catch (error) {
    abortController.abort();
    showProgressBar(false);
    appendLog(`\nError: ${error.message}`);
    showStatus(MESSAGES.ERROR(error.message), CSS_CLASSES.ERROR);
    setText(DOM_ELEMENT_IDS.RESULT_SUMMARY, 'Request failed');
  } finally {
    isProcessing = false;
    setAIButtonState('idle');
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
  
  // Update Ollama status on load and periodically
  updateOllamaStatus();
  setInterval(updateOllamaStatus, TIMEOUTS.OLLAMA_CHECK_INTERVAL_MS);
  
  // Set up model selector change handler
  const modelSelect = getElement(DOM_ELEMENT_IDS.MODEL_SELECT);
  if (modelSelect) {
    modelSelect.addEventListener("change", handleModelChange);
  }
  
  // Hide result section initially
  showResultSection(false);
  
  // Set initial AI button state
  setAIButtonState('idle');
  
  // Set up event listeners for AI features
  const askBtn = getElement(DOM_ELEMENT_IDS.ASK_BTN);
  if (askBtn) {
    askBtn.addEventListener("click", handleAskButtonClick);
  }
  
  // Set up event listeners for quick actions
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
