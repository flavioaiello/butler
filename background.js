/**
 * Background Service Worker for Butler
 * Monitors HTTP requests and extracts Authorization Bearer tokens
 * Archives emails that have been replied to
 */

const MAX_TOKENS_STORED = 50;
const TOKEN_EXPIRY_MS = 24 * 60 * 60 * 1000; // 24 hours
const MAX_EMAILS_TO_PROCESS = 2000;

// Optional: include Inbox subfolders when scanning/archiving.
// Controlled via runtime message param `includeSubfolders: true`.
const MAX_FOLDERS_TO_PROCESS = 30; // includes Inbox
const MIN_EMAILS_PER_FOLDER = 50;
const MAX_EMAILS_PER_FOLDER = 2000;

// FindItem paging controls
const FINDITEM_PAGE_SIZE = 200;

// Microsoft domains that may have valid tokens
const MICROSOFT_TOKEN_DOMAINS = [
  'graph.microsoft.com',
  'outlook.office.com',
  'outlook.office365.com',
  'outlook.cloud.microsoft',
  'substrate.office.com'
];

// OWA API endpoint
const OWA_API_BASE = 'https://outlook.cloud.microsoft/owa/service.svc';

const REQUEST_TIMEOUT_MS = 60000;

// Track processing state
let processingInProgress = false;

// Debounce token storage
const pendingTokens = new Map();
let storeTokensTimeout = null;
const TOKEN_STORE_DEBOUNCE_MS = 1000;

/**
 * HTML entity decode - handles &lt; &gt; &amp; etc.
 * @param {string} str - HTML encoded string
 * @returns {string} - Decoded string
 */
function htmlDecode(str) {
  if (!str || typeof str !== 'string') return '';
  return str
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

/**
 * Parse References/In-Reply-To header into array of Message-IDs
 * Handles HTML-encoded values and extracts all <message-id> patterns
 * @param {string} headerValue - Raw header value (may be HTML encoded)
 * @returns {string[]} - Array of Message-IDs (with angle brackets)
 */
function parseMessageIdHeader(headerValue) {
  if (!headerValue || typeof headerValue !== 'string') return [];
  
  // Decode HTML entities first
  const decoded = htmlDecode(headerValue);
  
  // Extract all <...> patterns (Message-IDs are enclosed in angle brackets)
  const matches = decoded.match(/<[^>]+>/g);
  return matches || [];
}

/**
 * Validates and sanitizes a Bearer token
 */
function isValidBearerToken(token) {
  if (typeof token !== 'string') return false;
  const MAX_TOKEN_LENGTH = 8192;
  if (token.length === 0 || token.length > MAX_TOKEN_LENGTH) return false;
  const validTokenPattern = /^[A-Za-z0-9\-_\.]+$/;
  return validTokenPattern.test(token);
}

/**
 * Extracts Bearer token from Authorization header
 */
function extractBearerToken(headerValue) {
  if (typeof headerValue !== 'string') return null;
  const BEARER_PREFIX_LENGTH = 7;
  if (!headerValue.toLowerCase().startsWith('bearer ')) return null;
  const token = headerValue.substring(BEARER_PREFIX_LENGTH).trim();
  if (!isValidBearerToken(token)) return null;
  return token;
}

/**
 * Stores a captured token with metadata
 */
async function storeToken(tokenData) {
  try {
    const result = await chrome.storage.local.get(['capturedTokens']);
    let tokens = result.capturedTokens || [];
    
    const existingIndex = tokens.findIndex(t => t.token === tokenData.token);
    if (existingIndex !== -1) {
      tokens[existingIndex].lastSeen = tokenData.timestamp;
      tokens[existingIndex].count = (tokens[existingIndex].count || 1) + 1;
    } else {
      tokens.unshift({
        ...tokenData,
        count: 1,
        lastSeen: tokenData.timestamp
      });
    }
    
    const now = Date.now();
    tokens = tokens.filter(t => (now - t.timestamp) < TOKEN_EXPIRY_MS);
    
    if (tokens.length > MAX_TOKENS_STORED) {
      tokens = tokens.slice(0, MAX_TOKENS_STORED);
    }
    
    await chrome.storage.local.set({ capturedTokens: tokens });
  } catch (error) {
    console.error('Failed to store token:', error.message);
  }
}

/**
 * Fetches inbox messages using OWA FindItem API
 */
async function fetchInboxMessages(token, folderName = 'inbox', maxCount = MAX_EMAILS_TO_PROCESS, offset = 0) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  
  try {
    const requestBody = {
      "__type": "FindItemJsonRequest:#Exchange",
      "Header": {
        "__type": "JsonRequestHeaders:#Exchange",
        "RequestServerVersion": "V2018_01_08",
        "TimeZoneContext": {
          "__type": "TimeZoneContext:#Exchange",
          "TimeZoneDefinition": {
            "__type": "TimeZoneDefinitionType:#Exchange",
            "Id": "UTC"
          }
        }
      },
      "Body": {
        "__type": "FindItemRequest:#Exchange",
        "ItemShape": {
          "__type": "ItemResponseShape:#Exchange",
          "BaseShape": "IdOnly",
          "AdditionalProperties": [
            { "__type": "PropertyUri:#Exchange", "FieldURI": "Subject" },
            { "__type": "PropertyUri:#Exchange", "FieldURI": "DateTimeReceived" },
            { "__type": "PropertyUri:#Exchange", "FieldURI": "From" },
            { "__type": "PropertyUri:#Exchange", "FieldURI": "InternetMessageId" },
            { "__type": "ExtendedPropertyUri:#Exchange", "PropertyTag": "0x1042", "PropertyType": "String" },
            { "__type": "ExtendedPropertyUri:#Exchange", "PropertyTag": "0x1039", "PropertyType": "String" }
          ]
        },
        "ParentFolderIds": [
          {
            "__type": "DistinguishedFolderId:#Exchange",
            "Id": folderName
          }
        ],
        "Traversal": "Shallow",
        "Paging": {
          "__type": "IndexedPageView:#Exchange",
          "BasePoint": "Beginning",
          "Offset": offset,
          "MaxEntriesReturned": maxCount
        },
        "SortOrder": [
          {
            "__type": "SortResults:#Exchange",
            "Order": "Descending",
            "Path": {
              "__type": "PropertyUri:#Exchange",
              "FieldURI": "DateTimeReceived"
            }
          }
        ]
      }
    };
    
    // Use x-owa-urlpostdata format like the real OWA client
    const urlEncodedData = encodeURIComponent(JSON.stringify(requestBody));
    
    const response = await fetch(`${OWA_API_BASE}?action=FindItem&app=Mail`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json',
        'Action': 'FindItem',
        'x-owa-urlpostdata': urlEncodedData
      },
      body: null,
      signal: controller.signal
    });
    
    clearTimeout(timeoutId);
    
    if (!response.ok) {
      const errorText = await response.text().catch(() => 'Unknown error');
      console.error('API Error Response:', response.status, errorText);
      throw new Error(`API error: ${response.status} - ${errorText.substring(0, 200)}`);
    }
    
    const data = await response.json();
    
    // Check for OWA error response
    if (data?.Body?.ResponseMessages?.Items?.[0]?.ResponseClass === 'Error') {
      const errorMsg = data.Body.ResponseMessages.Items[0].MessageText || 'Unknown OWA error';
      throw new Error(`OWA Error: ${errorMsg}`);
    }
    
    // Log the total count from response
    const rootFolder = data?.Body?.ResponseMessages?.Items?.[0]?.RootFolder;
    
    const items = rootFolder?.Items || [];
    
    return items.map(item => {
      // Extract In-Reply-To (0x1042) and References (0x1039) extended properties
      let inReplyToRaw = '';
      let referencesRaw = '';
      
      if (item.ExtendedProperty) {
        for (const prop of item.ExtendedProperty) {
          const propTag = prop.ExtendedFieldURI?.PropertyTag || prop.PropertyTag || '';
          const propTagStr = String(propTag).toLowerCase();
          
          // In-Reply-To: 0x1042 = 4162
          if (propTagStr === '0x1042' || propTagStr === '4162' || propTag === 4162) {
            inReplyToRaw = prop.Value || '';
          }
          // References: 0x1039 = 4153
          if (propTagStr === '0x1039' || propTagStr === '4153' || propTag === 4153) {
            referencesRaw = prop.Value || '';
          }
        }
      }
      
      // Parse and decode the headers
      const inReplyToArr = parseMessageIdHeader(inReplyToRaw);
      const referencesArr = parseMessageIdHeader(referencesRaw);
      
      // Also decode the InternetMessageId (may be HTML encoded)
      const messageId = htmlDecode(item.InternetMessageId || '');
      
      return {
        id: item.ItemId?.Id,
        changeKey: item.ItemId?.ChangeKey,
        subject: item.Subject || '(No Subject)',
        from: item.From?.Mailbox?.EmailAddress || '',
        receivedDateTime: item.DateTimeReceived,
        messageId: messageId,
        inReplyTo: inReplyToArr,
        references: referencesArr,
        sourceFolder: folderName
      };
    });
  } catch (error) {
    clearTimeout(timeoutId);
    throw error;
  }
}

/**
 * Fetch messages for a specific folder ID (non-distinguished).
 */
async function fetchFolderMessagesById(token, folderId, folderLabel, maxCount = MAX_EMAILS_TO_PROCESS, offset = 0) {
  if (typeof folderId !== 'string' || folderId.length === 0) {
    throw new Error('Invalid folderId');
  }

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

  try {
    const requestBody = {
      "__type": "FindItemJsonRequest:#Exchange",
      "Header": {
        "__type": "JsonRequestHeaders:#Exchange",
        "RequestServerVersion": "V2018_01_08",
        "TimeZoneContext": {
          "__type": "TimeZoneContext:#Exchange",
          "TimeZoneDefinition": {
            "__type": "TimeZoneDefinitionType:#Exchange",
            "Id": "UTC"
          }
        }
      },
      "Body": {
        "__type": "FindItemRequest:#Exchange",
        "ItemShape": {
          "__type": "ItemResponseShape:#Exchange",
          "BaseShape": "IdOnly",
          "AdditionalProperties": [
            { "__type": "PropertyUri:#Exchange", "FieldURI": "Subject" },
            { "__type": "PropertyUri:#Exchange", "FieldURI": "DateTimeReceived" },
            { "__type": "PropertyUri:#Exchange", "FieldURI": "From" },
            { "__type": "PropertyUri:#Exchange", "FieldURI": "InternetMessageId" },
            { "__type": "ExtendedPropertyUri:#Exchange", "PropertyTag": "0x1042", "PropertyType": "String" },
            { "__type": "ExtendedPropertyUri:#Exchange", "PropertyTag": "0x1039", "PropertyType": "String" }
          ]
        },
        "ParentFolderIds": [
          {
            "__type": "FolderId:#Exchange",
            "Id": folderId
          }
        ],
        "Traversal": "Shallow",
        "Paging": {
          "__type": "IndexedPageView:#Exchange",
          "BasePoint": "Beginning",
          "Offset": offset,
          "MaxEntriesReturned": maxCount
        },
        "SortOrder": [
          {
            "__type": "SortResults:#Exchange",
            "Order": "Descending",
            "Path": {
              "__type": "PropertyUri:#Exchange",
              "FieldURI": "DateTimeReceived"
            }
          }
        ]
      }
    };

    const urlEncodedData = encodeURIComponent(JSON.stringify(requestBody));

    const response = await fetch(`${OWA_API_BASE}?action=FindItem&app=Mail`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json',
        'Action': 'FindItem',
        'x-owa-urlpostdata': urlEncodedData
      },
      body: null,
      signal: controller.signal
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      const errorText = await response.text().catch(() => 'Unknown error');
      throw new Error(`API error: ${response.status} - ${errorText.substring(0, 200)}`);
    }

    const data = await response.json();

    if (data?.Body?.ResponseMessages?.Items?.[0]?.ResponseClass === 'Error') {
      const errorMsg = data.Body.ResponseMessages.Items[0].MessageText || 'Unknown OWA error';
      throw new Error(`OWA Error: ${errorMsg}`);
    }

    const rootFolder = data?.Body?.ResponseMessages?.Items?.[0]?.RootFolder;
    const items = rootFolder?.Items || [];

    const safeFolderLabel = (typeof folderLabel === 'string' && folderLabel.trim().length > 0)
      ? folderLabel
      : '(Unknown folder)';

    return items.map(item => {
      let inReplyToRaw = '';
      let referencesRaw = '';

      if (item.ExtendedProperty) {
        for (const prop of item.ExtendedProperty) {
          const propTag = prop.ExtendedFieldURI?.PropertyTag || prop.PropertyTag || '';
          const propTagStr = String(propTag).toLowerCase();

          if (propTagStr === '0x1042' || propTagStr === '4162' || propTag === 4162) {
            inReplyToRaw = prop.Value || '';
          }
          if (propTagStr === '0x1039' || propTagStr === '4153' || propTag === 4153) {
            referencesRaw = prop.Value || '';
          }
        }
      }

      const inReplyToArr = parseMessageIdHeader(inReplyToRaw);
      const referencesArr = parseMessageIdHeader(referencesRaw);
      const messageId = htmlDecode(item.InternetMessageId || '');

      return {
        id: item.ItemId?.Id,
        changeKey: item.ItemId?.ChangeKey,
        subject: item.Subject || '(No Subject)',
        from: item.From?.Mailbox?.EmailAddress || '',
        receivedDateTime: item.DateTimeReceived,
        messageId: messageId,
        inReplyTo: inReplyToArr,
        references: referencesArr,
        sourceFolder: safeFolderLabel
      };
    });
  } catch (error) {
    clearTimeout(timeoutId);
    throw error;
  }
}

/**
 * List Inbox subfolders (deep traversal), bounded for safety.
 * Returns array of { id, name }.
 */
async function listInboxSubfolders(token) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

  try {
    const requestData = {
      "__type": "FindFolderJsonRequest:#Exchange",
      "Header": {
        "__type": "JsonRequestHeaders:#Exchange",
        "RequestServerVersion": "Exchange2016",
        "TimeZoneContext": {
          "__type": "TimeZoneContext:#Exchange",
          "TimeZoneDefinition": {
            "__type": "TimeZoneDefinitionType:#Exchange",
            "Id": "UTC"
          }
        }
      },
      "Body": {
        "__type": "FindFolderRequest:#Exchange",
        "FolderShape": {
          "__type": "FolderResponseShape:#Exchange",
          "BaseShape": "Default"
        },
        "ParentFolderIds": [
          {
            "__type": "DistinguishedFolderId:#Exchange",
            "Id": "inbox"
          }
        ],
        "Traversal": "Deep",
        "Paging": {
          "__type": "IndexedPageView:#Exchange",
          "BasePoint": "Beginning",
          "Offset": 0,
          "MaxEntriesReturned": 500
        }
      }
    };

    const urlEncodedData = encodeURIComponent(JSON.stringify(requestData));
    const response = await fetch(`${OWA_API_BASE}?action=FindFolder&app=Mail`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json',
        'Action': 'FindFolder',
        'x-owa-urlpostdata': urlEncodedData
      },
      body: null,
      signal: controller.signal
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      const errorText = await response.text().catch(() => '');
      console.error(`[Butler] FindFolder (inbox subfolders) HTTP ${response.status}: ${errorText.substring(0, 200)}`);
      return [];
    }

    const responseText = await response.text().catch(() => '');
    if (!responseText || responseText.trim().length === 0) {
      console.error('[Butler] FindFolder (inbox subfolders) returned empty response body');
      return [];
    }

    let data;
    try {
      data = JSON.parse(responseText);
    } catch (e) {
      console.error(`[Butler] FindFolder (inbox subfolders) non-JSON response: ${responseText.substring(0, 200)}`);
      return [];
    }

    if (data?.Body?.ErrorCode || data?.Body?.ExceptionName) {
      return [];
    }

    const folders = data?.Body?.ResponseMessages?.Items?.[0]?.RootFolder?.Folders || [];
    const mapped = folders
      .map(f => ({
        id: f?.FolderId?.Id || '',
        name: f?.DisplayName || ''
      }))
      .filter(f => typeof f.id === 'string' && f.id.length > 0);

    return mapped;
  } catch (error) {
    clearTimeout(timeoutId);
    console.error(`[Butler] Error listing inbox subfolders: ${error.message}`);
    return [];
  }
}

function calculatePerFolderLimit(folderCount) {
  const safeFolderCount = Number.isFinite(folderCount) && folderCount > 0 ? folderCount : 1;
  const raw = Math.floor(MAX_EMAILS_TO_PROCESS / safeFolderCount);
  const bounded = Math.min(MAX_EMAILS_PER_FOLDER, Math.max(MIN_EMAILS_PER_FOLDER, raw));
  return bounded;
}

async function fetchMessagesAcrossInboxAndSubfolders(token, includeSubfolders) {
  const folders = [{ kind: 'distinguished', id: 'inbox', name: 'Inbox' }];

  if (includeSubfolders === true) {
    const subfolders = await listInboxSubfolders(token);
    const remainingSlots = Math.max(0, MAX_FOLDERS_TO_PROCESS - 1);
    subfolders.slice(0, remainingSlots).forEach(f => {
      folders.push({ kind: 'folderId', id: f.id, name: f.name || '(Unnamed)' });
    });
  }

  const messages = [];
  const seenItemIds = new Set();
  const folderStats = [];

  const perFolderMax = Math.min(MAX_EMAILS_PER_FOLDER, MAX_EMAILS_TO_PROCESS);
  const maxPagesPerFolder = Math.ceil(perFolderMax / FINDITEM_PAGE_SIZE);

  for (const folder of folders) {
    let fetchedFromFolder = 0;
    let includedFromFolder = 0;
    let pageError = null;
    let offset = 0;
    let pageIndex = 0;

    while (messages.length < MAX_EMAILS_TO_PROCESS && fetchedFromFolder < perFolderMax && pageIndex < maxPagesPerFolder) {
      const remainingForFolder = perFolderMax - fetchedFromFolder;
      const remainingGlobal = MAX_EMAILS_TO_PROCESS - messages.length;
      const pageSize = Math.min(FINDITEM_PAGE_SIZE, remainingForFolder, remainingGlobal);
      if (pageSize <= 0) break;

      let pageMessages;
      try {
        if (folder.kind === 'distinguished') {
          pageMessages = await fetchInboxMessages(token, folder.id, pageSize, offset);
        } else {
          pageMessages = await fetchFolderMessagesById(token, folder.id, folder.name, pageSize, offset);
        }
      } catch (e) {
        pageError = e.message;
        console.error(`[Butler] Failed fetching folder "${folder.name}" (page ${pageIndex}): ${e.message}`);
        break;
      }

      fetchedFromFolder += pageMessages.length;

      for (const msg of pageMessages) {
        if (!msg?.id || typeof msg.id !== 'string') continue;
        if (seenItemIds.has(msg.id)) continue;
        seenItemIds.add(msg.id);
        messages.push(msg);
        includedFromFolder++;
        if (messages.length >= MAX_EMAILS_TO_PROCESS) break;
      }

      // If the server returned fewer than requested, we've hit the end.
      if (pageMessages.length < pageSize) {
        break;
      }

      offset += pageSize;
      pageIndex++;
    }

    if (pageError) {
      folderStats.push({ folder: folder.name, fetched: fetchedFromFolder, included: includedFromFolder, error: pageError });
    } else {
      const truncated = fetchedFromFolder >= perFolderMax || messages.length >= MAX_EMAILS_TO_PROCESS;
      folderStats.push({ folder: folder.name, fetched: fetchedFromFolder, included: includedFromFolder, truncated: truncated });
    }

    if (messages.length >= MAX_EMAILS_TO_PROCESS) {
      return { messages, folderStats };
    }
  }

  return { messages, folderStats };
}

function countByFolder(messages) {
  const counts = new Map();
  for (const msg of messages) {
    const folder = (typeof msg?.sourceFolder === 'string' && msg.sourceFolder.trim().length > 0)
      ? msg.sourceFolder
      : '(Unknown folder)';
    counts.set(folder, (counts.get(folder) || 0) + 1);
  }
  return Array.from(counts.entries())
    .map(([folder, count]) => ({ folder, count }))
    .sort((a, b) => b.count - a.count);
}

function appendFolderStatsToLog(log, title, stats) {
  if (!Array.isArray(stats) || stats.length === 0) return;
  log.push(`${new Date().toISOString()}: ${title}`);
  for (const row of stats) {
    if (!row || typeof row.folder !== 'string') continue;
    if (typeof row.count === 'number') {
      log.push(`${new Date().toISOString()}:   - ${row.folder}: ${row.count}`);
    } else if (typeof row.fetched === 'number' && typeof row.included === 'number') {
      const errSuffix = row.error ? ` (error: ${row.error})` : '';
      log.push(`${new Date().toISOString()}:   - ${row.folder}: fetched ${row.fetched}, included ${row.included}${errSuffix}`);
    }
  }
}

/**
 * Moves a message to a folder using OWA's native format
 * Uses x-owa-urlpostdata header like the real OWA client
 */
async function moveToFolder(token, itemId, changeKey, folderId) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  
  try {
    // Build request matching OWA's exact format
    const requestData = {
      "__type": "MoveItemJsonRequest:#Exchange",
      "Header": {
        "__type": "JsonRequestHeaders:#Exchange",
        "RequestServerVersion": "V2018_01_08",
        "TimeZoneContext": {
          "__type": "TimeZoneContext:#Exchange",
          "TimeZoneDefinition": {
            "__type": "TimeZoneDefinitionType:#Exchange",
            "Id": "UTC"
          }
        }
      },
      "Body": {
        "__type": "MoveItemRequest:#Exchange",
        "ToFolderId": {
          "__type": "TargetFolderId:#Exchange",
          "BaseFolderId": {
            "__type": "FolderId:#Exchange",
            "Id": folderId
          }
        },
        "ItemIds": [
          {
            "__type": "ItemId:#Exchange",
            "Id": itemId
          }
        ],
        "ReturnNewItemIds": true
      }
    };
    
    // URL-encode the request data like OWA does
    const urlEncodedData = encodeURIComponent(JSON.stringify(requestData));
    
    console.log('MoveItem request for item:', itemId.substring(0, 50) + '...');
    
    const response = await fetch(`${OWA_API_BASE}?action=MoveItem&app=Mail`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json',
        'Action': 'MoveItem',
        'x-owa-urlpostdata': urlEncodedData
      },
      body: null,
      signal: controller.signal
    });
    
    clearTimeout(timeoutId);
    
    const responseData = await response.json().catch(() => ({}));
    
    if (!response.ok) {
      const errMsg = responseData?.Body?.ExceptionName || responseData?.Body?.FaultMessage || `Status ${response.status}`;
      throw new Error(`Move failed: ${errMsg}`);
    }
    
    // Check for error in response body
    if (responseData?.Body?.ResponseMessages?.Items?.[0]?.ResponseClass === 'Error') {
      const errMsg = responseData.Body.ResponseMessages.Items[0].MessageText || 'Unknown error';
      throw new Error(`Move error: ${errMsg}`);
    }
    
    return true;
  } catch (error) {
    clearTimeout(timeoutId);
    throw error;
  }
}

/**
 * Finds or creates a folder by name under mailbox root
 * @param {string} token - Bearer token
 * @param {string} folderName - Name of folder to find or create
 * @returns {string} - Folder ID
 */
async function findOrCreateFolder(token, folderName) {
  // Step 1: Try to find the folder
  console.log(`[Butler] Looking for folder "${folderName}"...`);
  let folderId = await findFolderByName(token, folderName);
  
  if (folderId) {
    console.log(`[Butler] Found existing folder "${folderName}" with ID: ${folderId}`);
    return folderId;
  }
  
  // Step 2: Folder doesn't exist, create it
  console.log(`[Butler] Folder "${folderName}" not found, creating...`);
  folderId = await createFolder(token, folderName);
  
  if (folderId) {
    console.log(`[Butler] Created folder "${folderName}" with ID: ${folderId}`);
    return folderId;
  }
  
  // Step 3: Creation failed - maybe folder already exists but we couldn't find it
  // Try finding again with a fresh request
  console.log(`[Butler] Create failed, trying to find folder again...`);
  folderId = await findFolderByName(token, folderName);
  
  if (folderId) {
    console.log(`[Butler] Found folder "${folderName}" on retry with ID: ${folderId}`);
    return folderId;
  }
  
  throw new Error(`Could not find or create "${folderName}" folder`);
}

/**
 * Find a folder by name
 */
async function findFolderByName(token, folderName) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  
  try {
    const requestData = {
      "__type": "FindFolderJsonRequest:#Exchange",
      "Header": {
        "__type": "JsonRequestHeaders:#Exchange",
        "RequestServerVersion": "Exchange2016",
        "TimeZoneContext": {
          "__type": "TimeZoneContext:#Exchange",
          "TimeZoneDefinition": {
            "__type": "TimeZoneDefinitionType:#Exchange",
            "Id": "UTC"
          }
        }
      },
      "Body": {
        "__type": "FindFolderRequest:#Exchange",
        "FolderShape": {
          "__type": "FolderResponseShape:#Exchange",
          "BaseShape": "Default"
        },
        "ParentFolderIds": [
          {
            "__type": "DistinguishedFolderId:#Exchange",
            "Id": "msgfolderroot"
          }
        ],
        "Traversal": "Deep",
        "Paging": {
          "__type": "IndexedPageView:#Exchange",
          "BasePoint": "Beginning",
          "Offset": 0,
          "MaxEntriesReturned": 500
        }
      }
    };
    
    const urlEncodedData = encodeURIComponent(JSON.stringify(requestData));
    
    console.log(`[Butler] FindFolder request for parent: msgfolderroot`);
    
    const response = await fetch(`${OWA_API_BASE}?action=FindFolder&app=Mail`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json',
        'Action': 'FindFolder',
        'x-owa-urlpostdata': urlEncodedData
      },
      body: null,
      signal: controller.signal
    });
    
    clearTimeout(timeoutId);

    if (!response.ok) {
      const errorText = await response.text().catch(() => '');
      console.error(`[Butler] FindFolder HTTP ${response.status}: ${errorText.substring(0, 200)}`);
      return null;
    }

    const responseText = await response.text().catch(() => '');
    if (!responseText || responseText.trim().length === 0) {
      console.error('[Butler] FindFolder returned empty response body');
      return null;
    }

    let data;
    try {
      data = JSON.parse(responseText);
    } catch (e) {
      console.error(`[Butler] FindFolder non-JSON response: ${responseText.substring(0, 200)}`);
      return null;
    }
    
    console.log('[Butler] FindFolder raw response:', JSON.stringify(data, null, 2));
    
    // Check for error response
    if (data?.Body?.ErrorCode || data?.Body?.ExceptionName) {
      console.error(`[Butler] FindFolder error: ${data.Body.ExceptionName || data.Body.ResponseCode}`);
      return null;
    }
    
    // Extract folders from response - handle various response structures
    let folders = data?.Body?.ResponseMessages?.Items?.[0]?.RootFolder?.Folders || [];
    
    // Try alternate paths if empty
    if (!folders || folders.length === 0) {
      folders = data?.Body?.Folders || [];
    }
    if (!folders || folders.length === 0) {
      // Sometimes folders are directly in Items
      const items = data?.Body?.ResponseMessages?.Items || [];
      for (const item of items) {
        if (item.RootFolder?.Folders) {
          folders = item.RootFolder.Folders;
          break;
        }
      }
    }
    
    console.log(`[Butler] Found ${folders.length} folders:`, folders.map(f => f.DisplayName));
    
    // Find the folder by name
    const folder = folders.find(f => f.DisplayName === folderName);
    if (folder) {
      console.log(`[Butler] Found folder "${folderName}" with ID: ${folder.FolderId?.Id}`);
    } else {
      console.log(`[Butler] Folder "${folderName}" not found in list`);
    }
    return folder?.FolderId?.Id || null;
    
  } catch (error) {
    clearTimeout(timeoutId);
    console.error(`[Butler] Error finding folder: ${error.message}`);
    return null;
  }
}

/**
 * Create a new folder
 */
async function createFolder(token, folderName) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  
  try {
    const requestData = {
      "__type": "CreateFolderJsonRequest:#Exchange",
      "Header": {
        "__type": "JsonRequestHeaders:#Exchange",
        "RequestServerVersion": "V2018_01_08",
        "TimeZoneContext": {
          "__type": "TimeZoneContext:#Exchange",
          "TimeZoneDefinition": {
            "__type": "TimeZoneDefinitionType:#Exchange",
            "Id": "UTC"
          }
        }
      },
      "Body": {
        "__type": "CreateFolderRequest:#Exchange",
        "ParentFolderId": {
          "__type": "TargetFolderId:#Exchange",
          "BaseFolderId": {
            "__type": "DistinguishedFolderId:#Exchange",
            "Id": "msgfolderroot"
          }
        },
        "Folders": [
          {
            "__type": "Folder:#Exchange",
            "DisplayName": folderName
          }
        ]
      }
    };
    
    const urlEncodedData = encodeURIComponent(JSON.stringify(requestData));
    
    const response = await fetch(`${OWA_API_BASE}?action=CreateFolder&app=Mail`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json',
        'Action': 'CreateFolder',
        'x-owa-urlpostdata': urlEncodedData
      },
      body: null,
      signal: controller.signal
    });
    
    clearTimeout(timeoutId);
    const data = await response.json();
    
    return data?.Body?.ResponseMessages?.Items?.[0]?.Folders?.[0]?.FolderId?.Id || null;
    
  } catch (error) {
    clearTimeout(timeoutId);
    console.error(`[Butler] Error creating folder: ${error.message}`);
    return null;
  }
}

/**
 * Gets a folder ID by distinguished folder name (e.g., "archive", "inbox")
 * @param {string} token - Bearer token
 * @param {string} distinguishedFolderId - The distinguished folder ID
 * @returns {string} - Folder ID
 */
async function getDistinguishedFolderId(token, distinguishedFolderId) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  
  try {
    const requestData = {
      "__type": "GetFolderJsonRequest:#Exchange",
      "Header": {
        "__type": "JsonRequestHeaders:#Exchange",
        "RequestServerVersion": "V2018_01_08",
        "TimeZoneContext": {
          "__type": "TimeZoneContext:#Exchange",
          "TimeZoneDefinition": {
            "__type": "TimeZoneDefinitionType:#Exchange",
            "Id": "UTC"
          }
        }
      },
      "Body": {
        "__type": "GetFolderRequest:#Exchange",
        "FolderShape": {
          "__type": "FolderResponseShape:#Exchange",
          "BaseShape": "IdOnly"
        },
        "FolderIds": [
          {
            "__type": "DistinguishedFolderId:#Exchange",
            "Id": distinguishedFolderId
          }
        ]
      }
    };
    
    const urlEncodedData = encodeURIComponent(JSON.stringify(requestData));
    
    console.log(`[Butler] GetFolder request for: ${distinguishedFolderId}`);
    
    const response = await fetch(`${OWA_API_BASE}?action=GetFolder&app=Mail`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json',
        'Action': 'GetFolder',
        'x-owa-urlpostdata': urlEncodedData
      },
      body: null,
      signal: controller.signal
    });
    
    clearTimeout(timeoutId);
    
    const data = await response.json();
    console.log(`[Butler] GetFolder response:`, JSON.stringify(data, null, 2));
    
    const folderId = data?.Body?.ResponseMessages?.Items?.[0]?.Folders?.[0]?.FolderId?.Id;
    if (!folderId) {
      throw new Error(`Could not find folder: ${distinguishedFolderId}`);
    }
    
    return folderId;
  } catch (error) {
    clearTimeout(timeoutId);
    throw error;
  }
}

/**
 * Gets a folder ID - tries distinguished folder first, then searches by name
 * @param {string} token - Bearer token
 * @param {string} folderName - Folder name (can be distinguished ID like "archive" or display name like "Duplicates")
 * @param {boolean} isDistinguished - Whether this is a distinguished folder ID
 * @returns {string|null} - Folder ID or null if not found
 */
async function getFolderId(token, folderName, isDistinguished = false) {
  if (isDistinguished) {
    try {
      return await getDistinguishedFolderId(token, folderName);
    } catch (error) {
      console.log(`[Butler] Distinguished folder "${folderName}" not found: ${error.message}`);
      return null;
    }
  }
  
  // Search by display name
  return await findFolderByName(token, folderName);
}

/**
 * Main function to process and archive replied-to emails
 * @param {string} token - Bearer token
 * @param {boolean} dryRun - If true, only scan and report what would be archived
 */
async function processAndArchiveEmails(token, dryRun = false, includeSubfolders = false) {
  if (processingInProgress) {
    return { success: false, error: 'Processing already in progress' };
  }
  
  processingInProgress = true;
  const log = [];
  
  try {
    log.push(`${new Date().toISOString()}: Starting email processing ${dryRun ? '(DRY RUN)' : ''}`);
    
    // Fetch inbox messages (optionally includes subfolders)
    log.push(`${new Date().toISOString()}: Fetching up to ${MAX_EMAILS_TO_PROCESS} messages from inbox${includeSubfolders ? ' (including subfolders)' : ''}`);
    const initialFetch = await fetchMessagesAcrossInboxAndSubfolders(token, includeSubfolders);
    const messages = initialFetch.messages;
    log.push(`${new Date().toISOString()}: Found ${messages.length} messages`);
    appendFolderStatsToLog(log, 'Scan breakdown (unique included per folder):', initialFetch.folderStats);
    
    if (messages.length === 0) {
      log.push(`${new Date().toISOString()}: No messages to process`);
      return { success: true, archivedCount: 0, totalScanned: 0, archivedSubjects: [], log };
    }
    
    // Build a set of all Message-IDs that have replies (from In-Reply-To and References headers)
    const repliedToIds = new Set();
    
    for (const msg of messages) {
      // In-Reply-To contains the Message-ID of the email being replied to
      if (msg.inReplyTo && Array.isArray(msg.inReplyTo)) {
        msg.inReplyTo.forEach(ref => repliedToIds.add(ref));
      }
      // References contains all Message-IDs in the reply chain
      if (msg.references && Array.isArray(msg.references)) {
        msg.references.forEach(ref => repliedToIds.add(ref));
      }
    }
    
    log.push(`${new Date().toISOString()}: Found ${repliedToIds.size} unique message references (In-Reply-To + References)`);
    
    // Find messages whose Message-ID is in the replied-to set
    const toArchive = messages.filter(msg => {
      if (!msg.messageId) return false;
      return repliedToIds.has(msg.messageId);
    });
    
    log.push(`${new Date().toISOString()}: ${toArchive.length} messages have been replied to${dryRun ? '' : ' and will be archived'}`);

    const toArchiveByFolder = countByFolder(toArchive);
    appendFolderStatsToLog(log, 'Replied-to emails by folder:', toArchiveByFolder);
    
    // Count duplicates by Message-ID
    const emailGroups = new Map();
    for (const msg of messages) {
      const messageId = msg.messageId || '';
      if (!messageId) continue;
      
      if (!emailGroups.has(messageId)) {
        emailGroups.set(messageId, []);
      }
      emailGroups.get(messageId).push(msg);
    }
    
    let duplicateCount = 0;
    const duplicateGroups = [];
    for (const [messageId, emails] of emailGroups) {
      if (emails.length > 1) {
        duplicateCount += emails.length - 1;
        duplicateGroups.push({
          messageId: messageId,
          subject: emails[0].subject || '(No Subject)',
          from: emails[0].from || 'unknown',
          count: emails.length
        });
      }
    }
    duplicateGroups.sort((a, b) => b.count - a.count);
    
    log.push(`${new Date().toISOString()}: Found ${duplicateCount} duplicate emails (same Message-ID) in ${duplicateGroups.length} groups`);
    console.log(`[Butler] Found ${duplicateCount} duplicates in ${duplicateGroups.length} groups`);
    
    // DRY RUN: Return the list of emails that would be archived + duplicate info
    if (dryRun) {
      const foundSubjects = toArchive.map(msg => msg.subject);
      return { 
        success: true, 
        foundCount: toArchive.length, 
        totalScanned: messages.length, 
        foundSubjects: foundSubjects,
        duplicateCount: duplicateCount,
        duplicateGroups: duplicateGroups,
        folderStats: initialFetch.folderStats,
        toArchiveByFolder: toArchiveByFolder,
        log: log 
      };
    }
    
    // STEP 1: Move duplicates to "Duplicates" folder (only if folder exists)
    let duplicatesMovedCount = 0;
    let duplicatesErrorCount = 0;
    
    console.log(`[Butler] duplicateCount > 0 check: ${duplicateCount > 0}`);
    
    if (duplicateCount > 0) {
      log.push(`${new Date().toISOString()}: Checking for "Duplicates" folder...`);
      console.log('[Butler] Looking for existing Duplicates folder...');
      
      // Only look for existing folder, don't create it
      const duplicatesFolderId = await getFolderId(token, 'Duplicates', false);
      
      if (duplicatesFolderId) {
        log.push(`${new Date().toISOString()}: Found "Duplicates" folder, moving duplicates...`);
        console.log(`[Butler] Found Duplicates folder ID: ${duplicatesFolderId}`);
        
        // For each group with duplicates, keep the first one and move the rest
        for (const [messageId, emails] of emailGroups) {
          if (emails.length > 1) {
            console.log(`[Butler] Processing duplicate group: ${messageId}, ${emails.length} emails`);
            // Keep the first email (most recent by position), move the rest
            const toMove = emails.slice(1);
            for (const msg of toMove) {
              try {
                console.log(`[Butler] Moving duplicate: ${msg.subject}`);
                await moveToFolder(token, msg.id, msg.changeKey, duplicatesFolderId);
                duplicatesMovedCount++;
                log.push(`${new Date().toISOString()}: Moved duplicate: ${msg.subject}`);
              } catch (error) {
                duplicatesErrorCount++;
                log.push(`${new Date().toISOString()}: Failed to move duplicate "${msg.subject}": ${error.message}`);
                console.error(`[Butler] Failed to move duplicate:`, error);
              }
            }
          }
        }
        log.push(`${new Date().toISOString()}: Moved ${duplicatesMovedCount} duplicates, ${duplicatesErrorCount} errors`);
        console.log(`[Butler] Finished moving duplicates: ${duplicatesMovedCount} moved, ${duplicatesErrorCount} errors`);
      } else {
        log.push(`${new Date().toISOString()}: "Duplicates" folder not found - skipping duplicate handling. Create the folder manually to enable.`);
        console.log('[Butler] Duplicates folder not found, skipping duplicate move');
      }
    }
    
    // STEP 2: Re-fetch and rebuild archive list after duplicates were moved
    log.push(`${new Date().toISOString()}: Re-fetching after duplicate removal...`);
    const afterDedupFetch = await fetchMessagesAcrossInboxAndSubfolders(token, includeSubfolders);
    const messagesAfterDedup = afterDedupFetch.messages;
    log.push(`${new Date().toISOString()}: Found ${messagesAfterDedup.length} messages after deduplication`);
    appendFolderStatsToLog(log, 'Scan breakdown after dedup (unique included per folder):', afterDedupFetch.folderStats);
    
    // Rebuild replied-to set from fresh data
    const repliedToIdsRefresh = new Set();
    for (const msg of messagesAfterDedup) {
      if (msg.inReplyTo && Array.isArray(msg.inReplyTo)) {
        msg.inReplyTo.forEach(ref => repliedToIdsRefresh.add(ref));
      }
      if (msg.references && Array.isArray(msg.references)) {
        msg.references.forEach(ref => repliedToIdsRefresh.add(ref));
      }
    }
    
    // Find messages to archive from fresh list
    const toArchiveRefresh = messagesAfterDedup.filter(msg => {
      if (!msg.messageId) return false;
      return repliedToIdsRefresh.has(msg.messageId);
    });
    
    log.push(`${new Date().toISOString()}: ${toArchiveRefresh.length} emails to archive`);

    const toArchiveRefreshByFolder = countByFolder(toArchiveRefresh);
    appendFolderStatsToLog(log, 'Emails to archive by folder:', toArchiveRefreshByFolder);
    
    // Check if there are still emails to archive
    if (toArchiveRefresh.length === 0) {
      return { 
        success: true, 
        archivedCount: 0,
        duplicatesMovedCount: duplicatesMovedCount,
        totalScanned: messages.length, 
        archivedSubjects: [],
        errors: duplicatesErrorCount, 
        log: log 
      };
    }
    
    // STEP 3: Get the archive folder ID
    log.push(`${new Date().toISOString()}: Getting archive folder ID`);
    let archiveFolderId;
    try {
      archiveFolderId = await getFolderId(token, 'archive', true);
      if (!archiveFolderId) {
        throw new Error('Archive folder not found');
      }
      log.push(`${new Date().toISOString()}: Archive folder found`);
    } catch (error) {
      log.push(`${new Date().toISOString()}: Could not get archive folder: ${error.message}`);
      return { success: false, error: `Could not get archive folder: ${error.message}`, log };
    }
    
    // STEP 4: Archive the replied-to messages
    let archivedCount = 0;
    let errorCount = 0;
    const archivedSubjects = [];
    const archivedByFolderCounter = new Map();
    
    for (const msg of toArchiveRefresh) {
      try {
        await moveToFolder(token, msg.id, msg.changeKey, archiveFolderId);
        archivedCount++;
        archivedSubjects.push(msg.subject);
        const folder = (typeof msg?.sourceFolder === 'string' && msg.sourceFolder.trim().length > 0)
          ? msg.sourceFolder
          : '(Unknown folder)';
        archivedByFolderCounter.set(folder, (archivedByFolderCounter.get(folder) || 0) + 1);
        log.push(`${new Date().toISOString()}: Archived: ${msg.subject}`);
      } catch (error) {
        errorCount++;
        log.push(`${new Date().toISOString()}: Failed to archive "${msg.subject}": ${error.message}`);
      }
    }

    const archivedByFolder = Array.from(archivedByFolderCounter.entries())
      .map(([folder, count]) => ({ folder, count }))
      .sort((a, b) => b.count - a.count);
    appendFolderStatsToLog(log, 'Archived emails by source folder:', archivedByFolder);
    
    log.push(`${new Date().toISOString()}: Done. Archived ${archivedCount} messages, moved ${duplicatesMovedCount} duplicates, ${errorCount + duplicatesErrorCount} total errors`);
    
    // Store the result
    await chrome.storage.local.set({
      lastProcessingResult: {
        timestamp: Date.now(),
        totalMessages: messages.length,
        archived: archivedCount,
        duplicatesMoved: duplicatesMovedCount,
        errors: errorCount + duplicatesErrorCount,
        log: log
      }
    });
    
    return { 
      success: true, 
      archivedCount: archivedCount,
      duplicatesMovedCount: duplicatesMovedCount,
      totalScanned: messages.length, 
      archivedSubjects: archivedSubjects,
      folderStats: afterDedupFetch.folderStats,
      toArchiveByFolder: toArchiveRefreshByFolder,
      archivedByFolder: archivedByFolder,
      errors: errorCount + duplicatesErrorCount, 
      log: log 
    };
    
  } catch (error) {
    log.push(`${new Date().toISOString()}: Error: ${error.message}`);
    return { success: false, error: error.message, log };
  } finally {
    processingInProgress = false;
  }
}

/**
 * Listener for web requests to capture Authorization headers
 */
chrome.webRequest.onBeforeSendHeaders.addListener(
  (details) => {
    if (!details.requestHeaders) return;
    
    for (const header of details.requestHeaders) {
      if (header.name.toLowerCase() === 'authorization' && header.value) {
        const token = extractBearerToken(header.value);
        
        if (token) {
          const url = new URL(details.url);
          
          pendingTokens.set(token, {
            token: token,
            url: details.url,
            domain: url.hostname,
            method: details.method,
            timestamp: Date.now(),
            tabId: details.tabId
          });
          
          if (storeTokensTimeout) clearTimeout(storeTokensTimeout);
          storeTokensTimeout = setTimeout(flushPendingTokens, TOKEN_STORE_DEBOUNCE_MS);
        }
      }
    }
  },
  {
    urls: [
      'https://graph.microsoft.com/*',
      'https://outlook.office.com/*',
      'https://outlook.office365.com/*',
      'https://outlook.cloud.microsoft/*',
      'https://substrate.office.com/*'
    ]
  },
  ['requestHeaders']
);

async function flushPendingTokens() {
  if (pendingTokens.size === 0) return;
  
  const tokensToStore = Array.from(pendingTokens.values());
  pendingTokens.clear();
  storeTokensTimeout = null;
  
  for (const tokenData of tokensToStore) {
    await storeToken(tokenData);
  }
}

/**
 * Validates message sender
 */
function isValidSender(sender) {
  if (!sender || !sender.id) return false;
  if (sender.id !== chrome.runtime.id) return false;
  if (sender.tab) return false;
  return true;
}

/**
 * Message handler
 */
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!isValidSender(sender)) {
    sendResponse({ error: 'Unauthorized sender' });
    return true;
  }
  
  if (message.action === 'getTokens') {
    chrome.storage.local.get(['capturedTokens'], (result) => {
      sendResponse({ tokens: result.capturedTokens || [] });
    });
    return true;
  }
  
  if (message.action === 'clearTokens') {
    chrome.storage.local.set({ capturedTokens: [] }, () => {
      sendResponse({ success: true });
    });
    return true;
  }
  
  if (message.action === 'getLastResult') {
    chrome.storage.local.get(['lastProcessingResult'], (result) => {
      sendResponse({ result: result.lastProcessingResult || null });
    });
    return true;
  }
  
  if (message.action === 'archiveRepliedEmails') {
    const dryRun = message.dryRun === true;
    const includeSubfolders = message.includeSubfolders === true;
    
    chrome.storage.local.get(['capturedTokens'], async (result) => {
      const tokens = result.capturedTokens || [];
      const msTokens = tokens.filter(t => MICROSOFT_TOKEN_DOMAINS.includes(t.domain));
      
      // Prioritize outlook.cloud.microsoft
      msTokens.sort((a, b) => {
        if (a.domain === 'outlook.cloud.microsoft') return -1;
        if (b.domain === 'outlook.cloud.microsoft') return 1;
        return 0;
      });
      
      if (msTokens.length === 0) {
        sendResponse({ success: false, error: 'No Microsoft token found. Visit outlook.office.com first.' });
        return;
      }
      
      // Try tokens until one works
      for (const tokenData of msTokens) {
        const result = await processAndArchiveEmails(tokenData.token, dryRun, includeSubfolders);
        sendResponse(result);
        return;
      }
    });
    return true;
  }
});
