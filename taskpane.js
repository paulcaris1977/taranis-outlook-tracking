/**
 * Taranis Email Tracking - Outlook Add-in
 * Injects a tracking pixel into outgoing emails
 */

// ============================================================================
// Configuration
// ============================================================================

const TRACKING_SERVER = 'https://web-production-205416.up.railway.app';
const SENDER_EMAIL    = 'contact@querceo.com'; // Default sender

// ============================================================================
// State
// ============================================================================

let currentTrackingId  = null;
let currentPixelUrl    = null;
let isTracking         = false;

// ============================================================================
// Office.js Initialization
// ============================================================================

Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    console.log('Taranis Tracking Add-in ready');
    setStatus('idle', 'Ready to track');
  }
});

// ============================================================================
// Main: Add Tracking
// ============================================================================

async function addTracking() {
  setStatus('loading', 'Setting up tracking...');
  setButtonState('loading');

  try {
    // 1. Get email details from Outlook
    const item = Office.context.mailbox.item;

    // Get subject
    const subject = await getSubject(item);

    // Get recipients (To field)
    const recipients = await getRecipients(item);
    const recipientStr = recipients.join(', ') || 'unknown';

    // 2. Call tracking server to create record
    const response = await fetch(`${TRACKING_SERVER}/track/create`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        subject:   subject,
        recipient: recipientStr,
        sender:    SENDER_EMAIL
      })
    });

    if (!response.ok) {
      throw new Error(`Server error: ${response.status}`);
    }

    const data = await response.json();

    if (!data.success) {
      throw new Error(data.error || 'Unknown error');
    }

    // 3. Store tracking info
    currentTrackingId = data.tracking_id;
    currentPixelUrl   = data.pixel_url;
    isTracking        = true;

    // 4. Inject pixel into email body
    await injectPixel(item, data.pixel_html);

    // 5. Update UI
    setStatus('active', 'Tracking active ✓');
    setButtonState('active');
    showInfo(recipientStr, subject, currentTrackingId);

    console.log(`Tracking pixel injected: ${currentTrackingId}`);

  } catch (error) {
    console.error('Tracking error:', error);
    setStatus('error', `Error: ${error.message}`);
    setButtonState('idle');
  }
}

// ============================================================================
// Remove Tracking
// ============================================================================

async function removeTracking() {
  if (!isTracking) return;

  setStatus('loading', 'Removing tracking...');

  try {
    const item = Office.context.mailbox.item;

    // Get current body
    const body = await getBody(item);

    // Remove pixel from body
    const cleanBody = removePixelFromBody(body);

    // Set clean body
    await setBody(item, cleanBody);

    // Reset state
    currentTrackingId = null;
    currentPixelUrl   = null;
    isTracking        = false;

    setStatus('idle', 'Tracking removed');
    setButtonState('idle');
    hideInfo();

  } catch (error) {
    console.error('Remove tracking error:', error);
    setStatus('error', `Error: ${error.message}`);
  }
}

// ============================================================================
// Dashboard
// ============================================================================

function openDashboard() {
  Office.context.ui.openBrowserWindow(`${TRACKING_SERVER}/dashboard`);
}

// ============================================================================
// Office.js Helpers
// ============================================================================

function getSubject(item) {
  return new Promise((resolve, reject) => {
    item.subject.getAsync(function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || '(no subject)');
      } else {
        resolve('(no subject)');
      }
    });
  });
}

function getRecipients(item) {
  return new Promise((resolve) => {
    item.to.getAsync(function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const emails = result.value.map(r => r.emailAddress);
        resolve(emails);
      } else {
        resolve([]);
      }
    });
  });
}

function getBody(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || '');
      } else {
        reject(new Error('Could not get email body'));
      }
    });
  });
}

function setBody(item, htmlContent) {
  return new Promise((resolve, reject) => {
    item.body.setAsync(htmlContent, { coercionType: Office.CoercionType.Html }, function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error('Could not set email body'));
      }
    });
  });
}

async function injectPixel(item, pixelHtml) {
  // Get current body
  const body = await getBody(item);

  // Remove any existing tracking pixel first
  const cleanBody = removePixelFromBody(body);

  // Inject new pixel just before </body> or at the end
  let newBody;
  if (cleanBody.toLowerCase().includes('</body>')) {
    newBody = cleanBody.replace(/<\/body>/i, `${pixelHtml}</body>`);
  } else {
    newBody = cleanBody + pixelHtml;
  }

  // Set new body
  await setBody(item, newBody);
}

function removePixelFromBody(body) {
  // Remove any img tag pointing to our tracking server
  const pattern = new RegExp(
    `<img[^>]*src="${TRACKING_SERVER.replace(/\./g, '\\.')}/open/[^"]*"[^>]*/?>`,
    'gi'
  );
  return body.replace(pattern, '');
}

// ============================================================================
// UI Helpers
// ============================================================================

function setStatus(type, message) {
  const box  = document.getElementById('statusBox');
  const dot  = document.getElementById('statusDot');
  const text = document.getElementById('statusText');

  // Reset classes
  box.className = `status-box ${type}`;

  // Dot color
  const dotColors = {
    idle: 'gray', active: 'green',
    loading: 'blue', error: 'red'
  };
  dot.className = `dot ${dotColors[type] || 'gray'}`;

  text.textContent = message;
}

function setButtonState(state) {
  const btnTrack   = document.getElementById('btnTrack');
  const btnRemove  = document.getElementById('btnRemove');

  if (state === 'active') {
    btnTrack.style.display  = 'none';
    btnRemove.style.display = 'block';
  } else if (state === 'loading') {
    btnTrack.disabled = true;
    btnTrack.textContent = '⏳ Adding tracking...';
  } else {
    // idle or error
    btnTrack.style.display  = 'block';
    btnTrack.disabled       = false;
    btnTrack.textContent    = '🎯 Track this email';
    btnRemove.style.display = 'none';
  }
}

function showInfo(recipient, subject, trackingId) {
  document.getElementById('infoSection').style.display = 'block';
  document.getElementById('infoRecipient').textContent = recipient;
  document.getElementById('infoSubject').textContent   = subject;
  document.getElementById('infoStatus').textContent    = '✅ Pixel injected';
  document.getElementById('infoTrackingId').textContent = `ID: ${trackingId}`;
}

function hideInfo() {
  document.getElementById('infoSection').style.display = 'none';
}
