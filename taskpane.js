/**
 * Taranis Email Tracking - Outlook Add-in
 * Injects a tracking pixel into outgoing emails
 */

// ============================================================================
// Configuration
// ============================================================================

const TRACKING_SERVER = 'https://web-production-205416.up.railway.app';

// ============================================================================
// State
// ============================================================================

let currentTrackingId = null;
let currentPixelUrl   = null;
let isTracking        = false;
let senderEmail       = 'unknown@taranis.com';

// ============================================================================
// Office.js Initialization
// ============================================================================

Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    // Récupérer automatiquement l'email de l'utilisateur connecté
    senderEmail = Office.context.mailbox.userProfile.emailAddress || 'unknown@taranis.com';
    console.log(`Taranis Tracking Add-in ready | Sender: ${senderEmail}`);
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
    const item = Office.context.mailbox.item;

    // Get subject
    const subject = await getSubject(item);

    // Get recipients (To field)
    const recipients = await getRecipients(item);
    const recipientStr = recipients.join(', ') || 'unknown';

    // Call tracking server to create record
    // senderEmail est maintenant automatique via Office.js
    const response = await fetch(`${TRACKING_SERVER}/track/create`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        subject:   subject,
        recipient: recipientStr,
        sender:    senderEmail
      })
    });

    if (!response.ok) {
      throw new Error(`Server error: ${response.status}`);
    }

    const data = await response.json();

    if (!data.success) {
      throw new Error(data.error || 'Unknown error');
    }

    // Store tracking info
    currentTrackingId = data.tracking_id;
    currentPixelUrl   = data.pixel_url;
    isTracking        = true;

    // Inject pixel into email body
    await injectPixel(item, data.pixel_html);

    // Update UI
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

    const body     = await getBody(item);
    const cleanBody = removePixelFromBody(body);
    await setBody(item, cleanBody);

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
  return new Promise((resolve) => {
    item.subject.getAsync(function(result) {
      resolve(result.status === Office.AsyncResultStatus.Succeeded
        ? result.value || '(no subject)'
        : '(no subject)');
    });
  });
}

function getRecipients(item) {
  return new Promise((resolve) => {
    item.to.getAsync(function(result) {
      resolve(result.status === Office.AsyncResultStatus.Succeeded
        ? result.value.map(r => r.emailAddress)
        : []);
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
  const body      = await getBody(item);
  const cleanBody = removePixelFromBody(body);

  const newBody = cleanBody.toLowerCase().includes('</body>')
    ? cleanBody.replace(/<\/body>/i, `${pixelHtml}</body>`)
    : cleanBody + pixelHtml;

  await setBody(item, newBody);
}

function removePixelFromBody(body) {
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

  box.className = `status-box ${type}`;

  const dotColors = { idle: 'gray', active: 'green', loading: 'blue', error: 'red' };
  dot.className = `dot ${dotColors[type] || 'gray'}`;

  text.textContent = message;
}

function setButtonState(state) {
  const btnTrack  = document.getElementById('btnTrack');
  const btnRemove = document.getElementById('btnRemove');

  if (state === 'active') {
    btnTrack.style.display  = 'none';
    btnRemove.style.display = 'block';
  } else if (state === 'loading') {
    btnTrack.disabled     = true;
    btnTrack.textContent  = '⏳ Adding tracking...';
  } else {
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
