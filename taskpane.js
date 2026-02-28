/* ============================================
   Email Feed Add-in — Task Pane Logic
   ============================================ */

// ---------- State ----------
const state = {
  emails: [],
  restToken: null,
  restUrl: null,
  isLoading: false,
  repliedIds: new Set(),
};

// ---------- Icons (SVG strings) ----------
const icons = {
  feed: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>`,
  inbox: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/><path d="M5.45 5.11L2 12v6a2 2 0 002 2h16a2 2 0 002-2v-6l-3.45-6.89A2 2 0 0016.76 4H7.24a2 2 0 00-1.79 1.11z"/></svg>`,
  reply: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 17 4 12 9 7"/><path d="M20 18v-2a4 4 0 00-4-4H4"/></svg>`,
  replyAll: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="7 17 2 12 7 7"/><polyline points="12 17 7 12 12 7"/><path d="M22 18v-2a4 4 0 00-4-4H7"/></svg>`,
  forward: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 17 20 12 15 7"/><path d="M4 18v-2a4 4 0 014-4h12"/></svg>`,
  send: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>`,
  check: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>`,
  refresh: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10"/></svg>`,
  expand: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>`,
  collapse: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"/></svg>`,
  close: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>`,
  warning: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`,
};

// ---------- Office.js Init ----------
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    renderApp();
  }
});

// ---------- Render App Shell ----------
function renderApp() {
  const app = document.getElementById("app");
  app.innerHTML = `
    <div class="app-container">
      <div class="header">
        <div class="header-left">
          <div class="header-icon">${icons.feed}</div>
          <span class="header-title">Email Feed</span>
          <span class="header-count" id="emailCount">0</span>
        </div>
        <button class="btn-icon" id="refreshBtn" title="Reload selected emails">
          ${icons.refresh}
        </button>
      </div>
      <div class="progress-bar" id="progressBar">
        <div class="progress-bar-fill" id="progressFill"></div>
      </div>
      <div class="feed-container" id="feedContainer">
        <div class="empty-state" id="emptyState">
          <div class="empty-state-icon">${icons.inbox}</div>
          <h3>No emails loaded</h3>
          <p>Select one or more emails in your inbox, then click the button below to load them into your feed.</p>
          <button class="btn btn-primary" id="loadBtn">
            ${icons.inbox} Load Selected Emails
          </button>
        </div>
      </div>
    </div>
    <div class="toast" id="toast"></div>
  `;

  document.getElementById("loadBtn").addEventListener("click", loadSelectedEmails);
  document.getElementById("refreshBtn").addEventListener("click", loadSelectedEmails);
}

// ---------- Load Selected Emails ----------
async function loadSelectedEmails() {
  if (state.isLoading) return;
  state.isLoading = true;

  showProgress(true);
  updateProgress(10);

  try {
    // 1. Get REST token
    const token = await getRestToken();
    state.restToken = token;
    state.restUrl = Office.context.mailbox.restUrl;
    updateProgress(20);

    // 2. Get selected items
    const selectedItems = await getSelectedItems();

    if (!selectedItems || selectedItems.length === 0) {
      showToast("No emails selected. Please select emails first.", "error");
      renderEmptyState();
      return;
    }

    updateProgress(30);
    updateEmailCount(selectedItems.length);

    // 3. Fetch full email data for each
    const emails = [];
    for (let i = 0; i < selectedItems.length; i++) {
      const item = selectedItems[i];
      const restId = Office.context.mailbox.convertToRestId(
        item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
      );
      try {
        const email = await fetchEmailDetails(restId);
        emails.push(email);
      } catch (err) {
        console.warn(`Failed to fetch email ${i}:`, err);
      }
      updateProgress(30 + ((i + 1) / selectedItems.length) * 60);
    }

    state.emails = emails;
    updateProgress(100);

    // 4. Render feed
    setTimeout(() => {
      renderFeed(emails);
      showProgress(false);
      showToast(`Loaded ${emails.length} email${emails.length !== 1 ? "s" : ""}`, "success");
    }, 200);

  } catch (err) {
    console.error("Error loading emails:", err);
    showToast("Failed to load emails. " + (err.message || ""), "error");
    showProgress(false);
  } finally {
    state.isLoading = false;
  }
}

// ---------- Office.js Helpers ----------
function getRestToken() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error?.message || "Failed to get token"));
      }
    });
  });
}

function getSelectedItems() {
  return new Promise((resolve, reject) => {
    if (Office.context.mailbox.getSelectedItemsAsync) {
      Office.context.mailbox.getSelectedItemsAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.error?.message || "Failed to get selected items"));
        }
      });
    } else {
      // Fallback: use the single currently-selected item
      const item = Office.context.mailbox.item;
      if (item) {
        resolve([{
          itemId: item.itemId,
          subject: item.subject,
          itemType: Office.MailboxEnums.ItemType.Message,
        }]);
      } else {
        resolve([]);
      }
    }
  });
}

// ---------- REST API Calls ----------
async function fetchEmailDetails(restId) {
  const url = `${state.restUrl}/v2.0/me/messages/${restId}?$select=Id,Subject,From,ToRecipients,CcRecipients,ReceivedDateTime,Body,BodyPreview,IsRead,ConversationId`;

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${state.restToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  }

  return response.json();
}

async function sendReply(restId, comment, replyAll = false) {
  const endpoint = replyAll ? "replyall" : "reply";
  const url = `${state.restUrl}/v2.0/me/messages/${restId}/${endpoint}`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${state.restToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      Comment: comment,
    }),
  });

  if (!response.ok) {
    const errorBody = await response.text();
    throw new Error(`HTTP ${response.status}: ${errorBody}`);
  }

  return true;
}

async function forwardEmail(restId, toRecipients, comment) {
  const url = `${state.restUrl}/v2.0/me/messages/${restId}/forward`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${state.restToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      Comment: comment,
      ToRecipients: toRecipients.map((email) => ({
        EmailAddress: { Address: email.trim() },
      })),
    }),
  });

  if (!response.ok) {
    const errorBody = await response.text();
    throw new Error(`HTTP ${response.status}: ${errorBody}`);
  }

  return true;
}

// ---------- Render Feed ----------
function renderFeed(emails) {
  const container = document.getElementById("feedContainer");

  if (emails.length === 0) {
    renderEmptyState();
    return;
  }

  container.innerHTML = emails.map((email, idx) => renderEmailCard(email, idx)).join("");

  // Attach event listeners
  emails.forEach((email, idx) => {
    // Reply button
    const replyBtn = document.getElementById(`reply-btn-${idx}`);
    if (replyBtn) {
      replyBtn.addEventListener("click", () => toggleComposer(idx, "reply"));
    }

    // Reply All button
    const replyAllBtn = document.getElementById(`replyall-btn-${idx}`);
    if (replyAllBtn) {
      replyAllBtn.addEventListener("click", () => toggleComposer(idx, "replyAll"));
    }

    // Forward button
    const fwdBtn = document.getElementById(`forward-btn-${idx}`);
    if (fwdBtn) {
      fwdBtn.addEventListener("click", () => toggleComposer(idx, "forward"));
    }

    // Expand/collapse body
    const expandBtn = document.getElementById(`expand-btn-${idx}`);
    if (expandBtn) {
      expandBtn.addEventListener("click", () => toggleExpand(idx));
    }

    // Send button
    const sendBtn = document.getElementById(`send-btn-${idx}`);
    if (sendBtn) {
      sendBtn.addEventListener("click", () => handleSend(idx));
    }

    // Cancel button
    const cancelBtn = document.getElementById(`cancel-btn-${idx}`);
    if (cancelBtn) {
      cancelBtn.addEventListener("click", () => closeComposer(idx));
    }

    // Reply type tabs
    ["reply", "replyAll", "forward"].forEach((type) => {
      const tab = document.getElementById(`tab-${type}-${idx}`);
      if (tab) {
        tab.addEventListener("click", () => switchReplyType(idx, type));
      }
    });

    // Keyboard shortcut: Cmd/Ctrl + Enter to send
    const textarea = document.getElementById(`reply-text-${idx}`);
    if (textarea) {
      textarea.addEventListener("keydown", (e) => {
        if ((e.metaKey || e.ctrlKey) && e.key === "Enter") {
          e.preventDefault();
          handleSend(idx);
        }
      });
    }
  });
}

function renderEmailCard(email, idx) {
  const from = email.From?.EmailAddress || {};
  const senderName = from.Name || from.Address || "Unknown";
  const senderEmail = from.Address || "";
  const initials = getInitials(senderName);
  const avatarColor = getAvatarColor(senderEmail);
  const subject = escapeHtml(email.Subject || "(No Subject)");
  const time = formatTime(email.ReceivedDateTime);
  const bodyHtml = sanitizeHtml(email.Body?.Content || email.BodyPreview || "");
  const toList = (email.ToRecipients || [])
    .map((r) => r.EmailAddress?.Name || r.EmailAddress?.Address)
    .join(", ");
  const isReplied = state.repliedIds.has(email.Id);

  return `
    <div class="email-card" id="card-${idx}">
      <div class="email-card-header">
        <div class="sender-avatar" style="background:${avatarColor}20; color:${avatarColor}">
          ${initials}
        </div>
        <div class="email-meta">
          <div class="sender-row">
            <span class="sender-name" title="${escapeHtml(senderEmail)}">${escapeHtml(senderName)}</span>
            <span class="email-time">${time}</span>
          </div>
          <div class="email-subject">${subject}</div>
          ${toList ? `<div class="email-recipients">To: ${escapeHtml(toList)}</div>` : ""}
        </div>
      </div>

      <div class="email-body-wrapper">
        <div class="email-body" id="body-${idx}">${bodyHtml}</div>
        <button class="expand-toggle" id="expand-btn-${idx}">
          Show more ${icons.expand}
        </button>
      </div>

      <div class="email-actions">
        ${isReplied
          ? `<span class="reply-sent-badge">${icons.check} Replied</span>`
          : `
            <button class="btn btn-ghost" id="reply-btn-${idx}" title="Reply">
              ${icons.reply} Reply
            </button>
            <button class="btn btn-ghost" id="replyall-btn-${idx}" title="Reply All">
              ${icons.replyAll} Reply All
            </button>
            <button class="btn btn-ghost" id="forward-btn-${idx}" title="Forward">
              ${icons.forward} Forward
            </button>
          `
        }
      </div>

      <div class="reply-composer" id="composer-${idx}" data-type="reply">
        <div class="reply-type-tabs">
          <button class="reply-type-tab active" id="tab-reply-${idx}">Reply</button>
          <button class="reply-type-tab" id="tab-replyAll-${idx}">Reply All</button>
          <button class="reply-type-tab" id="tab-forward-${idx}">Forward</button>
        </div>
        <div id="forward-to-${idx}" style="display:none; margin-bottom:8px;">
          <input type="text" placeholder="To: email@example.com (comma-separated)"
                 id="forward-input-${idx}"
                 style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:var(--radius-sm);font-family:inherit;font-size:12px;outline:none;"
          />
        </div>
        <div class="reply-input-wrapper">
          <textarea class="reply-textarea" id="reply-text-${idx}"
                    placeholder="Type your reply…"></textarea>
          <div class="reply-toolbar">
            <div class="reply-toolbar-left">
              <span class="keyboard-hint"><kbd>⌘</kbd><kbd>↵</kbd> to send</span>
            </div>
            <div style="display:flex;gap:6px;align-items:center;">
              <button class="btn btn-ghost" id="cancel-btn-${idx}" style="font-size:11px;">Cancel</button>
              <button class="btn btn-send" id="send-btn-${idx}">
                ${icons.send} Send
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  `;
}

// ---------- Composer Actions ----------
function toggleComposer(idx, type) {
  const composer = document.getElementById(`composer-${idx}`);
  const isOpen = composer.classList.contains("open");

  if (isOpen && composer.dataset.type === type) {
    closeComposer(idx);
    return;
  }

  composer.classList.add("open");
  composer.dataset.type = type;
  switchReplyType(idx, type);

  const textarea = document.getElementById(`reply-text-${idx}`);
  setTimeout(() => textarea.focus(), 100);
}

function closeComposer(idx) {
  const composer = document.getElementById(`composer-${idx}`);
  composer.classList.remove("open");
  document.getElementById(`reply-text-${idx}`).value = "";
}

function switchReplyType(idx, type) {
  const composer = document.getElementById(`composer-${idx}`);
  composer.dataset.type = type;

  // Update tab active states
  ["reply", "replyAll", "forward"].forEach((t) => {
    const tab = document.getElementById(`tab-${t}-${idx}`);
    tab.classList.toggle("active", t === type);
  });

  // Show/hide forward To field
  const forwardTo = document.getElementById(`forward-to-${idx}`);
  forwardTo.style.display = type === "forward" ? "block" : "none";

  // Update placeholder
  const textarea = document.getElementById(`reply-text-${idx}`);
  textarea.placeholder = type === "forward"
    ? "Add a message (optional)…"
    : "Type your reply…";
}

async function handleSend(idx) {
  const composer = document.getElementById(`composer-${idx}`);
  const type = composer.dataset.type;
  const textarea = document.getElementById(`reply-text-${idx}`);
  const text = textarea.value.trim();
  const email = state.emails[idx];
  const sendBtn = document.getElementById(`send-btn-${idx}`);

  if (type === "forward") {
    const forwardInput = document.getElementById(`forward-input-${idx}`);
    const recipients = forwardInput.value.split(",").filter((e) => e.trim());
    if (recipients.length === 0) {
      showToast("Please enter at least one recipient.", "error");
      forwardInput.focus();
      return;
    }
  }

  if (!text && type !== "forward") {
    showToast("Please type a reply before sending.", "error");
    textarea.focus();
    return;
  }

  // Disable send button
  sendBtn.disabled = true;
  sendBtn.innerHTML = `<span class="spinner" style="width:14px;height:14px;border-width:2px;"></span>`;

  try {
    // Need to refresh token if it may have expired
    const token = await getRestToken();
    state.restToken = token;

    const restId = email.Id;

    if (type === "forward") {
      const forwardInput = document.getElementById(`forward-input-${idx}`);
      const recipients = forwardInput.value.split(",").filter((e) => e.trim());
      await forwardEmail(restId, recipients, text);
      showToast("Email forwarded!", "success");
    } else {
      const replyAll = type === "replyAll";
      await sendReply(restId, text, replyAll);
      showToast(replyAll ? "Reply all sent!" : "Reply sent!", "success");
    }

    state.repliedIds.add(email.Id);
    closeComposer(idx);

    // Update the action area to show "Replied" badge
    const actionsEl = document.querySelector(`#card-${idx} .email-actions`);
    actionsEl.innerHTML = `<span class="reply-sent-badge">${icons.check} ${type === "forward" ? "Forwarded" : "Replied"}</span>`;

    // Auto-scroll to next card
    const nextCard = document.getElementById(`card-${idx + 1}`);
    if (nextCard) {
      setTimeout(() => {
        nextCard.scrollIntoView({ behavior: "smooth", block: "start" });
      }, 300);
    }

  } catch (err) {
    console.error("Send error:", err);
    showToast("Failed to send. " + (err.message || ""), "error");
  } finally {
    sendBtn.disabled = false;
    sendBtn.innerHTML = `${icons.send} Send`;
  }
}

// ---------- Expand / Collapse ----------
function toggleExpand(idx) {
  const body = document.getElementById(`body-${idx}`);
  const btn = document.getElementById(`expand-btn-${idx}`);
  const isExpanded = body.classList.toggle("expanded");

  btn.innerHTML = isExpanded
    ? `Show less ${icons.collapse}`
    : `Show more ${icons.expand}`;
}

// ---------- Empty State ----------
function renderEmptyState() {
  const container = document.getElementById("feedContainer");
  container.innerHTML = `
    <div class="empty-state">
      <div class="empty-state-icon">${icons.inbox}</div>
      <h3>No emails loaded</h3>
      <p>Select one or more emails in your inbox, then click the button below to load them into your feed.</p>
      <button class="btn btn-primary" id="loadBtn">
        ${icons.inbox} Load Selected Emails
      </button>
    </div>
  `;
  document.getElementById("loadBtn").addEventListener("click", loadSelectedEmails);
}

// ---------- UI Helpers ----------
function updateEmailCount(count) {
  const el = document.getElementById("emailCount");
  el.textContent = count;
  el.classList.toggle("visible", count > 0);
}

function showProgress(show) {
  const bar = document.getElementById("progressBar");
  if (bar) bar.classList.toggle("active", show);
  if (!show) updateProgress(0);
}

function updateProgress(pct) {
  const fill = document.getElementById("progressFill");
  if (fill) fill.style.width = `${Math.min(100, pct)}%`;
}

function showToast(message, type = "info") {
  const toast = document.getElementById("toast");
  const icon = type === "success" ? icons.check
    : type === "error" ? icons.warning
    : "";

  toast.className = `toast ${type}`;
  toast.innerHTML = `${icon} ${escapeHtml(message)}`;

  // Trigger reflow for animation
  void toast.offsetWidth;
  toast.classList.add("visible");

  setTimeout(() => toast.classList.remove("visible"), 3000);
}

// ---------- Utility ----------
function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

function sanitizeHtml(html) {
  // Basic sanitization: strip script tags but preserve email formatting
  return html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "")
    .replace(/on\w+="[^"]*"/gi, "")
    .replace(/on\w+='[^']*'/gi, "");
}

function getInitials(name) {
  if (!name) return "?";
  const parts = name.trim().split(/\s+/);
  if (parts.length >= 2) {
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }
  return name.slice(0, 2).toUpperCase();
}

function getAvatarColor(email) {
  const colors = [
    "#0062cc", "#7c3aed", "#059669", "#d97706",
    "#dc2626", "#0891b2", "#c026d3", "#4f46e5",
    "#16a34a", "#ea580c", "#2563eb", "#9333ea",
  ];
  let hash = 0;
  for (let i = 0; i < email.length; i++) {
    hash = email.charCodeAt(i) + ((hash << 5) - hash);
  }
  return colors[Math.abs(hash) % colors.length];
}

function formatTime(dateStr) {
  if (!dateStr) return "";
  const date = new Date(dateStr);
  const now = new Date();
  const diff = now - date;

  const mins = Math.floor(diff / 60000);
  if (mins < 1) return "Just now";
  if (mins < 60) return `${mins}m ago`;

  const hours = Math.floor(mins / 60);
  if (hours < 24) return `${hours}h ago`;

  const days = Math.floor(hours / 24);
  if (days < 7) return `${days}d ago`;

  return date.toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
    year: date.getFullYear() !== now.getFullYear() ? "numeric" : undefined,
  });
}
