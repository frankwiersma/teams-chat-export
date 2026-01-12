// Teams Chat Export - Content Script
// Runs on teams.microsoft.com pages

(function() {
  'use strict';

  // ============ HELPER FUNCTIONS ============

  function getChatName() {
    const chatTitle = document.querySelector('[data-tid="chat-title"]');
    if (chatTitle) {
      return chatTitle.textContent.trim();
    }

    const h2 = document.querySelector('h2');
    if (h2 && h2.textContent.trim().length > 0) {
      return h2.textContent.trim();
    }

    const pageTitle = document.title;
    const match = pageTitle.match(/Chat \| (.+?) \| Microsoft Teams/);
    if (match) {
      return match[1].replace(/^\(\d+\)\s*/, '');
    }

    return 'Teams Chat';
  }

  function findScrollContainer() {
    const selectors = [
      '[data-tid="message-pane-list-container"]',
      '[data-tid="chat-pane-list"]',
      'div[class*="fui-ChatMessageList"]',
      'div[class*="message-list"]'
    ];

    for (const sel of selectors) {
      const found = document.querySelector(sel);
      if (found) return found;
    }

    const allElements = document.querySelectorAll('*');
    for (const elem of allElements) {
      const style = getComputedStyle(elem);
      if ((style.overflowY === 'auto' || style.overflowY === 'scroll') &&
          elem.scrollHeight > elem.clientHeight &&
          elem.querySelector('[data-tid*="message"]')) {
        return elem;
      }
    }

    return null;
  }

  function countMessages() {
    return document.querySelectorAll('[data-tid="chat-pane-message"]').length;
  }

  function formatDateTime(isoString) {
    if (!isoString) return '';
    try {
      const date = new Date(isoString);
      const day = String(date.getDate()).padStart(2, '0');
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const year = date.getFullYear();
      const hours = String(date.getHours()).padStart(2, '0');
      const minutes = String(date.getMinutes()).padStart(2, '0');
      return `${day}/${month}/${year} ${hours}:${minutes}`;
    } catch (e) {
      return isoString;
    }
  }

  function extractAllMessages() {
    const messages = [];
    let currentSender = '';
    let currentTime = '';

    const messageElements = document.querySelectorAll('[data-tid="chat-pane-message"]');

    messageElements.forEach(msgEl => {
      const data = {
        sender: '',
        timestamp: '',
        datetime: '',
        content: ''
      };

      // Get message content
      data.content = msgEl.innerText.trim();
      // Remove reaction text
      data.content = data.content.replace(/\d+\s+(Like|Party popper|Heart|Laugh|Surprised|Sad|Angry)\s+reactions?\.?/gi, '').trim();

      // Walk up DOM to find sender and time
      let parent = msgEl.parentElement;
      let searchDepth = 0;
      let foundSender = false;
      let foundTime = false;

      while (parent && searchDepth < 10 && (!foundSender || !foundTime)) {
        if (!foundSender) {
          const senderEl = parent.querySelector('[data-tid="message-author-name"]');
          if (senderEl) {
            data.sender = senderEl.textContent.trim();
            foundSender = true;
          }
        }

        if (!foundTime) {
          const timeEl = parent.querySelector('time[datetime]');
          if (timeEl) {
            data.datetime = timeEl.getAttribute('datetime');
            data.timestamp = formatDateTime(data.datetime);
            foundTime = true;
          }
        }

        // Check previous siblings
        let prevSibling = parent.previousElementSibling;
        while (prevSibling && (!foundSender || !foundTime)) {
          if (!foundSender) {
            let sibSender = prevSibling.querySelector('[data-tid="message-author-name"]');
            if (!sibSender && prevSibling.getAttribute('data-tid') === 'message-author-name') {
              sibSender = prevSibling;
            }
            if (sibSender) {
              data.sender = sibSender.textContent.trim();
              foundSender = true;
            }
          }
          if (!foundTime) {
            let sibTime = prevSibling.querySelector('time[datetime]');
            if (!sibTime && prevSibling.tagName === 'TIME') {
              sibTime = prevSibling;
            }
            if (sibTime) {
              data.datetime = sibTime.getAttribute('datetime');
              data.timestamp = formatDateTime(data.datetime);
              foundTime = true;
            }
          }
          prevSibling = prevSibling.previousElementSibling;
        }

        parent = parent.parentElement;
        searchDepth++;
      }

      // Use last known values for grouped messages
      if (!data.sender && currentSender) {
        data.sender = currentSender;
      }
      if (!data.timestamp && currentTime) {
        data.timestamp = currentTime;
      }

      if (data.sender) currentSender = data.sender;
      if (data.timestamp) currentTime = data.timestamp;

      if (data.content) {
        messages.push(data);
      }
    });

    return messages;
  }

  // Store collected messages during scroll
  let collectedMessages = [];

  function collectCurrentMessages() {
    const messages = extractAllMessages();
    messages.forEach(msg => {
      // Check if message already exists (by content + sender + time)
      const key = `${msg.sender}|${msg.timestamp}|${msg.content.substring(0, 50)}`;
      const exists = collectedMessages.some(m =>
        `${m.sender}|${m.timestamp}|${m.content.substring(0, 50)}` === key
      );
      if (!exists) {
        collectedMessages.push(msg);
      }
    });
  }

  async function scrollToTop() {
    const container = findScrollContainer();
    if (!container) {
      return { success: false, error: 'Could not find scroll container' };
    }

    // Reset collected messages
    collectedMessages = [];

    // First collect messages at current position (bottom)
    collectCurrentMessages();

    let lastScrollTop = container.scrollTop;
    let samePositionCount = 0;

    // Scroll to top while collecting messages
    for (let i = 0; i < 300; i++) {
      container.scrollTop = 0;
      await new Promise(r => setTimeout(r, 350));

      // Collect messages at this scroll position
      collectCurrentMessages();

      if (container.scrollTop === lastScrollTop) {
        samePositionCount++;
        if (samePositionCount >= 3) break;
      } else {
        samePositionCount = 0;
      }
      lastScrollTop = container.scrollTop;
    }

    // Now scroll back down to collect any messages we might have missed
    await new Promise(r => setTimeout(r, 500));

    for (let i = 0; i < 50; i++) {
      container.scrollTop = container.scrollHeight;
      await new Promise(r => setTimeout(r, 300));
      collectCurrentMessages();

      // Check if we've reached the bottom
      if (container.scrollTop + container.clientHeight >= container.scrollHeight - 10) {
        break;
      }
    }

    // Sort collected messages by datetime
    collectedMessages.sort((a, b) => {
      if (a.datetime && b.datetime) {
        return new Date(a.datetime) - new Date(b.datetime);
      }
      return 0;
    });

    return { success: true, messageCount: collectedMessages.length };
  }

  function getCollectedMessages() {
    return collectedMessages.length > 0 ? collectedMessages : extractAllMessages();
  }

  function escapeHtml(text) {
    if (!text) return '';
    return text.replace(/&/g, '&amp;')
               .replace(/</g, '&lt;')
               .replace(/>/g, '&gt;')
               .replace(/"/g, '&quot;');
  }

  function downloadFile(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function exportMessages(format) {
    // Use collected messages if available, otherwise extract current
    const messages = getCollectedMessages();
    const chatName = getChatName();

    if (messages.length === 0) {
      return { success: false, error: 'No messages found' };
    }

    let content = '';
    const safeFileName = chatName.replace(/[^a-zA-Z0-9\-_ ]/g, '').replace(/\s+/g, '-').substring(0, 50);
    const filename = `${safeFileName}-export.${format}`;
    let mimeType = 'text/plain';

    if (format === 'html') {
      mimeType = 'text/html';
      content = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${escapeHtml(chatName)} - Export</title>`;
      content += '<style>body{font-family:Segoe UI,sans-serif;max-width:900px;margin:0 auto;padding:20px;background:#f5f5f5}';
      content += '.msg{background:#fff;padding:15px;margin:10px 0;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,0.1)}';
      content += '.header{display:flex;justify-content:space-between;margin-bottom:8px;flex-wrap:wrap}';
      content += '.sender{font-weight:bold;color:#6264a7}.time{color:#666;font-size:0.85em}';
      content += '.content{white-space:pre-wrap;line-height:1.5}</style></head><body>';
      content += `<h1 style="color:#6264a7">${escapeHtml(chatName)}</h1>`;
      content += `<p style="color:#666">Exported: ${new Date().toLocaleString()} | ${messages.length} messages</p>`;

      messages.forEach(m => {
        content += '<div class="msg">';
        content += '<div class="header">';
        content += `<span class="sender">${escapeHtml(m.sender)}</span>`;
        content += `<span class="time">${escapeHtml(m.timestamp)}</span>`;
        content += '</div>';
        content += `<div class="content">${escapeHtml(m.content)}</div>`;
        content += '</div>';
      });
      content += '</body></html>';
    }
    else if (format === 'txt') {
      content = `${chatName}\n`;
      content += `Exported: ${new Date().toLocaleString()}\n`;
      content += '='.repeat(50) + '\n\n';
      messages.forEach(m => {
        content += `[${m.timestamp}] ${m.sender}\n`;
        content += `${m.content}\n`;
        content += '\n---\n\n';
      });
    }
    else if (format === 'json') {
      mimeType = 'application/json';
      content = JSON.stringify({
        chatName: chatName,
        exportDate: new Date().toISOString(),
        messageCount: messages.length,
        messages: messages
      }, null, 2);
    }
    else if (format === 'csv') {
      mimeType = 'text/csv';
      content = '\uFEFF'; // BOM for Excel UTF-8
      content += 'Timestamp,Sender,Content\n';
      messages.forEach(m => {
        content += `"${(m.timestamp || '').replace(/"/g, '""')}",`;
        content += `"${(m.sender || '').replace(/"/g, '""')}",`;
        content += `"${(m.content || '').replace(/"/g, '""').replace(/\n/g, ' ')}"\n`;
      });
    }

    downloadFile(content, filename, mimeType);
    return { success: true, messageCount: messages.length };
  }

  // ============ MESSAGE LISTENER ============

  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'getStatus') {
      sendResponse({
        chatName: getChatName(),
        messageCount: countMessages()
      });
    }
    else if (request.action === 'scrollToTop') {
      scrollToTop().then(result => {
        sendResponse(result);
      });
      return true; // Async response
    }
    else if (request.action === 'export') {
      const result = exportMessages(request.format);
      sendResponse(result);
    }
    return false;
  });

  console.log('Teams Chat Export extension loaded');
})();
