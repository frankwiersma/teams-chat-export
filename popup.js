// Popup script - communicates with content script

document.addEventListener('DOMContentLoaded', async () => {
  const statusEl = document.getElementById('status');
  const chatNameEl = document.getElementById('chatName');
  const countEl = document.getElementById('messageCount');
  const scrollBtn = document.getElementById('scrollBtn');
  const refreshBtn = document.getElementById('refreshBtn');
  const exportHtmlBtn = document.getElementById('exportHtml');
  const exportTxtBtn = document.getElementById('exportTxt');
  const exportJsonBtn = document.getElementById('exportJson');
  const exportCsvBtn = document.getElementById('exportCsv');

  function setStatus(msg, type = '') {
    statusEl.textContent = msg;
    statusEl.className = 'status' + (type ? ' ' + type : '');
  }

  function setCount(count) {
    countEl.textContent = count;
  }

  function setChatName(name) {
    chatNameEl.textContent = name || 'Teams Chat';
  }

  async function sendMessage(action, data = {}) {
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

      if (!tab.url || !tab.url.includes('teams.microsoft.com')) {
        setStatus('Please open a Teams chat first', 'error');
        return null;
      }

      const response = await chrome.tabs.sendMessage(tab.id, { action, ...data });
      return response;
    } catch (error) {
      console.error('Error:', error);
      setStatus('Error communicating with Teams page. Refresh the page and try again.', 'error');
      return null;
    }
  }

  // Initialize - get current state
  async function initialize() {
    const response = await sendMessage('getStatus');
    if (response) {
      setChatName(response.chatName);
      setCount(response.messageCount);
      setStatus('Ready to export', 'success');
    }
  }

  // Scroll to top
  scrollBtn.addEventListener('click', async () => {
    scrollBtn.disabled = true;
    setStatus('Scrolling to load messages...');

    const response = await sendMessage('scrollToTop');

    if (response) {
      setCount(response.messageCount);
      setStatus('Loaded ' + response.messageCount + ' messages', 'success');
    }

    scrollBtn.disabled = false;
  });

  // Refresh count
  refreshBtn.addEventListener('click', async () => {
    setStatus('Refreshing...');
    const response = await sendMessage('getStatus');
    if (response) {
      setChatName(response.chatName);
      setCount(response.messageCount);
      setStatus('Ready to export', 'success');
    }
  });

  // Export functions
  async function exportChat(format) {
    setStatus('Exporting as ' + format.toUpperCase() + '...');
    const response = await sendMessage('export', { format });
    if (response && response.success) {
      setStatus('Exported ' + response.messageCount + ' messages!', 'success');
    } else if (response && response.error) {
      setStatus(response.error, 'error');
    }
  }

  exportHtmlBtn.addEventListener('click', () => exportChat('html'));
  exportTxtBtn.addEventListener('click', () => exportChat('txt'));
  exportJsonBtn.addEventListener('click', () => exportChat('json'));
  exportCsvBtn.addEventListener('click', () => exportChat('csv'));

  // Start
  initialize();
});
