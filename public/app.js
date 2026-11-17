// Outlook Webmail IMAP Frontend Application
(function () {
  'use strict';

  // Intercept fetch to attach active session credentials on Vercel Serverless
  const originalFetch = window.fetch.bind(window);
  window.fetch = function (input, init = {}) {
    const url = typeof input === 'string' ? input : (input && input.url) || '';
    if (url.startsWith('/api/')) {
      const sUser = sessionStorage.getItem('outlook_session_user');
      const sPass = sessionStorage.getItem('outlook_session_pass_b64');
      if (sUser && sPass) {
        const headers = new Headers(init.headers || {});
        if (!headers.has('x-imap-user')) headers.set('x-imap-user', sUser);
        if (!headers.has('x-imap-pass')) headers.set('x-imap-pass', sPass);
        init.headers = headers;
      }
    }
    return originalFetch(input, init);
  };

  // State
  let currentFolder = 'INBOX';
  let currentFilter = 'all';
  let currentSearch = '';
  let messages = [];
  let selectedMessage = null;
  let selectedUids = new Set();
  let folders = [];
  let attachedFiles = [];

  // Signature State
  let sigState = {
    enabled: true,
    autoNewMail: true,
    autoReplyForward: true,
    template: 'executive',
    themeColor: '#0078d4',
    fullName: 'Clodoaldo José da Silva',
    jobTitle: 'Direção / Coordenação',
    company: 'ITEC Ensino - Volta Redonda',
    phone: '',
    email: 'clodoaldo.silva@itecensino.com.br',
    website: 'www.itecensino.com.br',
    logoDataUrl: '',
    html: '<div style="font-family: \'Segoe UI\', Tahoma, sans-serif; font-size: 13px; color: #323130; border-left: 3px solid #0078d4; padding-left: 12px; margin-top: 12px;"><div style="font-size: 15px; font-weight: 700; color: #0078d4;">Clodoaldo José da Silva</div><div style="font-size: 12px; font-weight: 600; color: #605e5c;">Direção / Coordenação | ITEC Ensino - Volta Redonda</div><div style="margin-top: 4px; font-size: 12px; color: #323130;">✉️ <a href="mailto:clodoaldo.silva@itecensino.com.br" style="color: #0078d4; text-decoration: none;">clodoaldo.silva@itecensino.com.br</a> &nbsp;|&nbsp; 🌐 www.itecensino.com.br</div></div>'
  };

  // DOM Elements
  const elFoldersList = document.getElementById('foldersList');
  const elFavoritesList = document.getElementById('favoritesList');
  const elMessageListScroll = document.getElementById('messageListScroll');
  const elCurrentFolderTitle = document.getElementById('currentFolderTitle');
  const elMessageCountLabel = document.getElementById('messageCountLabel');
  const elSearchInput = document.getElementById('searchInput');
  const elBadgeInboxCount = document.getElementById('badgeInboxCount');

  // Reading Pane Elements
  const elNoSelectionView = document.getElementById('noSelectionView');
  const elReadingContentView = document.getElementById('readingContentView');
  const elReadSubject = document.getElementById('readSubject');
  const elReadSenderAvatar = document.getElementById('readSenderAvatar');
  const elReadSenderName = document.getElementById('readSenderName');
  const elReadSenderEmail = document.getElementById('readSenderEmail');
  const elReadRecipients = document.getElementById('readRecipients');
  const elReadDate = document.getElementById('readDate');
  const elReadBody = document.getElementById('readBody');
  const elAttachmentsBar = document.getElementById('attachmentsBar');
  const elAttachmentsChips = document.getElementById('attachmentsChips');

  // Ribbon and Action buttons
  const elBtnNewEmail = document.getElementById('btnNewEmail');
  const elBtnRibbonDelete = document.getElementById('btnRibbonDelete');
  const elBtnRibbonToggleRead = document.getElementById('btnRibbonToggleRead');
  const elBtnRibbonFlag = document.getElementById('btnRibbonFlag');
  const elBtnRibbonMove = document.getElementById('btnRibbonMove');
  const elBtnRibbonRefresh = document.getElementById('btnRibbonRefresh');
  const elBtnDeleteCurrent = document.getElementById('btnDeleteCurrent');
  const elBtnToggleReadCurrent = document.getElementById('btnToggleReadCurrent');
  const elBtnFlagCurrent = document.getElementById('btnFlagCurrent');
  const elBtnReply = document.getElementById('btnReply');
  const elBtnReplyAll = document.getElementById('btnReplyAll');
  const elBtnForward = document.getElementById('btnForward');

  // Compose Modal Elements
  const elComposeModal = document.getElementById('composeModal');
  const elComposeForm = document.getElementById('composeForm');
  const elComposeTo = document.getElementById('composeTo');
  const elComposeCc = document.getElementById('composeCc');
  const elComposeSubject = document.getElementById('composeSubject');
  const elComposeEditor = document.getElementById('composeEditor');
  const elComposeFileInput = document.getElementById('composeFileInput');
  const elComposeAttachedFilesCount = document.getElementById('composeAttachedFilesCount');
  const elBtnCloseCompose = document.getElementById('btnCloseCompose');
  const elBtnDiscardMail = document.getElementById('btnDiscardMail');
  const elBtnSendMail = document.getElementById('btnSendMail');
  const elComposeStatusMsg = document.getElementById('composeStatusMsg');

  // Settings Modal Elements
  const elSettingsModal = document.getElementById('settingsModal');
  const elBtnOpenSettings = document.getElementById('btnOpenSettings');
  const elRailSettingsBtn = document.getElementById('railSettingsBtn');
  const elBtnCloseSettings = document.getElementById('btnCloseSettings');
  const elBtnCancelSettings = document.getElementById('btnCancelSettings');
  const elBtnSaveSettings = document.getElementById('btnSaveSettings');
  const elBtnTestImap = document.getElementById('btnTestImap');
  const elTestResultBox = document.getElementById('testResultBox');
  const elImapStatusBadge = document.getElementById('imapStatusBadge');
  const elImapStatusText = document.getElementById('imapStatusText');

  // Login Screen Elements
  const elLoginOverlay = document.getElementById('loginOverlay');
  const elLoginForm = document.getElementById('loginForm');
  const elLoginEmailInput = document.getElementById('loginEmailInput');
  const elLoginPasswordInput = document.getElementById('loginPasswordInput');
  const elLoginRememberEmail = document.getElementById('loginRememberEmail');
  const elLoginErrorBox = document.getElementById('loginErrorBox');
  const elBtnLoginSubmit = document.getElementById('btnLoginSubmit');
  const elSavedEmailsSection = document.getElementById('savedEmailsSection');
  const elSavedEmailsList = document.getElementById('savedEmailsList');
  const elSavedEmailsDatalist = document.getElementById('savedEmailsDatalist');
  const elActiveUserEmailLabel = document.getElementById('activeUserEmailLabel');
  const elBtnSwitchAccount = document.getElementById('btnSwitchAccount');

  let savedEmails = [];
  let readingPaneLayout = localStorage.getItem('outlook_reading_pane_layout') || 'right';

  function applyReadingPaneLayout(mode) {
    readingPaneLayout = (mode === 'bottom') ? 'bottom' : 'right';
    try {
      localStorage.setItem('outlook_reading_pane_layout', readingPaneLayout);
    } catch (e) {}

    const elContentSplit = document.getElementById('contentSplit');
    if (elContentSplit) {
      if (readingPaneLayout === 'bottom') {
        elContentSplit.classList.add('layout-bottom');
      } else {
        elContentSplit.classList.remove('layout-bottom');
      }
    }

    const optRight = document.getElementById('optLayoutRight');
    const optBottom = document.getElementById('optLayoutBottom');
    if (optRight && optBottom) {
      if (readingPaneLayout === 'bottom') {
        optBottom.classList.add('active');
        optRight.classList.remove('active');
      } else {
        optRight.classList.add('active');
        optBottom.classList.remove('active');
      }
    }

    if (messages.length > 0) {
      renderMessageList();
    }
  }

  // Initialize
  async function init() {
    applyReadingPaneLayout(readingPaneLayout);
    if (elSearchInput) {
      elSearchInput.value = '';
      currentSearch = '';
    }
    setupEventListeners();
    await loadConfig();

    // Show login screen on startup unless already authenticated in this browser tab session
    const isSessionAuth = sessionStorage.getItem('outlook_session_auth') === 'true';
    if (isSessionAuth) {
      if (elLoginOverlay) elLoginOverlay.style.display = 'none';
      await loadFolders();
      await loadMessages();
    } else {
      showLoginScreen();
    }

    // Prevent browser password manager from auto-filling the search box
    setTimeout(() => {
      if (elSearchInput && currentSearch === '') {
        elSearchInput.value = '';
      }
    }, 250);
    setTimeout(() => {
      if (elSearchInput && currentSearch === '') {
        elSearchInput.value = '';
      }
    }, 1000);
  }

  function getMergedSavedEmails(serverList = []) {
    let localList = [];
    try {
      localList = JSON.parse(localStorage.getItem('outlook_saved_emails') || '[]');
    } catch (e) {}
    const combined = [...localList, ...serverList].map(e => (e || '').trim().toLowerCase()).filter(Boolean);
    return Array.from(new Set(combined));
  }

  function saveEmailsLocally(list) {
    try {
      localStorage.setItem('outlook_saved_emails', JSON.stringify(list));
    } catch (e) {}
  }

  function renderSavedEmails() {
    if (!elSavedEmailsList || !elSavedEmailsSection) return;

    if (elSavedEmailsDatalist) {
      elSavedEmailsDatalist.innerHTML = savedEmails.map(em => `<option value="${escapeHtml(em)}"></option>`).join('');
    }

    if (savedEmails.length === 0) {
      elSavedEmailsSection.style.display = 'none';
      return;
    }

    elSavedEmailsSection.style.display = 'block';
    elSavedEmailsList.innerHTML = '';

    const currentTyped = (elLoginEmailInput?.value || '').trim().toLowerCase();

    savedEmails.forEach(email => {
      const item = document.createElement('div');
      item.className = `saved-email-item ${email.toLowerCase() === currentTyped ? 'active' : ''}`;
      const shortName = email.split('@')[0].replace(/[._-]+/g, ' ');

      item.innerHTML = `
        <div class="saved-email-left">
          <div class="saved-email-avatar">${getInitials(shortName)}</div>
          <span class="saved-email-address">${escapeHtml(email)}</span>
        </div>
        <button type="button" class="btn-remove-saved-email" title="Esquecer este e-mail">
          <i class="fa-solid fa-xmark"></i>
        </button>
      `;

      item.onclick = (e) => {
        if (e.target.closest('.btn-remove-saved-email')) {
          e.stopPropagation();
          removeSavedEmail(email);
          return;
        }
        if (elLoginEmailInput) elLoginEmailInput.value = email;
        renderSavedEmails();
        if (elLoginPasswordInput) elLoginPasswordInput.focus();
      };

      elSavedEmailsList.appendChild(item);
    });
  }

  async function removeSavedEmail(emailToRemove) {
    savedEmails = savedEmails.filter(e => e.toLowerCase() !== emailToRemove.toLowerCase());
    saveEmailsLocally(savedEmails);
    renderSavedEmails();
    try {
      await fetch(`/api/auth/saved-email?email=${encodeURIComponent(emailToRemove)}`, { method: 'DELETE' });
    } catch (e) {}
  }

  function showLoginScreen() {
    if (!elLoginOverlay) return;
    elLoginOverlay.style.display = 'flex';
    if (elLoginErrorBox) elLoginErrorBox.style.display = 'none';
    if (elLoginPasswordInput) elLoginPasswordInput.value = '';

    // Pre-fill email with the most recent saved email so user doesn't need to retype it
    const lastEmail = localStorage.getItem('outlook_last_email') || savedEmails[0] || document.getElementById('cfgImapUser')?.value || '';
    if (elLoginEmailInput && lastEmail) {
      elLoginEmailInput.value = lastEmail;
    }

    renderSavedEmails();

    setTimeout(() => {
      if (elLoginEmailInput && !elLoginEmailInput.value) {
        elLoginEmailInput.focus();
      } else if (elLoginPasswordInput) {
        elLoginPasswordInput.focus();
      }
    }, 100);
  }

  async function handleLoginSubmit(e) {
    e.preventDefault();
    const email = (elLoginEmailInput?.value || '').trim().toLowerCase();
    const password = elLoginPasswordInput?.value || '';
    const rememberEmail = elLoginRememberEmail ? elLoginRememberEmail.checked : true;

    if (!email || !password) return;

    if (elBtnLoginSubmit) {
      elBtnLoginSubmit.disabled = true;
      elBtnLoginSubmit.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>Autenticando no servidor IMAP...</span>';
    }
    if (elLoginErrorBox) elLoginErrorBox.style.display = 'none';

    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password, rememberEmail })
      });
      const data = await res.json();

      if (data.success) {
        sessionStorage.setItem('outlook_session_auth', 'true');
        sessionStorage.setItem('outlook_session_user', email);
        try {
          sessionStorage.setItem('outlook_session_pass_b64', btoa(unescape(encodeURIComponent(password))));
        } catch (e) {}

        if (rememberEmail) {
          localStorage.setItem('outlook_last_email', email);
          savedEmails = getMergedSavedEmails([email, ...(data.savedEmails || []), ...savedEmails]);
          saveEmailsLocally(savedEmails);
        }

        const localSig = localStorage.getItem('outlook_sig_' + email);
        if (localSig) {
          try { sigState = { ...sigState, ...JSON.parse(localSig) }; } catch (e) {}
        } else if (data.signature) {
          sigState = { ...sigState, ...data.signature };
        }

        await loadConfig();
        if (elLoginOverlay) elLoginOverlay.style.display = 'none';
        showToast(`Bem-vindo(a), ${email}!`);
        await loadFolders();
        await loadMessages();
      } else {
        if (elLoginErrorBox) {
          elLoginErrorBox.style.display = 'block';
          elLoginErrorBox.innerHTML = `<i class="fa-solid fa-circle-exclamation"></i> ${data.message || 'E-mail ou senha inválidos.'}`;
        }
      }
    } catch (err) {
      if (elLoginErrorBox) {
        elLoginErrorBox.style.display = 'block';
        elLoginErrorBox.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> Erro de conexão: ${err.message}`;
      }
    } finally {
      if (elBtnLoginSubmit) {
        elBtnLoginSubmit.disabled = false;
        elBtnLoginSubmit.innerHTML = '<i class="fa-solid fa-right-to-bracket"></i> <span>Entrar no Webmail</span>';
      }
    }
  }

  // Toast Helper
  function showToast(msg) {
    const toast = document.getElementById('toastNotification');
    const toastMsg = document.getElementById('toastMessage');
    toastMsg.textContent = msg;
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 3500);
  }

  // Date Formatting
  function formatDate(isoStr) {
    if (!isoStr) return '';
    const date = new Date(isoStr);
    if (isNaN(date.getTime())) return '';
    const now = new Date();
    const isToday = date.toDateString() === now.toDateString();
    
    const timeStr = date.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
    if (isToday) return `Hoje ${timeStr}`;

    const diffDays = Math.floor((now - date) / (1000 * 60 * 60 * 24));
    if (diffDays === 1) return `Ontem ${timeStr}`;
    if (diffDays < 7) {
      const weekdays = ['Dom.', 'Seg.', 'Ter.', 'Qua.', 'Qui.', 'Sex.', 'Sáb.'];
      return `${weekdays[date.getDay()]} ${timeStr}`;
    }

    return date.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' });
  }

  // Avatar Initials
  function getInitials(name) {
    if (!name || typeof name !== 'string') return '?';
    const parts = name.trim().split(/\s+/).filter(Boolean);
    if (parts.length === 0) return '?';
    if (parts.length === 1) return (parts[0].substring(0, 2) || '?').toUpperCase();
    const first = (parts[0] && parts[0][0]) || '';
    const last = (parts[parts.length - 1] && parts[parts.length - 1][0]) || '';
    return (first + last).toUpperCase() || '?';
  }

  // Load Folders
  async function loadFolders() {
    try {
      const res = await fetch('/api/imap/folders');
      const data = await res.json();
      if (data.success && data.folders) {
        folders = data.folders;
        renderFolders();
      }
    } catch (err) {
      console.error('Erro ao carregar pastas:', err);
    }
  }

  function getFolderIcon(folder) {
    const name = (folder.name || folder.path).toLowerCase();
    if (name.includes('entrada') || name.includes('inbox')) return 'fa-solid fa-inbox';
    if (name.includes('enviad') || name.includes('sent')) return 'fa-regular fa-paper-plane';
    if (name.includes('rascunh') || name.includes('draft')) return 'fa-regular fa-file-lines';
    if (name.includes('spam') || name.includes('junk') || name.includes('lixo eletr')) return 'fa-solid fa-shield-halved';
    if (name.includes('lixeira') || name.includes('excluíd') || name.includes('trash')) return 'fa-regular fa-trash-can';
    if (name.includes('arquivo') || name.includes('archive')) return 'fa-regular fa-folder-archive';
    return 'fa-regular fa-folder';
  }

  function renderFolders() {
    elFoldersList.innerHTML = '';
    
    // Update inbox badge
    const inbox = folders.find(f => f.path.toUpperCase() === 'INBOX');
    if (inbox && elBadgeInboxCount) {
      if (inbox.unread > 0) {
        elBadgeInboxCount.textContent = inbox.unread;
        elBadgeInboxCount.style.display = 'inline-block';
      } else {
        elBadgeInboxCount.style.display = 'none';
      }
    }

    folders.forEach(folder => {
      const item = document.createElement('a');
      item.href = '#';
      item.className = `folder-item ${folder.path.toLowerCase() === currentFolder.toLowerCase() ? 'active' : ''}`;
      item.setAttribute('data-path', folder.path);
      item.innerHTML = `
        <div class="folder-left">
          <i class="${getFolderIcon(folder)} folder-icon"></i>
          <span>${folder.name}</span>
        </div>
        ${folder.unread > 0 ? `<span class="folder-badge">${folder.unread}</span>` : ''}
      `;
      item.onclick = (e) => {
        e.preventDefault();
        selectFolder(folder.path, folder.name);
      };
      elFoldersList.appendChild(item);
    });

    // Update Favorites selection
    document.querySelectorAll('#favoritesList .folder-item').forEach(item => {
      const favFolder = item.getAttribute('data-folder');
      if (favFolder && favFolder.toLowerCase() === currentFolder.toLowerCase()) {
        item.classList.add('active');
      } else {
        item.classList.remove('active');
      }
    });
  }

  function selectFolder(path, name) {
    currentFolder = path;
    elCurrentFolderTitle.textContent = name || path;
    selectedMessage = null;
    selectedUids.clear();

    // Reset reading view
    const noSel = document.getElementById('noSelectionView');
    const readingView = document.getElementById('readingContentView');
    if (noSel) noSel.setAttribute('style', 'display: flex !important;');
    if (readingView) readingView.setAttribute('style', 'display: none !important;');

    try {
      updateRibbonState();
    } catch (e) {
      console.warn('updateRibbonState:', e);
    }
    renderFolders();
    loadMessages();
  }

  // Load Messages
  async function loadMessages() {
    selectedUids.clear();
    updateRibbonState();
    elMessageListScroll.innerHTML = '<div style="padding: 24px; text-align: center; color: var(--text-secondary);"><i class="fa-solid fa-spinner fa-spin"></i> Carregando emails...</div>';
    
    try {
      const query = new URLSearchParams({
        folder: currentFolder,
        filter: currentFilter,
        search: currentSearch
      });
      const res = await fetch(`/api/imap/messages?${query.toString()}`);
      const data = await res.json();
      
      if (data.success) {
        messages = data.messages || [];
        elMessageCountLabel.textContent = `${messages.length} ${messages.length === 1 ? 'mensagem' : 'mensagens'}`;
        renderMessageList();
      }
    } catch (err) {
      console.error('Erro ao buscar emails:', err);
      elMessageListScroll.innerHTML = '<div style="padding: 24px; text-align: center; color: #d13438;">Erro ao conectar com servidor IMAP.</div>';
    }
  }

  function syncSelectionUI() {
    const count = selectedUids.size;
    const badge = document.getElementById('selectedCountBadge');
    if (badge) {
      if (count > 0) {
        badge.textContent = `${count} selecionado${count > 1 ? 's' : ''}`;
        badge.style.display = 'inline-block';
      } else {
        badge.style.display = 'none';
      }
    }

    const allChecked = messages.length > 0 && messages.every(m => selectedUids.has(String(m.uid)));
    const chkTop = document.getElementById('chkSelectAllMessages');
    const chkCols = document.getElementById('chkSelectAllCols');
    if (chkTop) chkTop.checked = allChecked;
    if (chkCols) chkCols.checked = allChecked;

    document.querySelectorAll('.message-card').forEach(c => {
      const uid = c.getAttribute('data-uid');
      const isChecked = selectedUids.has(String(uid));
      if (isChecked) c.classList.add('checked-row');
      else c.classList.remove('checked-row');
      const chk = c.querySelector('.msg-select-chk');
      if (chk) chk.checked = isChecked;
    });

    updateRibbonState();
  }

  function renderMessageList() {
    if (messages.length === 0) {
      elMessageListScroll.innerHTML = `
        <div style="padding: 40px 20px; text-align: center; color: var(--text-secondary);">
          <i class="fa-regular fa-envelope-open" style="font-size: 36px; margin-bottom: 12px; opacity: 0.5;"></i>
          <p>Nenhuma mensagem nesta pasta.</p>
        </div>
      `;
      syncSelectionUI();
      return;
    }

    const isSentOrDraft = currentFolder.toLowerCase().includes('sent') || 
                          currentFolder.toLowerCase().includes('enviad') || 
                          currentFolder.toLowerCase().includes('draft') || 
                          currentFolder.toLowerCase().includes('rascunh');

    const elColHdrSender = document.getElementById('colHdrSender');
    if (elColHdrSender) {
      elColHdrSender.textContent = isSentOrDraft ? 'Para / E-mail' : 'De / E-mail';
    }

    const isBottomLayout = readingPaneLayout === 'bottom';

    elMessageListScroll.innerHTML = '';
    messages.forEach(msg => {
      const card = document.createElement('div');
      const uidStr = String(msg.uid);
      const isSelected = selectedMessage && String(selectedMessage.uid) === uidStr;
      const isChecked = selectedUids.has(uidStr);
      card.className = `message-card ${!msg.seen ? 'unread' : ''} ${isSelected ? 'selected' : ''} ${isChecked ? 'checked-row' : ''}`;
      card.setAttribute('data-uid', uidStr);
      card.style.cursor = 'pointer';

      // Direct click handler on the card
      card.onclick = (e) => {
        if (e.target.closest('.card-action-btn') || e.target.closest('.msg-select-chk')) return;
        selectMessage(msg);
      };
      
      let displayName = '';
      let displayEmail = '';
      if (isSentOrDraft) {
        if (Array.isArray(msg.to) && msg.to.length > 0) {
          displayName = msg.to.map(t => t.name || t.address).filter(Boolean).join(', ');
          displayEmail = msg.to[0]?.address || '';
        } else {
          displayName = '(Sem destinatário)';
        }
      } else {
        displayName = msg.from ? (msg.from.name || msg.from.address || 'Sem remetente') : 'Sem remetente';
        displayEmail = msg.from?.address || '';
      }

      const timeStr = formatDate(msg.date);
      const subjectStr = msg.subject || '(Sem assunto)';
      const snippetStr = (msg.snippet && msg.snippet.trim() !== subjectStr.trim()) ? msg.snippet : '';

      const actionsHtml = `
        <button type="button" class="card-action-btn btn-quick-delete" onclick="window.quickAction(event, '${msg.uid}', 'delete')" title="Excluir"><i class="fa-regular fa-trash-can"></i></button>
        <button type="button" class="card-action-btn btn-quick-read" onclick="window.quickAction(event, '${msg.uid}', 'read')" title="${msg.seen ? 'Marcar como não lido' : 'Marcar como lido'}">
          <i class="fa-regular ${msg.seen ? 'fa-envelope' : 'fa-envelope-open'}"></i>
        </button>
        <button type="button" class="card-action-btn ${msg.flagged ? 'flagged' : ''} btn-quick-flag" onclick="window.quickAction(event, '${msg.uid}', 'flag')" title="Sinalizar">
          <i class="fa-${msg.flagged ? 'solid' : 'regular'} fa-flag"></i>
        </button>
      `;

      const chkHtml = `<input type="checkbox" class="msg-checkbox msg-select-chk" data-uid="${uidStr}" ${isChecked ? 'checked' : ''} title="Selecionar e-mail">`;

      if (isBottomLayout) {
        // Compact Table-Row Layout with aligned columns: Check | Status/Clip | Email | Assunto | Data | Ações
        const statusIcons = [
          !msg.seen ? '<i class="fa-solid fa-circle" style="font-size: 7px; color: var(--ms-blue);" title="Não lido"></i>' : '<i class="fa-regular fa-envelope-open" style="opacity: 0.45;" title="Lido"></i>',
          msg.hasAttachments ? '<i class="fa-solid fa-paperclip" title="Contém anexo"></i>' : '',
          msg.flagged ? '<i class="fa-solid fa-flag" style="color: #d83b01;" title="Sinalizado"></i>' : ''
        ].filter(Boolean).join('');

        card.innerHTML = `
          <div class="message-row-check">${chkHtml}</div>
          <div class="message-row-status">${statusIcons}</div>
          <div class="message-row-sender" title="${escapeHtml(displayEmail || displayName)}">${escapeHtml(displayName)}</div>
          <div class="message-row-subject" title="${escapeHtml(subjectStr)}">
            <span>${escapeHtml(subjectStr)}</span>
            ${snippetStr ? `<span style="color: var(--text-secondary); font-weight: 400; margin-left: 6px;">— ${escapeHtml(snippetStr)}</span>` : ''}
          </div>
          <div class="message-row-date">${timeStr}</div>
          <div class="message-row-actions">${actionsHtml}</div>
        `;
      } else {
        // Compact Lateral Card Layout with left checkbox
        card.innerHTML = `
          <div class="message-card-lateral-wrap">
            ${chkHtml}
            <div class="message-card-lateral-body">
              <div class="message-card-top" style="margin-bottom: 2px;">
                <span class="message-sender">${escapeHtml(displayName)}</span>
                <span class="message-date">${timeStr}</span>
              </div>
              <div style="display: flex; align-items: center; justify-content: space-between; gap: 6px;">
                <div class="message-subject" style="margin-bottom: 0; flex: 1;">${escapeHtml(subjectStr)}</div>
                <div class="message-indicators" style="flex-shrink: 0;">
                  ${msg.hasAttachments ? '<i class="fa-solid fa-paperclip" title="Contém anexo"></i>' : ''}
                  ${msg.flagged ? '<i class="fa-solid fa-flag" style="color: #d83b01;" title="Sinalizado"></i>' : ''}
                </div>
                <div class="message-actions-hover">${actionsHtml}</div>
              </div>
              ${snippetStr ? `<div class="message-snippet" style="margin-top: 2px;">${escapeHtml(snippetStr)}</div>` : ''}
            </div>
          </div>
        `;
      }

      const chkEl = card.querySelector('.msg-select-chk');
      if (chkEl) {
        chkEl.onchange = (e) => {
          e.stopPropagation();
          if (chkEl.checked) selectedUids.add(uidStr);
          else selectedUids.delete(uidStr);
          syncSelectionUI();
        };
      }

      elMessageListScroll.appendChild(card);
    });
    syncSelectionUI();
  }

  // Global window dispatcher to ensure clicks ALWAYS work
  window.selectMessageByUid = function (uid) {
    let msg = messages.find(m => String(m.uid) === String(uid));
    if (!msg) {
      const card = document.querySelector(`.message-card[data-uid="${uid}"]`);
      if (card) {
        msg = {
          uid: uid,
          subject: card.querySelector('.message-subject')?.textContent || '',
          from: { name: card.querySelector('.message-sender')?.textContent || 'Remetente' },
          date: new Date().toISOString()
        };
      }
    }
    if (msg) {
      selectMessage(msg);
    }
  };

  window.quickAction = function (e, uid, action) {
    if (e && e.stopPropagation) e.stopPropagation();
    const msg = messages.find(m => String(m.uid) === String(uid));
    if (!msg) return;
    if (action === 'delete') deleteEmail(msg.uid);
    else if (action === 'read') toggleReadEmail(msg.uid, !msg.seen);
    else if (action === 'flag') toggleFlagEmail(msg.uid, !msg.flagged);
  };

  // Select Message and Fetch Details
  async function selectMessage(msg) {
    if (!msg) return;

    try {
      selectedMessage = msg;

      // 1. Immediately reveal reading pane and hide "No Selection" placeholder
      const noSel = document.getElementById('noSelectionView');
      const readingView = document.getElementById('readingContentView');
      if (noSel) {
        noSel.style.display = 'none';
        noSel.setAttribute('style', 'display: none !important;');
      }
      if (readingView) {
        readingView.style.display = 'flex';
        readingView.setAttribute('style', 'display: flex !important; flex-direction: column; height: 100%;');
      }

      // 2. Update selection highlight on cards
      document.querySelectorAll('.message-card').forEach(c => {
        if (c.getAttribute('data-uid') === String(msg.uid)) {
          c.classList.add('selected');
          c.classList.remove('unread');
        } else {
          c.classList.remove('selected');
        }
      });

      try {
        updateRibbonState();
      } catch (err) {
        console.warn('updateRibbonState:', err);
      }

      // 3. Fill immediate preview from the message summary
      const senderName = msg.from ? (msg.from.name || msg.from.address || 'Remetente') : 'Remetente';
      const elReadSubj = document.getElementById('readSubject');
      const elReadAvt = document.getElementById('readSenderAvatar');
      const elReadName = document.getElementById('readSenderName');
      const elReadEmail = document.getElementById('readSenderEmail');
      const elReadRecip = document.getElementById('readRecipients');
      const elReadDt = document.getElementById('readDate');
      const elReadBdy = document.getElementById('readBody');

      if (elReadSubj) elReadSubj.textContent = msg.subject || '(Sem assunto)';
      if (elReadAvt) elReadAvt.textContent = getInitials(senderName);
      if (elReadName) elReadName.textContent = senderName;
      if (elReadEmail) elReadEmail.textContent = msg.from?.address ? `<${msg.from.address}>` : '';

      let toStr = 'você';
      if (Array.isArray(msg.to) && msg.to.length > 0) {
        toStr = msg.to.map(t => escapeHtml(t.name || t.address || '')).filter(Boolean).join(', ') || 'você';
      }
      if (elReadRecip) elReadRecip.innerHTML = `<strong>Para:</strong> ${toStr}`;
      if (elReadDt) elReadDt.textContent = formatDate(msg.date);

      if (elReadBdy) {
        elReadBdy.innerHTML = `
          <div style="padding: 50px 20px; text-align: center; color: var(--text-secondary);">
            <i class="fa-solid fa-spinner fa-spin" style="font-size: 32px; color: var(--ms-blue); margin-bottom: 14px;"></i>
            <p style="font-size: 14px; font-weight: 500;">Carregando mensagem do servidor IMAP...</p>
          </div>
        `;
      }

      // Mark seen locally
      msg.seen = true;

      // 4. Fetch full email content and attachments from the server
      const targetFolder = msg.folder || currentFolder || 'INBOX';
      const res = await fetch(`/api/imap/message/${msg.uid}?folder=${encodeURIComponent(targetFolder)}`);
      const data = await res.json();
      
      if (data.success && data.message) {
        selectedMessage = data.message;
        renderReadingDetail(data.message);
      } else {
        if (elReadBdy) {
          elReadBdy.innerHTML = `
            <div style="padding: 24px; color: #d13438; background: #fff4f4; border: 1px solid #fde7e9; border-radius: 4px; margin: 16px;">
              <i class="fa-solid fa-circle-exclamation"></i> <strong>Não foi possível exibir o email:</strong><br>
              ${data.error || 'Erro desconhecido ao baixar dados do servidor.'}
            </div>
          `;
        }
      }
    } catch (err) {
      console.error('Erro em selectMessage:', err);
      const elReadBdy = document.getElementById('readBody');
      if (elReadBdy) {
        elReadBdy.innerHTML = `
          <div style="padding: 24px; color: #d13438;">
            <i class="fa-solid fa-triangle-exclamation"></i> Erro ao renderizar email: ${err.message}
          </div>
        `;
      }
    }
  }

  let activeAttBlobUrl = null;

  function getAttachmentIcon(filename, contentType) {
    const fn = (filename || '').toLowerCase();
    const ct = (contentType || '').toLowerCase();
    if (ct.includes('pdf') || fn.endsWith('.pdf')) return 'fa-regular fa-file-pdf" style="color: #d13438;';
    if (ct.startsWith('image/') || /\.(png|jpg|jpeg|gif|webp|bmp|svg)$/.test(fn)) return 'fa-regular fa-file-image" style="color: #0078d4;';
    if (/\.(xls|xlsx|csv)$/.test(fn)) return 'fa-regular fa-file-excel" style="color: #107c41;';
    if (/\.(doc|docx)$/.test(fn)) return 'fa-regular fa-file-word" style="color: #2b579a;';
    if (/\.(zip|rar|7z)$/.test(fn)) return 'fa-regular fa-file-zipper" style="color: #ca5010;';
    return 'fa-regular fa-file';
  }

  async function fetchAttachmentBlob(msg, att, download = false) {
    const targetFolder = msg.folder || currentFolder || 'INBOX';
    const url = `/api/imap/attachment/${msg.uid}/${att.partId}?folder=${encodeURIComponent(targetFolder)}${download ? '&download=1' : ''}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error('Falha ao carregar anexo do servidor');
    return await res.blob();
  }

  async function downloadAttachment(msg, att) {
    try {
      showToast(`Baixando ${att.filename}...`);
      const blob = await fetchAttachmentBlob(msg, att, true);
      const blobUrl = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = blobUrl;
      a.download = att.filename || 'anexo';
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(blobUrl), 5000);
      showToast(`Download concluído: ${att.filename}`);
    } catch (err) {
      alert('Não foi possível baixar o anexo: ' + err.message);
    }
  }

  async function previewAttachment(msg, att) {
    const modal = document.getElementById('attachmentPreviewModal');
    const titleEl = document.getElementById('attPreviewTitle');
    const sizeEl = document.getElementById('attPreviewSize');
    const bodyEl = document.getElementById('attPreviewBody');
    const btnDl = document.getElementById('btnAttPreviewDownload');
    const btnClose = document.getElementById('btnCloseAttPreview');

    if (!modal || !bodyEl) {
      downloadAttachment(msg, att);
      return;
    }

    if (activeAttBlobUrl) {
      URL.revokeObjectURL(activeAttBlobUrl);
      activeAttBlobUrl = null;
    }

    if (titleEl) titleEl.textContent = att.filename || 'Anexo';
    if (sizeEl) sizeEl.textContent = `(${formatBytes(att.size)})`;
    bodyEl.innerHTML = '<div style="text-align: center; color: var(--text-secondary);"><i class="fa-solid fa-spinner fa-spin" style="font-size: 28px; color: var(--ms-blue); margin-bottom: 10px;"></i><p>Carregando visualização do anexo...</p></div>';
    modal.style.display = 'flex';

    if (btnDl) btnDl.onclick = () => downloadAttachment(msg, att);

    const closeModal = () => {
      modal.style.display = 'none';
      bodyEl.innerHTML = '';
      if (activeAttBlobUrl) {
        URL.revokeObjectURL(activeAttBlobUrl);
        activeAttBlobUrl = null;
      }
    };
    if (btnClose) btnClose.onclick = closeModal;
    modal.onclick = (e) => {
      if (e.target === modal) closeModal();
    };

    try {
      const blob = await fetchAttachmentBlob(msg, att, false);
      const ct = (att.contentType || blob.type || '').toLowerCase();
      const fn = (att.filename || '').toLowerCase();

      // Fix generic octet-stream MIME for PDF/images by extension
      let effectiveBlob = blob;
      if (ct === 'application/octet-stream' || !ct) {
        if (fn.endsWith('.pdf')) effectiveBlob = new Blob([blob], { type: 'application/pdf' });
        else if (/\.(png)$/.test(fn)) effectiveBlob = new Blob([blob], { type: 'image/png' });
        else if (/\.(jpg|jpeg)$/.test(fn)) effectiveBlob = new Blob([blob], { type: 'image/jpeg' });
        else if (fn.endsWith('.gif')) effectiveBlob = new Blob([blob], { type: 'image/gif' });
        else if (fn.endsWith('.webp')) effectiveBlob = new Blob([blob], { type: 'image/webp' });
        else if (/\.(txt|csv|log)$/.test(fn)) effectiveBlob = new Blob([blob], { type: 'text/plain; charset=utf-8' });
      }

      activeAttBlobUrl = URL.createObjectURL(effectiveBlob);
      const isImage = effectiveBlob.type.startsWith('image/') || /\.(png|jpg|jpeg|gif|webp|bmp|svg)$/.test(fn);
      const isPdf = effectiveBlob.type.includes('pdf') || fn.endsWith('.pdf');
      const isText = effectiveBlob.type.startsWith('text/') || /\.(txt|csv|json|xml|log)$/.test(fn);

      if (isImage) {
        bodyEl.innerHTML = `<img src="${activeAttBlobUrl}" alt="${escapeHtml(att.filename)}" style="max-width: 100%; max-height: 100%; object-fit: contain; border-radius: 4px; box-shadow: 0 2px 12px rgba(0,0,0,0.15);">`;
      } else if (isPdf) {
        bodyEl.innerHTML = `<iframe src="${activeAttBlobUrl}" style="width: 100%; height: 100%; border: none; border-radius: 4px; background: #fff;"></iframe>`;
      } else if (isText) {
        const textContent = await effectiveBlob.text();
        bodyEl.innerHTML = `<pre style="width: 100%; height: 100%; margin: 0; padding: 16px; background: #fff; border-radius: 4px; overflow: auto; font-family: Consolas, monospace; font-size: 13px; white-space: pre-wrap;">${escapeHtml(textContent)}</pre>`;
      } else {
        bodyEl.innerHTML = `
          <div style="background: #fff; padding: 32px 40px; border-radius: 8px; text-align: center; box-shadow: 0 2px 10px rgba(0,0,0,0.08); max-width: 420px;">
            <i class="${getAttachmentIcon(att.filename, att.contentType)}" style="font-size: 48px; margin-bottom: 14px; display: block;"></i>
            <h4 style="font-size: 15px; margin-bottom: 6px; word-break: break-word;">${escapeHtml(att.filename)}</h4>
            <p style="font-size: 12px; color: var(--text-secondary); margin-bottom: 18px;">
              Tamanho: ${formatBytes(att.size)}<br>
              Pré-visualização direta indisponível para este formato no navegador.
            </p>
            <button type="button" class="action-btn primary" id="btnAttFallbackDownload" style="margin: 0 auto;">
              <i class="fa-solid fa-download"></i>
              <span>Baixar no Computador</span>
            </button>
          </div>
        `;
        const btnFallback = document.getElementById('btnAttFallbackDownload');
        if (btnFallback) btnFallback.onclick = () => downloadAttachment(msg, att);
      }
    } catch (err) {
      bodyEl.innerHTML = `<div style="color: #a80000; padding: 20px;"><i class="fa-solid fa-circle-exclamation"></i> Erro ao visualizar anexo: ${escapeHtml(err.message)}</div>`;
    }
  }

  function renderReadingDetail(msg) {
    if (!msg) return;

    // Attachments
    if (msg.attachments && msg.attachments.length > 0) {
      elAttachmentsBar.style.display = 'flex';
      elAttachmentsChips.innerHTML = '';
      msg.attachments.forEach(att => {
        const chip = document.createElement('div');
        chip.className = 'attachment-chip';
        chip.title = 'Clique para visualizar ou baixe no ícone ao lado';
        chip.innerHTML = `
          <i class="${getAttachmentIcon(att.filename, att.contentType)}"></i>
          <span><strong>${escapeHtml(att.filename)}</strong> (${formatBytes(att.size)})</span>
          <button type="button" class="card-action-btn btn-att-dl" title="Baixar no computador" style="margin-left: 4px; color: var(--ms-blue);">
            <i class="fa-solid fa-download"></i>
          </button>
        `;
        chip.onclick = (e) => {
          if (e.target.closest('.btn-att-dl')) {
            e.stopPropagation();
            downloadAttachment(msg, att);
            return;
          }
          previewAttachment(msg, att);
        };
        elAttachmentsChips.appendChild(chip);
      });
    } else {
      elAttachmentsBar.style.display = 'none';
    }

    // Body rendering
    if (!elReadBody) return;

    if (msg.html) {
      const iframe = document.createElement('iframe');
      iframe.style.width = '100%';
      iframe.style.height = '100%';
      iframe.style.border = 'none';
      iframe.style.backgroundColor = '#ffffff';
      iframe.srcdoc = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="utf-8">
          <style>
            body { 
              font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif; 
              font-size: 14px; 
              color: #201f1e; 
              margin: 0; 
              padding: 16px; 
              line-height: 1.6; 
              word-break: break-word;
            }
            a { color: #0078d4; text-decoration: underline; }
            img { max-width: 100%; height: auto; }
            table { max-width: 100%; border-collapse: collapse; }
          </style>
        </head>
        <body>${msg.html}</body>
        </html>
      `;
      elReadBody.innerHTML = '';
      elReadBody.appendChild(iframe);
    } else {
      elReadBody.innerHTML = `
        <div style="font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 14px; line-height: 1.6; white-space: pre-wrap; padding: 16px; color: #201f1e;">
          ${escapeHtml(msg.text || '(Mensagem sem conteúdo)')}
        </div>
      `;
    }
  }

  function updateRibbonState() {
    const bulkCount = selectedUids.size;
    const hasSelection = bulkCount > 0 || !!selectedMessage;

    if (elBtnRibbonDelete) {
      elBtnRibbonDelete.disabled = !hasSelection;
      const delSpan = elBtnRibbonDelete.querySelector('span');
      if (delSpan) {
        delSpan.textContent = bulkCount > 0 ? `Excluir (${bulkCount})` : 'Excluir';
      }
    }
    if (elBtnRibbonToggleRead) {
      elBtnRibbonToggleRead.disabled = !hasSelection;
      const span = elBtnRibbonToggleRead.querySelector('span');
      if (span) {
        if (bulkCount > 0) {
          const anyUnread = messages.some(m => selectedUids.has(String(m.uid)) && !m.seen);
          span.textContent = anyUnread ? `Marcar como lidos (${bulkCount})` : `Marcar não lidos (${bulkCount})`;
        } else if (selectedMessage) {
          span.textContent = selectedMessage.seen ? 'Marcar como não lido' : 'Marcar como lido';
        }
      }
    }
    if (elBtnRibbonFlag) elBtnRibbonFlag.disabled = !hasSelection;
    if (elBtnRibbonMove) elBtnRibbonMove.disabled = !hasSelection;
  }

  // Flags & Actions
  async function toggleReadEmail(uid, seen) {
    try {
      await fetch('/api/imap/flags', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ folder: currentFolder, uid, seen })
      });
      const target = messages.find(m => String(m.uid) === String(uid));
      if (target) target.seen = seen;
      renderMessageList();
      updateRibbonState();
      showToast(seen ? 'Marcado como lido' : 'Marcado como não lido');
    } catch (e) {
      console.error(e);
    }
  }

  async function toggleFlagEmail(uid, flagged) {
    try {
      await fetch('/api/imap/flags', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ folder: currentFolder, uid, flagged })
      });
      const target = messages.find(m => String(m.uid) === String(uid));
      if (target) target.flagged = flagged;
      renderMessageList();
      showToast(flagged ? 'Email sinalizado' : 'Sinalização removida');
    } catch (e) {
      console.error(e);
    }
  }

  async function deleteEmail(uid) {
    if (!confirm('Deseja realmente excluir este email?')) return;
    try {
      await fetch(`/api/imap/message/${uid}?folder=${encodeURIComponent(currentFolder)}`, { method: 'DELETE' });
      messages = messages.filter(m => String(m.uid) !== String(uid));
      selectedUids.delete(String(uid));
      if (selectedMessage && String(selectedMessage.uid) === String(uid)) {
        selectedMessage = null;
        elReadingContentView.style.display = 'none';
        elNoSelectionView.style.display = 'flex';
      }
      renderMessageList();
      updateRibbonState();
      showToast('Email excluído com sucesso');
    } catch (e) {
      console.error(e);
    }
  }

  async function bulkDeleteEmails(uids) {
    const list = Array.from(uids).map(String);
    if (list.length === 0) return;
    if (!confirm(`Deseja realmente excluir ${list.length} ${list.length === 1 ? 'e-mail selecionado' : 'e-mails selecionados'}?`)) return;

    try {
      showToast(`Excluindo ${list.length} e-mail(s)...`);
      const res = await fetch('/api/imap/bulk-delete', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ folder: currentFolder, uids: list })
      });
      const data = await res.json();
      if (data.success) {
        const uidSet = new Set(list);
        messages = messages.filter(m => !uidSet.has(String(m.uid)));
        selectedUids.clear();
        if (selectedMessage && uidSet.has(String(selectedMessage.uid))) {
          selectedMessage = null;
          if (elReadingContentView) elReadingContentView.setAttribute('style', 'display: none !important;');
          if (elNoSelectionView) elNoSelectionView.setAttribute('style', 'display: flex !important;');
        }
        elMessageCountLabel.textContent = `${messages.length} ${messages.length === 1 ? 'mensagem' : 'mensagens'}`;
        renderMessageList();
        updateRibbonState();
        showToast(`${list.length} ${list.length === 1 ? 'e-mail excluído' : 'e-mails excluídos'} com sucesso!`);
      } else {
        alert(data.error || 'Falha ao excluir e-mails selecionados');
      }
    } catch (e) {
      console.error(e);
      alert('Erro ao excluir e-mails: ' + e.message);
    }
  }

  async function bulkSetFlags(uids, { seen, flagged }) {
    const list = Array.from(uids).map(String);
    if (list.length === 0) return;
    try {
      const res = await fetch('/api/imap/bulk-flags', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ folder: currentFolder, uids: list, seen, flagged })
      });
      const data = await res.json();
      if (data.success) {
        const uidSet = new Set(list);
        messages.forEach(m => {
          if (uidSet.has(String(m.uid))) {
            if (seen !== undefined) m.seen = seen;
            if (flagged !== undefined) m.flagged = flagged;
          }
        });
        renderMessageList();
        updateRibbonState();
        if (seen !== undefined) {
          showToast(seen ? `${list.length} marcado(s) como lido(s)` : `${list.length} marcado(s) como não lido(s)`);
        } else if (flagged !== undefined) {
          showToast(flagged ? `${list.length} sinalizado(s)` : `Sinalização removida de ${list.length}`);
        }
      }
    } catch (e) {
      console.error(e);
    }
  }

  // Signature Builder Helpers
  function buildSignatureHtml(s) {
    const color = s.themeColor || '#0078d4';
    const name = escapeHtml(s.fullName || 'Clodoaldo José da Silva');
    const role = escapeHtml(s.jobTitle || '');
    const company = escapeHtml(s.company || '');
    const phone = escapeHtml(s.phone || '');
    const email = escapeHtml(s.email || '');
    const website = escapeHtml(s.website || '');
    const logo = s.logoDataUrl || '';

    const contactParts = [];
    if (phone) contactParts.push(`<span>📞 ${phone}</span>`);
    if (email) contactParts.push(`<span>✉️ <a href="mailto:${email}" style="color: ${color}; text-decoration: none;">${email}</a></span>`);
    if (website) {
      const href = website.startsWith('http') ? website : `https://${website}`;
      contactParts.push(`<span>🌐 <a href="${href}" target="_blank" style="color: ${color}; text-decoration: none;">${website}</a></span>`);
    }
    const contactsHtml = contactParts.join(' &nbsp;|&nbsp; ');

    if (s.template === 'corporate') {
      const initials = getInitials(s.fullName || 'CS');
      const leftVisual = logo
        ? `<img src="${logo}" alt="Logo" style="max-height: 64px; max-width: 110px; border-radius: 6px; display: block; object-fit: contain;">`
        : `<div style="width: 52px; height: 52px; border-radius: 50%; background: ${color}; color: #ffffff; font-size: 18px; font-weight: 700; display: flex; align-items: center; justify-content: center; text-align: center; line-height: 52px;">${initials}</div>`;

      return `<table cellpadding="0" cellspacing="0" border="0" style="font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 13px; color: #323130; margin-top: 12px; border-collapse: collapse;">
        <tr>
          <td style="padding-right: 14px; vertical-align: middle;">
            ${leftVisual}
          </td>
          <td style="border-left: 3px solid ${color}; padding-left: 14px; vertical-align: middle;">
            <div style="font-size: 15px; font-weight: 700; color: ${color}; line-height: 1.2;">${name}</div>
            ${(role || company) ? `<div style="font-size: 12px; font-weight: 600; color: #605e5c; margin-top: 2px;">${[role, company].filter(Boolean).join(' • ')}</div>` : ''}
            ${contactsHtml ? `<div style="font-size: 12px; color: #323130; margin-top: 6px;">${contactsHtml}</div>` : ''}
          </td>
        </tr>
      </table>`;
    }

    if (s.template === 'minimal') {
      return `<div style="font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 13px; color: #323130; line-height: 1.5; margin-top: 12px;">
        <div style="color: #605e5c;">Atenciosamente,</div>
        ${logo ? `<div style="margin: 6px 0;"><img src="${logo}" alt="Logo" style="max-height: 48px; max-width: 120px;"></div>` : ''}
        <div style="font-weight: 700; font-size: 14px; color: ${color}; margin-top: 4px;">${name}</div>
        ${(role || company) ? `<div style="font-size: 12px; color: #323130;">${[role, company].filter(Boolean).join(' - ')}</div>` : ''}
        ${contactsHtml ? `<div style="font-size: 11px; color: #605e5c; margin-top: 3px;">${contactsHtml}</div>` : ''}
      </div>`;
    }

    // Default: 'executive'
    return `<div style="font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 13px; color: #323130; border-left: 3px solid ${color}; padding-left: 12px; margin-top: 12px;">
      ${logo ? `<div style="margin-bottom: 6px;"><img src="${logo}" alt="Logo" style="max-height: 54px; max-width: 130px; display: block;"></div>` : ''}
      <div style="font-size: 15px; font-weight: 700; color: ${color};">${name}</div>
      ${(role || company) ? `<div style="font-size: 12px; font-weight: 600; color: #605e5c; margin-top: 1px;">${[role, company].filter(Boolean).join(' | ')}</div>` : ''}
      ${contactsHtml ? `<div style="margin-top: 5px; font-size: 12px; color: #323130;">${contactsHtml}</div>` : ''}
    </div>`;
  }

  function populateSignatureBuilderUI() {
    const elFullName = document.getElementById('sigFullName');
    const elJobTitle = document.getElementById('sigJobTitle');
    const elCompany = document.getElementById('sigCompany');
    const elPhone = document.getElementById('sigPhone');
    const elEmail = document.getElementById('sigEmail');
    const elWebsite = document.getElementById('sigWebsite');
    const elCustomColor = document.getElementById('sigCustomColor');
    const elSigEditor = document.getElementById('sigEditor');
    const elAutoNew = document.getElementById('sigAutoNewMail');
    const elAutoReply = document.getElementById('sigAutoReplyForward');
    const elBtnRemoveLogo = document.getElementById('btnRemoveSigLogo');
    const elLogoStatus = document.getElementById('sigLogoStatus');

    if (elFullName) elFullName.value = sigState.fullName || '';
    if (elJobTitle) elJobTitle.value = sigState.jobTitle || '';
    if (elCompany) elCompany.value = sigState.company || '';
    if (elPhone) elPhone.value = sigState.phone || '';
    if (elEmail) elEmail.value = sigState.email || '';
    if (elWebsite) elWebsite.value = sigState.website || '';
    if (elCustomColor) elCustomColor.value = sigState.themeColor || '#0078d4';
    if (elAutoNew) elAutoNew.checked = sigState.autoNewMail !== false;
    if (elAutoReply) elAutoReply.checked = sigState.autoReplyForward !== false;

    if (elBtnRemoveLogo && elLogoStatus) {
      if (sigState.logoDataUrl) {
        elBtnRemoveLogo.style.display = 'inline-flex';
        elLogoStatus.textContent = 'Logo carregada';
      } else {
        elBtnRemoveLogo.style.display = 'none';
        elLogoStatus.textContent = '';
      }
    }

    // Highlight active template card
    document.querySelectorAll('.sig-template-card').forEach(card => {
      if (card.getAttribute('data-template') === (sigState.template || 'executive')) {
        card.classList.add('active');
      } else {
        card.classList.remove('active');
      }
    });

    // Highlight active color dot
    document.querySelectorAll('.sig-color-dot').forEach(dot => {
      if (dot.getAttribute('data-color')?.toLowerCase() === (sigState.themeColor || '#0078d4').toLowerCase()) {
        dot.classList.add('active');
      } else {
        dot.classList.remove('active');
      }
    });

    if (elSigEditor) {
      elSigEditor.innerHTML = sigState.html || buildSignatureHtml(sigState);
    }
  }

  function readGeneratorFieldsIntoSigState() {
    sigState.fullName = document.getElementById('sigFullName')?.value || '';
    sigState.jobTitle = document.getElementById('sigJobTitle')?.value || '';
    sigState.company = document.getElementById('sigCompany')?.value || '';
    sigState.phone = document.getElementById('sigPhone')?.value || '';
    sigState.email = document.getElementById('sigEmail')?.value || '';
    sigState.website = document.getElementById('sigWebsite')?.value || '';
    sigState.autoNewMail = document.getElementById('sigAutoNewMail')?.checked !== false;
    sigState.autoReplyForward = document.getElementById('sigAutoReplyForward')?.checked !== false;
  }

  function generateSignatureIntoEditor() {
    readGeneratorFieldsIntoSigState();
    const generated = buildSignatureHtml(sigState);
    sigState.html = generated;
    const elSigEditor = document.getElementById('sigEditor');
    if (elSigEditor) {
      elSigEditor.innerHTML = generated;
    }
  }

  function insertSignatureIntoCompose() {
    if (!elComposeEditor) return;
    const sigHtml = sigState.html || buildSignatureHtml(sigState);
    if (!sigHtml || !sigHtml.trim()) {
      showToast('Nenhuma assinatura definida. Crie uma em Configurações!');
      return;
    }

    const existingSig = elComposeEditor.querySelector('[data-outlook-signature="true"]');
    if (existingSig) {
      existingSig.innerHTML = sigHtml;
    } else {
      const wrapper = document.createElement('div');
      wrapper.setAttribute('data-outlook-signature', 'true');
      wrapper.style.marginTop = '16px';
      wrapper.innerHTML = sigHtml;
      elComposeEditor.appendChild(wrapper);
    }
    showToast('Assinatura aplicada ao e-mail!');
  }

  // Compose Functions
  function openCompose(options = {}) {
    elComposeModal.style.display = 'flex';
    elComposeTo.value = options.to || '';
    elComposeCc.value = options.cc || '';
    elComposeSubject.value = options.subject || '';

    const isReplyOrForward = !!options.isReplyOrForward;
    const shouldAutoSign = isReplyOrForward ? (sigState.autoReplyForward !== false) : (sigState.autoNewMail !== false);
    const sigHtml = (sigState.html || '').trim();

    let finalBody = options.body || '';
    if (shouldAutoSign && sigHtml) {
      const sigBlock = `<div data-outlook-signature="true" style="margin-top: 14px; margin-bottom: 10px;">${sigHtml}</div>`;
      if (isReplyOrForward) {
        finalBody = `<p><br></p>${sigBlock}${finalBody}`;
      } else {
        finalBody = `${finalBody ? finalBody + '<br>' : '<p><br></p>'}${sigBlock}`;
      }
    }

    // Reset default editor font to Arial 15px
    const selFontFamily = document.getElementById('composeFontFamily');
    const selFontSize = document.getElementById('composeFontSize');
    if (selFontFamily) selFontFamily.value = 'Arial, Helvetica, sans-serif';
    if (selFontSize) selFontSize.value = '15';
    elComposeEditor.style.fontFamily = 'Arial, Helvetica, sans-serif';
    elComposeEditor.style.fontSize = '15px';

    elComposeEditor.innerHTML = finalBody;
    attachedFiles = [];
    elComposeAttachedFilesCount.textContent = '';
    elComposeStatusMsg.textContent = '';
    if (options.to) {
      elComposeEditor.focus();
    } else {
      elComposeTo.focus();
    }
  }

  function closeCompose() {
    elComposeModal.style.display = 'none';
    elComposeModal.classList.remove('maximized', 'minimized');
    const btnMax = document.getElementById('btnMaximizeCompose');
    if (btnMax) {
      btnMax.title = 'Maximizar';
      btnMax.innerHTML = '<i class="fa-regular fa-square"></i>';
    }
    elComposeForm.reset();
    elComposeEditor.innerHTML = '';
    attachedFiles = [];
  }

  async function handleSendEmail(e) {
    e.preventDefault();
    elBtnSendMail.disabled = true;
    elComposeStatusMsg.textContent = 'Enviando email via SMTP...';

    const baseFont = elComposeEditor.style.fontFamily || 'Arial, Helvetica, sans-serif';
    const baseSize = elComposeEditor.style.fontSize || '15px';
    const wrappedHtml = `<div style="font-family: ${baseFont}; font-size: ${baseSize};">${elComposeEditor.innerHTML}</div>`;

    const formData = new FormData();
    formData.append('to', elComposeTo.value);
    formData.append('cc', elComposeCc.value);
    formData.append('subject', elComposeSubject.value);
    formData.append('html', wrappedHtml);
    formData.append('text', elComposeEditor.innerText);

    attachedFiles.forEach(f => formData.append('attachments', f));

    try {
      const res = await fetch('/api/smtp/send', {
        method: 'POST',
        body: formData
      });
      const data = await res.json();
      if (data.success) {
        showToast(data.message || 'Email enviado com sucesso!');
        closeCompose();
        loadFolders();
        if (currentFolder.toLowerCase().includes('sent') || currentFolder.toLowerCase().includes('enviad')) {
          loadMessages();
        }
      } else {
        alert(data.message || 'Falha ao enviar email');
      }
    } catch (err) {
      alert('Erro de comunicação com o servidor: ' + err.message);
    } finally {
      elBtnSendMail.disabled = false;
      elComposeStatusMsg.textContent = '';
    }
  }

  // Settings & Config
  async function loadConfig() {
    try {
      const res = await fetch('/api/config');
      const data = await res.json();
      if (data.imap) {
        if (data.imap.hasPassword) {
          elImapStatusBadge.className = 'imap-badge connected';
          elImapStatusText.textContent = 'IMAP: Conectado';
        }
        document.getElementById('cfgImapHost').value = data.imap.host || 'h58.servidorhh.com';
        document.getElementById('cfgImapPort').value = data.imap.port || 993;
        document.getElementById('cfgImapUser').value = data.imap.user || '';
        document.getElementById('cfgImapSsl').checked = data.imap.secure !== false;
        
        // Header text & active user label
        const userShort = (data.imap.user || '').split('@')[0].replace(/[._-]+/g, ' ');
        document.getElementById('userAvatar').title = data.imap.user;
        document.getElementById('userAvatar').textContent = getInitials(userShort || 'Clodoaldo Silva');
        if (elActiveUserEmailLabel) {
          elActiveUserEmailLabel.textContent = data.imap.user || '';
        }
      }
      if (data.smtp) {
        document.getElementById('cfgSmtpHost').value = data.smtp.host || 'h58.servidorhh.com';
        document.getElementById('cfgSmtpPort').value = data.smtp.port || 465;
      }
      if (data.signature) {
        sigState = { ...sigState, ...data.signature };
      }
      if (!localStorage.getItem('outlook_reading_pane_layout') && data.layout?.readingPane) {
        readingPaneLayout = data.layout.readingPane;
      }
      applyReadingPaneLayout(readingPaneLayout);
      const activeEmail = (data.imap?.user || sessionStorage.getItem('outlook_session_user') || '').toLowerCase();
      if (activeEmail) {
        const localSig = localStorage.getItem('outlook_sig_' + activeEmail);
        if (localSig) {
          try { sigState = { ...sigState, ...JSON.parse(localSig) }; } catch (e) {}
        }
      }
      savedEmails = getMergedSavedEmails(data.savedEmails || (data.imap?.user ? [data.imap.user] : []));
      saveEmailsLocally(savedEmails);
      populateSignatureBuilderUI();
    } catch (e) {
      console.error('Erro ao ler config:', e);
    }
  }

  async function handleSaveSettings() {
    readGeneratorFieldsIntoSigState();
    const elSigEditor = document.getElementById('sigEditor');
    if (elSigEditor) {
      sigState.html = elSigEditor.innerHTML;
    }

    const activeUser = (document.getElementById('cfgImapUser')?.value || sessionStorage.getItem('outlook_session_user') || '').trim().toLowerCase();
    if (activeUser) {
      try {
        localStorage.setItem('outlook_sig_' + activeUser, JSON.stringify(sigState));
      } catch (e) {}
    }

    const configData = {
      imap: {
        host: document.getElementById('cfgImapHost').value,
        port: document.getElementById('cfgImapPort').value,
        user: document.getElementById('cfgImapUser').value,
        password: document.getElementById('cfgImapPass').value,
        secure: document.getElementById('cfgImapSsl').checked
      },
      smtp: {
        host: document.getElementById('cfgSmtpHost').value,
        port: document.getElementById('cfgSmtpPort').value,
        user: document.getElementById('cfgImapUser').value,
        password: document.getElementById('cfgImapPass').value
      },
      signature: sigState,
      layout: { readingPane: readingPaneLayout }
    };

    try {
      const res = await fetch('/api/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(configData)
      });
      const data = await res.json();
      if (data.success) {
        showToast('Configurações e Assinatura salvas com sucesso!');
        elSettingsModal.style.display = 'none';
        loadFolders();
        loadMessages();
      }
    } catch (e) {
      alert('Erro ao salvar: ' + e.message);
    }
  }

  async function handleTestImap() {
    const passInput = document.getElementById('cfgImapPass');
    const passVal = passInput ? passInput.value : '';

    elTestResultBox.style.display = 'block';

    if (!passVal) {
      elTestResultBox.className = 'test-result-box error';
      elTestResultBox.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> <strong>Atenção:</strong> Por favor, digite a senha do seu email no campo "Senha do Email" antes de testar a conexão.';
      passInput.focus();
      return;
    }

    elTestResultBox.className = 'test-result-box';
    elTestResultBox.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Testando entrada IMAP e saída SMTP...';

    const testPayload = {
      imap: {
        host: document.getElementById('cfgImapHost').value.trim(),
        port: document.getElementById('cfgImapPort').value.trim(),
        user: document.getElementById('cfgImapUser').value.trim(),
        password: passVal,
        secure: document.getElementById('cfgImapSsl').checked
      },
      smtp: {
        host: document.getElementById('cfgSmtpHost').value.trim(),
        port: document.getElementById('cfgSmtpPort').value.trim(),
        user: document.getElementById('cfgImapUser').value.trim(),
        password: passVal
      }
    };

    try {
      // Test IMAP and SMTP in parallel
      const [resImap, resSmtp] = await Promise.all([
        fetch('/api/imap/test', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ imap: testPayload.imap })
        }).then(r => r.json()).catch(e => ({ success: false, message: e.message })),
        fetch('/api/smtp/test', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ smtp: testPayload.smtp })
        }).then(r => r.json()).catch(e => ({ success: false, message: e.message }))
      ]);

      let html = '<div style="display: flex; flex-direction: column; gap: 8px;">';

      // IMAP result
      if (resImap.success) {
        html += `<div style="color: #107c41;"><i class="fa-solid fa-check-circle"></i> <strong>IMAP (Entrada):</strong> ${resImap.message}</div>`;
        elImapStatusBadge.className = 'imap-badge connected';
        elImapStatusText.textContent = 'IMAP: Conectado';
      } else {
        html += `<div style="color: #a80000;"><i class="fa-solid fa-circle-xmark"></i> <strong>IMAP (Entrada):</strong> ${resImap.message}</div>`;
      }

      // SMTP result
      if (resSmtp.success) {
        html += `<div style="color: #107c41;"><i class="fa-solid fa-check-circle"></i> <strong>SMTP (Saída):</strong> ${resSmtp.message}</div>`;
        if (resSmtp.suggestedPort) {
          document.getElementById('cfgSmtpPort').value = resSmtp.suggestedPort;
        }
      } else {
        html += `<div style="color: #a80000;"><i class="fa-solid fa-circle-xmark"></i> <strong>SMTP (Saída):</strong> ${resSmtp.message}</div>`;
      }

      html += '</div>';

      elTestResultBox.className = `test-result-box ${resImap.success && resSmtp.success ? 'success' : 'error'}`;
      elTestResultBox.innerHTML = html;
    } catch (err) {
      elTestResultBox.className = 'test-result-box error';
      elTestResultBox.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> <strong>Erro de comunicação com o servidor local:</strong> ${err.message}`;
    }
  }

  // Event Listeners Setup
  function setupEventListeners() {
    // New email
    elBtnNewEmail.onclick = () => openCompose();

    // Message list item click (Event Delegation)
    elMessageListScroll.onclick = (e) => {
      const actionBtn = e.target.closest('.card-action-btn');
      if (actionBtn) {
        e.stopPropagation();
        const card = actionBtn.closest('.message-card');
        const uid = card ? card.getAttribute('data-uid') : null;
        const msg = messages.find(m => String(m.uid) === String(uid));
        if (!msg) return;

        if (actionBtn.classList.contains('btn-quick-delete')) {
          deleteEmail(msg.uid);
        } else if (actionBtn.classList.contains('btn-quick-read')) {
          toggleReadEmail(msg.uid, !msg.seen);
        } else if (actionBtn.classList.contains('btn-quick-flag')) {
          toggleFlagEmail(msg.uid, !msg.flagged);
        }
        return;
      }

      const card = e.target.closest('.message-card');
      if (card) {
        const uid = card.getAttribute('data-uid');
        const msg = messages.find(m => String(m.uid) === String(uid));
        if (msg) {
          selectMessage(msg);
        }
      }
    };

    // Favorites click
    document.querySelectorAll('#favoritesList .folder-item').forEach(item => {
      item.onclick = (e) => {
        e.preventDefault();
        const fPath = item.getAttribute('data-folder');
        const fName = item.querySelector('span').textContent;
        selectFolder(fPath, fName);
      };
    });

    // Filter tabs
    document.querySelectorAll('.filter-tab').forEach(tab => {
      tab.onclick = () => {
        document.querySelectorAll('.filter-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        currentFilter = tab.getAttribute('data-filter');
        loadMessages();
      };
    });

    // Search input
    const elBtnClearSearch = document.getElementById('btnClearSearch');
    let searchTimeout = null;

    const updateClearBtn = () => {
      if (elBtnClearSearch) {
        elBtnClearSearch.style.display = elSearchInput.value ? 'block' : 'none';
      }
    };

    elSearchInput.oninput = () => {
      updateClearBtn();
      clearTimeout(searchTimeout);
      searchTimeout = setTimeout(() => {
        currentSearch = elSearchInput.value.trim();
        loadMessages();
      }, 350);
    };

    if (elBtnClearSearch) {
      elBtnClearSearch.onclick = () => {
        elSearchInput.value = '';
        currentSearch = '';
        updateClearBtn();
        loadMessages();
        elSearchInput.focus();
      };
    }

    // Ribbon buttons
    elBtnRibbonRefresh.onclick = () => {
      loadFolders();
      loadMessages();
      showToast('Sincronizado com o servidor IMAP.');
    };

    // Select All Checkboxes (Header & Columns Bar)
    const handleSelectAllChange = (checked) => {
      if (checked) {
        messages.forEach(m => selectedUids.add(String(m.uid)));
      } else {
        selectedUids.clear();
      }
      syncSelectionUI();
    };

    const chkSelectAllMessages = document.getElementById('chkSelectAllMessages');
    const chkSelectAllCols = document.getElementById('chkSelectAllCols');
    if (chkSelectAllMessages) {
      chkSelectAllMessages.onchange = () => handleSelectAllChange(chkSelectAllMessages.checked);
    }
    if (chkSelectAllCols) {
      chkSelectAllCols.onchange = () => handleSelectAllChange(chkSelectAllCols.checked);
    }

    elBtnRibbonDelete.onclick = () => {
      if (selectedUids.size > 0) {
        bulkDeleteEmails(selectedUids);
      } else if (selectedMessage) {
        deleteEmail(selectedMessage.uid);
      }
    };

    elBtnRibbonToggleRead.onclick = () => {
      if (selectedUids.size > 0) {
        const anyUnread = messages.some(m => selectedUids.has(String(m.uid)) && !m.seen);
        bulkSetFlags(selectedUids, { seen: anyUnread });
      } else if (selectedMessage) {
        toggleReadEmail(selectedMessage.uid, !selectedMessage.seen);
      }
    };

    elBtnRibbonFlag.onclick = () => {
      if (selectedUids.size > 0) {
        const anyUnflagged = messages.some(m => selectedUids.has(String(m.uid)) && !m.flagged);
        bulkSetFlags(selectedUids, { flagged: anyUnflagged });
      } else if (selectedMessage) {
        toggleFlagEmail(selectedMessage.uid, !selectedMessage.flagged);
      }
    };

    // Reading pane buttons
    elBtnDeleteCurrent.onclick = () => {
      if (selectedMessage) deleteEmail(selectedMessage.uid);
    };

    elBtnToggleReadCurrent.onclick = () => {
      if (selectedMessage) toggleReadEmail(selectedMessage.uid, !selectedMessage.seen);
    };

    elBtnFlagCurrent.onclick = () => {
      if (selectedMessage) toggleFlagEmail(selectedMessage.uid, !selectedMessage.flagged);
    };

    const btnMaximizeReader = document.getElementById('btnMaximizeReader');
    const elReadingPane = document.getElementById('readingPane');
    const setReaderMaximized = (max) => {
      if (!elReadingPane) return;
      if (max) elReadingPane.classList.add('maximized');
      else elReadingPane.classList.remove('maximized');

      if (btnMaximizeReader) {
        btnMaximizeReader.title = max ? 'Restaurar visualização (Esc)' : 'Maximizar mensagem na tela';
        btnMaximizeReader.innerHTML = max
          ? '<i class="fa-solid fa-compress"></i> <span>Restaurar</span>'
          : '<i class="fa-solid fa-expand"></i> <span>Maximizar</span>';
      }
    };

    if (btnMaximizeReader && elReadingPane) {
      btnMaximizeReader.onclick = () => {
        setReaderMaximized(!elReadingPane.classList.contains('maximized'));
      };
    }

    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && elReadingPane && elReadingPane.classList.contains('maximized')) {
        setReaderMaximized(false);
      }
    });

    elBtnReply.onclick = () => {
      if (!selectedMessage) return;
      openCompose({
        to: selectedMessage.from?.address || '',
        subject: `Re: ${selectedMessage.subject || ''}`,
        isReplyOrForward: true,
        body: `<br><hr><p style="color: #605e5c;">Em ${formatDate(selectedMessage.date)}, ${selectedMessage.from?.name || 'remetente'} escreveu:</p>${selectedMessage.html || selectedMessage.text || ''}`
      });
    };

    elBtnReplyAll.onclick = () => {
      if (!selectedMessage) return;
      const ccList = (selectedMessage.to || []).map(t => t.address).filter(a => a !== selectedMessage.from?.address).join(', ');
      openCompose({
        to: selectedMessage.from?.address || '',
        cc: ccList,
        subject: `Re: ${selectedMessage.subject || ''}`,
        isReplyOrForward: true,
        body: `<br><hr><p style="color: #605e5c;">Em ${formatDate(selectedMessage.date)}, ${selectedMessage.from?.name || 'remetente'} escreveu:</p>${selectedMessage.html || selectedMessage.text || ''}`
      });
    };

    elBtnForward.onclick = () => {
      if (!selectedMessage) return;
      openCompose({
        subject: `Enc: ${selectedMessage.subject || ''}`,
        isReplyOrForward: true,
        body: `<br><hr><p style="color: #605e5c;">---------- Mensagem encaminhada ----------<br>De: ${selectedMessage.from?.name} &lt;${selectedMessage.from?.address}&gt;<br>Data: ${formatDate(selectedMessage.date)}<br>Assunto: ${selectedMessage.subject}</p>${selectedMessage.html || selectedMessage.text || ''}`
      });
    };

    // Compose actions
    elComposeForm.onsubmit = handleSendEmail;
    elBtnCloseCompose.onclick = closeCompose;

    const btnMaximizeCompose = document.getElementById('btnMaximizeCompose');
    const btnMinimizeCompose = document.getElementById('btnMinimizeCompose');

    if (btnMaximizeCompose && elComposeModal) {
      btnMaximizeCompose.onclick = () => {
        elComposeModal.classList.remove('minimized');
        const isMax = elComposeModal.classList.toggle('maximized');
        btnMaximizeCompose.title = isMax ? 'Restaurar tamanho' : 'Maximizar';
        btnMaximizeCompose.innerHTML = isMax
          ? '<i class="fa-regular fa-clone"></i>'
          : '<i class="fa-regular fa-square"></i>';
      };
    }

    if (btnMinimizeCompose && elComposeModal) {
      btnMinimizeCompose.onclick = () => {
        elComposeModal.classList.remove('maximized');
        if (btnMaximizeCompose) {
          btnMaximizeCompose.title = 'Maximizar';
          btnMaximizeCompose.innerHTML = '<i class="fa-regular fa-square"></i>';
        }
        elComposeModal.classList.toggle('minimized');
      };
    }

    elBtnDiscardMail.onclick = () => {
      if (confirm('Deseja descartar este rascunho?')) closeCompose();
    };

    elComposeFileInput.onchange = (e) => {
      attachedFiles = Array.from(e.target.files);
      elComposeAttachedFilesCount.textContent = attachedFiles.length > 0 ? `${attachedFiles.length} anexo(s)` : '';
    };

    // Compose Toolbar Signature Buttons
    const btnInsertSignatureNow = document.getElementById('btnInsertSignatureNow');
    if (btnInsertSignatureNow) {
      btnInsertSignatureNow.onclick = () => insertSignatureIntoCompose();
    }

    // Compose Font Family, Size (px) and Text Color
    const composeFontFamily = document.getElementById('composeFontFamily');
    const composeFontSize = document.getElementById('composeFontSize');
    const composeTextColor = document.getElementById('composeTextColor');

    if (composeFontFamily && elComposeEditor) {
      composeFontFamily.onchange = () => {
        const fontVal = composeFontFamily.value;
        elComposeEditor.focus();
        const sel = window.getSelection();
        if (sel && !sel.isCollapsed) {
          document.execCommand('fontName', false, fontVal);
        } else {
          elComposeEditor.style.fontFamily = fontVal;
          document.execCommand('fontName', false, fontVal);
        }
      };
    }

    if (composeFontSize && elComposeEditor) {
      composeFontSize.onchange = () => {
        const px = composeFontSize.value || '15';
        elComposeEditor.focus();
        const sel = window.getSelection();
        if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {
          document.execCommand('fontSize', false, '7');
          elComposeEditor.querySelectorAll('font[size="7"]').forEach(fontEl => {
            fontEl.removeAttribute('size');
            fontEl.style.fontSize = `${px}px`;
          });
        } else {
          elComposeEditor.style.fontSize = `${px}px`;
          if (sel && sel.rangeCount > 0) {
            const range = sel.getRangeAt(0);
            const span = document.createElement('span');
            span.style.fontSize = `${px}px`;
            span.innerHTML = '&#8203;';
            range.insertNode(span);
            const newRange = document.createRange();
            newRange.setStart(span.firstChild, 1);
            newRange.collapse(true);
            sel.removeAllRanges();
            sel.addRange(newRange);
          }
        }
      };
    }

    if (composeTextColor && elComposeEditor) {
      composeTextColor.oninput = (e) => {
        elComposeEditor.focus();
        document.execCommand('foreColor', false, e.target.value);
      };
    }

    const switchSettingsTab = (tabId) => {
      document.querySelectorAll('.settings-tab').forEach(btn => {
        if (btn.getAttribute('data-tab') === tabId) {
          btn.classList.add('active');
        } else {
          btn.classList.remove('active');
        }
      });
      document.querySelectorAll('.settings-tab-pane').forEach(pane => {
        if (pane.id === tabId) {
          pane.classList.add('active');
        } else {
          pane.classList.remove('active');
        }
      });
      if (elBtnTestImap) {
        elBtnTestImap.style.display = tabId === 'tabServer' ? 'inline-flex' : 'none';
      }
    };

    document.querySelectorAll('.settings-tab').forEach(tabBtn => {
      tabBtn.onclick = () => {
        const targetTab = tabBtn.getAttribute('data-tab');
        if (targetTab) switchSettingsTab(targetTab);
      };
    });

    const btnOpenSigFromCompose = document.getElementById('btnOpenSigFromCompose');
    if (btnOpenSigFromCompose) {
      btnOpenSigFromCompose.onclick = () => {
        populateSignatureBuilderUI();
        elTestResultBox.style.display = 'none';
        elSettingsModal.style.display = 'flex';
        switchSettingsTab('tabSignature');
      };
    }

    // Signature Builder Studio Listeners
    document.querySelectorAll('.sig-template-card').forEach(card => {
      card.onclick = () => {
        document.querySelectorAll('.sig-template-card').forEach(c => c.classList.remove('active'));
        card.classList.add('active');
        sigState.template = card.getAttribute('data-template') || 'executive';
        generateSignatureIntoEditor();
      };
    });

    document.querySelectorAll('.sig-color-dot').forEach(dot => {
      dot.onclick = () => {
        document.querySelectorAll('.sig-color-dot').forEach(d => d.classList.remove('active'));
        dot.classList.add('active');
        const c = dot.getAttribute('data-color') || '#0078d4';
        sigState.themeColor = c;
        const customColor = document.getElementById('sigCustomColor');
        if (customColor) customColor.value = c;
        generateSignatureIntoEditor();
      };
    });

    const sigCustomColor = document.getElementById('sigCustomColor');
    if (sigCustomColor) {
      sigCustomColor.oninput = (e) => {
        document.querySelectorAll('.sig-color-dot').forEach(d => d.classList.remove('active'));
        sigState.themeColor = e.target.value;
        generateSignatureIntoEditor();
      };
    }

    const btnGenerateSignature = document.getElementById('btnGenerateSignature');
    if (btnGenerateSignature) {
      btnGenerateSignature.onclick = () => {
        generateSignatureIntoEditor();
        showToast('Assinatura gerada e atualizada no editor!');
      };
    }

    // Live update on generator input fields
    ['sigFullName', 'sigJobTitle', 'sigCompany', 'sigPhone', 'sigEmail', 'sigWebsite'].forEach(fieldId => {
      const el = document.getElementById(fieldId);
      if (el) {
        el.oninput = () => generateSignatureIntoEditor();
      }
    });

    // Logo Upload in Signature Generator
    const sigLogoUpload = document.getElementById('sigLogoUpload');
    const btnRemoveSigLogo = document.getElementById('btnRemoveSigLogo');
    const sigLogoStatus = document.getElementById('sigLogoStatus');

    if (sigLogoUpload) {
      sigLogoUpload.onchange = (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
          sigState.logoDataUrl = ev.target.result;
          if (btnRemoveSigLogo) btnRemoveSigLogo.style.display = 'inline-flex';
          if (sigLogoStatus) sigLogoStatus.textContent = file.name;
          generateSignatureIntoEditor();
        };
        reader.readAsDataURL(file);
      };
    }

    if (btnRemoveSigLogo) {
      btnRemoveSigLogo.onclick = () => {
        sigState.logoDataUrl = '';
        if (sigLogoUpload) sigLogoUpload.value = '';
        btnRemoveSigLogo.style.display = 'none';
        if (sigLogoStatus) sigLogoStatus.textContent = '';
        generateSignatureIntoEditor();
      };
    }

    // Rich Editor Toolbar for Signature
    const elSigEditor = document.getElementById('sigEditor');
    document.querySelectorAll('[data-sig-cmd]').forEach(btn => {
      btn.onclick = () => {
        const cmd = btn.getAttribute('data-sig-cmd');
        if (elSigEditor) elSigEditor.focus();
        document.execCommand(cmd, false, null);
      };
    });

    const sigTextColorPicker = document.getElementById('sigTextColorPicker');
    if (sigTextColorPicker) {
      sigTextColorPicker.oninput = (e) => {
        if (elSigEditor) elSigEditor.focus();
        document.execCommand('foreColor', false, e.target.value);
      };
    }

    const sigInlineImageInput = document.getElementById('sigInlineImageInput');
    if (sigInlineImageInput) {
      sigInlineImageInput.onchange = (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
          if (elSigEditor) {
            elSigEditor.focus();
            document.execCommand('insertHTML', false, `<img src="${ev.target.result}" style="max-height: 70px; max-width: 180px; border-radius: 4px;">`);
          }
        };
        reader.readAsDataURL(file);
      };
    }

    const btnSigClear = document.getElementById('btnSigClear');
    if (btnSigClear && elSigEditor) {
      btnSigClear.onclick = () => {
        elSigEditor.innerHTML = '';
        sigState.html = '';
      };
    }

    // Reading Pane Layout Option Cards
    document.querySelectorAll('.layout-option-card').forEach(card => {
      card.onclick = () => {
        const chosen = card.getAttribute('data-layout') || 'right';
        applyReadingPaneLayout(chosen);
        showToast(chosen === 'bottom' ? 'Layout alterado: Visualização Inferior' : 'Layout alterado: Visualização Lateral');
      };
    });

    // Settings Modal
    const openSettings = () => {
      populateSignatureBuilderUI();
      elTestResultBox.style.display = 'none';
      elSettingsModal.style.display = 'flex';
    };
    if (elBtnOpenSettings) elBtnOpenSettings.onclick = openSettings;
    if (elRailSettingsBtn) elRailSettingsBtn.onclick = openSettings;
    elBtnCloseSettings.onclick = () => elSettingsModal.style.display = 'none';
    elBtnCancelSettings.onclick = () => elSettingsModal.style.display = 'none';
    elBtnSaveSettings.onclick = handleSaveSettings;
    elBtnTestImap.onclick = handleTestImap;

    // Login Screen & Multi-User Account Switching
    if (elLoginForm) {
      elLoginForm.onsubmit = handleLoginSubmit;
    }

    if (elLoginEmailInput) {
      elLoginEmailInput.oninput = () => renderSavedEmails();
    }

    const triggerAccountSwitch = () => {
      sessionStorage.removeItem('outlook_session_auth');
      showLoginScreen();
    };

    if (elBtnSwitchAccount) {
      elBtnSwitchAccount.onclick = triggerAccountSwitch;
    }

    const elUserAvatar = document.getElementById('userAvatar');
    if (elUserAvatar) {
      elUserAvatar.onclick = triggerAccountSwitch;
    }

    const btnToggleLoginPass = document.getElementById('btnToggleLoginPass');
    const loginPassEyeIcon = document.getElementById('loginPassEyeIcon');
    if (btnToggleLoginPass && elLoginPasswordInput && loginPassEyeIcon) {
      btnToggleLoginPass.onclick = () => {
        if (elLoginPasswordInput.type === 'password') {
          elLoginPasswordInput.type = 'text';
          loginPassEyeIcon.className = 'fa-regular fa-eye-slash';
        } else {
          elLoginPasswordInput.type = 'password';
          loginPassEyeIcon.className = 'fa-regular fa-eye';
        }
      };
    }

    // Toggle password eye icon in Settings
    const btnTogglePassVisibility = document.getElementById('btnTogglePassVisibility');
    const cfgImapPass = document.getElementById('cfgImapPass');
    const passEyeIcon = document.getElementById('passEyeIcon');
    if (btnTogglePassVisibility && cfgImapPass && passEyeIcon) {
      btnTogglePassVisibility.onclick = () => {
        if (cfgImapPass.type === 'password') {
          cfgImapPass.type = 'text';
          passEyeIcon.className = 'fa-regular fa-eye-slash';
        } else {
          cfgImapPass.type = 'password';
          passEyeIcon.className = 'fa-regular fa-eye';
        }
      };
    }

    // Keyboard Shortcuts (Ctrl+E for search, Ctrl+N for new mail)
    window.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'e') {
        e.preventDefault();
        elSearchInput.focus();
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'n') {
        e.preventDefault();
        openCompose();
      }
    });
  }

  // Utilities
  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function formatBytes(bytes) {
    if (!bytes || bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  // Kickstart
  document.addEventListener('DOMContentLoaded', init);
})();

