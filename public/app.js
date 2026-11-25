// Outlook Webmail IMAP Frontend Application
(function () {
  'use strict';

  // Persistent daily/boot session store (localStorage survives browser close/reopen;
  // expires only on computer/server restart, new calendar day, or manual logout)
  function getTodayDateKey() {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  const SESSION_STORE = {
    getAuth:      () => localStorage.getItem('outlook_session_auth'),
    setAuth:      (v) => localStorage.setItem('outlook_session_auth', v),
    clearAuth:    () => localStorage.removeItem('outlook_session_auth'),
    getUser:      () => localStorage.getItem('outlook_session_user'),
    setUser:      (v) => localStorage.setItem('outlook_session_user', v),
    clearUser:    () => localStorage.removeItem('outlook_session_user'),
    getPass:      () => localStorage.getItem('outlook_session_pass_b64'),
    setPass:      (v) => localStorage.setItem('outlook_session_pass_b64', v),
    clearPass:    () => localStorage.removeItem('outlook_session_pass_b64'),
    getBootId:    () => localStorage.getItem('outlook_server_boot_id'),
    setBootId:    (v) => localStorage.setItem('outlook_server_boot_id', v),
    getLoginDate: () => localStorage.getItem('outlook_session_login_date'),
    setLoginDate: (v) => localStorage.setItem('outlook_session_login_date', v),
    getAccounts() {
      let list = [];
      try {
        const raw = localStorage.getItem('outlook_connected_accounts');
        if (raw) {
          const parsed = JSON.parse(raw);
          if (Array.isArray(parsed)) list = parsed.filter(a => a && a.email && a.passB64);
        }
      } catch (e) {}
      const curUser = (this.getUser() || '').trim().toLowerCase();
      const curPass = this.getPass() || '';
      if (curUser && curPass && !list.some(a => a.email.toLowerCase() === curUser)) {
        list.unshift({ email: curUser, passB64: curPass });
        this.saveAccounts(list);
      }
      return list;
    },
    saveAccounts(list) {
      try {
        localStorage.setItem('outlook_connected_accounts', JSON.stringify(list || []));
      } catch (e) {}
    },
    addAccount(acc) {
      if (!acc || !acc.email || !acc.passB64) return;
      const email = acc.email.trim().toLowerCase();
      const list = this.getAccounts().filter(a => a.email.toLowerCase() !== email);
      list.push({
        email,
        passB64: acc.passB64,
        imapHost: acc.imapHost || '',
        imapPort: acc.imapPort || '',
        smtpHost: acc.smtpHost || '',
        smtpPort: acc.smtpPort || ''
      });
      this.saveAccounts(list);
    },
    removeAccount(email) {
      const target = String(email || '').trim().toLowerCase();
      const list = this.getAccounts().filter(a => a.email.toLowerCase() !== target);
      this.saveAccounts(list);
      return list;
    },
    getActiveAccountObj() {
      const curUser = (this.getUser() || '').trim().toLowerCase();
      const list = this.getAccounts();
      return list.find(a => a.email.toLowerCase() === curUser) || (curUser ? { email: curUser, passB64: this.getPass() || '' } : null);
    },
    setActiveAccount(acc) {
      if (!acc || !acc.email) return;
      this.setUser(acc.email.trim().toLowerCase());
      if (acc.passB64) this.setPass(acc.passB64);
    },
    clearAll() {
      this.clearAuth();
      this.clearUser();
      this.clearPass();
      localStorage.removeItem('outlook_connected_accounts');
      localStorage.removeItem('outlook_server_boot_id');
      localStorage.removeItem('outlook_session_login_date');
      // Clean up any legacy sessionStorage keys
      sessionStorage.removeItem('outlook_session_auth');
      sessionStorage.removeItem('outlook_session_user');
      sessionStorage.removeItem('outlook_session_pass_b64');
      sessionStorage.removeItem('outlook_last_activity');
    }
  };

  let currentServerBootId = null;

  async function fetchServerBootId() {
    try {
      const res = await originalFetch('/api/server-boot-id', { cache: 'no-store' });
      if (res.ok) {
        const data = await res.json();
        if (data && data.bootId) {
          currentServerBootId = data.bootId;
          return data.bootId;
        }
      }
    } catch (e) {}
    return currentServerBootId;
  }

  // Intercept fetch to attach active account credentials
  const originalFetch = window.fetch.bind(window);
  window.fetch = function (input, init = {}) {
    const url = typeof input === 'string' ? input : (input && input.url) || '';
    if (url.startsWith('/api/')) {
      const activeAcc = SESSION_STORE.getActiveAccountObj();
      if (activeAcc && activeAcc.email && activeAcc.passB64) {
        const headers = new Headers(init.headers || {});
        if (!headers.has('x-imap-user')) {
          headers.set('x-imap-user', activeAcc.email);
          headers.set('x-imap-pass', activeAcc.passB64);
          if (activeAcc.imapHost) headers.set('x-imap-host', activeAcc.imapHost);
          if (activeAcc.imapPort) headers.set('x-imap-port', String(activeAcc.imapPort));
          if (activeAcc.smtpHost) headers.set('x-smtp-host', activeAcc.smtpHost);
          if (activeAcc.smtpPort) headers.set('x-smtp-port', String(activeAcc.smtpPort));
        }
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
  let emailTemplates = [];
  let editingTemplateId = null;
  let messageLikesMap = {};

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
    const elListPane = document.getElementById('messageListPane');
    if (elContentSplit) {
      if (readingPaneLayout === 'bottom') {
        elContentSplit.classList.add('layout-bottom');
        if (elListPane) {
          elListPane.style.width = '';
          const savedH = localStorage.getItem('outlook_split_height');
          elListPane.style.height = savedH ? `${savedH}px` : '';
        }
      } else {
        elContentSplit.classList.remove('layout-bottom');
        if (elListPane) {
          elListPane.style.height = '';
          const savedW = localStorage.getItem('outlook_split_width');
          elListPane.style.width = savedW ? `${savedW}px` : '';
        }
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

    // Single daily login / valid until computer (server) restart — no 30-min inactivity logout
    const serverBootId = await fetchServerBootId();
    const isSessionAuth = SESSION_STORE.getAuth() === 'true';
    const storedBootId = SESSION_STORE.getBootId();
    const storedLoginDate = SESSION_STORE.getLoginDate();
    const todayKey = getTodayDateKey();

    const serverRestarted = Boolean(serverBootId && storedBootId && storedBootId !== serverBootId);
    const isDifferentDay = Boolean(storedLoginDate && storedLoginDate !== todayKey);
    const missingBootOrDate = Boolean(isSessionAuth && (!storedBootId || !storedLoginDate));

    if (isSessionAuth && !serverRestarted && !isDifferentDay && !missingBootOrDate) {
      if (elLoginOverlay) elLoginOverlay.style.display = 'none';
      await loadFolders();
      await loadMessages();
      loadContacts();
      loadCalendarEvents();
      loadTodos();
      loadEmailTemplates();
      loadMessageLikes();
      startAutoSync();
    } else if (isSessionAuth && serverRestarted) {
      endUserSession('O computador/servidor foi reiniciado. Faça login novamente.');
    } else if (isSessionAuth && isDifferentDay) {
      endUserSession('Novo dia iniciado. Faça login para o acesso de hoje.');
    } else {
      if (isSessionAuth) SESSION_STORE.clearAll();
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

  function renderLoginAutocomplete() {
    const box = document.getElementById('loginEmailSuggestions');
    if (!box || !elLoginEmailInput) return;

    const typed = elLoginEmailInput.value.trim().toLowerCase();
    if (!typed) {
      box.style.display = 'none';
      box.innerHTML = '';
      return;
    }

    const matches = savedEmails.filter(em => em.toLowerCase().includes(typed) && em.toLowerCase() !== typed);
    if (matches.length === 0) {
      box.style.display = 'none';
      box.innerHTML = '';
      return;
    }

    box.innerHTML = '';
    matches.slice(0, 6).forEach((email, idx) => {
      const item = document.createElement('div');
      item.className = `login-autocomplete-item ${idx === 0 ? 'active' : ''}`;
      item.innerHTML = `
        <i class="fa-regular fa-envelope" style="color: var(--ms-blue);"></i>
        <span>${escapeHtml(email)}</span>
      `;
      item.onmousedown = (e) => {
        e.preventDefault();
        elLoginEmailInput.value = email;
        box.style.display = 'none';
        if (elLoginPasswordInput) elLoginPasswordInput.focus();
      };
      box.appendChild(item);
    });

    box.style.display = 'block';
  }

  function toggleMobileFoldersDrawer(forceState) {
    const foldersPane = document.getElementById('foldersPane');
    const backdrop = document.getElementById('mobileDrawerBackdrop');
    if (!foldersPane) return;
    const nextState = typeof forceState === 'boolean' ? forceState : !foldersPane.classList.contains('mobile-open');
    foldersPane.classList.toggle('mobile-open', nextState);
    if (backdrop) backdrop.classList.toggle('mobile-open', nextState);
  }

  function closeMobileReadingPane() {
    document.body.classList.remove('mobile-reading-open');
  }

  function endUserSession(reasonMessage = '') {
    SESSION_STORE.clearAll();

    if (autoSyncTimer) {
      clearInterval(autoSyncTimer);
      autoSyncTimer = null;
    }
    stopDraftAutoSaveTimer();
    toggleMobileFoldersDrawer(false);
    closeMobileReadingPane();

    // Close open modals and menus
    if (elSettingsModal) elSettingsModal.style.display = 'none';
    if (elComposeModal) elComposeModal.style.display = 'none';
    const profileMenu = document.getElementById('userProfileMenu');
    if (profileMenu) profileMenu.style.display = 'none';
    const attModal = document.getElementById('attachmentPreviewModal') || document.getElementById('attPreviewModal');
    if (attModal) attModal.style.display = 'none';
    const contactModal = document.getElementById('contactModal');
    if (contactModal) contactModal.style.display = 'none';
    const calModal = document.getElementById('calendarEventModal');
    if (calModal) calModal.style.display = 'none';

    // Clear active mailbox UI behind the login screen
    messages = [];
    selectedMessage = null;
    selectedUids.clear();
    if (elMessageListScroll) elMessageListScroll.innerHTML = '';
    if (elReadingContentView) elReadingContentView.style.display = 'none';
    if (elNoSelectionView) elNoSelectionView.style.display = 'flex';

    showLoginScreen(reasonMessage);
    if (reasonMessage) {
      showToast(reasonMessage);
    }
  }

  function showLoginScreen(reasonMessage = '') {
    if (!elLoginOverlay) return;
    elLoginOverlay.style.display = 'flex';
    if (elLoginErrorBox) {
      if (reasonMessage) {
        elLoginErrorBox.style.display = 'block';
        elLoginErrorBox.innerHTML = `<i class="fa-regular fa-clock"></i> ${escapeHtml(reasonMessage)}`;
      } else {
        elLoginErrorBox.style.display = 'none';
      }
    }
    if (elLoginEmailInput) elLoginEmailInput.value = '';
    if (elLoginPasswordInput) elLoginPasswordInput.value = '';

    const box = document.getElementById('loginEmailSuggestions');
    if (box) {
      box.style.display = 'none';
      box.innerHTML = '';
    }

    setTimeout(() => {
      if (elLoginEmailInput) elLoginEmailInput.focus();
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
        const bootId = await fetchServerBootId();
        const isGmailLogin = email.endsWith('@gmail.com') || email.endsWith('@googlemail.com') || (data.imapHost || '').toLowerCase().includes('gmail.com');
        const cleanPassLogin = isGmailLogin ? password.replace(/\s+/g, '') : password;
        let passB64 = '';
        try {
          passB64 = btoa(unescape(encodeURIComponent(cleanPassLogin)));
        } catch (e) {}
        SESSION_STORE.setAuth('true');
        SESSION_STORE.setUser(email);
        if (passB64) SESSION_STORE.setPass(passB64);
        SESSION_STORE.addAccount({
          email,
          passB64,
          imapHost: data.imapHost || '',
          imapPort: data.imapPort || '',
          smtpHost: data.smtpHost || '',
          smtpPort: data.smtpPort || ''
        });
        if (bootId) SESSION_STORE.setBootId(bootId);
        SESSION_STORE.setLoginDate(getTodayDateKey());

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
        loadContacts();
        loadCalendarEvents();
        loadTodos();
        loadEmailTemplates();
        loadMessageLikes();
        startAutoSync();
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

  // Multi-account folders state
  const accountFoldersMap = {};
  const collapsedAccounts = new Set();

  function getActiveAccountEmail() {
    return (SESSION_STORE.getUser() || document.getElementById('cfgImapUser')?.value || '').trim().toLowerCase();
  }

  async function fetchFoldersForAccount(acc) {
    if (!acc || !acc.email || !acc.passB64) return [];
    const headers = {
      'x-imap-user': acc.email,
      'x-imap-pass': acc.passB64
    };
    if (acc.imapHost) headers['x-imap-host'] = acc.imapHost;
    if (acc.imapPort) headers['x-imap-port'] = String(acc.imapPort);
    if (acc.smtpHost) headers['x-smtp-host'] = acc.smtpHost;
    if (acc.smtpPort) headers['x-smtp-port'] = String(acc.smtpPort);
    const res = await fetch('/api/imap/folders', { headers });
    const data = await res.json();
    return (data.success && Array.isArray(data.folders)) ? data.folders : [];
  }

  // Load Folders (active account + any additional connected accounts)
  async function loadFolders(forceAllAccounts = false) {
    try {
      const activeEmail = getActiveAccountEmail();
      const res = await fetch('/api/imap/folders');
      const data = await res.json();
      if (data.success && data.folders) {
        folders = data.folders;
        if (activeEmail) {
          accountFoldersMap[activeEmail] = data.folders;
        }
        renderFolders();
      }

      const accounts = SESSION_STORE.getAccounts();
      for (const acc of accounts) {
        const key = acc.email.toLowerCase();
        if (key === activeEmail) continue;
        if (forceAllAccounts || !Array.isArray(accountFoldersMap[key])) {
          try {
            accountFoldersMap[key] = await fetchFoldersForAccount(acc);
            renderFolders();
          } catch (accErr) {
            console.warn(`Erro ao carregar pastas da conta ${acc.email}:`, accErr.message);
          }
        }
      }
      renderProfileConnectedAccounts();
    } catch (err) {
      console.error('Erro ao carregar pastas:', err);
    }
  }

  async function switchActiveAccount(targetEmail, targetFolderPath = 'INBOX', targetFolderName = 'Caixa de entrada') {
    const cleanTarget = String(targetEmail || '').trim().toLowerCase();
    const accounts = SESSION_STORE.getAccounts();
    const acc = accounts.find(a => a.email.toLowerCase() === cleanTarget);
    if (!acc) return;

    const currentActive = getActiveAccountEmail();
    if (currentActive !== cleanTarget) {
      SESSION_STORE.setActiveAccount(acc);
      currentFolder = targetFolderPath || 'INBOX';

      const localSig = localStorage.getItem('outlook_sig_' + cleanTarget);
      if (localSig) {
        try { sigState = { ...sigState, ...JSON.parse(localSig) }; } catch (e) {}
      }

      await loadConfig();
      if (Array.isArray(accountFoldersMap[cleanTarget])) {
        folders = accountFoldersMap[cleanTarget];
      } else {
        await loadFolders();
      }
      renderProfileConnectedAccounts();
      loadContacts();
      loadCalendarEvents();
      loadTodos();
      loadEmailTemplates();
      showToast(`Conta ativa: ${cleanTarget}`);
    }

    selectFolder(targetFolderPath || 'INBOX', targetFolderName || 'Caixa de entrada');
  }

  async function disconnectAccount(targetEmail) {
    const cleanTarget = String(targetEmail || '').trim().toLowerCase();
    const accounts = SESSION_STORE.getAccounts();
    if (accounts.length <= 1) {
      endUserSession();
      return;
    }
    if (!confirm(`Deseja desconectar a conta "${cleanTarget}"?`)) return;

    const wasActive = getActiveAccountEmail() === cleanTarget;
    delete accountFoldersMap[cleanTarget];
    collapsedAccounts.delete(cleanTarget);
    const remaining = SESSION_STORE.removeAccount(cleanTarget);

    if (remaining.length === 0) {
      endUserSession();
      return;
    }

    if (wasActive) {
      await switchActiveAccount(remaining[0].email, 'INBOX', 'Caixa de entrada');
    } else {
      renderFolders();
      renderProfileConnectedAccounts();
    }
    showToast(`Conta "${cleanTarget}" desconectada.`);
  }

  function renderProfileConnectedAccounts() {
    const listEl = document.getElementById('profileConnectedAccountsList');
    if (!listEl) return;
    const accounts = SESSION_STORE.getAccounts();
    const activeEmail = getActiveAccountEmail();

    if (accounts.length <= 1) {
      listEl.style.display = 'none';
      listEl.innerHTML = '';
      return;
    }

    listEl.style.display = 'block';
    listEl.innerHTML = '';
    accounts.forEach(acc => {
      const isAct = acc.email.toLowerCase() === activeEmail;
      const row = document.createElement('div');
      row.className = `profile-account-row ${isAct ? 'active' : ''}`;
      row.innerHTML = `
        <div class="profile-account-email" title="${escapeHtml(acc.email)}">
          <i class="fa-solid ${isAct ? 'fa-circle-check' : 'fa-envelope'}" style="font-size: 11px; color: ${isAct ? '#107c41' : 'var(--text-secondary)'};"></i>
          <span>${escapeHtml(acc.email)}</span>
        </div>
        <button type="button" class="profile-account-disconnect" title="Desconectar ${escapeHtml(acc.email)}" aria-label="Desconectar conta">
          <i class="fa-solid fa-xmark"></i>
        </button>
      `;
      row.onclick = (e) => {
        if (e.target.closest('.profile-account-disconnect')) return;
        const menu = document.getElementById('userProfileMenu');
        if (menu) menu.style.display = 'none';
        if (!isAct) switchActiveAccount(acc.email, 'INBOX', 'Caixa de entrada');
      };
      row.querySelector('.profile-account-disconnect').onclick = (e) => {
        e.stopPropagation();
        disconnectAccount(acc.email);
      };
      listEl.appendChild(row);
    });
  }

  function openAddAccountModal() {
    const modal = document.getElementById('addAccountModal');
    const form = document.getElementById('addAccountForm');
    const errBox = document.getElementById('addAccErrorBox');
    const emailInput = document.getElementById('addAccEmailInput');
    if (form) form.reset();
    if (errBox) { errBox.style.display = 'none'; errBox.textContent = ''; }
    if (modal) modal.style.display = 'flex';
    setTimeout(() => emailInput?.focus(), 50);
  }

  function closeAddAccountModal() {
    const modal = document.getElementById('addAccountModal');
    if (modal) modal.style.display = 'none';
  }

  async function handleAddAccountSubmit(e) {
    e.preventDefault();
    const email = (document.getElementById('addAccEmailInput')?.value || '').trim().toLowerCase();
    const password = document.getElementById('addAccPasswordInput')?.value || '';
    const imapHost = (document.getElementById('addAccImapHost')?.value || '').trim();
    const imapPort = (document.getElementById('addAccImapPort')?.value || '').trim();
    const smtpHost = (document.getElementById('addAccSmtpHost')?.value || '').trim();
    const smtpPort = (document.getElementById('addAccSmtpPort')?.value || '').trim();
    const errBox = document.getElementById('addAccErrorBox');
    const btn = document.getElementById('btnSubmitAddAccount');

    if (!email || !password) return;
    if (errBox) errBox.style.display = 'none';
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>Conectando...</span>';
    }

    try {
      const res = await fetch('/api/auth/verify-account', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password, imapHost, imapPort, smtpHost, smtpPort })
      });
      const data = await res.json();
      if (data.success) {
        const isGmailAcc = email.endsWith('@gmail.com') || email.endsWith('@googlemail.com') || (data.imapHost || '').toLowerCase().includes('gmail.com');
        const cleanPassForStore = isGmailAcc ? password.replace(/\s+/g, '') : password;
        let passB64 = '';
        try { passB64 = btoa(unescape(encodeURIComponent(cleanPassForStore))); } catch (err) {}
        SESSION_STORE.addAccount({
          email,
          passB64,
          imapHost: data.imapHost || '',
          imapPort: data.imapPort || '',
          smtpHost: data.smtpHost || '',
          smtpPort: data.smtpPort || ''
        });
        accountFoldersMap[email] = Array.isArray(data.folders) ? data.folders : [];
        collapsedAccounts.delete(email);
        closeAddAccountModal();
        renderFolders();
        renderProfileConnectedAccounts();
        showToast(`Conta "${email}" conectada na árvore de pastas!`);
      } else {
        if (errBox) {
          errBox.style.display = 'block';
          errBox.innerHTML = `<i class="fa-solid fa-circle-exclamation"></i> ${escapeHtml(data.message || 'Falha na autenticação.')}`;
        }
      }
    } catch (err) {
      if (errBox) {
        errBox.style.display = 'block';
        errBox.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> Erro de conexão: ${escapeHtml(err.message)}`;
      }
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = '<i class="fa-solid fa-plug-circle-plus"></i> <span>Conectar Conta</span>';
      }
    }
  }

  const collapsedFolders = new Set();

  function getFolderIcon(folder) {
    if (!folder.isSystem) return 'fa-regular fa-folder';
    const name = (folder.name || folder.path).toLowerCase();
    if (name.includes('entrada') || name.includes('inbox')) return 'fa-solid fa-inbox';
    if (name.includes('enviad') || name.includes('sent')) return 'fa-regular fa-paper-plane';
    if (name.includes('rascunh') || name.includes('draft')) return 'fa-regular fa-file-lines';
    if (name.includes('spam') || name.includes('junk') || name.includes('lixo eletr')) return 'fa-solid fa-shield-halved';
    if (name.includes('lixeira') || name.includes('excluíd') || name.includes('trash')) return 'fa-regular fa-trash-can';
    if (name.includes('arquivo') || name.includes('archive')) return 'fa-regular fa-folder-archive';
    return 'fa-regular fa-folder';
  }

  function getSystemFolderOrder(folder) {
    if (!folder.isSystem) return 100;
    const name = (folder.name || '').toLowerCase();
    if (name.includes('entrada') || folder.path.toUpperCase() === 'INBOX') return 1;
    if (name.includes('enviad')) return 2;
    if (name.includes('rascunh')) return 3;
    if (name.includes('lixo eletr') || name.includes('spam') || name.includes('junk')) return 4;
    if (name.includes('excluíd') || name.includes('lixeira') || name.includes('trash')) return 5;
    if (name.includes('arquivo')) return 6;
    return 50;
  }

  function buildFolderTree(folderList) {
    const byPath = new Map();
    const nodes = folderList.map(f => {
      const node = { ...f, children: [] };
      byPath.set(f.path.toLowerCase(), node);
      return node;
    });

    const roots = [];
    nodes.forEach(node => {
      const pKey = (node.parentPath || '').toLowerCase();
      if (!node.isSystem && pKey && byPath.has(pKey) && pKey !== node.path.toLowerCase()) {
        byPath.get(pKey).children.push(node);
      } else {
        roots.push(node);
      }
    });

    const sortNodes = (arr, isRoot) => {
      arr.sort((a, b) => {
        if (isRoot) {
          const oa = getSystemFolderOrder(a);
          const ob = getSystemFolderOrder(b);
          if (oa !== ob) return oa - ob;
        }
        return (a.name || a.path).localeCompare(b.name || b.path, 'pt-BR', { sensitivity: 'base' });
      });
      arr.forEach(n => sortNodes(n.children, false));
    };
    sortNodes(roots, true);
    return roots;
  }

  function flattenFolderTree(nodes, depth = 0, out = []) {
    nodes.forEach(node => {
      out.push({ folder: node, depth });
      if (node.children && node.children.length > 0) {
        flattenFolderTree(node.children, depth + 1, out);
      }
    });
    return out;
  }

  function openCreateFolderModal(defaultParentPath = '') {
    const modal = document.getElementById('createFolderModal');
    const select = document.getElementById('newFolderParentSelect');
    const input = document.getElementById('newFolderNameInput');
    if (!modal || !select || !input) return;

    select.innerHTML = '<option value="">Raiz (nível principal)</option>';
    const flat = flattenFolderTree(buildFolderTree(folders));
    flat.forEach(({ folder, depth }) => {
      const opt = document.createElement('option');
      opt.value = folder.path;
      const prefix = depth > 0 ? `${'\u00A0\u00A0'.repeat(depth)}↳ ` : '';
      opt.textContent = `${prefix}${folder.name}`;
      select.appendChild(opt);
    });

    if (defaultParentPath) {
      select.value = defaultParentPath;
    } else {
      select.value = '';
    }

    input.value = '';
    modal.style.display = 'flex';
    setTimeout(() => input.focus(), 50);
  }

  function closeCreateFolderModal() {
    const modal = document.getElementById('createFolderModal');
    if (modal) modal.style.display = 'none';
  }

  async function handleCreateFolderSubmit(e) {
    e.preventDefault();
    const select = document.getElementById('newFolderParentSelect');
    const input = document.getElementById('newFolderNameInput');
    const btn = document.getElementById('btnSubmitCreateFolder');
    const name = (input?.value || '').trim();
    const parentPath = select?.value || '';

    if (!name) {
      showToast('Informe o nome da subpasta.');
      input?.focus();
      return;
    }

    if (btn) {
      btn.disabled = true;
      btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>Criando...</span>';
    }

    try {
      const res = await fetch('/api/imap/folders', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, parentPath })
      });
      const data = await res.json();
      if (data.success) {
        if (parentPath) collapsedFolders.delete(parentPath.toLowerCase());
        if (Array.isArray(data.folders)) {
          folders = data.folders;
          renderFolders();
        } else {
          await loadFolders();
        }
        closeCreateFolderModal();
        showToast(`Subpasta "${name}" criada com sucesso!`);
      } else {
        showToast(data.error || 'Erro ao criar subpasta.');
      }
    } catch (err) {
      showToast('Erro ao criar subpasta: ' + err.message);
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = '<i class="fa-solid fa-folder-plus"></i> <span>Criar Subpasta</span>';
      }
    }
  }

  async function handleDeleteFolder(folder) {
    if (!folder || folder.isSystem) return;
    if (!confirm(`Deseja excluir a subpasta "${folder.name}"?\n\nAtenção: as mensagens contidas nela também podem ser removidas no servidor.`)) {
      return;
    }

    try {
      const res = await fetch('/api/imap/folders', {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ path: folder.path })
      });
      const data = await res.json();
      if (data.success) {
        if (Array.isArray(data.folders)) {
          folders = data.folders;
        } else {
          await loadFolders();
        }
        if (currentFolder.toLowerCase() === folder.path.toLowerCase()) {
          selectFolder('INBOX', 'Caixa de entrada');
        } else {
          renderFolders();
        }
        showToast(`Subpasta "${folder.name}" excluída.`);
      } else {
        showToast(data.error || 'Erro ao excluir subpasta.');
      }
    } catch (err) {
      showToast('Erro ao excluir subpasta: ' + err.message);
    }
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

    const activeEmail = getActiveAccountEmail();
    const accounts = SESSION_STORE.getAccounts();
    const accountEntries = accounts.length > 0
      ? accounts.map(acc => ({
          email: acc.email.toLowerCase(),
          list: (acc.email.toLowerCase() === activeEmail)
            ? folders
            : (accountFoldersMap[acc.email.toLowerCase()] || [])
        }))
      : [{ email: activeEmail || 'conta', list: folders }];

    const renderNode = (accEmail, isAccActive, folder, depth) => {
      const hasChildren = folder.children && folder.children.length > 0;
      const collapseKey = `${accEmail}::${folder.path.toLowerCase()}`;
      const isCollapsed = collapsedFolders.has(collapseKey);

      const item = document.createElement('div');
      const isFolderActive = isAccActive && folder.path.toLowerCase() === currentFolder.toLowerCase();
      item.className = `folder-item ${isFolderActive ? 'active' : ''}`;
      item.setAttribute('data-path', folder.path);
      item.setAttribute('data-account', accEmail);
      item.style.paddingLeft = `${14 + depth * 15}px`;

      const left = document.createElement('div');
      left.className = 'folder-left';

      if (hasChildren) {
        const expandBtn = document.createElement('button');
        expandBtn.type = 'button';
        expandBtn.className = 'folder-expand-btn';
        expandBtn.title = isCollapsed ? 'Expandir subpastas' : 'Recolher subpastas';
        expandBtn.innerHTML = `<i class="fa-solid ${isCollapsed ? 'fa-chevron-right' : 'fa-chevron-down'}"></i>`;
        expandBtn.onclick = (e) => {
          e.stopPropagation();
          if (collapsedFolders.has(collapseKey)) collapsedFolders.delete(collapseKey);
          else collapsedFolders.add(collapseKey);
          renderFolders();
        };
        left.appendChild(expandBtn);
      } else {
        const spacer = document.createElement('span');
        spacer.className = 'folder-expand-spacer';
        left.appendChild(spacer);
      }

      const icon = document.createElement('i');
      icon.className = `${getFolderIcon(folder)} folder-icon`;
      left.appendChild(icon);

      const label = document.createElement('span');
      label.textContent = folder.name;
      label.title = folder.name;
      left.appendChild(label);

      const right = document.createElement('div');
      right.className = 'folder-right';

      const actions = document.createElement('div');
      actions.className = 'folder-row-actions';

      const addSubBtn = document.createElement('button');
      addSubBtn.type = 'button';
      addSubBtn.className = 'folder-action-btn';
      addSubBtn.title = `Criar subpasta em ${folder.name}`;
      addSubBtn.innerHTML = '<i class="fa-solid fa-plus"></i>';
      addSubBtn.onclick = async (e) => {
        e.stopPropagation();
        if (!isAccActive) {
          await switchActiveAccount(accEmail, folder.path, folder.name);
        }
        openCreateFolderModal(folder.path);
      };
      actions.appendChild(addSubBtn);

      if (!folder.isSystem) {
        const delBtn = document.createElement('button');
        delBtn.type = 'button';
        delBtn.className = 'folder-action-btn delete';
        delBtn.title = `Excluir subpasta ${folder.name}`;
        delBtn.innerHTML = '<i class="fa-regular fa-trash-can"></i>';
        delBtn.onclick = async (e) => {
          e.stopPropagation();
          if (!isAccActive) {
            await switchActiveAccount(accEmail, 'INBOX', 'Caixa de entrada');
          }
          handleDeleteFolder(folder);
        };
        actions.appendChild(delBtn);
      }

      right.appendChild(actions);

      if (folder.unread > 0) {
        const badge = document.createElement('span');
        badge.className = 'folder-badge';
        badge.textContent = folder.unread;
        right.appendChild(badge);
      }

      item.appendChild(left);
      item.appendChild(right);

      item.onclick = (e) => {
        e.preventDefault();
        if (!isAccActive) {
          switchActiveAccount(accEmail, folder.path, folder.name);
        } else {
          selectFolder(folder.path, folder.name);
        }
      };

      elFoldersList.appendChild(item);

      if (hasChildren && !isCollapsed) {
        folder.children.forEach(child => renderNode(accEmail, isAccActive, child, depth + 1));
      }
    };

    accountEntries.forEach(({ email, list }) => {
      const isAccActive = email === activeEmail;
      const isAccCollapsed = collapsedAccounts.has(email);

      const accHeader = document.createElement('div');
      accHeader.className = `folder-account-header ${isAccActive ? 'active-account' : ''}`;
      accHeader.title = email;

      const accLeft = document.createElement('div');
      accLeft.className = 'folder-account-left';
      accLeft.innerHTML = `
        <i class="fa-solid ${isAccCollapsed ? 'fa-chevron-right' : 'fa-chevron-down'}" style="font-size: 9px; width: 12px; text-align: center;"></i>
        <i class="fa-solid fa-envelope-circle-check" style="font-size: 12px; color: ${isAccActive ? 'var(--ms-blue)' : 'var(--text-secondary)'};"></i>
        <span class="folder-account-label">${escapeHtml(email)}</span>
      `;

      const accActions = document.createElement('div');
      accActions.className = 'folder-account-actions';

      const accAddBtn = document.createElement('button');
      accAddBtn.type = 'button';
      accAddBtn.className = 'folder-action-btn';
      accAddBtn.title = `Nova pasta em ${email}`;
      accAddBtn.innerHTML = '<i class="fa-solid fa-folder-plus"></i>';
      accAddBtn.onclick = async (e) => {
        e.stopPropagation();
        if (!isAccActive) await switchActiveAccount(email, 'INBOX', 'Caixa de entrada');
        openCreateFolderModal('INBOX');
      };
      accActions.appendChild(accAddBtn);

      if (accountEntries.length > 1) {
        const accDelBtn = document.createElement('button');
        accDelBtn.type = 'button';
        accDelBtn.className = 'folder-action-btn delete';
        accDelBtn.title = `Desconectar ${email}`;
        accDelBtn.innerHTML = '<i class="fa-solid fa-xmark"></i>';
        accDelBtn.onclick = (e) => {
          e.stopPropagation();
          disconnectAccount(email);
        };
        accActions.appendChild(accDelBtn);
      }

      accHeader.appendChild(accLeft);
      accHeader.appendChild(accActions);

      accHeader.onclick = () => {
        if (collapsedAccounts.has(email)) collapsedAccounts.delete(email);
        else collapsedAccounts.add(email);
        renderFolders();
      };

      elFoldersList.appendChild(accHeader);

      if (!isAccCollapsed) {
        const tree = buildFolderTree(list);
        tree.forEach(rootNode => renderNode(email, isAccActive, rootNode, 0));
      }
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

  function adjustFolderUnreadCount(folderPath, delta) {
    if (!folderPath || !delta) return;
    const f = folders.find(item => item.path.toLowerCase() === folderPath.toLowerCase());
    if (f) {
      f.unread = Math.max(0, (f.unread || 0) + delta);
      renderFolders();
    }
  }

  function selectFolder(path, name) {
    currentFolder = path;
    elCurrentFolderTitle.textContent = name || path;
    selectedMessage = null;
    selectedUids.clear();
    toggleMobileFoldersDrawer(false);
    closeMobileReadingPane();

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

  let autoSyncTimer = null;
  let isBackgroundSyncing = false;

  async function backgroundSyncEmails() {
    if (isBackgroundSyncing) return;
    if (SESSION_STORE.getAuth() !== 'true') return;
    if (elLoginOverlay && elLoginOverlay.style.display === 'flex') return;

    isBackgroundSyncing = true;
    try {
      const folderBefore = currentFolder;
      const prevUids = new Set(messages.map(m => String(m.uid)));
      const prevUnreadCount = messages.filter(m => !m.seen).length;

      // Sync folders and current folder messages in parallel
      const query = new URLSearchParams({
        folder: folderBefore,
        filter: currentFilter,
        search: currentSearch
      });

      const [foldersRes, msgsRes] = await Promise.all([
        fetch('/api/imap/folders').then(r => r.json()).catch(() => null),
        fetch(`/api/imap/messages?${query.toString()}`).then(r => r.json()).catch(() => null)
      ]);

      if (foldersRes && foldersRes.success && Array.isArray(foldersRes.folders)) {
        folders = foldersRes.folders;
        renderFolders();
      }

      // Only update message list if user is still on the same folder
      if (msgsRes && msgsRes.success && currentFolder === folderBefore) {
        const freshMessages = msgsRes.messages || [];
        const newArrivals = freshMessages.filter(m => !prevUids.has(String(m.uid)));
        const hasChanges =
          newArrivals.length > 0 ||
          freshMessages.length !== messages.length ||
          freshMessages.filter(m => !m.seen).length !== prevUnreadCount;

        if (hasChanges) {
          messages = freshMessages;
          if (elMessageCountLabel) {
            elMessageCountLabel.textContent = `${messages.length} ${messages.length === 1 ? 'mensagem' : 'mensagens'}`;
          }
          renderMessageList();

          if (newArrivals.length > 0 && currentSearch === '') {
            showToast(`${newArrivals.length} ${newArrivals.length === 1 ? 'novo e-mail recebido' : 'novos e-mails recebidos'}!`);
          }
        }
      }
    } catch (err) {
      console.warn('[AutoSync 30s] Aviso na sincronização:', err.message);
    } finally {
      isBackgroundSyncing = false;
    }
  }

  function startAutoSync() {
    if (autoSyncTimer) clearInterval(autoSyncTimer);
    autoSyncTimer = setInterval(() => {
      backgroundSyncEmails();
    }, 30000); // Every 30 seconds
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

  function getLikeKeyForMessage(msg) {
    if (!msg) return '';
    const fromAddr = String(msg.from?.address || '').trim().toLowerCase();
    const cleanSubject = String(msg.subject || '(sem assunto)')
      .replace(/^(re|fwd|enc|res)\s*:\s*/i, '')
      .trim()
      .toLowerCase();
    return `${fromAddr}||${cleanSubject}`;
  }

  function getLikesForMessage(msg) {
    const key = getLikeKeyForMessage(msg);
    return (key && Array.isArray(messageLikesMap[key])) ? messageLikesMap[key] : [];
  }

  function isLikedByCurrentUser(msg) {
    const me = (SESSION_STORE.getUser() || '').trim().toLowerCase();
    if (!me) return false;
    return getLikesForMessage(msg).some(item => (item.user || '').toLowerCase() === me);
  }

  function formatLikesTooltip(likes) {
    if (!likes || likes.length === 0) return 'Curtir com joinha (confirmar leitura)';
    const parts = likes.map(l => {
      const who = l.userName || l.user || 'Usuário';
      const when = l.likedAt ? new Date(l.likedAt).toLocaleString('pt-BR') : '';
      return when ? `${who} em ${when}` : who;
    });
    return `👍 Leitura confirmada por: ${parts.join(', ')}`;
  }

  async function loadMessageLikes() {
    try {
      const res = await fetch('/api/messages/likes');
      const data = await res.json();
      if (data.success && data.likes) {
        messageLikesMap = data.likes;
        if (messages.length > 0) renderMessageList();
        if (selectedMessage) updateLikeButtonUI(selectedMessage);
      }
    } catch (e) {}
  }

  function updateLikeButtonUI(msg) {
    const btn = document.getElementById('btnLikeMessage');
    if (!btn) return;
    if (!msg) {
      btn.classList.remove('liked');
      btn.innerHTML = '<i class="fa-regular fa-thumbs-up"></i><span>Curtir</span>';
      return;
    }
    const likes = getLikesForMessage(msg);
    const likedByMe = isLikedByCurrentUser(msg);
    btn.classList.toggle('liked', likedByMe || likes.length > 0);
    btn.title = likes.length > 0
      ? formatLikesTooltip(likes)
      : 'Curtir com joinha (envia confirmação de leitura ao remetente)';
    const countSuffix = likes.length > 1 ? ` (${likes.length})` : '';
    btn.innerHTML = `<i class="fa-${(likedByMe || likes.length > 0) ? 'solid' : 'regular'} fa-thumbs-up"></i><span>${likedByMe ? 'Curtido' : 'Curtir'}${countSuffix}</span>`;
    btn.onclick = () => toggleLikeEmail(msg);
  }

  async function toggleLikeEmail(msg) {
    if (!msg) return;
    const btn = document.getElementById('btnLikeMessage');
    if (btn) btn.disabled = true;
    try {
      const res = await fetch('/api/messages/like', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fromEmail: msg.from?.address || '',
          fromName: msg.from?.name || '',
          subject: msg.subject || '',
          uid: msg.uid,
          folder: msg.folder || currentFolder || 'INBOX'
        })
      });
      const data = await res.json();
      if (data.success) {
        if (data.likes) messageLikesMap = data.likes;
        if (data.liked && !msg.seen) {
          msg.seen = true;
          const target = messages.find(m => String(m.uid) === String(msg.uid));
          if (target) target.seen = true;
        }
        renderMessageList();
        if (selectedMessage && String(selectedMessage.uid) === String(msg.uid)) {
          updateLikeButtonUI(selectedMessage);
        }
        if (data.liked) {
          showToast(
            data.notifiedSender
              ? '👍 Joinha enviado! O remetente recebeu a confirmação de leitura.'
              : '👍 E-mail curtido como lido!'
          );
        } else {
          showToast('Curtida removida.');
        }
      } else {
        showToast(data.error || 'Erro ao registrar curtida.');
      }
    } catch (err) {
      showToast('Erro ao enviar joinha: ' + err.message);
    } finally {
      if (btn) btn.disabled = false;
    }
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

    const isBottomLayout = readingPaneLayout === 'bottom' && window.innerWidth > 768;

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
        if (e.target.closest('.card-action-btn') || e.target.closest('.msg-select-chk') || e.target.closest('.row-like-btn')) return;
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

      const msgLikes = getLikesForMessage(msg);
      const isLiked = msgLikes.length > 0;
      const likeTooltip = formatLikesTooltip(msgLikes);

      const likeBtnHtml = `
        <button type="button" class="row-like-btn ${isLiked ? 'liked' : ''}" onclick="window.quickAction(event, '${msg.uid}', 'like')" title="${escapeHtml(likeTooltip)}" aria-label="Curtir com joinha">
          <i class="fa-${isLiked ? 'solid' : 'regular'} fa-thumbs-up"></i>
        </button>
      `;

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
        // Compact Table-Row Layout: Check | Carta + Joinha + Clipe + Flag | De/E-mail | Assunto | Data | Ações
        const statusIcons = [
          !msg.seen ? '<i class="fa-solid fa-circle" style="font-size: 7px; color: var(--ms-blue);" title="Não lido"></i>' : '<i class="fa-regular fa-envelope-open" style="opacity: 0.45;" title="Lido"></i>',
          likeBtnHtml,
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
        // Compact Lateral Card Layout with left checkbox + like button before sender
        card.innerHTML = `
          <div class="message-card-lateral-wrap">
            ${chkHtml}
            <div class="message-card-lateral-body">
              <div class="message-card-top" style="margin-bottom: 2px;">
                <span class="message-sender" style="display: inline-flex; align-items: center; gap: 5px;">${likeBtnHtml}<span>${escapeHtml(displayName)}</span></span>
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
    else if (action === 'like') toggleLikeEmail(msg);
  };

  // Select Message and Fetch Details
  async function selectMessage(msg) {
    if (!msg) return;

    try {
      selectedMessage = msg;
      toggleMobileFoldersDrawer(false);
      document.body.classList.add('mobile-reading-open');

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

      // Mark seen locally and update unread badge immediately
      const wasUnread = !msg.seen;
      const targetFolder = msg.folder || currentFolder || 'INBOX';
      msg.seen = true;
      if (wasUnread) {
        adjustFolderUnreadCount(targetFolder, -1);
        renderMessageList();
      }

      // 4. Fetch full email content and attachments from the server
      const res = await fetch(`/api/imap/message/${msg.uid}?folder=${encodeURIComponent(targetFolder)}`);
      const data = await res.json();
      
      if (data.success && data.message) {
        selectedMessage = data.message;
        renderReadingDetail(data.message);
        if (wasUnread) {
          loadFolders();
        }
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

  const externalScriptPromises = {};
  function loadExternalScriptOnce(src) {
    if (externalScriptPromises[src]) return externalScriptPromises[src];
    externalScriptPromises[src] = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.async = true;
      s.onload = () => resolve();
      s.onerror = () => {
        delete externalScriptPromises[src];
        reject(new Error('Falha ao carregar biblioteca de visualização.'));
      };
      document.head.appendChild(s);
    });
    return externalScriptPromises[src];
  }

  let activePreviewSessionId = 0;

  async function renderPdfInCanvas(bodyEl, arrayBuffer, sessionId) {
    if (!window.pdfjsLib) {
      await loadExternalScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js');
    }
    const pdfjsLib = window.pdfjsLib;
    if (!pdfjsLib) {
      throw new Error('Biblioteca PDF indisponível.');
    }
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    const loadingTask = pdfjsLib.getDocument({
      data: new Uint8Array(arrayBuffer.slice(0)),
      cMapUrl: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/cmaps/',
      cMapPacked: true
    });
    const pdfDoc = await loadingTask.promise;
    if (sessionId !== activePreviewSessionId) return;

    bodyEl.classList.add('is-document-mode');
    bodyEl.innerHTML = `
      <div class="att-viewer-shell">
        <div class="att-viewer-toolbar">
          <span><i class="fa-regular fa-file-pdf" style="color: #d13438; margin-right: 4px;"></i><strong id="pdfPageInfo">${pdfDoc.numPages} ${pdfDoc.numPages === 1 ? 'página' : 'páginas'}</strong></span>
          <div class="att-viewer-zoom-controls">
            <button type="button" class="att-viewer-zoom-btn" id="btnPdfZoomOut" title="Diminuir zoom"><i class="fa-solid fa-minus"></i></button>
            <button type="button" class="att-viewer-zoom-btn" id="btnPdfZoomReset" title="Ajustar à largura da tela">100%</button>
            <button type="button" class="att-viewer-zoom-btn" id="btnPdfZoomIn" title="Aumentar zoom"><i class="fa-solid fa-plus"></i></button>
          </div>
        </div>
        <div class="att-viewer-pages" id="pdfPagesContainer"></div>
      </div>
    `;

    const pagesContainer = document.getElementById('pdfPagesContainer');
    const btnZoomOut = document.getElementById('btnPdfZoomOut');
    const btnZoomReset = document.getElementById('btnPdfZoomReset');
    const btnZoomIn = document.getElementById('btnPdfZoomIn');

    let zoomFactor = 1.0;
    let renderRunId = 0;

    const renderAllPages = async () => {
      const currentRun = ++renderRunId;
      if (!pagesContainer || sessionId !== activePreviewSessionId) return;

      pagesContainer.innerHTML = '';
      if (btnZoomReset) btnZoomReset.textContent = `${Math.round(zoomFactor * 100)}%`;

      const firstPage = await pdfDoc.getPage(1);
      const unscaledVp = firstPage.getViewport({ scale: 1 });
      const availWidth = Math.max(240, (pagesContainer.clientWidth || window.innerWidth) - 24);
      const fitScale = Math.min(availWidth / unscaledVp.width, 1.45);
      const effectiveScale = fitScale * zoomFactor;
      const dpr = Math.min(window.devicePixelRatio || 1, 2.5);

      pagesContainer.style.alignItems = (unscaledVp.width * effectiveScale > availWidth) ? 'flex-start' : 'center';

      for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
        if (currentRun !== renderRunId || sessionId !== activePreviewSessionId) return;

        const page = pageNum === 1 ? firstPage : await pdfDoc.getPage(pageNum);
        const viewport = page.getViewport({ scale: effectiveScale });

        const wrap = document.createElement('div');
        wrap.className = 'att-pdf-page-wrap';

        const canvas = document.createElement('canvas');
        canvas.className = 'att-pdf-canvas';
        canvas.width = Math.floor(viewport.width * dpr);
        canvas.height = Math.floor(viewport.height * dpr);
        canvas.style.width = `${Math.floor(viewport.width)}px`;
        canvas.style.height = `${Math.floor(viewport.height)}px`;

        const badge = document.createElement('span');
        badge.className = 'att-pdf-page-badge';
        badge.textContent = `${pageNum} / ${pdfDoc.numPages}`;

        wrap.appendChild(canvas);
        if (pdfDoc.numPages > 1) wrap.appendChild(badge);
        pagesContainer.appendChild(wrap);

        const ctx = canvas.getContext('2d', { alpha: false });
        const transform = dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : null;
        await page.render({ canvasContext: ctx, viewport, transform }).promise;
      }
    };

    if (btnZoomOut) {
      btnZoomOut.onclick = () => {
        zoomFactor = Math.max(0.5, Math.round((zoomFactor - 0.25) * 100) / 100);
        renderAllPages();
      };
    }
    if (btnZoomIn) {
      btnZoomIn.onclick = () => {
        zoomFactor = Math.min(3.0, Math.round((zoomFactor + 0.25) * 100) / 100);
        renderAllPages();
      };
    }
    if (btnZoomReset) {
      btnZoomReset.onclick = () => {
        zoomFactor = 1.0;
        renderAllPages();
      };
    }

    await renderAllPages();
  }

  function renderImageInViewer(bodyEl, blobUrl, filename) {
    bodyEl.classList.add('is-document-mode');
    bodyEl.innerHTML = `
      <div class="att-viewer-shell">
        <div class="att-viewer-toolbar">
          <span><i class="fa-regular fa-file-image" style="color: var(--ms-blue); margin-right: 4px;"></i>Visualização de Imagem</span>
          <div class="att-viewer-zoom-controls">
            <button type="button" class="att-viewer-zoom-btn" id="btnImgZoomOut" title="Diminuir zoom"><i class="fa-solid fa-minus"></i></button>
            <button type="button" class="att-viewer-zoom-btn" id="btnImgZoomReset" title="Ajustar à tela">100%</button>
            <button type="button" class="att-viewer-zoom-btn" id="btnImgZoomIn" title="Aumentar zoom"><i class="fa-solid fa-plus"></i></button>
          </div>
        </div>
        <div class="att-viewer-pages" id="imgPreviewWrap" style="justify-content: center;">
          <img id="attZoomableImg" src="${blobUrl}" alt="${escapeHtml(filename)}" style="max-width: 100%; max-height: 100%; object-fit: contain; border-radius: 4px; box-shadow: 0 2px 12px rgba(0,0,0,0.18); transition: width 0.15s;">
        </div>
      </div>
    `;

    const imgEl = document.getElementById('attZoomableImg');
    const wrapEl = document.getElementById('imgPreviewWrap');
    const btnOut = document.getElementById('btnImgZoomOut');
    const btnReset = document.getElementById('btnImgZoomReset');
    const btnIn = document.getElementById('btnImgZoomIn');
    let imgZoom = 1.0;

    const applyImgZoom = () => {
      if (!imgEl || !wrapEl) return;
      if (btnReset) btnReset.textContent = `${Math.round(imgZoom * 100)}%`;
      if (imgZoom === 1.0) {
        imgEl.style.maxWidth = '100%';
        imgEl.style.maxHeight = '100%';
        imgEl.style.width = 'auto';
        wrapEl.style.alignItems = 'center';
        wrapEl.style.justifyContent = 'center';
      } else {
        imgEl.style.maxWidth = 'none';
        imgEl.style.maxHeight = 'none';
        imgEl.style.width = `${Math.round(imgZoom * 100)}%`;
        wrapEl.style.alignItems = imgZoom > 1 ? 'flex-start' : 'center';
        wrapEl.style.justifyContent = 'flex-start';
      }
    };

    if (btnOut) btnOut.onclick = () => { imgZoom = Math.max(0.5, Math.round((imgZoom - 0.25) * 100) / 100); applyImgZoom(); };
    if (btnIn) btnIn.onclick = () => { imgZoom = Math.min(3.5, Math.round((imgZoom + 0.25) * 100) / 100); applyImgZoom(); };
    if (btnReset) btnReset.onclick = () => { imgZoom = 1.0; applyImgZoom(); };
  }

  function parseCsvRows(text) {
    const firstLine = (text.split(/\r?\n/)[0] || '');
    const semiCount = (firstLine.match(/;/g) || []).length;
    const commaCount = (firstLine.match(/,/g) || []).length;
    const tabCount = (firstLine.match(/\t/g) || []).length;
    const sep = tabCount > semiCount && tabCount > commaCount ? '\t' : (semiCount > commaCount ? ';' : ',');

    return text.split(/\r?\n/).filter(l => l.trim().length > 0).slice(0, 500).map(line => {
      const cells = [];
      let cur = '';
      let inQuotes = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          if (inQuotes && line[i + 1] === '"') { cur += '"'; i++; }
          else inQuotes = !inQuotes;
        } else if (ch === sep && !inQuotes) {
          cells.push(cur);
          cur = '';
        } else {
          cur += ch;
        }
      }
      cells.push(cur);
      return cells;
    });
  }

  function renderCsvInViewer(bodyEl, text) {
    const rows = parseCsvRows(text);
    const tableHtml = rows.map((row, rIdx) => {
      const tag = rIdx === 0 ? 'th' : 'td';
      return `<tr>${row.map(c => `<${tag}>${escapeHtml(c)}</${tag}>`).join('')}</tr>`;
    }).join('');
    bodyEl.classList.add('is-document-mode');
    bodyEl.innerHTML = `
      <div class="att-viewer-shell">
        <div class="att-sheet-table-wrap">
          <table class="att-sheet-table">${tableHtml}</table>
        </div>
      </div>
    `;
  }

  async function renderExcelInViewer(bodyEl, arrayBuffer, sessionId) {
    if (!window.XLSX) {
      await loadExternalScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
    }
    if (sessionId !== activePreviewSessionId) return;
    const wb = window.XLSX.read(arrayBuffer, { type: 'array' });
    const sheetNames = wb.SheetNames || [];
    if (sheetNames.length === 0) throw new Error('Planilha vazia.');

    bodyEl.classList.add('is-document-mode');
    bodyEl.innerHTML = `
      <div class="att-viewer-shell">
        <div class="att-sheet-tabs" id="attSheetTabs"></div>
        <div class="att-sheet-table-wrap" id="attSheetContent"></div>
      </div>
    `;
    const tabsEl = document.getElementById('attSheetTabs');
    const contentEl = document.getElementById('attSheetContent');

    const showSheet = (name) => {
      if (!contentEl) return;
      const ws = wb.Sheets[name];
      const html = window.XLSX.utils.sheet_to_html(ws, { editable: false });
      contentEl.innerHTML = html;
      const tbl = contentEl.querySelector('table');
      if (tbl) tbl.className = 'att-sheet-table';
      if (tabsEl) {
        tabsEl.querySelectorAll('.att-sheet-tab').forEach(b => {
          b.classList.toggle('active', b.getAttribute('data-sheet') === name);
        });
      }
    };

    sheetNames.forEach((name, idx) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `att-sheet-tab ${idx === 0 ? 'active' : ''}`;
      btn.setAttribute('data-sheet', name);
      btn.textContent = name;
      btn.onclick = () => showSheet(name);
      if (tabsEl) tabsEl.appendChild(btn);
    });
    if (sheetNames.length <= 1 && tabsEl) tabsEl.style.display = 'none';
    showSheet(sheetNames[0]);
  }

  async function renderDocxInViewer(bodyEl, arrayBuffer, sessionId) {
    if (!window.mammoth) {
      await loadExternalScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js');
    }
    if (sessionId !== activePreviewSessionId) return;
    const result = await window.mammoth.convertToHtml({ arrayBuffer });
    if (sessionId !== activePreviewSessionId) return;

    bodyEl.classList.add('is-document-mode');
    bodyEl.innerHTML = `
      <div class="att-viewer-shell">
        <div class="att-viewer-pages">
          <div class="att-docx-paper">${result.value || '<p style="color: #605e5c;">(Documento sem conteúdo de texto)</p>'}</div>
        </div>
      </div>
    `;
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

    const sessionId = ++activePreviewSessionId;

    if (activeAttBlobUrl) {
      URL.revokeObjectURL(activeAttBlobUrl);
      activeAttBlobUrl = null;
    }

    bodyEl.classList.remove('is-document-mode');
    if (titleEl) titleEl.textContent = att.filename || 'Anexo';
    if (sizeEl) sizeEl.textContent = `(${formatBytes(att.size)})`;
    bodyEl.innerHTML = '<div style="padding: 24px; text-align: center; color: var(--text-secondary);"><i class="fa-solid fa-spinner fa-spin" style="font-size: 28px; color: var(--ms-blue); margin-bottom: 10px;"></i><p>Carregando visualização do anexo...</p></div>';
    modal.style.display = 'flex';

    if (btnDl) btnDl.onclick = () => downloadAttachment(msg, att);

    const closeModal = () => {
      activePreviewSessionId++;
      modal.style.display = 'none';
      bodyEl.classList.remove('is-document-mode');
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
      if (sessionId !== activePreviewSessionId) return;

      const ct = (att.contentType || blob.type || '').toLowerCase();
      const fn = (att.filename || '').toLowerCase();

      // Fix generic octet-stream MIME for common formats by extension
      let effectiveBlob = blob;
      if (ct === 'application/octet-stream' || !ct) {
        if (fn.endsWith('.pdf')) effectiveBlob = new Blob([blob], { type: 'application/pdf' });
        else if (fn.endsWith('.png')) effectiveBlob = new Blob([blob], { type: 'image/png' });
        else if (/\.(jpg|jpeg)$/.test(fn)) effectiveBlob = new Blob([blob], { type: 'image/jpeg' });
        else if (fn.endsWith('.gif')) effectiveBlob = new Blob([blob], { type: 'image/gif' });
        else if (fn.endsWith('.webp')) effectiveBlob = new Blob([blob], { type: 'image/webp' });
        else if (fn.endsWith('.svg')) effectiveBlob = new Blob([blob], { type: 'image/svg+xml' });
        else if (fn.endsWith('.mp4')) effectiveBlob = new Blob([blob], { type: 'video/mp4' });
        else if (fn.endsWith('.mp3')) effectiveBlob = new Blob([blob], { type: 'audio/mpeg' });
        else if (/\.(txt|csv|tsv|log|json|xml|md)$/.test(fn)) effectiveBlob = new Blob([blob], { type: 'text/plain; charset=utf-8' });
      }

      const isPdf = effectiveBlob.type.includes('pdf') || fn.endsWith('.pdf');
      const isImage = effectiveBlob.type.startsWith('image/') || /\.(png|jpg|jpeg|gif|webp|bmp|svg|avif|ico)$/.test(fn);
      const isCsv = fn.endsWith('.csv') || fn.endsWith('.tsv') || ct.includes('csv');
      const isExcel = /\.(xlsx|xls|ods)$/.test(fn) || ct.includes('spreadsheet') || ct.includes('excel');
      const isDocx = fn.endsWith('.docx') || ct.includes('wordprocessingml');
      const isVideo = effectiveBlob.type.startsWith('video/') || /\.(mp4|webm|mov|ogv)$/.test(fn);
      const isAudio = effectiveBlob.type.startsWith('audio/') || /\.(mp3|wav|ogg|m4a|aac|flac)$/.test(fn);
      const isHtml = /\.(html|htm)$/.test(fn) || ct === 'text/html';
      const isText = effectiveBlob.type.startsWith('text/') || /\.(txt|json|xml|log|md|ini|yml|yaml|sql|js|css|py)$/.test(fn);

      if (isPdf) {
        const arrayBuf = await effectiveBlob.arrayBuffer();
        if (sessionId !== activePreviewSessionId) return;
        await renderPdfInCanvas(bodyEl, arrayBuf, sessionId);
      } else if (isImage) {
        activeAttBlobUrl = URL.createObjectURL(effectiveBlob);
        renderImageInViewer(bodyEl, activeAttBlobUrl, att.filename || 'Imagem');
      } else if (isCsv) {
        const csvText = await effectiveBlob.text();
        if (sessionId !== activePreviewSessionId) return;
        renderCsvInViewer(bodyEl, csvText);
      } else if (isExcel) {
        const arrayBuf = await effectiveBlob.arrayBuffer();
        if (sessionId !== activePreviewSessionId) return;
        await renderExcelInViewer(bodyEl, arrayBuf, sessionId);
      } else if (isDocx) {
        const arrayBuf = await effectiveBlob.arrayBuffer();
        if (sessionId !== activePreviewSessionId) return;
        await renderDocxInViewer(bodyEl, arrayBuf, sessionId);
      } else if (isVideo) {
        activeAttBlobUrl = URL.createObjectURL(effectiveBlob);
        bodyEl.innerHTML = `<video src="${activeAttBlobUrl}" controls playsinline style="max-width: 100%; max-height: 100%; border-radius: 6px; background: #000;"></video>`;
      } else if (isAudio) {
        activeAttBlobUrl = URL.createObjectURL(effectiveBlob);
        bodyEl.innerHTML = `
          <div style="background: #fff; color: #201f1e; padding: 28px 24px; border-radius: 8px; text-align: center; width: min(420px, 90vw); box-shadow: 0 2px 12px rgba(0,0,0,0.12);">
            <i class="fa-solid fa-music" style="font-size: 36px; color: var(--ms-blue); margin-bottom: 12px;"></i>
            <div style="font-weight: 600; font-size: 14px; margin-bottom: 14px; word-break: break-word;">${escapeHtml(att.filename)}</div>
            <audio src="${activeAttBlobUrl}" controls style="width: 100%;"></audio>
          </div>
        `;
      } else if (isHtml) {
        const htmlContent = await effectiveBlob.text();
        if (sessionId !== activePreviewSessionId) return;
        bodyEl.classList.add('is-document-mode');
        bodyEl.innerHTML = `<iframe sandbox="" style="width: 100%; height: 100%; border: none; background: #fff;"></iframe>`;
        const ifr = bodyEl.querySelector('iframe');
        if (ifr) ifr.srcdoc = htmlContent;
      } else if (isText) {
        let textContent = await effectiveBlob.text();
        if (sessionId !== activePreviewSessionId) return;
        if (fn.endsWith('.json')) {
          try { textContent = JSON.stringify(JSON.parse(textContent), null, 2); } catch (e) {}
        }
        bodyEl.classList.add('is-document-mode');
        bodyEl.innerHTML = `<pre style="width: 100%; height: 100%; margin: 0; padding: 16px; background: #fff; color: #201f1e; overflow: auto; font-family: Consolas, monospace; font-size: 13px; white-space: pre-wrap; word-break: break-word;">${escapeHtml(textContent)}</pre>`;
      } else {
        bodyEl.innerHTML = `
          <div style="background: #fff; color: #201f1e; padding: 28px 24px; border-radius: 8px; text-align: center; box-shadow: 0 2px 10px rgba(0,0,0,0.08); max-width: min(420px, 90vw);">
            <i class="${getAttachmentIcon(att.filename, att.contentType)}" style="font-size: 48px; margin-bottom: 14px; display: block;"></i>
            <h4 style="font-size: 15px; margin-bottom: 6px; word-break: break-word;">${escapeHtml(att.filename)}</h4>
            <p style="font-size: 12px; color: #605e5c; margin-bottom: 18px;">
              Tamanho: ${formatBytes(att.size)}<br>
              Este formato requer aplicativo externo para abertura.
            </p>
            <button type="button" class="action-btn primary" id="btnAttFallbackDownload" style="margin: 0 auto;">
              <i class="fa-solid fa-download"></i>
              <span>Baixar Arquivo</span>
            </button>
          </div>
        `;
        const btnFallback = document.getElementById('btnAttFallbackDownload');
        if (btnFallback) btnFallback.onclick = () => downloadAttachment(msg, att);
      }
    } catch (err) {
      if (sessionId !== activePreviewSessionId) return;
      bodyEl.classList.remove('is-document-mode');
      bodyEl.innerHTML = `<div style="background: #fff; color: #a80000; padding: 20px; border-radius: 6px; margin: 16px; text-align: center;"><i class="fa-solid fa-circle-exclamation"></i> Erro ao visualizar anexo: ${escapeHtml(err.message)}</div>`;
    }
  }

  function isDraftsFolderActive() {
    const f = (currentFolder || '').toLowerCase();
    return f.includes('draft') || f.includes('rascunh');
  }

  function renderReadingDetail(msg) {
    if (!msg) return;
    updateLikeButtonUI(msg);

    const btnEditDraft = document.getElementById('btnEditDraft');
    if (btnEditDraft) {
      if (isDraftsFolderActive()) {
        btnEditDraft.style.display = 'inline-flex';
        btnEditDraft.onclick = () => {
          const toList = Array.isArray(msg.to) ? msg.to.map(t => t.address).filter(Boolean).join(', ') : '';
          const ccList = Array.isArray(msg.cc) ? msg.cc.map(t => t.address).filter(Boolean).join(', ') : '';
          const subj = (msg.subject && msg.subject !== '(Rascunho sem assunto)' && msg.subject !== '(Sem assunto)') ? msg.subject : '';
          openCompose({
            to: toList,
            cc: ccList,
            subject: subj,
            body: msg.html || (msg.text ? `<p>${escapeHtml(msg.text).replace(/\n/g, '<br>')}</p>` : ''),
            draftUid: msg.uid,
            skipAutoSignature: true
          });
        };
      } else {
        btnEditDraft.style.display = 'none';
      }
    }

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
              font-size: 13.5px;
              color: #201f1e;
              margin: 0;
              padding: 8px 12px;
              line-height: 1.35;
              word-break: break-word;
            }
            p { margin: 0 0 6px 0; line-height: 1.35; }
            p:last-child { margin-bottom: 0; }
            div { line-height: 1.35; }
            ul, ol { margin: 4px 0 6px 0; padding-left: 20px; }
            li { margin-bottom: 2px; line-height: 1.35; }
            h1, h2, h3, h4 { margin: 8px 0 4px 0; line-height: 1.25; }
            blockquote { margin: 6px 0; padding-left: 10px; border-left: 2px solid #d1d5db; }
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
        <div style="font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 13.5px; line-height: 1.35; white-space: pre-wrap; padding: 10px 14px; color: #201f1e;">
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
      if (target) {
        const wasUnread = !target.seen;
        target.seen = seen;
        if (wasUnread && seen) adjustFolderUnreadCount(currentFolder, -1);
        else if (!wasUnread && !seen) adjustFolderUnreadCount(currentFolder, 1);
      }
      renderMessageList();
      updateRibbonState();
      loadFolders();
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

  function isTrashFolderActive() {
    return /^(trash|deleted|lixeira|itens exclu)/i.test(currentFolder || '');
  }

  async function deleteEmail(uid) {
    const inTrash = isTrashFolderActive();
    const confirmMsg = inTrash
      ? 'Este e-mail já está em Itens Excluídos. Deseja apagá-lo definitivamente?'
      : 'Deseja mover este e-mail para Itens Excluídos?';
    if (!confirm(confirmMsg)) return;

    try {
      const res = await fetch(`/api/imap/message/${uid}?folder=${encodeURIComponent(currentFolder)}`, { method: 'DELETE' });
      const data = await res.json();
      if (data.success) {
        messages = messages.filter(m => String(m.uid) !== String(uid));
        selectedUids.delete(String(uid));
        if (selectedMessage && String(selectedMessage.uid) === String(uid)) {
          selectedMessage = null;
          closeMobileReadingPane();
          if (elReadingContentView) elReadingContentView.setAttribute('style', 'display: none !important;');
          if (elNoSelectionView) elNoSelectionView.setAttribute('style', 'display: flex !important;');
        }
        elMessageCountLabel.textContent = `${messages.length} ${messages.length === 1 ? 'mensagem' : 'mensagens'}`;
        renderMessageList();
        updateRibbonState();
        loadFolders();
        showToast(data.movedToTrash ? 'E-mail movido para Itens Excluídos!' : 'E-mail excluído definitivamente!');
      } else {
        alert(data.error || 'Falha ao excluir e-mail');
      }
    } catch (e) {
      console.error(e);
    }
  }

  async function bulkDeleteEmails(uids) {
    const list = Array.from(uids).map(String);
    if (list.length === 0) return;

    const inTrash = isTrashFolderActive();
    const confirmMsg = inTrash
      ? `Estes ${list.length} e-mail(s) já estão em Itens Excluídos. Deseja apagá-los definitivamente?`
      : `Deseja mover ${list.length} ${list.length === 1 ? 'e-mail selecionado' : 'e-mails selecionados'} para Itens Excluídos?`;
    if (!confirm(confirmMsg)) return;

    try {
      showToast(inTrash ? `Excluindo ${list.length} e-mail(s)...` : `Movendo ${list.length} e-mail(s) para Itens Excluídos...`);
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
          closeMobileReadingPane();
          if (elReadingContentView) elReadingContentView.setAttribute('style', 'display: none !important;');
          if (elNoSelectionView) elNoSelectionView.setAttribute('style', 'display: flex !important;');
        }
        elMessageCountLabel.textContent = `${messages.length} ${messages.length === 1 ? 'mensagem' : 'mensagens'}`;
        renderMessageList();
        updateRibbonState();
        loadFolders();
        showToast(
          data.movedToTrash
            ? `${list.length} ${list.length === 1 ? 'e-mail movido' : 'e-mails movidos'} para Itens Excluídos!`
            : `${list.length} ${list.length === 1 ? 'e-mail excluído' : 'e-mails excluídos'} definitivamente!`
        );
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
        let unreadDelta = 0;
        messages.forEach(m => {
          if (uidSet.has(String(m.uid))) {
            if (seen !== undefined) {
              if (!m.seen && seen) unreadDelta -= 1;
              else if (m.seen && !seen) unreadDelta += 1;
              m.seen = seen;
            }
            if (flagged !== undefined) m.flagged = flagged;
          }
        });
        if (unreadDelta !== 0) adjustFolderUnreadCount(currentFolder, unreadDelta);
        renderMessageList();
        updateRibbonState();
        if (seen !== undefined) loadFolders();
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

  // ========================
  // Contacts Management
  // ========================
  let contactsList = [];
  let contactGroupsList = [];
  let contactsSearchQuery = '';
  let currentContactsTab = 'all'; // 'all' | 'contacts' | 'groups'

  function getContactsStorageKey() {
    const user = (SESSION_STORE.getUser() || document.getElementById('cfgImapUser')?.value || 'default').trim().toLowerCase();
    return 'outlook_contacts_' + user;
  }

  function updateContactsDatalist() {
    const dl = document.getElementById('contactsDatalist');
    if (!dl) return;
    // Individual contacts
    const contactOptions = contactsList.map(c =>
      `<option value="${escapeHtml(c.email)}">${escapeHtml(c.name || c.email)}</option>`
    );
    // Groups — value = group name so it can be expanded later
    const groupOptions = contactGroupsList.map(g =>
      `<option value="${escapeHtml(g.name)}" data-group-id="${escapeHtml(g.id)}">${escapeHtml(g.name)} (${g.members.length} membros)</option>`
    );
    dl.innerHTML = [...groupOptions, ...contactOptions].join('');
  }

  function renderContactsList() {
    const container = document.getElementById('contactsListContainer');
    const countEl = document.getElementById('contactsCountLabel');
    if (!container) return;

    const q = contactsSearchQuery.trim().toLowerCase();

    // Filter individual contacts
    const filteredContacts = q
      ? contactsList.filter(c =>
          (c.name && c.name.toLowerCase().includes(q)) ||
          (c.email && c.email.toLowerCase().includes(q)) ||
          (c.company && c.company.toLowerCase().includes(q)) ||
          (c.phone && c.phone.toLowerCase().includes(q))
        )
      : contactsList;

    // Filter groups
    const filteredGroups = q
      ? contactGroupsList.filter(g =>
          (g.name && g.name.toLowerCase().includes(q)) ||
          (g.description && g.description.toLowerCase().includes(q)) ||
          (g.members && g.members.some(m => m.toLowerCase().includes(q)))
        )
      : contactGroupsList;

    // Update tab badge counts
    const tabAll = document.getElementById('contactsTabAllCount');
    const tabIndiv = document.getElementById('contactsTabIndivCount');
    const tabGroups = document.getElementById('contactsTabGroupsCount');
    if (tabAll) tabAll.textContent = filteredContacts.length + filteredGroups.length;
    if (tabIndiv) tabIndiv.textContent = filteredContacts.length;
    if (tabGroups) tabGroups.textContent = filteredGroups.length;

    // Update total label
    const totalShown = currentContactsTab === 'groups'
      ? filteredGroups.length
      : currentContactsTab === 'contacts'
      ? filteredContacts.length
      : filteredContacts.length + filteredGroups.length;

    if (countEl) {
      countEl.textContent = `${filteredContacts.length} contato(s), ${filteredGroups.length} grupo(s)`;
    }

    // Decide what to render
    const showContacts = currentContactsTab === 'all' || currentContactsTab === 'contacts';
    const showGroups   = currentContactsTab === 'all' || currentContactsTab === 'groups';

    const toRenderContacts = showContacts ? filteredContacts : [];
    const toRenderGroups   = showGroups   ? filteredGroups   : [];

    if (toRenderContacts.length === 0 && toRenderGroups.length === 0) {
      container.innerHTML = `
        <div style="padding: 48px 20px; text-align: center; color: var(--text-secondary);">
          <i class="fa-regular fa-address-book" style="font-size: 40px; opacity: 0.4; margin-bottom: 12px; display: block;"></i>
          <p style="font-size: 14px; margin-bottom: 12px;">${q ? 'Nenhum resultado encontrado para essa busca.' : 'Nenhum contato ou grupo salvo ainda.'}</p>
          ${!q ? `<button type="button" class="action-btn primary" id="btnEmptySyncContacts" style="margin: 0 auto;"><i class="fa-solid fa-rotate"></i> <span>Sincronizar Contatos do Servidor</span></button>` : ''}
        </div>
      `;
      const emptyBtn = document.getElementById('btnEmptySyncContacts');
      if (emptyBtn) emptyBtn.onclick = () => syncContactsFromServer();
      return;
    }

    container.innerHTML = '';

    // Render groups first
    toRenderGroups.forEach(group => {
      const row = document.createElement('div');
      row.className = 'contact-row';
      row.style.background = 'linear-gradient(to right, rgba(92,45,145,0.03), transparent)';
      const membersPreview = group.members.slice(0, 3).join('; ') + (group.members.length > 3 ? ` +${group.members.length - 3} mais` : '');
      row.innerHTML = `
        <div class="ct-col ct-col-name">
          <div class="ct-avatar group-avatar" title="Grupo de Envio"><i class="fa-solid fa-users" style="font-size:12px;"></i></div>
          <span>
            ${escapeHtml(group.name)}
            <span class="ct-badge-group"><i class="fa-solid fa-users" style="font-size:9px;"></i> Grupo</span>
          </span>
        </div>
        <div class="ct-col ct-col-email" title="${escapeHtml(group.members.join('; '))}">
          <span class="ct-group-members-count">${group.members.length} membro(s)</span>
        </div>
        <div class="ct-col ct-col-phone ct-group-members-preview" style="font-size:11px; white-space:normal; line-height:1.3;" title="${escapeHtml(group.members.join('; '))}">
          ${escapeHtml(membersPreview || '—')}
        </div>
        <div class="ct-col ct-col-company">
          <span class="ct-badge-group-pill">${group.description ? escapeHtml(group.description.slice(0, 30)) + (group.description.length > 30 ? '...' : '') : 'Lista de distribuição'}</span>
        </div>
        <div class="ct-col ct-col-actions">
          <button type="button" class="action-btn btn-grp-mail" title="Enviar e-mail para este grupo" style="padding: 4px 8px;">
            <i class="fa-regular fa-paper-plane" style="color: #5c2d91;"></i>
          </button>
          <button type="button" class="action-btn btn-grp-edit" title="Editar grupo" style="padding: 4px 8px;">
            <i class="fa-regular fa-pen-to-square"></i>
          </button>
          <button type="button" class="action-btn btn-grp-del" title="Excluir grupo" style="padding: 4px 8px; color: #a80000;">
            <i class="fa-regular fa-trash-can"></i>
          </button>
        </div>
      `;

      row.querySelector('.btn-grp-mail').onclick = () => {
        openCompose({ to: group.members.join(', ') });
      };
      row.querySelector('.btn-grp-edit').onclick = () => {
        openContactGroupModal(group);
      };
      row.querySelector('.btn-grp-del').onclick = () => {
        deleteContactGroup(group.id, group.name);
      };

      container.appendChild(row);
    });

    // Render individual contacts
    toRenderContacts.forEach(contact => {
      const row = document.createElement('div');
      row.className = 'contact-row';
      row.innerHTML = `
        <div class="ct-col ct-col-name">
          <div class="ct-avatar">${getInitials(contact.name || contact.email)}</div>
          <span title="${escapeHtml(contact.name || '')}">${escapeHtml(contact.name || contact.email.split('@')[0])}</span>
        </div>
        <div class="ct-col ct-col-email" title="${escapeHtml(contact.email)}">${escapeHtml(contact.email)}</div>
        <div class="ct-col ct-col-phone">${escapeHtml(contact.phone || '—')}</div>
        <div class="ct-col ct-col-company">
          ${contact.company ? escapeHtml(contact.company) : `<span class="ct-badge-source">${contact.source === 'servidor' ? 'Servidor IMAP' : 'Manual'}</span>`}
        </div>
        <div class="ct-col ct-col-actions">
          <button type="button" class="action-btn btn-ct-mail" title="Enviar e-mail para ${escapeHtml(contact.email)}" style="padding: 4px 8px;">
            <i class="fa-regular fa-paper-plane" style="color: var(--ms-blue);"></i>
          </button>
          <button type="button" class="action-btn btn-ct-edit" title="Editar contato" style="padding: 4px 8px;">
            <i class="fa-regular fa-pen-to-square"></i>
          </button>
          <button type="button" class="action-btn btn-ct-del" title="Excluir contato" style="padding: 4px 8px; color: #a80000;">
            <i class="fa-regular fa-trash-can"></i>
          </button>
        </div>
      `;

      row.querySelector('.btn-ct-mail').onclick = () => {
        openCompose({ to: contact.email });
      };
      row.querySelector('.btn-ct-edit').onclick = () => {
        openContactModal(contact);
      };
      row.querySelector('.btn-ct-del').onclick = () => {
        deleteContact(contact.email);
      };

      container.appendChild(row);
    });
  }

  async function loadContacts() {
    try {
      const res = await fetch('/api/contacts');
      const data = await res.json();
      if (data.success && Array.isArray(data.contacts) && data.contacts.length > 0) {
        contactsList = data.contacts;
        try { localStorage.setItem(getContactsStorageKey(), JSON.stringify(contactsList)); } catch (e) {}
      } else {
        try {
          const cached = JSON.parse(localStorage.getItem(getContactsStorageKey()) || '[]');
          if (Array.isArray(cached) && cached.length > 0) contactsList = cached;
        } catch (e) {}
      }
      updateContactsDatalist();
      await loadContactGroups();
      renderContactsList();
    } catch (err) {
      console.warn('Erro ao carregar contatos:', err);
    }
  }

  async function loadContactGroups() {
    try {
      const res = await fetch('/api/contact-groups');
      const data = await res.json();
      if (data.success && Array.isArray(data.groups)) {
        contactGroupsList = data.groups;
      }
    } catch (err) {
      console.warn('Erro ao carregar grupos de contatos:', err);
    }
  }

  function renderContactGroupsPicker(selectedEmails = []) {
    const picker = document.getElementById('groupContactsPicker');
    if (!picker) return;
    const filterVal = (document.getElementById('groupPickerFilter')?.value || '').toLowerCase();
    const filtered = filterVal
      ? contactsList.filter(c =>
          (c.name && c.name.toLowerCase().includes(filterVal)) ||
          (c.email && c.email.toLowerCase().includes(filterVal))
        )
      : contactsList;

    if (filtered.length === 0) {
      picker.innerHTML = `<div style="padding: 10px; color: var(--text-secondary); font-size: 12px; text-align:center;">Nenhum contato disponível.</div>`;
      return;
    }

    picker.innerHTML = filtered.map(c => {
      const isChecked = selectedEmails.includes(c.email.toLowerCase());
      return `
        <label class="group-contact-item" for="picker_${CSS.escape(c.email)}">
          <input type="checkbox" id="picker_${CSS.escape(c.email)}" value="${escapeHtml(c.email)}" ${isChecked ? 'checked' : ''}>
          <div class="group-contact-info">
            <span class="group-contact-name">${escapeHtml(c.name || c.email.split('@')[0])}</span>
            <span class="group-contact-email">${escapeHtml(c.email)}</span>
          </div>
        </label>
      `;
    }).join('');

    // Sync picker checkboxes with group member chips
    picker.querySelectorAll('input[type="checkbox"]').forEach(cb => {
      cb.addEventListener('change', () => {
        const email = cb.value;
        if (cb.checked) {
          groupMemberManager && groupMemberManager.add(email);
        } else {
          // Remove from chips by rebuilding from remaining checked
          if (groupMemberManager) {
            const allChecked = Array.from(picker.querySelectorAll('input[type="checkbox"]:checked')).map(el => el.value);
            const currentChips = groupMemberManager.getAllEmails();
            // remove unchecked
            currentChips.forEach(em => {
              if (em.toLowerCase() === email.toLowerCase()) {
                // find chip to remove — rebuild
              }
            });
            // simplest: remove and re-add all checked
            groupMemberManager.clear();
            allChecked.forEach(em => groupMemberManager.add(em));
          }
        }
        updateGroupMembersCount();
      });
    });
  }

  let groupMemberManager = null;

  function updateGroupMembersCount() {
    const badge = document.getElementById('groupMembersCountBadge');
    if (!badge) return;
    const count = groupMemberManager ? groupMemberManager.getAllEmails().length : 0;
    badge.textContent = count;
  }

  function openContactGroupModal(group = null) {
    const modal = document.getElementById('contactGroupModal');
    const titleEl = document.getElementById('contactGroupModalTitle');
    const idEl = document.getElementById('contactGroupId');
    const nameEl = document.getElementById('contactGroupNameInput');
    const descEl = document.getElementById('contactGroupDescInput');
    const btnDel = document.getElementById('btnDeleteContactGroup');
    if (!modal) return;

    if (group) {
      if (titleEl) titleEl.innerHTML = '<i class="fa-solid fa-users" style="color:#5c2d91;margin-right:8px;"></i>Editar Grupo de Envio';
      if (idEl) idEl.value = group.id || '';
      if (nameEl) nameEl.value = group.name || '';
      if (descEl) descEl.value = group.description || '';
      if (btnDel) btnDel.style.display = '';
    } else {
      if (titleEl) titleEl.innerHTML = '<i class="fa-solid fa-users" style="color:#5c2d91;margin-right:8px;"></i>Novo Grupo de Envio';
      if (idEl) idEl.value = '';
      if (nameEl) nameEl.value = '';
      if (descEl) descEl.value = '';
      if (btnDel) btnDel.style.display = 'none';
    }

    // Init group member chip manager
    if (groupMemberManager) groupMemberManager.clear();
    groupMemberManager = createRecipientChipManager({
      containerId: 'containerGroupMembers',
      inputId: 'groupMemberInput',
      hiddenId: null,
      nextFocusId: null,
      onUpdate: updateGroupMembersCount
    });

    // Pre-load existing members
    if (group && group.members) {
      group.members.forEach(email => groupMemberManager.add(email));
    }
    updateGroupMembersCount();

    // Render contacts picker
    const selectedEmails = (group && group.members) ? group.members.map(e => e.toLowerCase()) : [];
    renderContactGroupsPicker(selectedEmails);

    // Clear picker filter
    const pickerFilter = document.getElementById('groupPickerFilter');
    if (pickerFilter) pickerFilter.value = '';

    modal.style.display = 'flex';
    if (nameEl) nameEl.focus();
  }

  async function saveContactGroup(e) {
    e.preventDefault();
    const idEl = document.getElementById('contactGroupId');
    const nameEl = document.getElementById('contactGroupNameInput');
    const descEl = document.getElementById('contactGroupDescInput');
    const id = idEl?.value?.trim() || '';
    const name = nameEl?.value?.trim() || '';
    const description = descEl?.value?.trim() || '';

    if (!name) { alert('O nome do grupo é obrigatório.'); return; }

    // Collect members from chip manager
    const members = groupMemberManager ? groupMemberManager.getAllEmails() : [];
    if (members.length === 0) {
      if (!confirm('O grupo não tem membros ainda. Salvar assim mesmo?')) return;
    }

    const payload = { name, description, members };
    if (id) payload.id = id;

    try {
      const res = await fetch('/api/contact-groups', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (data.success) {
        contactGroupsList = data.groups || [];
        updateContactsDatalist();
        renderContactsList();
        const modal = document.getElementById('contactGroupModal');
        if (modal) modal.style.display = 'none';
        showToast(`Grupo "${name}" salvo com sucesso! (${members.length} membro(s))`);
      } else {
        alert(data.error || 'Não foi possível salvar o grupo.');
      }
    } catch (err) {
      alert('Erro ao salvar grupo: ' + err.message);
    }
  }

  async function deleteContactGroup(id, name) {
    if (!confirm(`Deseja realmente excluir o grupo "${name}"?`)) return;
    try {
      const res = await fetch(`/api/contact-groups/${encodeURIComponent(id)}`, { method: 'DELETE' });
      const data = await res.json();
      if (data.success) {
        contactGroupsList = data.groups || [];
        updateContactsDatalist();
        renderContactsList();
        const modal = document.getElementById('contactGroupModal');
        if (modal) modal.style.display = 'none';
        showToast(`Grupo "${name}" excluído.`);
      } else {
        alert(data.error || 'Erro ao excluir grupo.');
      }
    } catch (err) {
      alert('Erro ao excluir grupo: ' + err.message);
    }
  }

  async function syncContactsFromServer() {
    const btn = document.getElementById('btnSyncContacts');
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>Sincronizando...</span>';
    }
    try {
      const res = await fetch('/api/contacts/sync', { method: 'POST' });
      const data = await res.json();
      if (data.success) {
        contactsList = data.contacts || [];
        try { localStorage.setItem(getContactsStorageKey(), JSON.stringify(contactsList)); } catch (e) {}
        updateContactsDatalist();
        renderContactsList();
        showToast(`Sincronizado! ${data.total} contato(s) disponível(is) (${data.added} novo(s)).`);
      } else {
        alert(data.error || 'Erro ao sincronizar contatos com o servidor.');
      }
    } catch (err) {
      alert('Erro ao sincronizar contatos: ' + err.message);
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = '<i class="fa-solid fa-rotate" style="color: var(--ms-blue);"></i> <span>Sincronizar com Servidor</span>';
      }
    }
  }

  function openContactModal(contact = null) {
    const modal = document.getElementById('contactModal');
    const titleEl = document.getElementById('contactModalTitle');
    const origEl = document.getElementById('contactOriginalEmail');
    const nameEl = document.getElementById('contactNameInput');
    const emailEl = document.getElementById('contactEmailInput');
    const phoneEl = document.getElementById('contactPhoneInput');
    const compEl = document.getElementById('contactCompanyInput');
    if (!modal) return;

    if (contact) {
      if (titleEl) titleEl.innerHTML = '<i class="fa-solid fa-user-pen" style="color: var(--ms-blue); margin-right: 8px;"></i>Editar Contato';
      if (origEl) origEl.value = contact.email || '';
      if (nameEl) nameEl.value = contact.name || '';
      if (emailEl) emailEl.value = contact.email || '';
      if (phoneEl) phoneEl.value = contact.phone || '';
      if (compEl) compEl.value = contact.company || '';
    } else {
      if (titleEl) titleEl.innerHTML = '<i class="fa-solid fa-user-plus" style="color: var(--ms-blue); margin-right: 8px;"></i>Novo Contato';
      if (origEl) origEl.value = '';
      if (nameEl) nameEl.value = '';
      if (emailEl) emailEl.value = '';
      if (phoneEl) phoneEl.value = '';
      if (compEl) compEl.value = '';
    }

    modal.style.display = 'flex';
    if (nameEl) nameEl.focus();
  }

  async function saveContact(e) {
    e.preventDefault();
    const payload = {
      originalEmail: document.getElementById('contactOriginalEmail')?.value || '',
      name: document.getElementById('contactNameInput')?.value || '',
      email: document.getElementById('contactEmailInput')?.value || '',
      phone: document.getElementById('contactPhoneInput')?.value || '',
      company: document.getElementById('contactCompanyInput')?.value || ''
    };

    try {
      const res = await fetch('/api/contacts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (data.success) {
        contactsList = data.contacts || [];
        try { localStorage.setItem(getContactsStorageKey(), JSON.stringify(contactsList)); } catch (err) {}
        updateContactsDatalist();
        renderContactsList();
        const modal = document.getElementById('contactModal');
        if (modal) modal.style.display = 'none';
        showToast('Contato salvo com sucesso!');
      } else {
        alert(data.error || 'Não foi possível salvar o contato.');
      }
    } catch (err) {
      alert('Erro ao salvar contato: ' + err.message);
    }
  }

  async function deleteContact(email) {
    if (!confirm(`Deseja realmente excluir o contato "${email}"?`)) return;
    try {
      const res = await fetch(`/api/contacts?email=${encodeURIComponent(email)}`, { method: 'DELETE' });
      const data = await res.json();
      if (data.success) {
        contactsList = data.contacts || [];
        try { localStorage.setItem(getContactsStorageKey(), JSON.stringify(contactsList)); } catch (err) {}
        updateContactsDatalist();
        renderContactsList();
        showToast('Contato excluído com sucesso!');
      }
    } catch (err) {
      alert('Erro ao excluir contato: ' + err.message);
    }
  }

  // ========================
  // Calendar Management
  // ========================
  let calendarEvents = [];
  let calCurrentDate = new Date();
  let calSelectedDateStr = toLocalDateStr(new Date());

  function toLocalDateStr(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function getCalendarStorageKey() {
    const user = (SESSION_STORE.getUser() || document.getElementById('cfgImapUser')?.value || 'default').trim().toLowerCase();
    return 'outlook_calendar_' + user;
  }

  async function loadCalendarEvents() {
    try {
      const res = await fetch('/api/calendar');
      const data = await res.json();
      if (data.success && Array.isArray(data.events) && data.events.length > 0) {
        calendarEvents = data.events;
        try { localStorage.setItem(getCalendarStorageKey(), JSON.stringify(calendarEvents)); } catch (e) {}
      } else {
        try {
          const cached = JSON.parse(localStorage.getItem(getCalendarStorageKey()) || '[]');
          if (Array.isArray(cached) && cached.length > 0) calendarEvents = cached;
        } catch (e) {}
      }
      renderCalendar();
    } catch (err) {
      console.warn('Erro ao carregar calendário:', err);
    }
  }

  function renderCalendar() {
    const grid = document.getElementById('calendarGrid');
    const titleEl = document.getElementById('calMonthTitle');
    if (!grid) return;

    const year = calCurrentDate.getFullYear();
    const month = calCurrentDate.getMonth();
    const monthNames = [
      'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
      'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
    ];
    if (titleEl) {
      titleEl.textContent = `${monthNames[month]} de ${year}`;
    }

    const firstDayOfMonth = new Date(year, month, 1);
    const startWeekday = firstDayOfMonth.getDay(); // 0=Sun..6=Sat
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const daysInPrevMonth = new Date(year, month, 0).getDate();
    const todayStr = toLocalDateStr(new Date());

    const totalCells = Math.ceil((startWeekday + daysInMonth) / 7) * 7;
    grid.innerHTML = '';

    for (let i = 0; i < totalCells; i++) {
      let cellDate;
      let isOtherMonth = false;

      if (i < startWeekday) {
        const dayNum = daysInPrevMonth - startWeekday + 1 + i;
        cellDate = new Date(year, month - 1, dayNum);
        isOtherMonth = true;
      } else if (i >= startWeekday + daysInMonth) {
        const dayNum = i - (startWeekday + daysInMonth) + 1;
        cellDate = new Date(year, month + 1, dayNum);
        isOtherMonth = true;
      } else {
        const dayNum = i - startWeekday + 1;
        cellDate = new Date(year, month, dayNum);
      }

      const dateStr = toLocalDateStr(cellDate);
      const dayEvents = calendarEvents.filter(ev => ev.date === dateStr);

      const cell = document.createElement('div');
      cell.className = `cal-day-cell ${isOtherMonth ? 'other-month' : ''} ${dateStr === todayStr ? 'is-today' : ''} ${dateStr === calSelectedDateStr ? 'selected-day' : ''}`;
      cell.innerHTML = `<div class="cal-day-number">${cellDate.getDate()}</div>`;

      dayEvents.slice(0, 3).forEach(ev => {
        const pill = document.createElement('div');
        pill.className = 'cal-event-pill';
        pill.style.backgroundColor = ev.color || '#0078d4';
        pill.textContent = `${ev.startTime ? ev.startTime + ' ' : ''}${ev.title}`;
        pill.title = `${ev.startTime} - ${ev.endTime}: ${ev.title}`;
        pill.onclick = (e) => {
          e.stopPropagation();
          calSelectedDateStr = dateStr;
          renderCalendar();
          openCalendarEventModal(ev);
        };
        cell.appendChild(pill);
      });

      if (dayEvents.length > 3) {
        const more = document.createElement('div');
        more.style.fontSize = '10px';
        more.style.color = 'var(--text-secondary)';
        more.style.fontWeight = '600';
        more.textContent = `+${dayEvents.length - 3} mais`;
        cell.appendChild(more);
      }

      cell.onclick = () => {
        calSelectedDateStr = dateStr;
        renderCalendar();
      };

      cell.ondblclick = () => {
        calSelectedDateStr = dateStr;
        openCalendarEventModal(null, dateStr);
      };

      grid.appendChild(cell);
    }

    renderSelectedDayAgenda();
  }

  function renderSelectedDayAgenda() {
    const titleEl = document.getElementById('calSelectedDayTitle');
    const listEl = document.getElementById('calSelectedDayList');
    if (!listEl) return;

    const parts = calSelectedDateStr.split('-');
    const formatted = parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : calSelectedDateStr;
    if (titleEl) {
      titleEl.textContent = `Agenda • ${formatted}`;
    }

    const dayEvents = calendarEvents.filter(ev => ev.date === calSelectedDateStr);
    if (dayEvents.length === 0) {
      listEl.innerHTML = `
        <div style="text-align: center; padding: 30px 12px; color: var(--text-secondary);">
          <i class="fa-regular fa-calendar-check" style="font-size: 30px; opacity: 0.4; margin-bottom: 8px; display: block;"></i>
          <p style="font-size: 12px; margin-bottom: 12px;">Nenhum compromisso agendado para este dia.</p>
          <button type="button" class="action-btn primary" id="btnAgendaQuickAdd" style="margin: 0 auto;">
            <i class="fa-solid fa-plus"></i> <span>Agendar neste dia</span>
          </button>
        </div>
      `;
      const quickAdd = document.getElementById('btnAgendaQuickAdd');
      if (quickAdd) quickAdd.onclick = () => openCalendarEventModal(null, calSelectedDateStr);
      return;
    }

    listEl.innerHTML = '';
    dayEvents.forEach(ev => {
      const card = document.createElement('div');
      card.className = 'cal-agenda-card';
      card.style.borderLeftColor = ev.color || '#0078d4';
      card.innerHTML = `
        <div style="font-size: 11px; font-weight: 700; color: ${ev.color || '#0078d4'}; margin-bottom: 2px;">
          <i class="fa-regular fa-clock"></i> ${escapeHtml(ev.startTime || '')} - ${escapeHtml(ev.endTime || '')}
        </div>
        <div style="font-size: 13px; font-weight: 600; color: var(--text-primary); margin-bottom: 4px;">
          ${escapeHtml(ev.title)}
        </div>
        ${ev.location ? `<div style="font-size: 11px; color: var(--text-secondary);"><i class="fa-solid fa-location-dot"></i> ${escapeHtml(ev.location)}</div>` : ''}
        ${ev.description ? `<div style="font-size: 11px; color: var(--text-secondary); margin-top: 4px;">${escapeHtml(ev.description)}</div>` : ''}
      `;
      card.onclick = () => openCalendarEventModal(ev);
      listEl.appendChild(card);
    });
  }

  function openCalendarEventModal(ev = null, defaultDate = null) {
    const modal = document.getElementById('calendarEventModal');
    const titleHeader = document.getElementById('calEventModalTitle');
    const idEl = document.getElementById('calEventId');
    const titleEl = document.getElementById('calEventTitle');
    const dateEl = document.getElementById('calEventDate');
    const startEl = document.getElementById('calEventStart');
    const endEl = document.getElementById('calEventEnd');
    const locEl = document.getElementById('calEventLocation');
    const colorEl = document.getElementById('calEventColor');
    const attEl = document.getElementById('calEventAttendees');
    const inviteEl = document.getElementById('calEventSendInvite');
    const descEl = document.getElementById('calEventDesc');
    const btnDel = document.getElementById('btnDeleteCalEvent');

    if (!modal) return;

    if (ev) {
      if (titleHeader) titleHeader.innerHTML = '<i class="fa-regular fa-calendar-check" style="color: var(--ms-blue); margin-right: 8px;"></i>Editar Evento';
      if (idEl) idEl.value = ev.id || '';
      if (titleEl) titleEl.value = ev.title || '';
      if (dateEl) dateEl.value = ev.date || calSelectedDateStr;
      if (startEl) startEl.value = ev.startTime || '09:00';
      if (endEl) endEl.value = ev.endTime || '10:00';
      if (locEl) locEl.value = ev.location || '';
      if (colorEl) colorEl.value = ev.color || '#0078d4';
      if (attEl) attEl.value = ev.attendees || '';
      if (inviteEl) inviteEl.checked = false;
      if (descEl) descEl.value = ev.description || '';
      if (btnDel) {
        btnDel.style.display = 'inline-flex';
        btnDel.onclick = () => deleteCalendarEvent(ev.id);
      }
    } else {
      if (titleHeader) titleHeader.innerHTML = '<i class="fa-regular fa-calendar-plus" style="color: var(--ms-blue); margin-right: 8px;"></i>Novo Evento';
      if (idEl) idEl.value = '';
      if (titleEl) titleEl.value = '';
      if (dateEl) dateEl.value = defaultDate || calSelectedDateStr || toLocalDateStr(new Date());
      if (startEl) startEl.value = '09:00';
      if (endEl) endEl.value = '10:00';
      if (locEl) locEl.value = '';
      if (colorEl) colorEl.value = '#0078d4';
      if (attEl) attEl.value = '';
      if (inviteEl) inviteEl.checked = false;
      if (descEl) descEl.value = '';
      if (btnDel) btnDel.style.display = 'none';
    }

    modal.style.display = 'flex';
    if (titleEl) titleEl.focus();
  }

  async function saveCalendarEvent(e) {
    e.preventDefault();
    const payload = {
      id: document.getElementById('calEventId')?.value || '',
      title: document.getElementById('calEventTitle')?.value || '',
      date: document.getElementById('calEventDate')?.value || '',
      startTime: document.getElementById('calEventStart')?.value || '09:00',
      endTime: document.getElementById('calEventEnd')?.value || '10:00',
      location: document.getElementById('calEventLocation')?.value || '',
      color: document.getElementById('calEventColor')?.value || '#0078d4',
      attendees: document.getElementById('calEventAttendees')?.value || '',
      sendInvite: document.getElementById('calEventSendInvite')?.checked === true,
      description: document.getElementById('calEventDesc')?.value || ''
    };

    try {
      const res = await fetch('/api/calendar', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (data.success) {
        calendarEvents = data.events || [];
        try { localStorage.setItem(getCalendarStorageKey(), JSON.stringify(calendarEvents)); } catch (err) {}
        calSelectedDateStr = payload.date;
        renderCalendar();
        const modal = document.getElementById('calendarEventModal');
        if (modal) modal.style.display = 'none';
        showToast(payload.sendInvite && payload.attendees ? 'Evento salvo e convite enviado por e-mail!' : 'Evento salvo no calendário!');
      } else {
        alert(data.error || 'Não foi possível salvar o evento.');
      }
    } catch (err) {
      alert('Erro ao salvar evento: ' + err.message);
    }
  }

  async function deleteCalendarEvent(eventId) {
    if (!eventId || !confirm('Deseja realmente excluir este evento do calendário?')) return;
    try {
      const res = await fetch(`/api/calendar/${encodeURIComponent(eventId)}`, { method: 'DELETE' });
      const data = await res.json();
      if (data.success) {
        calendarEvents = data.events || [];
        try { localStorage.setItem(getCalendarStorageKey(), JSON.stringify(calendarEvents)); } catch (err) {}
        renderCalendar();
        const modal = document.getElementById('calendarEventModal');
        if (modal) modal.style.display = 'none';
        showToast('Evento excluído do calendário.');
      }
    } catch (err) {
      alert('Erro ao excluir evento: ' + err.message);
    }
  }

  // ========================
  // Microsoft To Do Tasks
  // ========================
  let todoTasks = [];
  let todoFilter = 'today'; // 'today' | 'important' | 'all' | 'completed'
  let todoQuickImportant = false;

  function getTodosStorageKey() {
    const user = (SESSION_STORE.getUser() || document.getElementById('cfgImapUser')?.value || 'default').trim().toLowerCase();
    return 'outlook_todos_' + user;
  }

  async function loadTodos() {
    try {
      const res = await fetch('/api/todos');
      const data = await res.json();
      if (data.success && Array.isArray(data.todos) && data.todos.length > 0) {
        todoTasks = data.todos;
        try { localStorage.setItem(getTodosStorageKey(), JSON.stringify(todoTasks)); } catch (e) {}
      } else {
        try {
          const cached = JSON.parse(localStorage.getItem(getTodosStorageKey()) || '[]');
          if (Array.isArray(cached) && cached.length > 0) todoTasks = cached;
        } catch (e) {}
      }
      renderTodos();
    } catch (err) {
      console.warn('Erro ao carregar To Do:', err);
    }
  }

  function renderTodos() {
    const container = document.getElementById('todoListContainer');
    const titleEl = document.getElementById('todoListTitle');
    const subEl = document.getElementById('todoListSubtitle');
    if (!container) return;

    const todayStr = toLocalDateStr(new Date());

    // Update sidebar counts
    const countToday = todoTasks.filter(t => !t.completed && (!t.dueDate || t.dueDate <= todayStr)).length;
    const countImp = todoTasks.filter(t => !t.completed && t.important).length;
    const countAll = todoTasks.filter(t => !t.completed).length;
    const countDone = todoTasks.filter(t => t.completed).length;

    const elCT = document.getElementById('todoCountToday');
    const elCI = document.getElementById('todoCountImportant');
    const elCA = document.getElementById('todoCountAll');
    const elCD = document.getElementById('todoCountCompleted');
    if (elCT) elCT.textContent = countToday;
    if (elCI) elCI.textContent = countImp;
    if (elCA) elCA.textContent = countAll;
    if (elCD) elCD.textContent = countDone;

    const titles = {
      today: 'Meu Dia',
      important: 'Importante',
      all: 'Todas as Tarefas',
      completed: 'Concluídas'
    };
    if (titleEl) titleEl.textContent = titles[todoFilter] || 'Tarefas';
    if (subEl) {
      subEl.textContent = new Date().toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'long' });
    }

    let filtered = [];
    if (todoFilter === 'today') {
      filtered = todoTasks.filter(t => !t.completed && (!t.dueDate || t.dueDate <= todayStr));
    } else if (todoFilter === 'important') {
      filtered = todoTasks.filter(t => !t.completed && t.important);
    } else if (todoFilter === 'completed') {
      filtered = todoTasks.filter(t => t.completed);
    } else {
      filtered = todoTasks.slice();
    }

    // Sort: incomplete first, then important, then newest
    filtered.sort((a, b) => {
      if (a.completed !== b.completed) return a.completed ? 1 : -1;
      if (a.important !== b.important) return a.important ? -1 : 1;
      return 0;
    });

    if (filtered.length === 0) {
      container.innerHTML = `
        <div style="text-align: center; padding: 48px 20px; color: var(--text-secondary);">
          <i class="fa-regular fa-circle-check" style="font-size: 40px; opacity: 0.4; margin-bottom: 12px; display: block;"></i>
          <p style="font-size: 14px;">Nenhuma tarefa nesta lista.</p>
        </div>
      `;
      return;
    }

    container.innerHTML = '';
    filtered.forEach(todo => {
      const card = document.createElement('div');
      card.className = `todo-item-card ${todo.completed ? 'completed' : ''}`;

      let dueHtml = '';
      if (todo.dueDate) {
        const parts = todo.dueDate.split('-');
        const fmt = parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : todo.dueDate;
        const isOverdue = !todo.completed && todo.dueDate < todayStr;
        dueHtml = `<span class="todo-due-badge ${isOverdue ? 'overdue' : ''}"><i class="fa-regular fa-calendar"></i> ${isOverdue ? 'Atrasada: ' : 'Prazo: '}${fmt}</span>`;
      }

      card.innerHTML = `
        <button type="button" class="todo-circle-chk" title="${todo.completed ? 'Reabrir tarefa' : 'Concluir tarefa'}">
          ${todo.completed ? '<i class="fa-solid fa-check"></i>' : ''}
        </button>
        <div class="todo-item-body">
          <div class="todo-item-title">${escapeHtml(todo.title)}</div>
          ${(dueHtml || todo.notes) ? `<div class="todo-item-meta">${dueHtml}${todo.notes ? `<span><i class="fa-regular fa-note-sticky"></i> ${escapeHtml(todo.notes)}</span>` : ''}</div>` : ''}
        </div>
        <button type="button" class="todo-star-btn ${todo.important ? 'active' : ''}" title="${todo.important ? 'Remover importância' : 'Marcar como importante'}">
          <i class="fa-${todo.important ? 'solid' : 'regular'} fa-star"></i>
        </button>
        <button type="button" class="card-action-btn btn-todo-del" title="Excluir tarefa" style="color: #a80000; padding: 6px;">
          <i class="fa-regular fa-trash-can"></i>
        </button>
      `;

      card.querySelector('.todo-circle-chk').onclick = () => updateTodoTask({ ...todo, completed: !todo.completed });
      card.querySelector('.todo-star-btn').onclick = () => updateTodoTask({ ...todo, important: !todo.important });
      card.querySelector('.btn-todo-del').onclick = () => deleteTodoTask(todo.id);

      container.appendChild(card);
    });
  }

  async function addTodoTaskFromForm(e) {
    e.preventDefault();
    const titleEl = document.getElementById('todoQuickTitle');
    const dateEl = document.getElementById('todoQuickDate');
    const starBtn = document.getElementById('todoQuickStar');
    const title = (titleEl?.value || '').trim();
    if (!title) return;

    const payload = {
      title,
      dueDate: dateEl?.value || '',
      important: todoQuickImportant || todoFilter === 'important',
      completed: false
    };

    const ok = await updateTodoTask(payload);
    if (ok) {
      if (titleEl) titleEl.value = '';
      if (dateEl) dateEl.value = '';
      todoQuickImportant = false;
      if (starBtn) {
        starBtn.classList.remove('active');
        starBtn.innerHTML = '<i class="fa-regular fa-star"></i>';
      }
      showToast('Tarefa adicionada ao To Do!');
    }
  }

  async function updateTodoTask(payload) {
    try {
      const res = await fetch('/api/todos', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (data.success) {
        todoTasks = data.todos || [];
        try { localStorage.setItem(getTodosStorageKey(), JSON.stringify(todoTasks)); } catch (e) {}
        renderTodos();
        return true;
      }
    } catch (err) {
      console.error('Erro ao salvar tarefa:', err);
    }
    return false;
  }

  async function deleteTodoTask(id) {
    try {
      const res = await fetch(`/api/todos/${encodeURIComponent(id)}`, { method: 'DELETE' });
      const data = await res.json();
      if (data.success) {
        todoTasks = data.todos || [];
        try { localStorage.setItem(getTodosStorageKey(), JSON.stringify(todoTasks)); } catch (e) {}
        renderTodos();
        showToast('Tarefa removida.');
      }
    } catch (err) {
      console.error('Erro ao excluir tarefa:', err);
    }
  }

  function switchAppSection(section) {
    toggleMobileFoldersDrawer(false);
    closeMobileReadingPane();
    const railMail = document.getElementById('railMailBtn');
    const railCalendar = document.getElementById('railCalendarBtn');
    const railContacts = document.getElementById('railContactsBtn');
    const railTodo = document.getElementById('railTodoBtn');
    const foldersPane = document.getElementById('foldersPane');
    const contentSplit = document.getElementById('contentSplit');
    const contactsView = document.getElementById('contactsView');
    const calendarView = document.getElementById('calendarView');
    const todoView = document.getElementById('todoView');

    if (railMail) railMail.classList.remove('active');
    if (railCalendar) railCalendar.classList.remove('active');
    if (railContacts) railContacts.classList.remove('active');
    if (railTodo) railTodo.classList.remove('active');
    if (foldersPane) foldersPane.style.display = 'none';
    if (contentSplit) contentSplit.style.display = 'none';
    if (contactsView) contactsView.style.display = 'none';
    if (calendarView) calendarView.style.display = 'none';
    if (todoView) todoView.style.display = 'none';

    if (section === 'contacts') {
      if (railContacts) railContacts.classList.add('active');
      if (contactsView) contactsView.style.display = 'flex';
      renderContactsList();
      if (contactsList.length === 0) {
        syncContactsFromServer();
      }
    } else if (section === 'calendar') {
      if (railCalendar) railCalendar.classList.add('active');
      if (calendarView) calendarView.style.display = 'flex';
      renderCalendar();
    } else if (section === 'todo') {
      if (railTodo) railTodo.classList.add('active');
      if (todoView) todoView.style.display = 'flex';
      renderTodos();
    } else {
      if (railMail) railMail.classList.add('active');
      if (foldersPane) foldersPane.style.display = 'flex';
      if (contentSplit) contentSplit.style.display = 'flex';
    }
  }

  // =========================================================
  // Recipient Chip Input Manager (Para & Cc Multiple Emails)
  // =========================================================
  function isValidEmailAddress(email) {
    if (!email || typeof email !== 'string') return false;
    let clean = email.trim();
    if (clean.includes('<') && clean.includes('>')) {
      const match = clean.match(/<([^>]+)>/);
      if (match && match[1]) clean = match[1].trim();
    }
    return /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$/.test(clean);
  }

  function cleanRecipientAddress(raw) {
    if (!raw) return '';
    let clean = raw.trim().replace(/^[,;\s]+|[,;\s]+$/g, '');
    if (clean.includes('<') && clean.includes('>')) {
      const match = clean.match(/<([^>]+)>/);
      if (match && match[1]) return match[1].trim();
    }
    return clean;
  }

  function createRecipientChipManager({ containerId, inputId, hiddenId, nextFocusId, onUpdate }) {
    const container = document.getElementById(containerId);
    const input = document.getElementById(inputId);
    const hidden = document.getElementById(hiddenId);
    let chips = [];
    let idCounter = 0;

    function syncHiddenValue() {
      const val = chips.map(c => c.text).join(', ');
      if (hidden) hidden.value = val;
      if (onUpdate) onUpdate(val);
    }

    function renderChipElement(chip) {
      const el = document.createElement('div');
      el.className = `recipient-chip ${chip.isValid ? 'valid' : 'invalid'}`;
      el.setAttribute('data-chip-id', String(chip.id));
      el.title = chip.isValid
        ? 'E-mail válido (clique para editar)'
        : 'Formato de e-mail inválido! Clique para corrigir ou no "×" para remover.';

      el.innerHTML = `
        <i class="chip-status-icon fa-solid ${chip.isValid ? 'fa-circle-check' : 'fa-circle-exclamation'}"></i>
        <span class="chip-text">${escapeHtml(chip.text)}</span>
        <button type="button" class="chip-remove-btn" title="Remover">&times;</button>
      `;

      el.onclick = (e) => {
        e.stopPropagation();
        if (e.target.closest('.chip-remove-btn')) {
          removeChip(chip.id);
        } else {
          const currentText = chip.text;
          removeChip(chip.id);
          if (input) {
            input.value = currentText;
            input.focus();
            input.select();
          }
        }
      };

      return el;
    }

    function addChip(rawText) {
      if (!rawText) return;
      const parts = String(rawText).split(/[,;\n\r\t]+/).map(p => cleanRecipientAddress(p)).filter(Boolean);
      parts.forEach(clean => {
        // Check if this token matches a contact group name → expand into member emails
        if (typeof contactGroupsList !== 'undefined' && contactGroupsList.length > 0) {
          const matchedGroup = contactGroupsList.find(g => g.name.toLowerCase() === clean.toLowerCase());
          if (matchedGroup && matchedGroup.members && matchedGroup.members.length > 0) {
            matchedGroup.members.forEach(memberEmail => {
              if (!chips.some(c => c.text.toLowerCase() === memberEmail.toLowerCase())) {
                const chip = { id: ++idCounter, text: memberEmail, isValid: isValidEmailAddress(memberEmail) };
                chips.push(chip);
                if (container && input) container.insertBefore(renderChipElement(chip), input);
              }
            });
            syncHiddenValue();
            return;
          }
        }
        if (chips.some(c => c.text.toLowerCase() === clean.toLowerCase())) {
          return;
        }
        const isValid = isValidEmailAddress(clean);
        const chip = { id: ++idCounter, text: clean, isValid };
        chips.push(chip);
        if (container && input) {
          container.insertBefore(renderChipElement(chip), input);
        }
      });
      syncHiddenValue();
    }

    function removeChip(chipId) {
      chips = chips.filter(c => c.id !== chipId);
      if (container) {
        const el = container.querySelector(`.recipient-chip[data-chip-id="${chipId}"]`);
        if (el) el.remove();
      }
      syncHiddenValue();
    }

    function clear() {
      chips = [];
      if (container) {
        container.querySelectorAll('.recipient-chip').forEach(c => c.remove());
      }
      if (input) input.value = '';
      syncHiddenValue();
    }

    function setFromText(text) {
      clear();
      if (text) addChip(text);
    }

    function commitInput() {
      if (!input) return;
      const val = input.value.trim();
      if (val) {
        addChip(val);
        input.value = '';
      }
    }

    if (input) {
      input.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && (e.key === 'Enter' || e.code === 'NumpadEnter')) {
          return; // Allow Ctrl+Enter to submit/send the email
        }
        if (e.key === 'Enter' || e.key === ';' || e.key === ',') {
          e.preventDefault();
          const val = input.value.trim();
          if (val) {
            addChip(val);
            input.value = '';
          } else if (e.key === 'Enter' && nextFocusId) {
            const nextEl = document.getElementById(nextFocusId);
            if (nextEl) nextEl.focus();
          }
        } else if (e.key === 'Tab') {
          const val = input.value.trim();
          if (val) {
            e.preventDefault();
            addChip(val);
            input.value = '';
          }
        } else if (e.key === 'Backspace' && !input.value && chips.length > 0) {
          const lastChip = chips[chips.length - 1];
          removeChip(lastChip.id);
          input.value = lastChip.text;
          input.select();
        }
      });

      input.addEventListener('paste', (e) => {
        const text = (e.clipboardData || window.clipboardData)?.getData('text') || '';
        if (text && (text.includes(';') || text.includes(',') || text.includes('\n') || text.includes(' '))) {
          e.preventDefault();
          addChip(text);
          input.value = '';
        }
      });

      input.addEventListener('input', () => {
        const val = input.value;
        if (val.includes(';') || val.includes(',')) {
          addChip(val);
          input.value = '';
        }
      });

      input.addEventListener('change', () => {
        commitInput();
      });

      input.addEventListener('blur', () => {
        commitInput();
      });
    }

    if (container) {
      container.addEventListener('click', (e) => {
        if (input && !e.target.closest('.recipient-chip')) {
          input.focus();
        }
      });
    }

    return {
      add: addChip,
      remove: removeChip,
      clear,
      setFromText,
      commitInput,
      focus: () => { if (input) input.focus(); },
      getAllEmails: () => chips.map(c => c.text),
      getValidEmails: () => chips.filter(c => c.isValid).map(c => c.text),
      getInvalidEmails: () => chips.filter(c => !c.isValid).map(c => c.text)
    };
  }

  let recipientManagerTo = null;
  let recipientManagerCc = null;

  function initRecipientChipManagers() {
    recipientManagerTo = createRecipientChipManager({
      containerId: 'containerComposeTo',
      inputId: 'composeToInput',
      hiddenId: 'composeTo',
      nextFocusId: 'composeCcInput'
    });
    recipientManagerCc = createRecipientChipManager({
      containerId: 'containerComposeCc',
      inputId: 'composeCcInput',
      hiddenId: 'composeCc',
      nextFocusId: 'composeSubject'
    });
  }

  // Compose & 2-Minute Draft Auto-Save Functions
  const DRAFT_AUTOSAVE_MS = 2 * 60 * 1000; // 2 minutes
  let draftAutoSaveTimer = null;
  let currentDraftUid = null;
  let initialComposeBodyHtml = '';
  let lastSavedDraftSnapshot = '';
  let isSavingDraft = false;

  function stopDraftAutoSaveTimer() {
    if (draftAutoSaveTimer) {
      clearInterval(draftAutoSaveTimer);
      draftAutoSaveTimer = null;
    }
  }

  function getComposeSnapshot() {
    if (recipientManagerTo) recipientManagerTo.commitInput();
    if (recipientManagerCc) recipientManagerCc.commitInput();
    return JSON.stringify({
      to: (elComposeTo?.value || '').trim(),
      cc: (elComposeCc?.value || '').trim(),
      subject: (elComposeSubject?.value || '').trim(),
      html: (elComposeEditor?.innerHTML || '').trim(),
      files: attachedFiles.map(f => `${f.name}:${f.size}`).join('|')
    });
  }

  function hasUnsentComposeContent() {
    const toInputVal = (document.getElementById('composeToInput')?.value || '').trim();
    const ccInputVal = (document.getElementById('composeCcInput')?.value || '').trim();
    if ((elComposeTo?.value || '').trim() || toInputVal) return true;
    if ((elComposeCc?.value || '').trim() || ccInputVal) return true;
    if ((elComposeSubject?.value || '').trim()) return true;
    if (attachedFiles.length > 0) return true;
    const currentHtml = (elComposeEditor?.innerHTML || '').trim();
    const currentText = (elComposeEditor?.innerText || '').trim();
    return currentHtml !== initialComposeBodyHtml && currentText.length > 0;
  }

  async function saveComposeDraftNow() {
    if (!elComposeModal || elComposeModal.style.display !== 'flex') return;
    if (elBtnSendMail?.disabled || isSavingDraft) return;
    if (recipientManagerTo) recipientManagerTo.commitInput();
    if (recipientManagerCc) recipientManagerCc.commitInput();
    if (!hasUnsentComposeContent()) return;

    const snapshot = getComposeSnapshot();
    if (snapshot === lastSavedDraftSnapshot) return;

    isSavingDraft = true;
    if (elComposeStatusMsg) {
      elComposeStatusMsg.textContent = 'Salvando rascunho...';
    }

    const baseFont = elComposeEditor.style.fontFamily || 'Arial, Helvetica, sans-serif';
    const baseSize = elComposeEditor.style.fontSize || '15px';
    const rawHtml = elComposeEditor.innerHTML || '';
    const wrappedHtml = rawHtml.includes('font-family:')
      ? rawHtml
      : `<div style="font-family: ${baseFont}; font-size: ${baseSize};">${rawHtml}</div>`;

    const formData = new FormData();
    formData.append('to', elComposeTo.value || '');
    formData.append('cc', elComposeCc.value || '');
    formData.append('subject', elComposeSubject.value || '');
    formData.append('html', wrappedHtml);
    formData.append('text', elComposeEditor.innerText || '');
    if (currentDraftUid) {
      formData.append('previousDraftUid', String(currentDraftUid));
    }
    attachedFiles.forEach(f => formData.append('attachments', f));

    try {
      const res = await fetch('/api/imap/draft', {
        method: 'POST',
        body: formData
      });
      const data = await res.json();
      if (data && data.success) {
        if (data.uid) currentDraftUid = data.uid;
        lastSavedDraftSnapshot = snapshot;
        const hhmm = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
        if (elComposeStatusMsg) {
          elComposeStatusMsg.textContent = `Salvo em Rascunhos às ${hhmm}`;
        }
        showToast('Mensagem salva automaticamente na pasta Rascunhos.');
        loadFolders();
        if (isDraftsFolderActive()) {
          loadMessages();
        }
      } else if (elComposeStatusMsg) {
        elComposeStatusMsg.textContent = '';
      }
    } catch (err) {
      console.warn('[AutoSave Draft]:', err.message);
      if (elComposeStatusMsg) elComposeStatusMsg.textContent = '';
    } finally {
      isSavingDraft = false;
    }
  }

  function startDraftAutoSaveTimer() {
    stopDraftAutoSaveTimer();
    draftAutoSaveTimer = setInterval(() => {
      saveComposeDraftNow();
    }, DRAFT_AUTOSAVE_MS);
  }

  function openCompose(options = {}) {
    elComposeModal.style.display = 'flex';
    if (recipientManagerTo) recipientManagerTo.setFromText(options.to || '');
    else if (elComposeTo) elComposeTo.value = options.to || '';
    if (recipientManagerCc) recipientManagerCc.setFromText(options.cc || '');
    else if (elComposeCc) elComposeCc.value = options.cc || '';
    elComposeSubject.value = options.subject || '';

    const isReplyOrForward = !!options.isReplyOrForward;
    const skipAutoSignature = !!options.skipAutoSignature;
    const shouldAutoSign = !skipAutoSignature && (isReplyOrForward ? (sigState.autoReplyForward !== false) : (sigState.autoNewMail !== false));
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
    renderComposeAttachments();
    populateComposeTemplateSelect();
    const tplSel = document.getElementById('composeTemplateSelect');
    if (tplSel) tplSel.value = '';
    elComposeStatusMsg.textContent = options.draftUid ? 'Editando rascunho salvo' : '';

    currentDraftUid = options.draftUid || null;
    initialComposeBodyHtml = elComposeEditor.innerHTML.trim();
    lastSavedDraftSnapshot = options.draftUid ? getComposeSnapshot() : '';
    startDraftAutoSaveTimer();

    if (options.to) {
      elComposeEditor.focus();
    } else {
      if (recipientManagerTo) recipientManagerTo.focus();
      else if (elComposeTo) elComposeTo.focus();
    }
  }

  function closeCompose() {
    stopDraftAutoSaveTimer();
    currentDraftUid = null;
    initialComposeBodyHtml = '';
    lastSavedDraftSnapshot = '';
    elComposeModal.style.display = 'none';
    elComposeModal.classList.remove('maximized', 'minimized');
    const btnMax = document.getElementById('btnMaximizeCompose');
    if (btnMax) {
      btnMax.title = 'Maximizar';
      btnMax.innerHTML = '<i class="fa-regular fa-square"></i>';
    }
    elComposeForm.reset();
    if (recipientManagerTo) recipientManagerTo.clear();
    if (recipientManagerCc) recipientManagerCc.clear();
    elComposeEditor.innerHTML = '';
    if (elComposeStatusMsg) elComposeStatusMsg.textContent = '';
    attachedFiles = [];
    renderComposeAttachments();
  }

  function renderComposeAttachments() {
    const listEl = document.getElementById('composeAttachmentsList');
    if (elComposeAttachedFilesCount) {
      elComposeAttachedFilesCount.textContent = attachedFiles.length > 0 ? `${attachedFiles.length} anexo(s)` : '';
    }
    if (!listEl) return;
    listEl.innerHTML = '';
    if (attachedFiles.length === 0) {
      listEl.style.display = 'none';
      return;
    }
    listEl.style.display = 'flex';
    attachedFiles.forEach((file, idx) => {
      const chip = document.createElement('div');
      chip.className = 'compose-att-chip';
      chip.title = `${file.name} (${formatBytes(file.size)})`;
      chip.innerHTML = `
        <i class="${getAttachmentIcon(file.name, file.type)}"></i>
        <span class="compose-att-name">${escapeHtml(file.name)}</span>
        <span class="compose-att-size">(${formatBytes(file.size)})</span>
        <button type="button" class="compose-att-remove" title="Remover anexo" aria-label="Remover anexo">
          <i class="fa-solid fa-xmark"></i>
        </button>
      `;
      chip.querySelector('.compose-att-remove').onclick = () => {
        attachedFiles.splice(idx, 1);
        renderComposeAttachments();
      };
      listEl.appendChild(chip);
    });
  }

  function addComposeFiles(fileList) {
    const incoming = Array.from(fileList || []);
    if (incoming.length === 0) return;
    let added = 0;
    incoming.forEach(file => {
      const exists = attachedFiles.some(
        f => f.name === file.name && f.size === file.size && f.lastModified === file.lastModified
      );
      if (!exists) {
        attachedFiles.push(file);
        added++;
      }
    });
    renderComposeAttachments();
    if (added > 0) {
      showToast(`${added} arquivo(s) anexado(s)!`);
    }
  }

  // Shared Email Templates (Global across all users)
  async function loadEmailTemplates() {
    try {
      const res = await fetch('/api/email-templates');
      const data = await res.json();
      if (data.success && Array.isArray(data.templates)) {
        emailTemplates = data.templates;
        populateComposeTemplateSelect();
        renderSharedTemplatesList();
      }
    } catch (e) {
      console.warn('Erro ao carregar modelos de e-mail:', e);
    }
  }

  function populateComposeTemplateSelect() {
    const sel = document.getElementById('composeTemplateSelect');
    if (!sel) return;
    sel.innerHTML = '<option value="">Modelo padrão...</option>';
    emailTemplates.forEach(tpl => {
      const opt = document.createElement('option');
      opt.value = tpl.id;
      opt.textContent = tpl.shared === false ? `${tpl.name} (Privado)` : tpl.name;
      sel.appendChild(opt);
    });
  }

  function resetEmailTemplateForm() {
    editingTemplateId = null;
    const idEl = document.getElementById('tplEditId');
    const nameEl = document.getElementById('tplNameInput');
    const subjEl = document.getElementById('tplSubjectInput');
    const bodyEl = document.getElementById('tplBodyEditor');
    const sharedEl = document.getElementById('tplSharedCheckbox');
    const delBtn = document.getElementById('btnDeleteEmailTemplate');
    if (idEl) idEl.value = '';
    if (nameEl) nameEl.value = '';
    if (subjEl) subjEl.value = '';
    if (bodyEl) bodyEl.innerHTML = '';
    if (sharedEl) sharedEl.checked = true;
    if (delBtn) delBtn.style.display = 'none';
    renderSharedTemplatesList();
    nameEl?.focus();
  }

  function loadTemplateIntoEditor(tpl) {
    if (!tpl) return;
    editingTemplateId = tpl.id;
    const idEl = document.getElementById('tplEditId');
    const nameEl = document.getElementById('tplNameInput');
    const subjEl = document.getElementById('tplSubjectInput');
    const bodyEl = document.getElementById('tplBodyEditor');
    const sharedEl = document.getElementById('tplSharedCheckbox');
    const delBtn = document.getElementById('btnDeleteEmailTemplate');
    if (idEl) idEl.value = tpl.id;
    if (nameEl) nameEl.value = tpl.name || '';
    if (subjEl) subjEl.value = tpl.subject || '';
    if (bodyEl) bodyEl.innerHTML = tpl.html || '';
    if (sharedEl) sharedEl.checked = tpl.shared !== false;
    if (delBtn) delBtn.style.display = 'inline-flex';
    renderSharedTemplatesList();
  }

  function renderSharedTemplatesList() {
    const listEl = document.getElementById('sharedTemplatesList');
    if (!listEl) return;
    listEl.innerHTML = '';
    if (!emailTemplates || emailTemplates.length === 0) {
      listEl.innerHTML = '<div class="templates-empty-state">Nenhum modelo criado ainda.<br>Crie o primeiro ao lado.</div>';
      return;
    }
    emailTemplates.forEach(tpl => {
      const isShared = tpl.shared !== false;
      const card = document.createElement('div');
      card.className = `template-item-card ${editingTemplateId === tpl.id ? 'active' : ''}`;
      card.innerHTML = `
        <div class="template-item-name" title="${escapeHtml(tpl.name)}">${escapeHtml(tpl.name)}</div>
        <div class="template-item-sub">${tpl.subject ? escapeHtml(tpl.subject) : 'Sem assunto padrão'}</div>
        <div class="template-scope-badge ${isShared ? 'shared' : 'private'}">
          <i class="fa-solid ${isShared ? 'fa-users' : 'fa-lock'}"></i>
          <span>${isShared ? 'Compartilhado' : 'Somente eu'}</span>
        </div>
      `;
      card.onclick = () => loadTemplateIntoEditor(tpl);
      listEl.appendChild(card);
    });
  }

  function applyEmailTemplateToCompose(templateId) {
    if (!templateId) return;
    const tpl = emailTemplates.find(t => t.id === templateId);
    if (!tpl) return;

    if (tpl.subject && elComposeSubject && !elComposeSubject.value.trim()) {
      elComposeSubject.value = tpl.subject;
    }

    if (elComposeEditor) {
      const sigEl = elComposeEditor.querySelector('[data-outlook-signature="true"]');
      const sigHtml = sigEl ? sigEl.outerHTML : '';
      elComposeEditor.innerHTML = `${tpl.html || ''}${sigHtml ? '<p><br></p>' + sigHtml : ''}`;
      elComposeEditor.focus();
    }

    showToast(`Modelo "${tpl.name}" aplicado!`);
  }

  async function handleSendEmail(e) {
    e.preventDefault();
    if (recipientManagerTo) recipientManagerTo.commitInput();
    if (recipientManagerCc) recipientManagerCc.commitInput();

    const toEmails = recipientManagerTo ? recipientManagerTo.getAllEmails() : (elComposeTo?.value || '').split(',').map(s => s.trim()).filter(Boolean);
    if (toEmails.length === 0) {
      alert('Por favor, informe pelo menos um destinatário no campo "Para:".');
      if (recipientManagerTo) recipientManagerTo.focus();
      return;
    }

    const invalidTo = recipientManagerTo ? recipientManagerTo.getInvalidEmails() : [];
    const invalidCc = recipientManagerCc ? recipientManagerCc.getInvalidEmails() : [];
    const allInvalid = [...invalidTo, ...invalidCc];

    if (allInvalid.length > 0) {
      const msg = allInvalid.length === 1
        ? `O endereço "${allInvalid[0]}" não possui um formato de e-mail válido.\nPor favor, corrija ou remova antes de enviar.`
        : `Os seguintes endereços não possuem formato de e-mail válido:\n• ${allInvalid.join('\n• ')}\n\nPor favor, corrija ou remova antes de enviar.`;
      alert(msg);
      if (invalidTo.length > 0 && recipientManagerTo) recipientManagerTo.focus();
      else if (recipientManagerCc) recipientManagerCc.focus();
      return;
    }

    stopDraftAutoSaveTimer();
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
    if (currentDraftUid) {
      formData.append('draftUid', String(currentDraftUid));
    }

    attachedFiles.forEach(f => formData.append('attachments', f));

    let sentOk = false;
    try {
      const res = await fetch('/api/smtp/send', {
        method: 'POST',
        body: formData
      });
      const data = await res.json();
      if (data.success) {
        sentOk = true;
        showToast(data.message || 'Email enviado com sucesso!');
        closeCompose();
        loadFolders();
        if (currentFolder.toLowerCase().includes('sent') || currentFolder.toLowerCase().includes('enviad') || isDraftsFolderActive()) {
          loadMessages();
        }
      } else {
        alert(data.message || 'Falha ao enviar email');
      }
    } catch (err) {
      alert('Erro de comunicação com o servidor: ' + err.message);
    } finally {
      elBtnSendMail.disabled = false;
      if (!sentOk && elComposeModal.style.display === 'flex') {
        elComposeStatusMsg.textContent = '';
        startDraftAutoSaveTimer();
      }
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
        const elProfileMenuEmail = document.getElementById('profileMenuEmail');
        if (elProfileMenuEmail) {
          elProfileMenuEmail.textContent = data.imap.user || '';
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
      const activeEmail = (data.imap?.user || SESSION_STORE.getUser() || '').toLowerCase();
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

    const activeUser = (document.getElementById('cfgImapUser')?.value || SESSION_STORE.getUser() || '').trim().toLowerCase();
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
    initRecipientChipManagers();

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

    // Draggable Pane Resizer (Lateral Width or Bottom Height)
    const elPaneResizer = document.getElementById('paneResizer');
    const elContentSplit = document.getElementById('contentSplit');
    const elListPane = document.getElementById('messageListPane');
    if (elPaneResizer && elContentSplit && elListPane) {
      let isResizing = false;

      elPaneResizer.addEventListener('mousedown', (e) => {
        e.preventDefault();
        isResizing = true;
        elPaneResizer.classList.add('dragging');
        document.body.classList.add('is-resizing');
      });

      window.addEventListener('mousemove', (e) => {
        if (!isResizing) return;
        const splitRect = elContentSplit.getBoundingClientRect();

        if (readingPaneLayout === 'bottom') {
          const newHeight = Math.max(120, Math.min(splitRect.height - 150, e.clientY - splitRect.top));
          elListPane.style.height = `${newHeight}px`;
        } else {
          const newWidth = Math.max(240, Math.min(splitRect.width - 280, e.clientX - splitRect.left));
          elListPane.style.width = `${newWidth}px`;
        }
      });

      window.addEventListener('mouseup', () => {
        if (!isResizing) return;
        isResizing = false;
        elPaneResizer.classList.remove('dragging');
        document.body.classList.remove('is-resizing');

        if (readingPaneLayout === 'bottom' && elListPane.style.height) {
          localStorage.setItem('outlook_split_height', String(parseInt(elListPane.style.height, 10)));
        } else if (readingPaneLayout !== 'bottom' && elListPane.style.width) {
          localStorage.setItem('outlook_split_width', String(parseInt(elListPane.style.width, 10)));
        }
      });

      elPaneResizer.addEventListener('dblclick', () => {
        if (readingPaneLayout === 'bottom') {
          elListPane.style.height = '';
          localStorage.removeItem('outlook_split_height');
        } else {
          elListPane.style.width = '';
          localStorage.removeItem('outlook_split_width');
        }
        showToast('Tamanho do painel restaurado para o padrão');
      });
    }

    // App Rail Navigation (Mail <-> Calendar <-> Contacts <-> To Do) & Actions
    const railMailBtn = document.getElementById('railMailBtn');
    const railCalendarBtn = document.getElementById('railCalendarBtn');
    const railContactsBtn = document.getElementById('railContactsBtn');
    const railTodoBtn = document.getElementById('railTodoBtn');
    if (railMailBtn) railMailBtn.onclick = () => switchAppSection('mail');
    if (railCalendarBtn) railCalendarBtn.onclick = () => switchAppSection('calendar');
    if (railContactsBtn) railContactsBtn.onclick = () => switchAppSection('contacts');
    if (railTodoBtn) railTodoBtn.onclick = () => switchAppSection('todo');

    // Mobile Portrait Navigation Controls
    const btnMobileFolders = document.getElementById('btnMobileFolders');
    const mobileDrawerBackdrop = document.getElementById('mobileDrawerBackdrop');
    const btnMobileBackToList = document.getElementById('btnMobileBackToList');
    if (btnMobileFolders) {
      btnMobileFolders.onclick = (e) => {
        e.stopPropagation();
        const contentSplit = document.getElementById('contentSplit');
        if (contentSplit && contentSplit.style.display === 'none') {
          switchAppSection('mail');
        }
        toggleMobileFoldersDrawer();
      };
    }
    if (mobileDrawerBackdrop) {
      mobileDrawerBackdrop.onclick = () => toggleMobileFoldersDrawer(false);
    }
    if (btnMobileBackToList) {
      btnMobileBackToList.onclick = () => closeMobileReadingPane();
    }
    if (window.matchMedia) {
      const mobileMq = window.matchMedia('(max-width: 768px)');
      const onViewportChange = () => {
        if (!mobileMq.matches) {
          toggleMobileFoldersDrawer(false);
        }
        if (messages.length > 0) {
          renderMessageList();
        }
      };
      if (mobileMq.addEventListener) mobileMq.addEventListener('change', onViewportChange);
      else if (mobileMq.addListener) mobileMq.addListener(onViewportChange);
    }

    // To Do Event Listeners
    const todoAddForm = document.getElementById('todoAddForm');
    const todoQuickStar = document.getElementById('todoQuickStar');
    const btnEmailToTodo = document.getElementById('btnEmailToTodo');
    if (todoAddForm) todoAddForm.onsubmit = addTodoTaskFromForm;
    if (todoQuickStar) {
      todoQuickStar.onclick = () => {
        todoQuickImportant = !todoQuickImportant;
        todoQuickStar.classList.toggle('active', todoQuickImportant);
        todoQuickStar.innerHTML = `<i class="fa-${todoQuickImportant ? 'solid' : 'regular'} fa-star"></i>`;
      };
    }
    document.querySelectorAll('[data-todo-filter]').forEach(btn => {
      btn.onclick = () => {
        document.querySelectorAll('[data-todo-filter]').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        todoFilter = btn.getAttribute('data-todo-filter') || 'today';
        renderTodos();
      };
    });
    if (btnEmailToTodo) {
      btnEmailToTodo.onclick = async () => {
        if (!selectedMessage) return;
        const sender = selectedMessage.from?.name || selectedMessage.from?.address || '';
        const subj = selectedMessage.subject || '(Sem assunto)';
        const ok = await updateTodoTask({
          title: subj,
          dueDate: toLocalDateStr(new Date()),
          important: true,
          completed: false,
          notes: sender ? `E-mail de: ${sender}` : ''
        });
        if (ok) showToast('E-mail adicionado às tarefas do To Do!');
      };
    }

    const btnNewCalEvent = document.getElementById('btnNewCalEvent');
    const btnCalToday = document.getElementById('btnCalToday');
    const btnCalPrevMonth = document.getElementById('btnCalPrevMonth');
    const btnCalNextMonth = document.getElementById('btnCalNextMonth');
    const btnAddEventOnSelectedDay = document.getElementById('btnAddEventOnSelectedDay');
    const calendarEventForm = document.getElementById('calendarEventForm');
    const btnCloseCalEventModal = document.getElementById('btnCloseCalEventModal');
    const btnCancelCalEventModal = document.getElementById('btnCancelCalEventModal');
    const calendarEventModal = document.getElementById('calendarEventModal');

    if (btnNewCalEvent) btnNewCalEvent.onclick = () => openCalendarEventModal(null, calSelectedDateStr);
    if (btnAddEventOnSelectedDay) btnAddEventOnSelectedDay.onclick = () => openCalendarEventModal(null, calSelectedDateStr);
    if (btnCalToday) {
      btnCalToday.onclick = () => {
        calCurrentDate = new Date();
        calSelectedDateStr = toLocalDateStr(new Date());
        renderCalendar();
      };
    }
    if (btnCalPrevMonth) {
      btnCalPrevMonth.onclick = () => {
        calCurrentDate = new Date(calCurrentDate.getFullYear(), calCurrentDate.getMonth() - 1, 1);
        renderCalendar();
      };
    }
    if (btnCalNextMonth) {
      btnCalNextMonth.onclick = () => {
        calCurrentDate = new Date(calCurrentDate.getFullYear(), calCurrentDate.getMonth() + 1, 1);
        renderCalendar();
      };
    }
    if (calendarEventForm) calendarEventForm.onsubmit = saveCalendarEvent;
    if (btnCloseCalEventModal && calendarEventModal) {
      btnCloseCalEventModal.onclick = () => { calendarEventModal.style.display = 'none'; };
    }
    if (btnCancelCalEventModal && calendarEventModal) {
      btnCancelCalEventModal.onclick = () => { calendarEventModal.style.display = 'none'; };
    }

    const btnNewContact = document.getElementById('btnNewContact');
    const btnNewContactGroup = document.getElementById('btnNewContactGroup');
    const btnSyncContacts = document.getElementById('btnSyncContacts');
    const contactsSearchInput = document.getElementById('contactsSearchInput');
    const contactForm = document.getElementById('contactForm');
    const btnCloseContactModal = document.getElementById('btnCloseContactModal');
    const btnCancelContactModal = document.getElementById('btnCancelContactModal');
    const contactModal = document.getElementById('contactModal');
    const contactGroupForm = document.getElementById('contactGroupForm');
    const btnCloseContactGroupModal = document.getElementById('btnCloseContactGroupModal');
    const btnCancelContactGroupModal = document.getElementById('btnCancelContactGroupModal');
    const contactGroupModal = document.getElementById('contactGroupModal');
    const btnDeleteContactGroup = document.getElementById('btnDeleteContactGroup');

    if (btnNewContact) btnNewContact.onclick = () => openContactModal(null);
    if (btnNewContactGroup) btnNewContactGroup.onclick = () => openContactGroupModal(null);
    if (btnSyncContacts) btnSyncContacts.onclick = () => syncContactsFromServer();
    if (contactsSearchInput) {
      contactsSearchInput.oninput = (e) => {
        contactsSearchQuery = e.target.value || '';
        renderContactsList();
      };
    }
    if (contactForm) contactForm.onsubmit = saveContact;
    if (btnCloseContactModal && contactModal) {
      btnCloseContactModal.onclick = () => { contactModal.style.display = 'none'; };
    }
    if (btnCancelContactModal && contactModal) {
      btnCancelContactModal.onclick = () => { contactModal.style.display = 'none'; };
    }
    if (contactGroupForm) contactGroupForm.onsubmit = saveContactGroup;
    if (btnCloseContactGroupModal && contactGroupModal) {
      btnCloseContactGroupModal.onclick = () => { contactGroupModal.style.display = 'none'; };
    }
    if (btnCancelContactGroupModal && contactGroupModal) {
      btnCancelContactGroupModal.onclick = () => { contactGroupModal.style.display = 'none'; };
    }
    if (btnDeleteContactGroup) {
      btnDeleteContactGroup.onclick = () => {
        const id = document.getElementById('contactGroupId')?.value || '';
        const name = document.getElementById('contactGroupNameInput')?.value || '';
        if (id) deleteContactGroup(id, name);
      };
    }

    // Clear group members button
    const btnClearGroupMembers = document.getElementById('btnClearGroupMembers');
    if (btnClearGroupMembers) {
      btnClearGroupMembers.onclick = () => {
        if (groupMemberManager) groupMemberManager.clear();
        updateGroupMembersCount();
        // Uncheck all picker checkboxes
        document.querySelectorAll('#groupContactsPicker input[type="checkbox"]').forEach(cb => { cb.checked = false; });
      };
    }

    // Group picker filter
    const groupPickerFilter = document.getElementById('groupPickerFilter');
    if (groupPickerFilter) {
      groupPickerFilter.oninput = () => {
        const selectedEmails = groupMemberManager ? groupMemberManager.getAllEmails().map(e => e.toLowerCase()) : [];
        renderContactGroupsPicker(selectedEmails);
      };
    }

    // Contacts filter tabs
    document.querySelectorAll('.contacts-tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.contacts-tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentContactsTab = btn.dataset.tab || 'all';
        renderContactsList();
      });
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

    // Folder / Subfolder creation modal listeners
    const btnCreateRootFolder = document.getElementById('btnCreateRootFolder');
    const createFolderForm = document.getElementById('createFolderForm');
    const btnCloseCreateFolderModal = document.getElementById('btnCloseCreateFolderModal');
    const btnCancelCreateFolderModal = document.getElementById('btnCancelCreateFolderModal');
    const createFolderModal = document.getElementById('createFolderModal');

    if (btnCreateRootFolder) {
      btnCreateRootFolder.onclick = (e) => {
        e.stopPropagation();
        openCreateFolderModal(currentFolder || 'INBOX');
      };
    }
    if (createFolderForm) createFolderForm.onsubmit = handleCreateFolderSubmit;
    if (btnCloseCreateFolderModal) btnCloseCreateFolderModal.onclick = closeCreateFolderModal;
    if (btnCancelCreateFolderModal) btnCancelCreateFolderModal.onclick = closeCreateFolderModal;
    if (createFolderModal) {
      createFolderModal.addEventListener('click', (e) => {
        if (e.target === createFolderModal) closeCreateFolderModal();
      });
    }

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
      if (confirm('Deseja descartar este rascunho?')) {
        const draftToRemove = currentDraftUid;
        closeCompose();
        if (draftToRemove) {
          fetch(`/api/imap/draft/${encodeURIComponent(draftToRemove)}`, { method: 'DELETE' })
            .then(() => {
              loadFolders();
              if (isDraftsFolderActive()) loadMessages();
            })
            .catch(() => {});
        }
      }
    };

    elComposeFileInput.onchange = (e) => {
      addComposeFiles(e.target.files);
      e.target.value = '';
    };

    // Drag & Drop files onto Compose Modal to attach
    (function setupComposeDragDrop() {
      const overlay = document.getElementById('composeDropOverlay');
      if (!elComposeModal || !overlay) return;

      let dragDepth = 0;
      const hasFiles = (e) => {
        const types = e.dataTransfer && e.dataTransfer.types;
        return types && Array.from(types).includes('Files');
      };
      const showOverlay = () => { overlay.style.display = 'flex'; };
      const hideOverlay = () => { dragDepth = 0; overlay.style.display = 'none'; };

      elComposeModal.addEventListener('dragenter', (e) => {
        if (!hasFiles(e)) return;
        e.preventDefault();
        e.stopPropagation();
        dragDepth++;
        showOverlay();
      });

      elComposeModal.addEventListener('dragover', (e) => {
        if (!hasFiles(e)) return;
        e.preventDefault();
        e.stopPropagation();
        if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
        showOverlay();
      });

      elComposeModal.addEventListener('dragleave', (e) => {
        if (!hasFiles(e)) return;
        e.preventDefault();
        e.stopPropagation();
        dragDepth--;
        if (dragDepth <= 0) hideOverlay();
      });

      elComposeModal.addEventListener('drop', (e) => {
        if (!hasFiles(e)) return;
        e.preventDefault();
        e.stopPropagation();
        hideOverlay();
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
          addComposeFiles(e.dataTransfer.files);
        }
      });

      // Prevent browser from opening a file if dropped slightly outside while compose is open
      window.addEventListener('dragover', (e) => {
        if (elComposeModal.style.display === 'flex' && hasFiles(e)) {
          e.preventDefault();
        }
      });
      window.addEventListener('drop', (e) => {
        if (elComposeModal.style.display === 'flex' && hasFiles(e)) {
          e.preventDefault();
          hideOverlay();
        }
      });
    })();

    // ── Insert Table Popover ──────────────────────────────────────
    (function setupTableInserter() {
      const ROWS = 6, COLS = 8;
      const wrap     = document.getElementById('tableInsertWrap');
      const popover  = document.getElementById('tablePickerPopover');
      const grid     = document.getElementById('tablePickerGrid');
      const label    = document.getElementById('tablePickerLabel');
      const btn      = document.getElementById('btnInsertTable');
      const manRows  = document.getElementById('tableManualRows');
      const manCols  = document.getElementById('tableManualCols');
      const btnManual= document.getElementById('btnInsertTableManual');
      if (!btn || !popover || !grid) return;

      // Build grid cells
      for (let r = 1; r <= ROWS; r++) {
        for (let c = 1; c <= COLS; c++) {
          const cell = document.createElement('div');
          cell.className = 'table-picker-cell';
          cell.dataset.row = r;
          cell.dataset.col = c;
          grid.appendChild(cell);
        }
      }

      function highlightCells(maxR, maxC) {
        grid.querySelectorAll('.table-picker-cell').forEach(cell => {
          const r = +cell.dataset.row, c = +cell.dataset.col;
          cell.classList.toggle('hovered', r <= maxR && c <= maxC);
        });
        if (label) label.textContent = `${maxR} × ${maxC}`;
      }

      grid.addEventListener('mouseover', e => {
        const cell = e.target.closest('.table-picker-cell');
        if (cell) highlightCells(+cell.dataset.row, +cell.dataset.col);
      });

      grid.addEventListener('mouseleave', () => highlightCells(0, 0));

      grid.addEventListener('click', e => {
        const cell = e.target.closest('.table-picker-cell');
        if (cell) {
          insertTable(+cell.dataset.row, +cell.dataset.col);
          closePopover();
        }
      });

      if (btnManual) {
        btnManual.onclick = () => {
          const r = Math.max(1, Math.min(20, parseInt(manRows?.value) || 3));
          const c = Math.max(1, Math.min(10, parseInt(manCols?.value) || 3));
          insertTable(r, c);
          closePopover();
        };
      }

      btn.addEventListener('click', e => {
        e.stopPropagation();
        const isOpen = popover.style.display !== 'none';
        closePopover();
        if (!isOpen) {
          popover.style.display = '';
          highlightCells(0, 0);
        }
      });

      function closePopover() {
        if (popover) popover.style.display = 'none';
      }

      document.addEventListener('click', e => {
        if (!wrap?.contains(e.target)) closePopover();
      });

      document.addEventListener('keydown', e => {
        if (e.key === 'Escape') closePopover();
      });

      function insertTable(rows, cols) {
        const editor = document.getElementById('composeEditor');
        if (!editor) return;
        editor.focus();

        // Build table HTML
        const headerRow = `<tr>${Array.from({length: cols}, () =>
          `<th style="border:1px solid #adb5bd;padding:6px 10px;background:#e9f0fb;font-weight:700;min-width:60px;">​</th>`
        ).join('')}</tr>`;
        const bodyRows = Array.from({length: rows - 1}, () =>
          `<tr>${Array.from({length: cols}, () =>
            `<td style="border:1px solid #adb5bd;padding:6px 10px;min-width:60px;">​</td>`
          ).join('')}</tr>`
        ).join('');

        const tableHtml = `<table style="border-collapse:collapse;width:auto;max-width:100%;margin:8px 0;"><tbody>${headerRow}${bodyRows}</tbody></table><p><br></p>`;

        // Restore saved selection or insert at cursor
        const sel = window.getSelection();
        if (sel && sel.rangeCount > 0) {
          const range = sel.getRangeAt(0);
          if (editor.contains(range.commonAncestorContainer)) {
            range.deleteContents();
            const frag = range.createContextualFragment(tableHtml);
            range.insertNode(frag);
            sel.collapseToEnd();
            return;
          }
        }
        // Fallback: append at end
        editor.insertAdjacentHTML('beforeend', tableHtml);
        editor.scrollTop = editor.scrollHeight;
      }
    })();
    // ── End Insert Table ──────────────────────────────────────────

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
      if (tabId === 'tabBackup') {
        populateBackupFolderSelects();
      }
      if (tabId === 'tabTemplates') {
        loadEmailTemplates();
      }
    };

    // Compose Optional Template Selector
    const composeTemplateSelect = document.getElementById('composeTemplateSelect');
    if (composeTemplateSelect) {
      composeTemplateSelect.onchange = (e) => {
        const val = e.target.value;
        if (val) {
          applyEmailTemplateToCompose(val);
        }
      };
    }

    // Shared Email Templates Tab Listeners
    const btnNewEmailTemplate = document.getElementById('btnNewEmailTemplate');
    const btnSaveEmailTemplate = document.getElementById('btnSaveEmailTemplate');
    const btnDeleteEmailTemplate = document.getElementById('btnDeleteEmailTemplate');
    const tplBodyEditor = document.getElementById('tplBodyEditor');

    document.querySelectorAll('[data-tpl-cmd]').forEach(btn => {
      btn.onclick = () => {
        if (tplBodyEditor) tplBodyEditor.focus();
        document.execCommand(btn.getAttribute('data-tpl-cmd'), false, null);
      };
    });

    if (btnNewEmailTemplate) {
      btnNewEmailTemplate.onclick = () => resetEmailTemplateForm();
    }

    if (btnSaveEmailTemplate) {
      btnSaveEmailTemplate.onclick = async () => {
        const id = (document.getElementById('tplEditId')?.value || '').trim();
        const name = (document.getElementById('tplNameInput')?.value || '').trim();
        const subject = (document.getElementById('tplSubjectInput')?.value || '').trim();
        const html = (tplBodyEditor?.innerHTML || '').trim();
        const shared = document.getElementById('tplSharedCheckbox')?.checked !== false;

        if (!name) {
          showToast('Informe o nome do modelo.');
          document.getElementById('tplNameInput')?.focus();
          return;
        }
        if (!html || !(tplBodyEditor?.innerText || '').trim()) {
          showToast('Digite o conteúdo do modelo de e-mail.');
          tplBodyEditor?.focus();
          return;
        }

        btnSaveEmailTemplate.disabled = true;
        try {
          const res = await fetch('/api/email-templates', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: id || undefined, name, subject, html, shared })
          });
          const data = await res.json();
          if (data.success) {
            emailTemplates = data.templates || [];
            editingTemplateId = data.template?.id || null;
            if (data.template) loadTemplateIntoEditor(data.template);
            populateComposeTemplateSelect();
            showToast(shared ? `Modelo "${name}" salvo e compartilhado com todos!` : `Modelo privado "${name}" salvo!`);
          } else {
            showToast(data.error || 'Erro ao salvar modelo.');
          }
        } catch (err) {
          showToast('Erro ao salvar modelo: ' + err.message);
        } finally {
          btnSaveEmailTemplate.disabled = false;
        }
      };
    }

    if (btnDeleteEmailTemplate) {
      btnDeleteEmailTemplate.onclick = async () => {
        const id = (document.getElementById('tplEditId')?.value || '').trim();
        if (!id) return;
        const tpl = emailTemplates.find(t => t.id === id);
        if (!confirm(`Deseja excluir o modelo compartilhado "${tpl?.name || ''}" para todos os usuários?`)) return;

        try {
          const res = await fetch(`/api/email-templates/${encodeURIComponent(id)}`, { method: 'DELETE' });
          const data = await res.json();
          if (data.success) {
            emailTemplates = data.templates || [];
            populateComposeTemplateSelect();
            resetEmailTemplateForm();
            showToast('Modelo excluído com sucesso.');
          } else {
            showToast(data.error || 'Erro ao excluir modelo.');
          }
        } catch (err) {
          showToast('Erro ao excluir modelo: ' + err.message);
        }
      };
    }

    // Backup Tab (Export / Import) Helpers & Listeners
    let backupSelectedFiles = [];

    function populateBackupFolderSelects() {
      const exportSel = document.getElementById('backupExportFolder');
      const importSel = document.getElementById('backupImportTargetFolder');
      const flat = flattenFolderTree(buildFolderTree(folders));

      if (exportSel) {
        const prev = exportSel.value || '__ALL__';
        exportSel.innerHTML = '<option value="__ALL__">Todas as pastas e subpastas</option>';
        flat.forEach(({ folder, depth }) => {
          const opt = document.createElement('option');
          opt.value = folder.path;
          const prefix = depth > 0 ? `${'\u00A0\u00A0'.repeat(depth)}↳ ` : '';
          opt.textContent = `${prefix}${folder.name}`;
          exportSel.appendChild(opt);
        });
        exportSel.value = Array.from(exportSel.options).some(o => o.value === prev) ? prev : '__ALL__';
      }

      if (importSel) {
        const prev = importSel.value || '__ORIGINAL__';
        importSel.innerHTML = '<option value="__ORIGINAL__">Manter pastas originais do backup</option>';
        flat.forEach(({ folder, depth }) => {
          const opt = document.createElement('option');
          opt.value = folder.path;
          const prefix = depth > 0 ? `${'\u00A0\u00A0'.repeat(depth)}↳ ` : '';
          opt.textContent = `${prefix}${folder.name}`;
          importSel.appendChild(opt);
        });
        importSel.value = Array.from(importSel.options).some(o => o.value === prev) ? prev : '__ORIGINAL__';
      }
    }

    function setBackupStatus(type, messageHtml) {
      const box = document.getElementById('backupStatusBox');
      if (!box) return;
      if (!messageHtml) {
        box.style.display = 'none';
        box.innerHTML = '';
        return;
      }
      box.className = `backup-status-box ${type}`;
      box.innerHTML = messageHtml;
      box.style.display = 'flex';
    }

    function updateBackupFilesLabel(fileList) {
      backupSelectedFiles = Array.from(fileList || []);
      const label = document.getElementById('backupSelectedFilesLabel');
      if (!label) return;
      if (backupSelectedFiles.length === 0) {
        label.textContent = 'Clique ou arraste arquivos .json / .eml aqui';
      } else if (backupSelectedFiles.length === 1) {
        label.textContent = `${backupSelectedFiles[0].name} (${formatBytes(backupSelectedFiles[0].size)})`;
      } else {
        const totalSize = backupSelectedFiles.reduce((acc, f) => acc + (f.size || 0), 0);
        label.textContent = `${backupSelectedFiles.length} arquivos selecionados (${formatBytes(totalSize)})`;
      }
    }

    const backupDropzone = document.getElementById('backupDropzone');
    const backupImportFileInput = document.getElementById('backupImportFileInput');
    if (backupDropzone && backupImportFileInput) {
      backupDropzone.onclick = () => backupImportFileInput.click();
      backupImportFileInput.onchange = (e) => updateBackupFilesLabel(e.target.files);
      backupDropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        backupDropzone.classList.add('dragover');
      });
      backupDropzone.addEventListener('dragleave', () => {
        backupDropzone.classList.remove('dragover');
      });
      backupDropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        backupDropzone.classList.remove('dragover');
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
          updateBackupFilesLabel(e.dataTransfer.files);
        }
      });
    }

    const btnExportBackup = document.getElementById('btnExportBackup');
    if (btnExportBackup) {
      btnExportBackup.onclick = async () => {
        const folderPath = document.getElementById('backupExportFolder')?.value || '__ALL__';
        const limitPerFolder = Number(document.getElementById('backupExportLimit')?.value || 0);

        btnExportBackup.disabled = true;
        btnExportBackup.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>Exportando backup...</span>';
        setBackupStatus('info', '<i class="fa-solid fa-spinner fa-spin"></i> <span>Baixando mensagens do servidor IMAP para gerar o arquivo de backup. Aguarde...</span>');

        try {
          const res = await fetch('/api/backup/export', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ folderPath, limitPerFolder })
          });
          if (!res.ok) {
            const errData = await res.json().catch(() => ({}));
            throw new Error(errData.error || `Falha HTTP ${res.status}`);
          }

          const count = Number(res.headers.get('X-Backup-Count') || 0);
          const filename = res.headers.get('X-Backup-Filename') || `backup-emails-${new Date().toISOString().slice(0, 10)}.json`;
          const blob = await res.blob();

          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = filename;
          document.body.appendChild(a);
          a.click();
          a.remove();
          setTimeout(() => URL.revokeObjectURL(url), 5000);

          setBackupStatus('success', `<i class="fa-solid fa-circle-check"></i> <span>Backup concluído! <strong>${count}</strong> e-mail(s) exportado(s) para <strong>${escapeHtml(filename)}</strong> (${formatBytes(blob.size)}).</span>`);
          showToast(`Backup exportado com ${count} e-mail(s)!`);
        } catch (err) {
          setBackupStatus('error', `<i class="fa-solid fa-circle-exclamation"></i> <span>Erro ao exportar backup: ${escapeHtml(err.message)}</span>`);
          showToast('Erro ao exportar backup.');
        } finally {
          btnExportBackup.disabled = false;
          btnExportBackup.innerHTML = '<i class="fa-solid fa-download"></i> <span>Exportar Backup Agora</span>';
        }
      };
    }

    const btnImportBackup = document.getElementById('btnImportBackup');
    if (btnImportBackup) {
      btnImportBackup.onclick = async () => {
        if (!backupSelectedFiles || backupSelectedFiles.length === 0) {
          setBackupStatus('error', '<i class="fa-solid fa-circle-exclamation"></i> <span>Selecione primeiro um arquivo de backup (.json) ou arquivos .eml do computador.</span>');
          return;
        }

        const targetFolder = document.getElementById('backupImportTargetFolder')?.value || '__ORIGINAL__';
        const formData = new FormData();
        formData.append('targetFolder', targetFolder);
        backupSelectedFiles.forEach(f => formData.append('files', f));

        btnImportBackup.disabled = true;
        btnImportBackup.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>Importando...</span>';
        setBackupStatus('info', '<i class="fa-solid fa-spinner fa-spin"></i> <span>Importando mensagens para o servidor IMAP. Isso pode levar alguns instantes...</span>');

        try {
          const res = await fetch('/api/backup/import', {
            method: 'POST',
            body: formData
          });
          const data = await res.json();
          if (!res.ok || !data.success) {
            throw new Error(data.error || `Falha HTTP ${res.status}`);
          }

          if (Array.isArray(data.folders)) {
            folders = data.folders;
            renderFolders();
            populateBackupFolderSelects();
          } else {
            await loadFolders();
          }
          await loadMessages();

          updateBackupFilesLabel([]);
          if (backupImportFileInput) backupImportFileInput.value = '';

          const extra = [];
          if (data.foldersCreated > 0) extra.push(`${data.foldersCreated} pasta(s) criada(s)`);
          if (data.errors > 0) extra.push(`${data.errors} falha(s)`);
          const extraStr = extra.length > 0 ? ` (${extra.join(', ')})` : '';

          setBackupStatus('success', `<i class="fa-solid fa-circle-check"></i> <span>Importação concluída! <strong>${data.imported}</strong> e-mail(s) restaurado(s)${extraStr}.</span>`);
          showToast(`${data.imported} e-mail(s) importado(s) com sucesso!`);
        } catch (err) {
          setBackupStatus('error', `<i class="fa-solid fa-circle-exclamation"></i> <span>Erro ao importar backup: ${escapeHtml(err.message)}</span>`);
          showToast('Erro ao importar backup.');
        } finally {
          btnImportBackup.disabled = false;
          btnImportBackup.innerHTML = '<i class="fa-solid fa-upload"></i> <span>Importar Backup</span>';
        }
      };
    }

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
      switchSettingsTab('tabTemplates');
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
      elLoginEmailInput.oninput = () => renderLoginAutocomplete();
      elLoginEmailInput.onblur = () => {
        setTimeout(() => {
          const box = document.getElementById('loginEmailSuggestions');
          if (box) box.style.display = 'none';
        }, 150);
      };
      elLoginEmailInput.onkeydown = (e) => {
        const box = document.getElementById('loginEmailSuggestions');
        if (!box || box.style.display === 'none') return;
        const items = Array.from(box.querySelectorAll('.login-autocomplete-item'));
        if (items.length === 0) return;

        let activeIdx = items.findIndex(el => el.classList.contains('active'));
        if (e.key === 'ArrowDown') {
          e.preventDefault();
          if (activeIdx >= 0) items[activeIdx].classList.remove('active');
          activeIdx = (activeIdx + 1) % items.length;
          items[activeIdx].classList.add('active');
        } else if (e.key === 'ArrowUp') {
          e.preventDefault();
          if (activeIdx >= 0) items[activeIdx].classList.remove('active');
          activeIdx = (activeIdx - 1 + items.length) % items.length;
          items[activeIdx].classList.add('active');
        } else if (e.key === 'Tab' || (e.key === 'Enter' && activeIdx >= 0)) {
          e.preventDefault();
          const chosen = items[activeIdx >= 0 ? activeIdx : 0].querySelector('span')?.textContent;
          if (chosen) {
            elLoginEmailInput.value = chosen;
            box.style.display = 'none';
            if (elLoginPasswordInput) elLoginPasswordInput.focus();
          }
        } else if (e.key === 'Escape') {
          box.style.display = 'none';
        }
      };
    }

    const triggerAccountSwitch = () => {
      endUserSession();
    };

    const elBtnToggleTheme = document.getElementById('btnToggleTheme');
    const elThemeToggleIcon = document.getElementById('themeToggleIcon');
    const applyTheme = (theme) => {
      const isDark = theme === 'dark';
      document.body.classList.toggle('dark-mode', isDark);
      if (elThemeToggleIcon) {
        elThemeToggleIcon.className = isDark ? 'fa-solid fa-sun' : 'fa-solid fa-moon';
      }
      if (elBtnToggleTheme) {
        elBtnToggleTheme.title = isDark ? 'Alternar para modo claro' : 'Alternar para modo escuro';
      }
    };
    applyTheme(localStorage.getItem('outlook_theme') === 'dark' ? 'dark' : 'light');
    if (elBtnToggleTheme) {
      elBtnToggleTheme.onclick = () => {
        const nextTheme = document.body.classList.contains('dark-mode') ? 'light' : 'dark';
        localStorage.setItem('outlook_theme', nextTheme);
        applyTheme(nextTheme);
      };
    }

    const elUserAvatar = document.getElementById('userAvatar');
    const elUserProfileMenu = document.getElementById('userProfileMenu');
    const elBtnMenuLogout = document.getElementById('btnMenuLogout');

    if (elUserAvatar && elUserProfileMenu) {
      elUserAvatar.onclick = (e) => {
        e.stopPropagation();
        const isOpen = elUserProfileMenu.style.display === 'block';
        elUserProfileMenu.style.display = isOpen ? 'none' : 'block';
      };

      document.addEventListener('click', (e) => {
        if (!e.target.closest('.user-menu-wrapper')) {
          elUserProfileMenu.style.display = 'none';
        }
      });
    }

    if (elBtnMenuLogout) {
      elBtnMenuLogout.onclick = () => {
        if (elUserProfileMenu) elUserProfileMenu.style.display = 'none';
        triggerAccountSwitch();
      };
    }

    const elBtnMenuAddAccount = document.getElementById('btnMenuAddAccount');
    const elAddAccountModal = document.getElementById('addAccountModal');
    const elAddAccountForm = document.getElementById('addAccountForm');
    const elBtnCloseAddAccountModal = document.getElementById('btnCloseAddAccountModal');
    const elBtnCancelAddAccountModal = document.getElementById('btnCancelAddAccountModal');

    if (elBtnMenuAddAccount) {
      elBtnMenuAddAccount.onclick = () => {
        if (elUserProfileMenu) elUserProfileMenu.style.display = 'none';
        openAddAccountModal();
      };
    }
    if (elAddAccountForm) elAddAccountForm.onsubmit = handleAddAccountSubmit;
    const elAddAccEmailInput = document.getElementById('addAccEmailInput');
    if (elAddAccEmailInput) {
      elAddAccEmailInput.addEventListener('input', () => {
        const val = (elAddAccEmailInput.value || '').trim().toLowerCase();
        const isGmail = val.endsWith('@gmail.com') || val.endsWith('@googlemail.com');
        const hint = document.getElementById('addAccGmailHint');
        if (hint) hint.style.display = isGmail ? 'block' : 'none';
        if (isGmail) {
          const ih = document.getElementById('addAccImapHost');
          const ip = document.getElementById('addAccImapPort');
          const sh = document.getElementById('addAccSmtpHost');
          const sp = document.getElementById('addAccSmtpPort');
          if (ih && !ih.value) ih.value = 'imap.gmail.com';
          if (ip && !ip.value) ip.value = '993';
          if (sh && !sh.value) sh.value = 'smtp.gmail.com';
          if (sp && !sp.value) sp.value = '465';
        }
      });
    }
    if (elBtnCloseAddAccountModal) elBtnCloseAddAccountModal.onclick = closeAddAccountModal;
    if (elBtnCancelAddAccountModal) elBtnCancelAddAccountModal.onclick = closeAddAccountModal;
    if (elAddAccountModal) {
      elAddAccountModal.addEventListener('click', (e) => {
        if (e.target === elAddAccountModal) closeAddAccountModal();
      });
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

    // Keyboard Shortcuts (Ctrl+E for search, Ctrl+N for new mail, Ctrl+Enter to send, Delete to move to Trash)
    window.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && (e.key === 'Enter' || e.code === 'NumpadEnter')) {
        const isComposeOpen = elComposeModal && elComposeModal.style.display === 'flex';
        if (isComposeOpen && elComposeForm && !elBtnSendMail?.disabled) {
          e.preventDefault();
          e.stopPropagation();
          if (typeof elComposeForm.requestSubmit === 'function') {
            elComposeForm.requestSubmit();
          } else {
            handleSendEmail(e);
          }
          return;
        }
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'e') {
        e.preventDefault();
        elSearchInput.focus();
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'n') {
        e.preventDefault();
        openCompose();
      }
      if (e.key === 'Delete' && !e.ctrlKey && !e.metaKey && !e.altKey) {
        const target = e.target;
        const tag = target ? target.tagName : '';
        const isCheckbox = tag === 'INPUT' && target.type === 'checkbox';
        const isTypingField = (!isCheckbox && (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT')) ||
                              Boolean(target && target.isContentEditable);
        if (isTypingField) return;

        const openModalIds = [
          'composeModal', 'settingsModal', 'loginOverlay',
          'contactModal', 'contactGroupModal', 'calendarEventModal',
          'createFolderModal', 'attachmentPreviewModal', 'attPreviewModal'
        ];
        const anyModalOpen = openModalIds.some(id => {
          const el = document.getElementById(id);
          return el && window.getComputedStyle(el).display !== 'none';
        });
        if (anyModalOpen) return;

        // Only trigger when Mail list view is visible
        if (!elMessageListScroll || elMessageListScroll.offsetParent === null) return;

        if (selectedUids.size > 0) {
          e.preventDefault();
          bulkDeleteEmails(selectedUids);
        } else if (selectedMessage) {
          e.preventDefault();
          deleteEmail(selectedMessage.uid);
        }
      }
    }, true);
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

