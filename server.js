const express = require('express');
const cors = require('cors');
const fs = require('fs');
const os = require('os');
const path = require('path');
const multer = require('multer');

const ImapService = require('./imap-service');
const SmtpService = require('./smtp-service');

const app = express();
const PORT = process.env.PORT || 4000;
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Disable browser caching for real-time frontend updates
app.use((req, res, next) => {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  next();
});

app.use(express.static(path.join(__dirname, 'public'), { etag: false, maxAge: 0 }));

// Load configuration (supports both local disk and Vercel read-only /tmp environment)
const CONFIG_FILE = path.join(__dirname, 'config.json');
const TMP_CONFIG_FILE = path.join(os.tmpdir(), 'outlook_webmail_config.json');

function saveConfigSafe(cfg) {
  const jsonStr = JSON.stringify(cfg, null, 2);
  try {
    fs.writeFileSync(CONFIG_FILE, jsonStr, 'utf-8');
  } catch (err) {
    // On Vercel / Serverless, /var/task is read-only (EROFS). Save to /tmp instead!
    try {
      fs.writeFileSync(TMP_CONFIG_FILE, jsonStr, 'utf-8');
    } catch (tmpErr) {
      console.warn('[Config] Ambiente somente-leitura, mantendo configurações em memória:', tmpErr.message);
    }
  }
}

let appConfig = {
  imap: {
    host: 'h58.servidorhh.com',
    port: 993,
    secure: true,
    user: 'clodoaldo.silva@itecensino.com.br',
    password: ''
  },
  smtp: {
    host: 'h58.servidorhh.com',
    port: 465,
    secure: true,
    user: 'clodoaldo.silva@itecensino.com.br',
    password: ''
  },
  signature: {
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
  },
  savedEmails: ['clodoaldo.silva@itecensino.com.br'],
  signaturesByUser: {},
  contactsByUser: {},
  deletedContactsByUser: {},
  calendarByUser: {},
  todosByUser: {},
  layout: { readingPane: 'right' },
  demoModeIfDisconnected: true
};

// Load from /tmp (if running on Vercel warm instance) or local config.json
[CONFIG_FILE, TMP_CONFIG_FILE].forEach(file => {
  if (fs.existsSync(file)) {
    try {
      const raw = fs.readFileSync(file, 'utf-8');
      const parsed = JSON.parse(raw);
      appConfig = {
        ...appConfig,
        ...parsed,
        imap: { ...appConfig.imap, ...(parsed.imap || {}) },
        smtp: { ...appConfig.smtp, ...(parsed.smtp || {}) },
        signature: { ...appConfig.signature, ...(parsed.signature || {}) },
        savedEmails: Array.isArray(parsed.savedEmails) && parsed.savedEmails.length > 0
          ? parsed.savedEmails
          : [parsed.imap?.user || 'clodoaldo.silva@itecensino.com.br'].filter(Boolean),
        signaturesByUser: { ...(appConfig.signaturesByUser || {}), ...(parsed.signaturesByUser || {}) },
        contactsByUser: { ...(appConfig.contactsByUser || {}), ...(parsed.contactsByUser || {}) },
        deletedContactsByUser: { ...(appConfig.deletedContactsByUser || {}), ...(parsed.deletedContactsByUser || {}) },
        calendarByUser: { ...(appConfig.calendarByUser || {}), ...(parsed.calendarByUser || {}) },
        todosByUser: { ...(appConfig.todosByUser || {}), ...(parsed.todosByUser || {}) }
      };
    } catch (e) {
      console.error(`Erro ao carregar ${file}:`, e.message);
    }
  }
});

const imapService = new ImapService(appConfig);
const smtpService = new SmtpService(appConfig);

// Stateless Session Middleware for Vercel Serverless (restores user credentials across Lambda instances)
app.use((req, res, next) => {
  const hdrUser = req.headers['x-imap-user'];
  const hdrPassB64 = req.headers['x-imap-pass'];
  if (hdrUser && hdrPassB64) {
    try {
      const decodedPass = Buffer.from(String(hdrPassB64), 'base64').toString('utf8');
      const cleanUser = String(hdrUser).trim().toLowerCase();
      if (cleanUser && decodedPass && (appConfig.imap.user !== cleanUser || appConfig.imap.password !== decodedPass)) {
        appConfig.imap.user = cleanUser;
        appConfig.imap.password = decodedPass;
        appConfig.smtp.user = cleanUser;
        appConfig.smtp.password = decodedPass;
        imapService.updateConfig(appConfig);
        smtpService.updateConfig(appConfig);
      }
    } catch (e) {}
  }
  next();
});

// API Endpoints

// Login & Switch User Account
app.post('/api/auth/login', async (req, res) => {
  try {
    const email = (req.body.email || '').trim().toLowerCase();
    const password = req.body.password || '';
    const rememberEmail = req.body.rememberEmail !== false;

    if (!email || !password) {
      return res.status(400).json({ success: false, message: 'Por favor, informe o e-mail e a senha.' });
    }

    const tempConfig = {
      ...appConfig,
      imap: {
        ...appConfig.imap,
        user: email,
        password: password
      }
    };

    const testService = new ImapService(tempConfig);
    const testResult = await testService.testConnection();

    if (!testResult.success) {
      return res.status(401).json({
        success: false,
        message: testResult.message || 'Falha na autenticação. Verifique seu e-mail e senha.'
      });
    }

    // Update active credentials
    appConfig.imap.user = email;
    appConfig.imap.password = password;
    appConfig.smtp.user = email;
    appConfig.smtp.password = password;

    // Save email to savedEmails list
    if (!Array.isArray(appConfig.savedEmails)) {
      appConfig.savedEmails = [];
    }
    if (rememberEmail) {
      appConfig.savedEmails = [email, ...appConfig.savedEmails.filter(e => e.toLowerCase() !== email)];
    }

    // Load user-specific signature if available, or adapt default
    appConfig.signaturesByUser = appConfig.signaturesByUser || {};
    if (appConfig.signaturesByUser[email]) {
      appConfig.signature = appConfig.signaturesByUser[email];
    } else if (email !== 'clodoaldo.silva@itecensino.com.br') {
      const namePart = email.split('@')[0].replace(/[._-]+/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
      appConfig.signature = {
        ...appConfig.signature,
        fullName: namePart,
        email: email,
        html: `<div style="font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 13px; color: #323130; border-left: 3px solid #0078d4; padding-left: 12px; margin-top: 12px;"><div style="font-size: 15px; font-weight: 700; color: #0078d4;">${namePart}</div><div style="font-size: 12px; font-weight: 600; color: #605e5c;">ITEC Ensino - Volta Redonda</div><div style="margin-top: 4px; font-size: 12px; color: #323130;">✉️ <a href="mailto:${email}" style="color: #0078d4; text-decoration: none;">${email}</a> &nbsp;|&nbsp; 🌐 www.itecensino.com.br</div></div>`
      };
      appConfig.signaturesByUser[email] = appConfig.signature;
    }

    saveConfigSafe(appConfig);
    imapService.updateConfig(appConfig);
    smtpService.updateConfig(appConfig);

    res.json({
      success: true,
      user: email,
      savedEmails: appConfig.savedEmails,
      signature: appConfig.signature
    });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao autenticar: ' + err.message });
  }
});

// Remove a saved email from list
app.delete('/api/auth/saved-email', (req, res) => {
  try {
    const emailToRemove = (req.query.email || '').trim().toLowerCase();
    appConfig.savedEmails = (appConfig.savedEmails || []).filter(e => e.toLowerCase() !== emailToRemove);
    saveConfigSafe(appConfig);
    res.json({ success: true, savedEmails: appConfig.savedEmails });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Get config
app.get('/api/config', (req, res) => {
  res.json({
    imap: {
      host: appConfig.imap.host,
      port: appConfig.imap.port,
      secure: appConfig.imap.secure,
      user: appConfig.imap.user,
      hasPassword: !!appConfig.imap.password
    },
    smtp: {
      host: appConfig.smtp.host,
      port: appConfig.smtp.port,
      secure: appConfig.smtp.secure,
      user: appConfig.smtp.user,
      hasPassword: !!appConfig.smtp.password
    },
    signature: appConfig.signature,
    layout: appConfig.layout || { readingPane: 'right' },
    savedEmails: appConfig.savedEmails || []
  });
});

// Save config
app.post('/api/config', (req, res) => {
  try {
    const { imap, smtp, signature, layout } = req.body;
    if (imap) {
      appConfig.imap = {
        ...appConfig.imap,
        ...imap,
        password: imap.password !== undefined && imap.password !== '' ? imap.password : appConfig.imap.password
      };
    }
    if (smtp) {
      appConfig.smtp = {
        ...appConfig.smtp,
        ...smtp,
        password: smtp.password !== undefined && smtp.password !== '' ? smtp.password : appConfig.smtp.password
      };
    }
    if (signature) {
      appConfig.signature = {
        ...appConfig.signature,
        ...signature
      };
      if (appConfig.imap.user) {
        appConfig.signaturesByUser = appConfig.signaturesByUser || {};
        appConfig.signaturesByUser[appConfig.imap.user.toLowerCase()] = appConfig.signature;
      }
    }
    if (layout) {
      appConfig.layout = {
        ...(appConfig.layout || { readingPane: 'right' }),
        ...layout
      };
    }

    saveConfigSafe(appConfig);
    imapService.updateConfig(appConfig);
    smtpService.updateConfig(appConfig);

    res.json({ success: true, message: 'Configurações salvas com sucesso!', signature: appConfig.signature, layout: appConfig.layout });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Test IMAP Connection
app.post('/api/imap/test', async (req, res) => {
  if (req.body && req.body.imap) {
    const tempConfig = {
      ...appConfig,
      imap: {
        ...appConfig.imap,
        ...req.body.imap
      }
    };
    const testService = new ImapService(tempConfig);
    const result = await testService.testConnection();
    return res.json(result);
  }
  const result = await imapService.testConnection();
  res.json(result);
});

// Test SMTP Connection
app.post('/api/smtp/test', async (req, res) => {
  if (req.body && req.body.smtp) {
    const tempConfig = {
      ...appConfig,
      smtp: {
        ...appConfig.smtp,
        ...req.body.smtp
      }
    };
    const result = await smtpService.testConnection(tempConfig);
    return res.json(result);
  }
  const result = await smtpService.testConnection();
  res.json(result);
});

// Get folders
app.get('/api/imap/folders', async (req, res) => {
  try {
    const folders = await imapService.getFolders();
    res.json({ success: true, folders });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Get messages in folder
app.get('/api/imap/messages', async (req, res) => {
  try {
    const folder = req.query.folder || 'INBOX';
    const filter = req.query.filter || 'all';
    const search = req.query.search || '';
    const page = parseInt(req.query.page, 10) || 1;
    const limit = parseInt(req.query.limit, 10) || 50;

    const result = await imapService.getMessages(folder, { search, filter, page, limit });
    res.json({ success: true, ...result });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Get single message detail
app.get('/api/imap/message/:uid', async (req, res) => {
  try {
    const folder = req.query.folder || 'INBOX';
    const uid = req.params.uid;
    console.log(`[API /api/imap/message] Buscando mensagem UID: ${uid} (pasta: ${folder})`);
    const message = await imapService.getMessageDetail(folder, uid);
    if (!message) {
      console.warn(`[API /api/imap/message] Mensagem UID: ${uid} não encontrada`);
      return res.status(404).json({ success: false, error: 'Mensagem não encontrada' });
    }
    res.json({ success: true, message });
  } catch (err) {
    console.error(`[API /api/imap/message] Erro ao buscar mensagem:`, err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Download or preview single attachment
app.get('/api/imap/attachment/:uid/:partId', async (req, res) => {
  try {
    const folder = req.query.folder || 'INBOX';
    const uid = req.params.uid;
    const partId = req.params.partId;
    const isDownload = req.query.download === '1' || req.query.download === 'true';

    const att = await imapService.getAttachment(folder, uid, partId);
    if (!att || !att.content) {
      return res.status(404).json({ success: false, error: 'Anexo não encontrado' });
    }

    const safeFilename = encodeURIComponent(att.filename).replace(/['()]/g, escape).replace(/\*/g, '%2A');
    const disposition = isDownload ? 'attachment' : 'inline';

    res.setHeader('Content-Type', att.contentType || 'application/octet-stream');
    res.setHeader('Content-Disposition', `${disposition}; filename*=UTF-8''${safeFilename}`);
    res.send(att.content);
  } catch (err) {
    console.error(`[API /api/imap/attachment] Erro:`, err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Update flags (seen, flagged)
app.post('/api/imap/flags', async (req, res) => {
  try {
    const { folder = 'INBOX', uid, seen, flagged } = req.body;
    const result = await imapService.setFlags(folder, uid, { seen, flagged });
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Move message
app.post('/api/imap/move', async (req, res) => {
  try {
    const { fromFolder = 'INBOX', toFolder, uid } = req.body;
    const result = await imapService.moveMessage(fromFolder, toFolder, uid);
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Delete message
app.delete('/api/imap/message/:uid', async (req, res) => {
  try {
    const folder = req.query.folder || 'INBOX';
    const uid = req.params.uid;
    const result = await imapService.deleteMessage(folder, uid);
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Bulk delete messages
app.post('/api/imap/bulk-delete', async (req, res) => {
  try {
    const { folder = 'INBOX', uids = [] } = req.body;
    const result = await imapService.deleteMessages(folder, uids);
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Bulk update flags (seen, flagged)
app.post('/api/imap/bulk-flags', async (req, res) => {
  try {
    const { folder = 'INBOX', uids = [], seen, flagged } = req.body;
    const result = await imapService.setFlagsBulk(folder, uids, { seen, flagged });
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Send message via SMTP
app.post('/api/smtp/send', upload.array('attachments'), async (req, res) => {
  try {
    const { to, cc, bcc, subject, text, html } = req.body;
    const attachments = req.files || [];

    const result = await smtpService.sendMail({
      to,
      cc,
      bcc,
      subject,
      text,
      html,
      attachments
    });

    if (result.success && imapService.hasCredentials()) {
      imapService.appendSentMessage({ to, cc, subject, text, html }).catch(() => {});
    }

    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ========================
// Contacts Endpoints
// ========================
function getActiveUserKey() {
  return (appConfig.imap?.user || 'default').trim().toLowerCase();
}

app.get('/api/contacts', (req, res) => {
  const userKey = getActiveUserKey();
  const list = (appConfig.contactsByUser && appConfig.contactsByUser[userKey]) || [];
  res.json({ success: true, contacts: list });
});

app.post('/api/contacts/sync', async (req, res) => {
  try {
    const userKey = getActiveUserKey();
    const serverContacts = await imapService.extractContactsFromServer(100);

    appConfig.contactsByUser = appConfig.contactsByUser || {};
    appConfig.deletedContactsByUser = appConfig.deletedContactsByUser || {};

    const existing = appConfig.contactsByUser[userKey] || [];
    const deletedSet = new Set((appConfig.deletedContactsByUser[userKey] || []).map(e => e.toLowerCase()));
    const mergedMap = new Map();

    existing.forEach(c => {
      if (c && c.email) mergedMap.set(c.email.toLowerCase(), c);
    });

    let addedCount = 0;
    serverContacts.forEach(sc => {
      const em = sc.email.toLowerCase();
      if (deletedSet.has(em)) return;
      if (!mergedMap.has(em)) {
        mergedMap.set(em, sc);
        addedCount++;
      } else {
        const cur = mergedMap.get(em);
        if ((!cur.name || cur.name === em.split('@')[0]) && sc.name) {
          cur.name = sc.name;
        }
      }
    });

    const finalContacts = Array.from(mergedMap.values()).sort((a, b) =>
      (a.name || a.email).localeCompare(b.name || b.email, 'pt-BR')
    );

    appConfig.contactsByUser[userKey] = finalContacts;
    saveConfigSafe(appConfig);

    res.json({
      success: true,
      added: addedCount,
      total: finalContacts.length,
      contacts: finalContacts
    });
  } catch (err) {
    console.error('[Contacts Sync Error]', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/contacts', (req, res) => {
  try {
    const userKey = getActiveUserKey();
    const email = (req.body.email || '').trim().toLowerCase();
    const name = (req.body.name || '').trim();
    const phone = (req.body.phone || '').trim();
    const company = (req.body.company || '').trim();
    const originalEmail = (req.body.originalEmail || '').trim().toLowerCase();

    if (!email || !email.includes('@')) {
      return res.status(400).json({ success: false, error: 'Informe um endereço de e-mail válido.' });
    }

    appConfig.contactsByUser = appConfig.contactsByUser || {};
    appConfig.deletedContactsByUser = appConfig.deletedContactsByUser || {};

    let list = appConfig.contactsByUser[userKey] || [];

    // Remove old entry if editing email
    if (originalEmail && originalEmail !== email) {
      list = list.filter(c => c.email.toLowerCase() !== originalEmail);
    }

    const idx = list.findIndex(c => c.email.toLowerCase() === email);
    const contactObj = {
      email,
      name: name || email.split('@')[0],
      phone,
      company,
      source: 'manual'
    };

    if (idx >= 0) {
      list[idx] = { ...list[idx], ...contactObj };
    } else {
      list.push(contactObj);
    }

    // Remove from deleted list if user explicitly re-adds it
    if (Array.isArray(appConfig.deletedContactsByUser[userKey])) {
      appConfig.deletedContactsByUser[userKey] = appConfig.deletedContactsByUser[userKey].filter(
        e => e.toLowerCase() !== email
      );
    }

    list.sort((a, b) => (a.name || a.email).localeCompare(b.name || b.email, 'pt-BR'));
    appConfig.contactsByUser[userKey] = list;
    saveConfigSafe(appConfig);

    res.json({ success: true, contacts: list, contact: contactObj });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.delete('/api/contacts', (req, res) => {
  try {
    const userKey = getActiveUserKey();
    const email = (req.query.email || '').trim().toLowerCase();
    if (!email) {
      return res.status(400).json({ success: false, error: 'E-mail não informado.' });
    }

    appConfig.contactsByUser = appConfig.contactsByUser || {};
    appConfig.deletedContactsByUser = appConfig.deletedContactsByUser || {};

    const list = (appConfig.contactsByUser[userKey] || []).filter(c => c.email.toLowerCase() !== email);
    appConfig.contactsByUser[userKey] = list;

    const deletedList = appConfig.deletedContactsByUser[userKey] || [];
    if (!deletedList.includes(email)) deletedList.push(email);
    appConfig.deletedContactsByUser[userKey] = deletedList;

    saveConfigSafe(appConfig);
    res.json({ success: true, contacts: list });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ========================
// Calendar Events Endpoints
// ========================
app.get('/api/calendar', (req, res) => {
  const userKey = getActiveUserKey();
  const events = (appConfig.calendarByUser && appConfig.calendarByUser[userKey]) || [];
  res.json({ success: true, events });
});

app.post('/api/calendar', async (req, res) => {
  try {
    const userKey = getActiveUserKey();
    const {
      id,
      title,
      date,
      startTime = '09:00',
      endTime = '10:00',
      location = '',
      description = '',
      color = '#0078d4',
      attendees = '',
      sendInvite = false
    } = req.body;

    if (!title || !date) {
      return res.status(400).json({ success: false, error: 'Informe o título e a data do evento.' });
    }

    appConfig.calendarByUser = appConfig.calendarByUser || {};
    const list = appConfig.calendarByUser[userKey] || [];
    const eventId = id || ('evt_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7));

    const eventObj = {
      id: eventId,
      title: String(title).trim(),
      date: String(date).trim(),
      startTime: String(startTime).trim(),
      endTime: String(endTime).trim(),
      location: String(location).trim(),
      description: String(description).trim(),
      color: color || '#0078d4',
      attendees: String(attendees).trim()
    };

    const idx = list.findIndex(ev => ev.id === eventId);
    if (idx >= 0) {
      list[idx] = eventObj;
    } else {
      list.push(eventObj);
    }

    list.sort((a, b) => `${a.date} ${a.startTime}`.localeCompare(`${b.date} ${b.startTime}`));
    appConfig.calendarByUser[userKey] = list;
    saveConfigSafe(appConfig);

    // Optional SMTP invite notification
    if (sendInvite && eventObj.attendees && smtpService) {
      const formattedDate = eventObj.date.split('-').reverse().join('/');
      const inviteHtml = `
        <div style="font-family: Arial, sans-serif; font-size: 14px; color: #201f1e; border-left: 4px solid ${eventObj.color}; padding-left: 14px;">
          <h3 style="margin: 0 0 8px 0; color: ${eventObj.color};">Convite de Evento: ${eventObj.title}</h3>
          <p style="margin: 4px 0;"><strong>Data:</strong> ${formattedDate}</p>
          <p style="margin: 4px 0;"><strong>Horário:</strong> ${eventObj.startTime} às ${eventObj.endTime}</p>
          ${eventObj.location ? `<p style="margin: 4px 0;"><strong>Local / Link:</strong> ${eventObj.location}</p>` : ''}
          ${eventObj.description ? `<p style="margin: 8px 0 0 0;">${eventObj.description}</p>` : ''}
        </div>
      `;
      smtpService.sendMail({
        to: eventObj.attendees,
        subject: `Convite: ${eventObj.title} (${formattedDate} ${eventObj.startTime})`,
        html: inviteHtml,
        text: `Convite: ${eventObj.title}\nData: ${formattedDate}\nHorário: ${eventObj.startTime} - ${eventObj.endTime}\nLocal: ${eventObj.location}\n${eventObj.description}`
      }).catch(e => console.warn('[Calendar Invite SMTP warn]:', e.message));
    }

    res.json({ success: true, event: eventObj, events: list });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.delete('/api/calendar/:id', (req, res) => {
  try {
    const userKey = getActiveUserKey();
    const id = req.params.id;
    appConfig.calendarByUser = appConfig.calendarByUser || {};
    const list = (appConfig.calendarByUser[userKey] || []).filter(ev => ev.id !== id);
    appConfig.calendarByUser[userKey] = list;
    saveConfigSafe(appConfig);
    res.json({ success: true, events: list });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ========================
// To Do Tasks Endpoints
// ========================
app.get('/api/todos', (req, res) => {
  const userKey = getActiveUserKey();
  const todos = (appConfig.todosByUser && appConfig.todosByUser[userKey]) || [];
  res.json({ success: true, todos });
});

app.post('/api/todos', (req, res) => {
  try {
    const userKey = getActiveUserKey();
    const {
      id,
      title,
      dueDate = '',
      important = false,
      completed = false,
      notes = ''
    } = req.body;

    if (!title || !String(title).trim()) {
      return res.status(400).json({ success: false, error: 'Informe o título da tarefa.' });
    }

    appConfig.todosByUser = appConfig.todosByUser || {};
    const list = appConfig.todosByUser[userKey] || [];
    const taskId = id || ('todo_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7));

    const existingIdx = list.findIndex(t => t.id === taskId);
    const taskObj = {
      id: taskId,
      title: String(title).trim(),
      dueDate: String(dueDate || '').trim(),
      important: !!important,
      completed: !!completed,
      notes: String(notes || '').trim(),
      createdAt: existingIdx >= 0 ? (list[existingIdx].createdAt || new Date().toISOString()) : new Date().toISOString()
    };

    if (existingIdx >= 0) {
      list[existingIdx] = taskObj;
    } else {
      list.unshift(taskObj);
    }

    appConfig.todosByUser[userKey] = list;
    saveConfigSafe(appConfig);
    res.json({ success: true, todo: taskObj, todos: list });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.delete('/api/todos/:id', (req, res) => {
  try {
    const userKey = getActiveUserKey();
    const id = req.params.id;
    appConfig.todosByUser = appConfig.todosByUser || {};
    const list = (appConfig.todosByUser[userKey] || []).filter(t => t.id !== id);
    appConfig.todosByUser[userKey] = list;
    saveConfigSafe(appConfig);
    res.json({ success: true, todos: list });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`=======================================================`);
    console.log(` Sistema IMAP Outlook Webmail rodando na porta ${PORT} `);
    console.log(` Acesse: http://localhost:${PORT}                      `);
    console.log(` Servidor IMAP Alvo: ${appConfig.imap.host}:993 (SSL) `);
    console.log(` Usuário: ${appConfig.imap.user}                        `);
    console.log(`=======================================================`);
  });
}

module.exports = app;

