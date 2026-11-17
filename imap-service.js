const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');

// Mock data based on Roundcube screenshot for when no credentials are provided
const MOCK_MESSAGES = [
  {
    uid: 101,
    seq: 1,
    folder: 'INBOX',
    subject: 'Acesso Sels Ensino e Escola Web',
    from: { name: 'Daniele Henriques', address: 'daniele.henriques@itecensino.com.br' },
    to: [{ name: 'Clodoaldo Silva', address: 'clodoaldo.silva@itecensino.com.br' }, { name: 'Cristiane Lins', address: 'cristiane.lins@itecensino.com.br' }],
    date: new Date(Date.now() - 1000 * 60 * 60 * 3).toISOString(),
    seen: false,
    flagged: true,
    hasAttachments: true,
    attachments: [
      { filename: 'Manual_Acesso_Sels_Web.pdf', size: 1048576, contentType: 'application/pdf', partId: '1' }
    ],
    snippet: 'Prezados, seguem as orientações atualizadas para o acesso dos novos professores e alunos ao portal Sels Ensino e Escola Web...',
    html: `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; line-height: 1.6;">
        <h3 style="color: #0078d4;">Acesso Sels Ensino e Escola Web</h3>
        <p>Prezados,</p>
        <p>Seguem as orientações atualizadas para o acesso dos novos professores e alunos ao portal <strong>Sels Ensino</strong> e <strong>Escola Web</strong> para o semestre letivo.</p>
        <div style="background: #f3f4f6; border-left: 4px solid #0078d4; padding: 12px; margin: 15px 0;">
          <strong>Link do Portal:</strong> <a href="https://escolaweb.itecensino.com.br" target="_blank" style="color: #0078d4;">escolaweb.itecensino.com.br</a><br/>
          <strong>Credenciais padrão:</strong> Primeiro acesso com CPF e senha temporária cadastrada na secretaria.
        </div>
        <p>Qualquer inconsistência no login ou permissões da turma, favor responder a este chamado.</p>
        <br/>
        <p>Atenciosamente,</p>
        <p><strong>Daniele Henriques</strong><br/>Secretaria Acadêmica - ITEC Ensino</p>
      </div>
    `,
    text: 'Prezados, seguem as orientações atualizadas para o acesso dos novos professores e alunos ao portal Sels Ensino e Escola Web...'
  },
  {
    uid: 102,
    seq: 2,
    folder: 'INBOX',
    subject: 'Re: Confirmação de Frequência Escolar.',
    from: { name: 'Daniele Henriques', address: 'daniele.henriques@itecensino.com.br' },
    to: [{ name: 'Clodoaldo Silva', address: 'clodoaldo.silva@itecensino.com.br' }],
    date: new Date(Date.now() - 1000 * 60 * 60 * 26).toISOString(),
    seen: false,
    flagged: false,
    hasAttachments: false,
    attachments: [],
    snippet: 'Bom dia, Clodoaldo. Verifiquei em sistema e a frequência da aluna foi regularizada conforme solicitado...',
    html: `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; line-height: 1.6;">
        <p>Bom dia, Clodoaldo,</p>
        <p>Verifiquei em sistema e a frequência da turma foi regularizada conforme solicitado na ata da coordenação pedagógica.</p>
        <p>Ela ainda não constava na lista de chamada inicial porque o registro da matrícula estava em fase de homologação financeira. Agora já está tudo normalizado no diário eletrônico.</p>
        <br/>
        <p>Cordialmente,</p>
        <p><strong>Daniele Henriques</strong><br/>Coordenação de Registros - ITEC Ensino</p>
      </div>
    `,
    text: 'Bom dia, Clodoaldo. Verifiquei em sistema e a frequência da aluna foi regularizada conforme solicitado...'
  },
  {
    uid: 103,
    seq: 3,
    folder: 'INBOX',
    subject: 'Controle de ponto novembro/26 Iago',
    from: { name: 'Cristiane Lins', address: 'cristiane.lins@itecensino.com.br' },
    to: [{ name: 'Clodoaldo Silva', address: 'clodoaldo.silva@itecensino.com.br' }, { name: 'Daniele Henriques', address: 'daniele.henriques@itecensino.com.br' }],
    date: new Date(Date.now() - 1000 * 60 * 60 * 48).toISOString(),
    seen: true,
    flagged: true,
    hasAttachments: true,
    attachments: [
      { filename: 'Folha_Ponto_Novembro_2026.xlsx', size: 45056, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', partId: '1' }
    ],
    snippet: 'Olá a todos. Lembramos do fechamento do espelho de ponto referente ao mês de novembro...',
    html: `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; line-height: 1.6;">
        <p>Olá a todos,</p>
        <p>Lembramos do fechamento do espelho de ponto referente ao mês de novembro da unidade Lago.</p>
        <p>Por favor, enviem as assinaturas e eventuais justificativas de ausência até sexta-feira às 17h para processamento da folha.</p>
        <br/>
        <p>Atenciosamente,</p>
        <p><strong>Cristiane Lins</strong><br/>Recursos Humanos - ITEC Ensino</p>
      </div>
    `,
    text: 'Olá a todos. Lembramos do fechamento do espelho de ponto referente ao mês de novembro...'
  },
  {
    uid: 104,
    seq: 4,
    folder: 'INBOX',
    subject: 'Renovação de matrículas - 78I e 78I.2',
    from: { name: 'Albert Espilari', address: 'albert.espilari@itecensino.com.br' },
    to: [{ name: 'Clodoaldo Silva', address: 'clodoaldo.silva@itecensino.com.br' }, { name: 'Debora...', address: 'debora@itecensino.com.br' }],
    date: new Date(Date.now() - 1000 * 60 * 60 * 72).toISOString(),
    seen: true,
    flagged: false,
    hasAttachments: false,
    attachments: [],
    snippet: 'Prezada equipe, iniciamos o período de rematrícula para as turmas técnicas 78I e 78I.2...',
    html: `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; line-height: 1.6;">
        <p>Prezada equipe,</p>
        <p>Iniciamos o período de rematrícula para as turmas técnicas 78I e 78I.2. A adesão dos alunos está em 85% no primeiro ciclo.</p>
        <p>Pedimos atenção especial aos atendimentos com pendências contratuais no portal.</p>
        <br/>
        <p>Abraços,</p>
        <p><strong>Albert Espilari</strong><br/>Diretoria Acadêmica</p>
      </div>
    `,
    text: 'Prezada equipe, iniciamos o período de rematrícula para as turmas técnicas 78I e 78I.2...'
  },
  {
    uid: 105,
    seq: 5,
    folder: 'INBOX',
    subject: 'Aluno como status Pre Matricula',
    from: { name: 'Financeiro ITEC', address: 'financeiro@itecensino.com.br' },
    to: [{ name: 'Clodoaldo Silva', address: 'clodoaldo.silva@itecensino.com.br' }],
    date: new Date(Date.now() - 1000 * 60 * 60 * 96).toISOString(),
    seen: true,
    flagged: false,
    hasAttachments: false,
    attachments: [],
    snippet: 'Identificamos que 14 inscrições ainda constam em Pré Matrícula aguardando confirmação do pagamento da primeira parcela...',
    html: `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; line-height: 1.6;">
        <p>Bom dia Clodoaldo,</p>
        <p>Identificamos que 14 inscrições ainda constam em <em>Pré Matrícula</em> aguardando confirmação do pagamento da primeira parcela.</p>
        <p>O setor de cobrança já efetuou o primeiro contato telefônico.</p>
        <br/>
        <p>Atenciosamente,</p>
        <p><strong>Setor Financeiro</strong><br/>ITEC Ensino</p>
      </div>
    `,
    text: 'Identificamos que 14 inscrições ainda constam em Pré Matrícula...'
  },
  {
    uid: 106,
    seq: 6,
    folder: 'Sent',
    subject: 'Re: Confirmação de Frequência Escolar.',
    from: { name: 'Clodoaldo Silva', address: 'clodoaldo.silva@itecensino.com.br' },
    to: [{ name: 'Daniele Henriques', address: 'daniele.henriques@itecensino.com.br' }],
    date: new Date(Date.now() - 1000 * 60 * 60 * 27).toISOString(),
    seen: true,
    flagged: false,
    hasAttachments: false,
    attachments: [],
    snippet: 'Daniele, por gentileza verificar a lista de frequência da aluna referente ao módulo II...',
    html: `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; line-height: 1.6;">
        <p>Daniele,</p>
        <p>Por gentileza verificar a lista de frequência da aluna referente ao módulo II, pois os pais entraram em contato com a coordenação.</p>
        <br/>
        <p>Obrigado,</p>
        <p><strong>Clodoaldo Silva</strong><br/>ITEC Ensino</p>
      </div>
    `,
    text: 'Daniele, por gentileza verificar a lista de frequência da aluna referente ao módulo II...'
  }
];

class ImapService {
  constructor(config) {
    this.config = config || {};
    this.mockMessages = [...MOCK_MESSAGES];
    this.activeClient = null;
    this.isConnecting = false;
  }

  hasCredentials() {
    return !!(this.config.imap?.host && this.config.imap?.user && this.config.imap?.password);
  }

  updateConfig(config) {
    this.config = { ...this.config, ...config };
    if (this.activeClient) {
      try { this.activeClient.logout(); } catch (e) {}
      this.activeClient = null;
    }
  }

  createClient() {
    if (!this.hasCredentials()) {
      return null;
    }

    const port = parseInt(this.config.imap.port, 10) || 993;
    const isSecure = port === 993 ? (this.config.imap.secure !== false) : false;

    return new ImapFlow({
      host: this.config.imap.host,
      port: port,
      secure: isSecure,
      auth: {
        user: this.config.imap.user.trim(),
        pass: this.config.imap.password
      },
      logger: false,
      tls: {
        rejectUnauthorized: false
      }
    });
  }

  // Persistent connection manager to eliminate rate limits and ETIMEDOUT
  async getClient() {
    if (!this.hasCredentials()) return null;

    if (this.activeClient && this.activeClient.usable) {
      return this.activeClient;
    }

    if (this.activeClient) {
      try { await this.activeClient.logout(); } catch (e) {}
      this.activeClient = null;
    }

    const client = this.createClient();
    if (!client) return null;

    await client.connect();
    this.activeClient = client;

    client.on('close', () => {
      if (this.activeClient === client) this.activeClient = null;
    });

    client.on('error', (err) => {
      console.warn('[IMAP Persistent Connection Warning]:', err.message);
      if (this.activeClient === client) this.activeClient = null;
    });

    return this.activeClient;
  }

  formatError(err) {
    if (!err) return 'Erro desconhecido ao comunicar com o servidor IMAP.';

    const serverText = err.responseText || (err.response && String(err.response)) || '';
    const isAuthError = err.authenticationFailed || 
                        err.serverResponseCode === 'AUTHENTICATIONFAILED' || 
                        /authentication failed|login failed|auth|credential|invalid user|password/i.test(serverText);

    if (isAuthError) {
      return `Falha de autenticação: O servidor ${this.config.imap?.host || 'IMAP'} rejeitou o usuário ou a senha ("Authentication failed"). Por favor, confirme se a senha digitada é exatamente a mesma que você usa no Webmail Roundcube.`;
    }

    if (err.code === 'ETIMEDOUT') {
      return `Tempo limite esgotado ao tentar alcançar o servidor ${this.config.imap?.host}:${this.config.imap?.port}.`;
    }

    if (err.code === 'ECONNREFUSED') {
      return `Conexão recusada em ${this.config.imap?.host}:${this.config.imap?.port}.`;
    }

    if (err.code === 'ENOTFOUND') {
      return `Servidor não encontrado: "${this.config.imap?.host}".`;
    }

    if (serverText) {
      return `Resposta do servidor: ${serverText}`;
    }

    return err.message || 'Erro inesperado no protocolo IMAP.';
  }

  async testConnection() {
    if (this.activeClient) {
      try { await this.activeClient.logout(); } catch (e) {}
      this.activeClient = null;
    }

    const client = this.createClient();
    if (!client) {
      return { 
        success: false, 
        message: 'Por favor, preencha o Host, o Usuário e a Senha do seu email antes de testar a conexão.' 
      };
    }

    try {
      await client.connect();
      await client.logout();
      return { 
        success: true, 
        message: `Conexão IMAP estabelecida e autenticada com sucesso no servidor ${this.config.imap.host}!` 
      };
    } catch (err) {
      console.error('[IMAP testConnection Error]', err);
      return { 
        success: false, 
        message: this.formatError(err)
      };
    }
  }

  async getFolders() {
    let client = null;
    try {
      client = await this.getClient();
    } catch (err) {
      console.warn('[IMAP getFolders connect error]:', err.message);
    }

    if (!client) {
      if (this.hasCredentials()) return [];
      return [
        { path: 'INBOX', name: 'Caixa de entrada', specialUse: '\\Inbox', unread: 2, total: 5 },
        { path: 'Sent', name: 'Itens Enviados', specialUse: '\\Sent', unread: 0, total: 1 },
        { path: 'Drafts', name: 'Rascunhos', specialUse: '\\Drafts', unread: 0, total: 0 },
        { path: 'Junk', name: 'Lixo Eletrônico', specialUse: '\\Junk', unread: 0, total: 0 },
        { path: 'Trash', name: 'Itens Excluídos', specialUse: '\\Trash', unread: 0, total: 0 }
      ];
    }

    try {
      const list = await client.list();
      const folders = [];

      for (const item of list) {
        let unread = 0;
        let total = 0;
        try {
          const status = await client.status(item.path, { messages: true, unseen: true });
          unread = status.unseen || 0;
          total = status.messages || 0;
        } catch (e) {
          // ignore status on special unselectable folders
        }

        let friendlyName = item.name;
        if (item.specialUse === '\\Inbox' || item.name.toUpperCase() === 'INBOX') friendlyName = 'Caixa de entrada';
        else if (item.specialUse === '\\Sent' || /sent|enviados/i.test(item.name)) friendlyName = 'Itens Enviados';
        else if (item.specialUse === '\\Drafts' || /drafts|rascunhos/i.test(item.name)) friendlyName = 'Rascunhos';
        else if (item.specialUse === '\\Junk' || /junk|spam/i.test(item.name)) friendlyName = 'Lixo Eletrônico';
        else if (item.specialUse === '\\Trash' || /trash|lixeira|excluídos/i.test(item.name)) friendlyName = 'Itens Excluídos';
        else if (item.specialUse === '\\Archive' || /archive|arquivo/i.test(item.name)) friendlyName = 'Arquivo Morto';

        folders.push({
          path: item.path,
          name: friendlyName,
          originalName: item.name,
          specialUse: item.specialUse || '',
          unread,
          total
        });
      }

      return folders;
    } catch (err) {
      console.warn('[IMAP getFolders Error]:', err.message);
      return [];
    }
  }

  getMockMessages(folder = 'INBOX', options = {}) {
    const { search = '', filter = 'all' } = options;
    let msgs = this.mockMessages.filter(m => {
      if (folder.toUpperCase() === 'INBOX') return m.folder === 'INBOX';
      if (folder.toUpperCase() === 'SENT') return m.folder === 'Sent';
      return m.folder.toLowerCase() === folder.toLowerCase();
    });

    if (filter === 'unread') msgs = msgs.filter(m => !m.seen);
    if (filter === 'flagged') msgs = msgs.filter(m => m.flagged);
    if (filter === 'attachments') msgs = msgs.filter(m => m.hasAttachments);
    if (search) {
      const q = search.toLowerCase();
      msgs = msgs.filter(m => 
        (m.subject && m.subject.toLowerCase().includes(q)) ||
        (m.snippet && m.snippet.toLowerCase().includes(q)) ||
        (m.from && (m.from.name.toLowerCase().includes(q) || m.from.address.toLowerCase().includes(q)))
      );
    }

    return { folder, total: msgs.length, messages: msgs };
  }

  async getMessages(folder = 'INBOX', options = {}) {
    const { search = '', filter = 'all', limit = 50 } = options;

    let client = null;
    try {
      client = await this.getClient();
    } catch (err) {
      console.warn(`[IMAP getMessages client error]:`, err.message);
    }

    if (!client) {
      if (!this.hasCredentials()) {
        return this.getMockMessages(folder, options);
      }
      return { folder, total: 0, messages: [], error: 'Falha na conexão IMAP' };
    }

    try {
      const lock = await client.getMailboxLock(folder);
      const messages = [];

      try {
        const mailboxStatus = client.mailbox;
        const total = mailboxStatus.exists || 0;

        if (total === 0) {
          return { folder, total: 0, messages: [] };
        }

        // Fetch recent messages descending (last messages first)
        const fetchRange = `${Math.max(1, total - limit + 1)}:*`;

        for await (const message of client.fetch(fetchRange, {
          envelope: true,
          flags: true,
          bodyStructure: true,
          internalDate: true,
          uid: true
        })) {
          const envelope = message.envelope || {};
          const flags = message.flags ? Array.from(message.flags) : [];
          const seen = flags.includes('\\Seen');
          const flagged = flags.includes('\\Flagged');

          let fromObj = { name: 'Desconhecido', address: '' };
          if (envelope.from && envelope.from[0]) {
            fromObj = {
              name: envelope.from[0].name || envelope.from[0].address,
              address: envelope.from[0].address
            };
          }

          let toList = [];
          if (envelope.to) {
            toList = envelope.to.map(t => ({ name: t.name || t.address, address: t.address }));
          }

          let hasAttachments = false;
          const checkAttachments = (structure) => {
            if (!structure) return;
            if (structure.disposition === 'attachment' || (structure.parameters && structure.parameters.name)) {
              hasAttachments = true;
            }
            if (structure.childNodes) {
              structure.childNodes.forEach(checkAttachments);
            }
          };
          checkAttachments(message.bodyStructure);

          messages.push({
            uid: message.uid,
            seq: message.seq,
            folder: folder,
            subject: envelope.subject || '(Sem assunto)',
            from: fromObj,
            to: toList,
            date: envelope.date || message.internalDate || new Date().toISOString(),
            seen,
            flagged,
            hasAttachments,
            snippet: envelope.subject || ''
          });
        }
      } finally {
        lock.release(); // Keep connection alive for other requests
      }

      // Sort descending by date
      messages.sort((a, b) => new Date(b.date) - new Date(a.date));

      let filtered = messages;
      if (filter === 'unread') filtered = filtered.filter(m => !m.seen);
      if (filter === 'flagged') filtered = filtered.filter(m => m.flagged);
      if (filter === 'attachments') filtered = filtered.filter(m => m.hasAttachments);
      if (search) {
        const q = search.toLowerCase();
        filtered = filtered.filter(m => 
          (m.subject && m.subject.toLowerCase().includes(q)) ||
          (m.from && (m.from.name.toLowerCase().includes(q) || m.from.address.toLowerCase().includes(q)))
        );
      }

      return {
        folder,
        total: filtered.length,
        messages: filtered
      };
    } catch (err) {
      console.warn(`[IMAP] Erro ao buscar emails da pasta "${folder}":`, err.message);
      return { folder, total: 0, messages: [], error: err.message };
    }
  }

  async getMessageDetail(folder = 'INBOX', uid) {
    const numUid = parseInt(uid, 10);

    let client = null;
    try {
      client = await this.getClient();
    } catch (err) {
      console.warn(`[IMAP getMessageDetail client error]:`, err.message);
    }

    if (!client) {
      if (!this.hasCredentials()) {
        const found = this.mockMessages.find(m => m.uid === numUid);
        if (found) found.seen = true;
        return found || null;
      }
      return null;
    }

    try {
      console.log(`[IMAP] Abrindo pasta "${folder}" para ler email UID: ${numUid}...`);
      const lock = await client.getMailboxLock(folder);
      let parsed = null;

      try {
        // Fast direct fetchOne with raw source
        const msgResult = await client.fetchOne(String(numUid), { source: true, flags: true, envelope: true }, { uid: true });
        if (msgResult && msgResult.source) {
          parsed = await simpleParser(msgResult.source);
        } else {
          // Fallback to download
          const dlResult = await client.download(String(numUid), undefined, { uid: true });
          if (dlResult && dlResult.content) {
            parsed = await simpleParser(dlResult.content);
          }
        }

        // Mark as seen non-blocking
        try {
          await client.messageFlagsAdd(String(numUid), ['\\Seen'], { uid: true });
        } catch (e) {}
      } finally {
        lock.release(); // Keep connection alive!
      }

      if (!parsed) {
        console.warn(`[IMAP] Não foi possível fazer parse do email UID ${numUid} na pasta ${folder}`);
        return null;
      }

      const attachments = (parsed.attachments || []).map((att, idx) => ({
        filename: att.filename || `anexo_${idx + 1}`,
        contentType: att.contentType || 'application/octet-stream',
        size: att.size || 0,
        partId: String(idx)
      }));

      console.log(`[IMAP] Email UID ${numUid} lido com sucesso! Assunto: "${parsed.subject}"`);

      return {
        uid: numUid,
        folder,
        subject: parsed.subject || '(Sem assunto)',
        from: parsed.from ? { name: parsed.from.value[0]?.name || parsed.from.value[0]?.address, address: parsed.from.value[0]?.address } : {},
        to: parsed.to ? parsed.to.value.map(t => ({ name: t.name || t.address, address: t.address })) : [],
        cc: parsed.cc ? parsed.cc.value.map(t => ({ name: t.name || t.address, address: t.address })) : [],
        date: parsed.date ? parsed.date.toISOString() : new Date().toISOString(),
        seen: true,
        html: parsed.html || (parsed.textAsHtml || `<pre style="white-space: pre-wrap; font-family: inherit;">${parsed.text || ''}</pre>`),
        text: parsed.text || '',
        hasAttachments: attachments.length > 0,
        attachments
      };
    } catch (err) {
      console.error(`[IMAP] Falha ao baixar detalhes do email UID ${numUid}:`, err.message);
      return null;
    }
  }

  async getAttachment(folder = 'INBOX', uid, partId) {
    const numUid = parseInt(uid, 10);
    const idx = parseInt(partId, 10);
    if (isNaN(numUid) || isNaN(idx)) return null;

    const client = await this.getClient();
    if (!client) return null;

    try {
      const lock = await client.getMailboxLock(folder);
      let parsed = null;
      try {
        const msgResult = await client.fetchOne(String(numUid), { source: true }, { uid: true });
        if (msgResult && msgResult.source) {
          parsed = await simpleParser(msgResult.source);
        } else {
          const dlResult = await client.download(String(numUid), undefined, { uid: true });
          if (dlResult && dlResult.content) {
            parsed = await simpleParser(dlResult.content);
          }
        }
      } finally {
        lock.release();
      }

      if (!parsed || !Array.isArray(parsed.attachments) || !parsed.attachments[idx]) {
        return null;
      }

      const att = parsed.attachments[idx];
      return {
        filename: att.filename || `anexo_${idx + 1}`,
        contentType: att.contentType || 'application/octet-stream',
        size: att.size || (att.content ? att.content.length : 0),
        content: att.content
      };
    } catch (err) {
      console.error(`[IMAP] Erro ao baixar anexo UID ${numUid} part ${idx}:`, err.message);
      return null;
    }
  }

  async setFlags(folder, uid, { seen, flagged }) {
    const client = await this.getClient();
    if (!client) return { success: true };

    try {
      const lock = await client.getMailboxLock(folder);
      try {
        if (seen === true) await client.messageFlagsAdd(String(uid), ['\\Seen'], { uid: true });
        if (seen === false) await client.messageFlagsRemove(String(uid), ['\\Seen'], { uid: true });
        if (flagged === true) await client.messageFlagsAdd(String(uid), ['\\Flagged'], { uid: true });
        if (flagged === false) await client.messageFlagsRemove(String(uid), ['\\Flagged'], { uid: true });
      } finally {
        lock.release();
      }
      return { success: true };
    } catch (err) {
      console.error('[IMAP setFlags Error]', err);
      return { success: false, error: err.message };
    }
  }

  async deleteMessage(folder, uid) {
    return this.deleteMessages(folder, [uid]);
  }

  async findTrashFolder(client) {
    try {
      const list = await client.list();
      const bySpecialUse = list.find(m => m.specialUse === '\\Trash');
      if (bySpecialUse) return bySpecialUse.path;
      const byName = list.find(m => /^(trash|deleted|lixeira|itens exclu)/i.test(m.name || m.path));
      if (byName) return byName.path;
    } catch (e) {}
    return 'Trash';
  }

  async deleteMessages(folder, uids) {
    const cleanUids = (Array.isArray(uids) ? uids : [uids]).filter(Boolean).map(String);
    if (cleanUids.length === 0) return { success: true };

    const client = await this.getClient();
    if (!client) return { success: true };

    try {
      const trashFolder = await this.findTrashFolder(client);
      const isAlreadyInTrash =
        folder.toLowerCase() === trashFolder.toLowerCase() ||
        /^(trash|deleted|lixeira|itens exclu)/i.test(folder);

      const lock = await client.getMailboxLock(folder);
      try {
        if (isAlreadyInTrash) {
          // Permanently delete only if already inside Trash / Itens Excluídos
          await client.messageDelete(cleanUids.join(','), { uid: true });
        } else {
          // Move to Trash / Itens Excluídos without permanent deletion
          await client.messageMove(cleanUids.join(','), trashFolder, { uid: true });
        }
      } finally {
        lock.release();
      }
      return {
        success: true,
        movedToTrash: !isAlreadyInTrash,
        trashFolder,
        count: cleanUids.length
      };
    } catch (err) {
      console.error('[IMAP deleteMessages Error]', err);
      return { success: false, error: err.message };
    }
  }

  async setFlagsBulk(folder, uids, { seen, flagged }) {
    const cleanUids = (Array.isArray(uids) ? uids : [uids]).filter(Boolean).map(String);
    if (cleanUids.length === 0) return { success: true };

    const client = await this.getClient();
    if (!client) return { success: true };

    try {
      const range = cleanUids.join(',');
      const lock = await client.getMailboxLock(folder);
      try {
        if (seen === true) await client.messageFlagsAdd(range, ['\\Seen'], { uid: true });
        if (seen === false) await client.messageFlagsRemove(range, ['\\Seen'], { uid: true });
        if (flagged === true) await client.messageFlagsAdd(range, ['\\Flagged'], { uid: true });
        if (flagged === false) await client.messageFlagsRemove(range, ['\\Flagged'], { uid: true });
      } finally {
        lock.release();
      }
      return { success: true, count: cleanUids.length };
    } catch (err) {
      console.error('[IMAP setFlagsBulk Error]', err);
      return { success: false, error: err.message };
    }
  }

  async moveMessage(fromFolder, toFolder, uid) {
    const client = await this.getClient();
    if (!client) return { success: true };

    try {
      const lock = await client.getMailboxLock(fromFolder);
      try {
        await client.messageMove(String(uid), toFolder, { uid: true });
      } finally {
        lock.release();
      }
      return { success: true };
    } catch (err) {
      console.error('[IMAP moveMessage Error]', err);
      return { success: false, error: err.message };
    }
  }
  async appendSentMessage({ to, cc, subject, text, html }) {
    const client = await this.getClient();
    if (!client) return;

    try {
      const user = this.config.imap?.user || '';
      const rfc822 = [
        `From: "Clodoaldo Silva" <${user}>`,
        `To: ${to}`,
        cc ? `Cc: ${cc}` : null,
        `Subject: ${subject || '(Sem assunto)'}`,
        `Date: ${new Date().toUTCString()}`,
        'MIME-Version: 1.0',
        'Content-Type: text/html; charset=utf-8',
        '',
        html || text || ''
      ].filter(line => line !== null).join('\r\n');

      const sentFolder = 'Sent';
      await client.append(sentFolder, Buffer.from(rfc822, 'utf8'), ['\\Seen']);
      console.log(`[IMAP] Email enviado copiado para a pasta "${sentFolder}".`);
    } catch (err) {
      console.warn('[IMAP appendSentMessage]:', err.message);
    }
  }

  async extractContactsFromServer(limitPerFolder = 100) {
    const client = await this.getClient();
    if (!client) return [];

    const ownEmail = (this.config.imap?.user || '').trim().toLowerCase();
    const map = new Map();

    const addAddress = (addrObj) => {
      if (!addrObj || !addrObj.address) return;
      const email = String(addrObj.address).trim().toLowerCase();
      if (!email || !email.includes('@') || email === ownEmail) return;
      const name = (addrObj.name || '').trim();
      const existing = map.get(email);
      if (!existing || (!existing.name && name)) {
        map.set(email, {
          email,
          name: name || existing?.name || email.split('@')[0],
          phone: existing?.phone || '',
          company: existing?.company || '',
          source: 'servidor'
        });
      }
    };

    const foldersToScan = ['INBOX', 'Sent'];
    for (const folder of foldersToScan) {
      try {
        const lock = await client.getMailboxLock(folder);
        try {
          const total = client.mailbox?.exists || 0;
          if (total > 0) {
            const range = `${Math.max(1, total - limitPerFolder + 1)}:*`;
            for await (const msg of client.fetch(range, { envelope: true })) {
              const env = msg.envelope || {};
              (env.from || []).forEach(addAddress);
              (env.to || []).forEach(addAddress);
              (env.cc || []).forEach(addAddress);
            }
          }
        } finally {
          lock.release();
        }
      } catch (err) {
        console.warn(`[IMAP Contacts] Aviso ao escanear pasta "${folder}":`, err.message);
      }
    }

    return Array.from(map.values()).sort((a, b) => (a.name || a.email).localeCompare(b.name || b.email, 'pt-BR'));
  }
}

module.exports = ImapService;
