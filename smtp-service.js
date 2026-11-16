const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const dns = require('dns');

const CONFIG_FILE = path.join(__dirname, 'config.json');

// Known resolved IP for h58.servidorhh.com to bypass local DNS latency
const KNOWN_HOST_IPS = {
  'h58.servidorhh.com': '177.234.150.58'
};

class SmtpService {
  constructor(config) {
    this.config = config || {};
  }

  updateConfig(config) {
    this.config = { ...this.config, ...config };
  }

  persistPort(port) {
    try {
      this.config.smtp = this.config.smtp || {};
      this.config.smtp.port = String(port);
      this.config.smtp.secure = (port === 465);

      if (fs.existsSync(CONFIG_FILE)) {
        const current = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8'));
        current.smtp = current.smtp || {};
        current.smtp.port = String(port);
        current.smtp.secure = (port === 465);
        fs.writeFileSync(CONFIG_FILE, JSON.stringify(current, null, 2), 'utf8');
        console.log(`[SMTP] Porta ${port} persistida com sucesso em config.json`);
      }
    } catch (e) {
      console.warn('[SMTP] Aviso ao salvar porta:', e.message);
    }
  }

  createTransporter(port = 465) {
    if (!this.config.smtp || !this.config.smtp.host || !this.config.smtp.user || !this.config.smtp.password) {
      return null;
    }

    const host = this.config.smtp.host.trim();
    const isSecure = port === 465;

    return nodemailer.createTransport({
      host: host,
      port: port,
      secure: isSecure,
      auth: {
        user: this.config.smtp.user.trim(),
        pass: this.config.smtp.password
      },
      tls: {
        servername: host,
        rejectUnauthorized: false
      },
      lookup: (hostname, options, callback) => {
        if (KNOWN_HOST_IPS[hostname]) {
          return callback(null, KNOWN_HOST_IPS[hostname], 4);
        }
        dns.lookup(hostname, options, callback);
      },
      connectionTimeout: 20000,
      greetingTimeout: 20000,
      socketTimeout: 30000
    });
  }

  formatError(err, port) {
    if (!err) return 'Erro desconhecido ao enviar email via SMTP.';

    if (err.code === 'EAUTH' || /authentication|invalid login|535/i.test(err.message || '')) {
      return 'Falha de autenticação SMTP: Usuário ou senha rejeitados pelo servidor de email. Verifique a senha configurada.';
    }

    if (err.code === 'ETIMEDOUT' || /timeout/i.test(err.message || '')) {
      return `O servidor SMTP (${this.config.smtp?.host || 'h58.servidorhh.com'}) demorou para responder na porta ${port}. Tentamos automaticamente as portas 465 e 587.`;
    }

    if (err.code === 'ECONNREFUSED') {
      return `Conexão recusada no servidor SMTP na porta ${port}.`;
    }

    if (err.response) {
      return `Servidor SMTP: ${err.response}`;
    }

    return err.message || 'Falha ao enviar e-mail via SMTP.';
  }

  async testConnection(customConfig) {
    const activeConfig = customConfig || this.config;
    if (!activeConfig.smtp || !activeConfig.smtp.host || !activeConfig.smtp.user || !activeConfig.smtp.password) {
      return { 
        success: false, 
        message: 'Preencha o Host, Usuário e Senha antes de testar a saída SMTP.' 
      };
    }

    const initialPort = parseInt(activeConfig.smtp.port, 10) || 465;
    const portsToTry = initialPort === 465 ? [465, 587] : [587, 465];

    let lastError = null;
    for (const port of portsToTry) {
      console.log(`[SMTP Test] Testando conexão na porta ${port}...`);
      const transporter = nodemailer.createTransport({
        host: activeConfig.smtp.host.trim(),
        port: port,
        secure: port === 465,
        auth: {
          user: activeConfig.smtp.user.trim(),
          pass: activeConfig.smtp.password
        },
        tls: {
          servername: activeConfig.smtp.host.trim(),
          rejectUnauthorized: false
        },
        lookup: (hostname, options, callback) => {
          if (KNOWN_HOST_IPS[hostname]) return callback(null, KNOWN_HOST_IPS[hostname], 4);
          dns.lookup(hostname, options, callback);
        },
        connectionTimeout: 12000,
        greetingTimeout: 12000,
        socketTimeout: 15000
      });

      try {
        await transporter.verify();
        this.persistPort(port);
        return {
          success: true,
          suggestedPort: port,
          message: `Servidor SMTP (${activeConfig.smtp.host}:${port}) conectado e autenticado com sucesso!`
        };
      } catch (err) {
        lastError = err;
        console.warn(`[SMTP Test] Porta ${port} falhou:`, err.message);
        if (err.code === 'EAUTH' || /authentication|invalid login|535/i.test(err.message)) {
          return { success: false, message: this.formatError(err, port) };
        }
      }
    }

    return {
      success: false,
      message: this.formatError(lastError, initialPort)
    };
  }

  async sendMail({ to, cc, bcc, subject, text, html, attachments = [] }) {
    if (!this.config.smtp || !this.config.smtp.host || !this.config.smtp.user || !this.config.smtp.password) {
      console.log('[SMTP Demo] Modo demonstração ativo.');
      return {
        success: true,
        messageId: 'mock-' + Date.now(),
        isDemo: true,
        message: 'Email enviado com sucesso (Modo Demonstração).'
      };
    }

    const initialPort = parseInt(this.config.smtp?.port, 10) || 465;
    const portsToTry = initialPort === 465 ? [465, 587] : [587, 465];

    const mailOptions = {
      from: `"${this.config.smtp?.user?.split('@')[0] || 'Clodoaldo Silva'}" <${this.config.smtp?.user}>`,
      to,
      cc,
      bcc,
      subject,
      text,
      html,
      attachments: attachments.map(att => ({
        filename: att.originalname || att.filename,
        content: att.buffer,
        contentType: att.mimetype
      }))
    };

    let lastError = null;

    // Try primary port (with 2 attempts), then fallback to alternate port
    for (const port of portsToTry) {
      for (let attempt = 1; attempt <= 2; attempt++) {
        console.log(`[SMTP] Tentativa ${attempt} de envio via porta ${port}...`);
        const transporter = this.createTransporter(port);

        try {
          const info = await transporter.sendMail(mailOptions);
          console.log(`[SMTP] Email enviado com sucesso via porta ${port}! ID: ${info.messageId}`);
          this.persistPort(port);
          return {
            success: true,
            messageId: info.messageId,
            isDemo: false,
            message: `Email enviado com sucesso via SMTP (porta ${port})!`
          };
        } catch (err) {
          lastError = err;
          console.warn(`[SMTP] Tentativa ${attempt} na porta ${port} falhou: ${err.message}`);

          // If authentication failed, do not retry because the credentials are wrong
          if (err.code === 'EAUTH' || /authentication|invalid login|535/i.test(err.message)) {
            return {
              success: false,
              error: err.message,
              message: this.formatError(err, port)
            };
          }

          // Small delay before retry
          await new Promise(r => setTimeout(r, 600));
        }
      }
    }

    console.error('[SMTP Send Final Error]', lastError);
    return {
      success: false,
      error: lastError?.message || 'Falha de conexão',
      message: this.formatError(lastError, initialPort)
    };
  }
}

module.exports = SmtpService;
