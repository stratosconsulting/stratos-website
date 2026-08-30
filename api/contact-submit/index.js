// POST /api/contact-submit
//
// Backend for the site's contact forms (contacto.html, evaluacion-gratuita.html,
// and any future form.stratos-form instance). Until now those forms only
// showed a fake "success" message client-side and never sent anything —
// this function is what actually delivers the submission.
//
// It sends ONE email via Microsoft Graph (app-only client-credentials,
// same app registration as /api/security-stats — needs the additional
// application permission Mail.Send, admin-consented) to up to two
// recipients:
//   MAIL_TO_CONSTANCIA — Miguel's own inbox, for record-keeping.
//   MAIL_TO_CRM        — Odoo CRM Sales Team's lead-catching email alias.
//                         Any mail sent here becomes a CRM lead automatically
//                         (Odoo's built-in "email to lead" feature). Leave
//                         unset until the alias is confirmed/created in Odoo
//                         — the function still sends to MAIL_TO_CONSTANCIA.
// The submitter's own address is set as replyTo, so replying to the email
// goes straight to them, not through Odoo or STRATOS's own mailbox.
//
// Required Application Settings (Azure Portal -> Static Web App ->
// Configuration -> Application settings), NOT committed to git:
//   GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET  (shared w/ security-stats)
//   MAIL_SENDER_UPN     — mailbox the app sends "as", e.g. no-reply@stratosconsultingpr.com
//                         or info@stratosconsultingpr.com. Must be a real, licensed mailbox.
//   MAIL_TO_CONSTANCIA  — Miguel's inbox address.
// Optional:
//   MAIL_TO_CRM         — Odoo CRM team alias. Omitted = CRM copy simply isn't sent yet.

const { ClientSecretCredential } = require("@azure/identity");

const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPE = "https://graph.microsoft.com/.default";

function getCredential() {
  return new ClientSecretCredential(
    process.env.GRAPH_TENANT_ID,
    process.env.GRAPH_CLIENT_ID,
    process.env.GRAPH_CLIENT_SECRET
  );
}

function escapeHtml(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function isEmail(s) {
  return typeof s === "string" && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
}

// Field maps per form type — keeps the function honest about exactly what
// each form on the site actually collects (see contacto.html / evaluacion-gratuita.html).
const FORM_LABELS = {
  contacto: {
    title: "Nueva solicitud de consulta — formulario de Contacto",
    fields: [
      ["nombre", "Nombre completo"],
      ["empresa", "Empresa"],
      ["correo", "Correo electrónico"],
      ["telefono", "Teléfono"],
      ["servicio", "Servicio de interés"],
      ["mensaje", "Mensaje"],
      ["pref", "Método de contacto preferido"],
    ],
  },
  evaluacion: {
    title: "Nueva solicitud de evaluación gratuita",
    fields: [
      ["nombre", "Nombre completo"],
      ["empresa", "Empresa"],
      ["correo", "Correo electrónico"],
      ["telefono", "Teléfono"],
    ],
  },
};

module.exports = async function (context, req) {
  const body = req.body || {};
  const formType = typeof body.formType === "string" ? body.formType : "contacto";
  const spec = FORM_LABELS[formType] || FORM_LABELS.contacto;

  const nombre = (body.nombre || "").toString().trim();
  const correo = (body.correo || "").toString().trim();

  if (!nombre || !isEmail(correo)) {
    context.res = {
      status: 400,
      headers: { "Content-Type": "application/json" },
      body: { ok: false, error: "Falta nombre o el correo electrónico no es válido." },
    };
    return;
  }

  const senderUpn = process.env.MAIL_SENDER_UPN;
  const toConstancia = process.env.MAIL_TO_CONSTANCIA;
  const toCrm = process.env.MAIL_TO_CRM;

  if (!senderUpn || !toConstancia) {
    context.log.error("contact-submit: missing MAIL_SENDER_UPN or MAIL_TO_CONSTANCIA app setting");
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: { ok: false, error: "El formulario no está disponible en este momento. Por favor llámenos o escríbanos directamente." },
    };
    return;
  }

  const recipients = [toConstancia, toCrm].filter(isEmail).map((a) => ({ emailAddress: { address: a } }));

  const rows = spec.fields
    .map(([key, label]) => {
      const val = (body[key] || "").toString().trim();
      if (!val) return "";
      return `<tr><td style="padding:4px 12px 4px 0;color:#666;white-space:nowrap;vertical-align:top;">${escapeHtml(label)}</td><td style="padding:4px 0;">${escapeHtml(val).replace(/\n/g, "<br>")}</td></tr>`;
    })
    .join("\n");

  const htmlBody = `<div style="font-family:sans-serif;font-size:14px;color:#222;">
<p><strong>${escapeHtml(spec.title)}</strong></p>
<table>${rows}</table>
<p style="color:#888;font-size:12px;margin-top:16px;">Enviado desde stratosconsultingpr.com — responder a este correo llega directamente a ${escapeHtml(correo)}.</p>
</div>`;

  const message = {
    message: {
      subject: `${spec.title} — ${nombre}`,
      body: { contentType: "HTML", content: htmlBody },
      toRecipients: recipients,
      replyTo: [{ emailAddress: { address: correo, name: nombre } }],
    },
    saveToSentItems: true,
  };

  try {
    const credential = getCredential();
    const token = await credential.getToken(SCOPE);
    const res = await fetch(`${GRAPH}/users/${encodeURIComponent(senderUpn)}/sendMail`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token.token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(message),
    });

    if (!res.ok) {
      const errBody = await res.text().catch(() => "");
      context.log.error(`contact-submit: Graph sendMail -> ${res.status} ${errBody.slice(0, 500)}`);
      context.res = {
        status: 502,
        headers: { "Content-Type": "application/json" },
        body: { ok: false, error: "No pudimos enviar el formulario. Intente de nuevo o escríbanos directamente." },
      };
      return;
    }

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: { ok: true },
    };
  } catch (err) {
    context.log.error("contact-submit: unexpected error", err);
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: { ok: false, error: "No pudimos enviar el formulario. Intente de nuevo o escríbanos directamente." },
    };
  }
};
