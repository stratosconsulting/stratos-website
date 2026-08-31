// GET /api/daily-visits-report?secret=... — NOT a timer trigger. Static
// Web Apps "Managed Functions" only support httpTrigger (confirmed the
// hard way: a timerTrigger version of this failed the whole deploy with
// "Currently, only httpTriggers are supported"). Instead this is a plain
// HTTP endpoint, called once a day by a GitHub Actions cron workflow (see
// .github/workflows/daily-visits-cron.yml) — same effect, no separate
// Azure Function App needed.
//
// Protected by a shared secret (CRON_SECRET) so a random visitor hitting
// this URL can't spam Miguel's inbox or rack up Application Insights query
// calls — GitHub Actions passes it as a query string param, sourced from
// a repo secret of the same name.
//
// Queries Application Insights (already wired to the Static Web App) for
// the last 24 hours of pageViews telemetry and emails Miguel a short daily
// summary via Microsoft Graph — same sendMail pattern as
// api/contact-submit, same app registration and Mail.Send permission.
//
// Required Application Settings, NOT committed to git:
//   GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET  (shared w/ other functions)
//   MAIL_SENDER_UPN     — shared w/ contact-submit
//   APPINSIGHTS_APP_ID  — Application Insights resource -> API Access -> Application ID
//   APPINSIGHTS_API_KEY — same screen -> Create API Key, "Read telemetry" permission
//   CRON_SECRET         — any random string; must match the GitHub repo secret
//                         of the same name (see the workflow file).
// Optional:
//   MAIL_TO_ANALYTICS   — who gets the daily report. Falls back to MAIL_TO_CONSTANCIA
//                         if unset, so no extra config is needed to get started.
//
// If Application Insights isn't connected yet (APPINSIGHTS_APP_ID/KEY unset),
// this logs a warning and exits quietly instead of emailing an empty/broken
// report every morning — nothing to see until it's actually wired up.

const { ClientSecretCredential } = require("@azure/identity");

const GRAPH = "https://graph.microsoft.com/v1.0";
const GRAPH_SCOPE = "https://graph.microsoft.com/.default";
const AI_QUERY_BASE = "https://api.applicationinsights.io/v1/apps";

function escapeHtml(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

async function runQuery(appId, apiKey, kql) {
  const url = `${AI_QUERY_BASE}/${appId}/query?query=${encodeURIComponent(kql)}`;
  const res = await fetch(url, { headers: { "x-api-key": apiKey } });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`App Insights query -> ${res.status} ${body.slice(0, 300)}`);
  }
  const json = await res.json();
  // Table 0's rows, as an array of column->value objects.
  const table = json.tables && json.tables[0];
  if (!table) return [];
  return table.rows.map((row) => {
    const obj = {};
    table.columns.forEach((col, i) => { obj[col.name] = row[i]; });
    return obj;
  });
}

module.exports = async function (context, req) {
  const expectedSecret = process.env.CRON_SECRET;
  const providedSecret = (req.query && req.query.secret) || req.headers["x-cron-secret"];
  if (!expectedSecret || providedSecret !== expectedSecret) {
    context.res = { status: 401, body: "unauthorized" };
    return;
  }

  const appId = process.env.APPINSIGHTS_APP_ID;
  const apiKey = process.env.APPINSIGHTS_API_KEY;

  if (!appId || !apiKey) {
    context.log.warn("daily-visits-report: APPINSIGHTS_APP_ID/APPINSIGHTS_API_KEY not set yet — skipping (connect Application Insights to the Static Web App first).");
    context.res = { status: 200, body: "skipped: Application Insights not configured yet" };
    return;
  }

  const senderUpn = process.env.MAIL_SENDER_UPN;
  const recipient = process.env.MAIL_TO_ANALYTICS || process.env.MAIL_TO_CONSTANCIA;
  if (!senderUpn || !recipient) {
    context.log.error("daily-visits-report: missing MAIL_SENDER_UPN or a recipient (MAIL_TO_ANALYTICS / MAIL_TO_CONSTANCIA)");
    context.res = { status: 500, body: "missing mail configuration" };
    return;
  }

  let totalsRow, topPages, topCountries;
  try {
    [totalsRow, topPages, topCountries] = await Promise.all([
      runQuery(appId, apiKey, `pageViews | where timestamp > ago(24h) | summarize Visits=count(), Visitantes=dcount(user_Id)`),
      runQuery(appId, apiKey, `pageViews | where timestamp > ago(24h) | summarize Visits=count() by name | order by Visits desc | take 8`),
      runQuery(appId, apiKey, `pageViews | where timestamp > ago(24h) and isnotempty(client_CountryOrRegion) | summarize Visits=count() by client_CountryOrRegion | order by Visits desc | take 5`),
    ]);
  } catch (err) {
    context.log.error("daily-visits-report: Application Insights query failed:", err.message);
    context.res = { status: 502, body: "Application Insights query failed: " + err.message };
    return;
  }

  const totals = (totalsRow && totalsRow[0]) || { Visits: 0, Visitantes: 0 };

  const pagesHtml = topPages.length
    ? topPages.map((r) => `<tr><td style="padding:3px 16px 3px 0;">${escapeHtml(r.name || "(sin nombre)")}</td><td style="padding:3px 0;text-align:right;">${r.Visits}</td></tr>`).join("\n")
    : `<tr><td colspan="2" style="color:#888;">Sin visitas registradas en las últimas 24 horas.</td></tr>`;

  const countriesHtml = topCountries.length
    ? topCountries.map((r) => `<tr><td style="padding:3px 16px 3px 0;">${escapeHtml(r.client_CountryOrRegion)}</td><td style="padding:3px 0;text-align:right;">${r.Visits}</td></tr>`).join("\n")
    : `<tr><td colspan="2" style="color:#888;">Sin datos de país.</td></tr>`;

  const today = new Date().toLocaleDateString("es-PR", { year: "numeric", month: "long", day: "numeric" });

  const htmlBody = `<div style="font-family:sans-serif;font-size:14px;color:#222;max-width:520px;">
<p><strong>Reporte diario de visitas — stratosconsultingpr.com</strong><br>
<span style="color:#888;font-size:12px;">${escapeHtml(today)} · últimas 24 horas</span></p>
<p style="font-size:22px;margin:12px 0 4px;">${totals.Visits} visitas <span style="font-size:14px;color:#888;font-weight:normal;">(${totals.Visitantes} visitantes únicos)</span></p>
<p style="margin:20px 0 6px;"><strong>Páginas más visitadas</strong></p>
<table style="border-collapse:collapse;width:100%;max-width:400px;">${pagesHtml}</table>
<p style="margin:20px 0 6px;"><strong>Por país</strong></p>
<table style="border-collapse:collapse;width:100%;max-width:400px;">${countriesHtml}</table>
<p style="color:#888;font-size:12px;margin-top:20px;">Datos de Application Insights. Este correo se genera automáticamente cada mañana — no hace falta responder.</p>
</div>`;

  const message = {
    message: {
      subject: `Visitas de ayer en stratosconsultingpr.com: ${totals.Visits}`,
      body: { contentType: "HTML", content: htmlBody },
      toRecipients: [{ emailAddress: { address: recipient } }],
    },
    saveToSentItems: false,
  };

  try {
    const credential = new ClientSecretCredential(
      process.env.GRAPH_TENANT_ID,
      process.env.GRAPH_CLIENT_ID,
      process.env.GRAPH_CLIENT_SECRET
    );
    const token = await credential.getToken(GRAPH_SCOPE);
    const res = await fetch(`${GRAPH}/users/${encodeURIComponent(senderUpn)}/sendMail`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token.token}`, "Content-Type": "application/json" },
      body: JSON.stringify(message),
    });
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      context.log.error(`daily-visits-report: Graph sendMail -> ${res.status} ${body.slice(0, 300)}`);
      context.res = { status: 502, body: "Graph sendMail failed: " + res.status };
      return;
    }
    context.log(`daily-visits-report: sent (${totals.Visits} visits, ${totals.Visitantes} unique)`);
    context.res = { status: 200, body: `sent: ${totals.Visits} visits, ${totals.Visitantes} unique visitors` };
  } catch (err) {
    context.log.error("daily-visits-report: unexpected error sending mail:", err.message);
    context.res = { status: 500, body: "unexpected error: " + err.message };
  }
};
