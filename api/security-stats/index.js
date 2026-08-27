// GET /api/security-stats
//
// Feeds the live security panel on the STRATOS homepage. Returns a small
// aggregated JSON — never per-client or per-device detail (see the
// technical integration doc, section 9, for the "what to publish" rule).
//
// v1 scope: reads Microsoft Graph for STRATOS's OWN tenant only, using the
// "Stratos Security Panel" app registration (single-tenant, client
// credentials flow). Multi-client aggregation via GDAP is a follow-up —
// see the doc's GDAP section for what that upgrade needs.
//
// Required Application Settings (Azure Portal -> Static Web App ->
// Configuration -> Application settings), NOT committed to git:
//   GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET
// Optional:
//   CACHE_TTL_SECONDS (default 1800 = 30 min)

const { ClientSecretCredential } = require("@azure/identity");

const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPE = "https://graph.microsoft.com/.default";

let cache = { data: null, expiresAt: 0 };

// A fresh credential per call means a fresh token per call — no risk of
// MSAL's internal cache handing back a token that was issued before a
// permission change propagated (that token simply won't carry the new
// scopes, no matter how long we wait, since scopes are baked in at
// issuance). We already rate-limit ourselves via `cache` above, so the
// extra token request per cache-miss is cheap.
function getCredential() {
  return new ClientSecretCredential(
    process.env.GRAPH_TENANT_ID,
    process.env.GRAPH_CLIENT_ID,
    process.env.GRAPH_CLIENT_SECRET
  );
}

async function getToken(context) {
  const token = await getCredential().getToken(SCOPE);
  return token.token;
}

// Follows @odata.nextLink until done or `cap` pages are read (safety valve
// so a bug or a huge tenant can't turn one request into an unbounded loop).
async function graphGetAll(url, token, cap = 25) {
  let results = [];
  let next = url;
  let pages = 0;
  while (next && pages < cap) {
    const res = await fetch(next, {
      headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
    });
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`Graph GET ${next} -> ${res.status} ${body.slice(0, 300)}`);
    }
    const json = await res.json();
    results = results.concat(json.value || []);
    next = json["@odata.nextLink"] || null;
    pages += 1;
  }
  return results;
}

async function graphGetOne(url, token) {
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Graph GET ${url} -> ${res.status} ${body.slice(0, 300)}`);
  }
  return res.json();
}

function round1(n) {
  return Math.round(n * 10) / 10;
}

async function computeMetrics(context, token, debugErrors) {
  const metrics = {};

  // 1. Endpoints protegidos — Intune managed devices
  try {
    const devices = await graphGetAll(
      `${GRAPH}/deviceManagement/managedDevices?$select=id`,
      token
    );
    metrics.endpoints_protected = {
      value: devices.length,
      note: "Dispositivos Intune",
    };
  } catch (err) {
    context.log.warn("endpoints_protected failed:", err.message);
    if (debugErrors) debugErrors.endpoints_protected = err.message;
  }

  // 2. Cumplimiento de políticas — Intune compliance state summary
  try {
    const summary = await graphGetOne(
      `${GRAPH}/deviceManagement/deviceCompliancePolicyDeviceStateSummary`,
      token
    );
    const compliant = summary.compliantDeviceCount || 0;
    const nonCompliant = summary.nonCompliantDeviceCount || 0;
    const known = compliant + nonCompliant;
    if (known > 0) {
      metrics.compliance_rate = {
        value: round1((compliant / known) * 100),
        suffix: "%",
        note: "Dispositivos Intune",
      };
    }
  } catch (err) {
    context.log.warn("compliance_rate failed:", err.message);
    if (debugErrors) debugErrors.compliance_rate = err.message;
  }

  // 3. Cuentas con MFA activo — Entra ID authentication methods report
  try {
    const regs = await graphGetAll(
      `${GRAPH}/reports/authenticationMethods/userRegistrationDetails?$select=isMfaRegistered`,
      token
    );
    if (regs.length > 0) {
      const mfaCount = regs.filter((u) => u.isMfaRegistered).length;
      metrics.mfa_coverage = {
        value: round1((mfaCount / regs.length) * 100),
        suffix: "%",
        note: "Microsoft Entra ID",
      };
    }
  } catch (err) {
    context.log.warn("mfa_coverage failed:", err.message);
    if (debugErrors) debugErrors.mfa_coverage = err.message;
  }

  // 4. Amenazas bloqueadas (30 días) — Defender alerts via Graph Security API
  try {
    const since = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
    const alerts = await graphGetAll(
      `${GRAPH}/security/alerts_v2?$filter=createdDateTime ge ${since}&$select=id`,
      token
    );
    metrics.threats_blocked_30d = {
      value: alerts.length,
      note: "Últimos 30 días",
    };
  } catch (err) {
    context.log.warn("threats_blocked_30d failed:", err.message);
    if (debugErrors) debugErrors.threats_blocked_30d = err.message;
  }

  return metrics;
}

module.exports = async function (context, req) {
  const ttlMs = (parseInt(process.env.CACHE_TTL_SECONDS, 10) || 1800) * 1000;
  const now = Date.now();
  // TEMPORARY debug aid: ?debug=1 bypasses the cache and includes the raw
  // per-metric error messages in the response, so failures can be diagnosed
  // without needing Application Insights. Remove this once everything is
  // green — it's harmless (no secrets, no per-client data) but noisy.
  const debug = req.query && req.query.debug === "1";

  if (!debug && cache.data && cache.expiresAt > now) {
    context.res = { status: 200, headers: { "Content-Type": "application/json" }, body: cache.data };
    return;
  }

  const debugErrors = debug ? {} : null;

  try {
    const token = await getToken(context);
    const metrics = await computeMetrics(context, token, debugErrors);

    if (Object.keys(metrics).length === 0 && !debug) {
      throw new Error("No metric succeeded — check app permissions/consent");
    }

    const payload = { updated_at: new Date().toISOString(), metrics };
    if (debug) payload._debug_errors = debugErrors;
    if (!debug) cache = { data: payload, expiresAt: now + ttlMs };

    context.res = { status: 200, headers: { "Content-Type": "application/json" }, body: payload };
  } catch (err) {
    context.log.error("security-stats failed:", err.message);
    // Fail closed but harmless: the frontend already falls back to sample
    // numbers when this endpoint doesn't return 200 (see template-v2.html).
    if (cache.data) {
      context.res = { status: 200, headers: { "Content-Type": "application/json" }, body: cache.data };
    } else {
      context.res = { status: 503, headers: { "Content-Type": "application/json" }, body: { error: "temporarily_unavailable" } };
    }
  }
};
