// GET /api/security-stats
//
// Feeds the live security panel on the STRATOS homepage. Returns a small
// aggregated JSON — never per-client or per-device detail (see the
// technical integration doc, section 9, for the "what to publish" rule).
//
// v2 scope: aggregates STRATOS's own tenant PLUS every client tenant that
// has completed the GDAP + app-consent setup (scripts/obo-delegated-setup.ps1).
// Own tenant uses application-only client-credentials (unchanged from v1).
// Client tenants use a delegated OBO token exchanged per-tenant — GDAP does
// not accept application-only tokens against a customer tenant. See
// delegatedAuth.js and tokenStore.js.
//
// Required Application Settings (Azure Portal -> Static Web App ->
// Configuration -> Application settings), NOT committed to git:
//   GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET
//   GRAPH_CLIENT_TENANT_IDS   — JSON array of client tenant IDs, e.g.
//                               ["5129a3d7-...", "5f5d7739-..."]
//                               (copy straight from scripts/client-tenant-ids.json)
//   GRAPH_OBO_REFRESH_TOKEN   — one-time seed, see scripts/obo-delegated-setup.ps1.
//                               Only read once; rotation lives in Table Storage
//                               after that.
// Optional:
//   CACHE_TTL_SECONDS (default 1800 = 30 min)

const { ClientSecretCredential } = require("@azure/identity");
const { getGraphTokenForTenant } = require("./delegatedAuth");

const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPE = "https://graph.microsoft.com/.default";

let cache = { data: null, expiresAt: 0 };

// A fresh credential per call means a fresh token per call — no risk of
// MSAL's internal cache handing back a token that was issued before a
// permission change propagated (that token simply won't carry the new
// scopes, no matter how long we wait, since scopes are baked in at
// issuance). We already rate-limit ourselves via `cache` above, so the
// extra token request per cache-miss is cheap.
function getOwnTenantCredential() {
  return new ClientSecretCredential(
    process.env.GRAPH_TENANT_ID,
    process.env.GRAPH_CLIENT_ID,
    process.env.GRAPH_CLIENT_SECRET
  );
}

function getClientTenantIds() {
  const raw = process.env.GRAPH_CLIENT_TENANT_IDS;
  if (!raw) return [];
  try {
    const list = JSON.parse(raw);
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
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

// Raw counts for ONE tenant. Nulls/omissions mean that metric couldn't be
// read for this tenant (missing consent, empty data, transient error) — the
// aggregator below just skips it for that tenant rather than failing.
async function computeTenantRaw(context, token, label) {
  const raw = {};

  try {
    const devices = await graphGetAll(`${GRAPH}/deviceManagement/managedDevices?$select=id`, token);
    raw.endpoints = devices.length;
  } catch (err) {
    context.log.warn(`[${label}] endpoints failed:`, err.message);
  }

  try {
    const summary = await graphGetOne(`${GRAPH}/deviceManagement/deviceCompliancePolicyDeviceStateSummary`, token);
    const compliant = summary.compliantDeviceCount || 0;
    const nonCompliant = summary.nonCompliantDeviceCount || 0;
    if (compliant + nonCompliant > 0) {
      raw.compliantDevices = compliant;
      raw.knownDevices = compliant + nonCompliant;
    }
  } catch (err) {
    context.log.warn(`[${label}] compliance failed:`, err.message);
  }

  try {
    const regs = await graphGetAll(
      `${GRAPH}/reports/authenticationMethods/userRegistrationDetails?$select=isMfaRegistered`,
      token
    );
    if (regs.length > 0) {
      raw.mfaUsers = regs.filter((u) => u.isMfaRegistered).length;
      raw.totalUsers = regs.length;
    }
  } catch (err) {
    context.log.warn(`[${label}] mfa failed:`, err.message);
  }

  try {
    const since = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
    const alerts = await graphGetAll(
      `${GRAPH}/security/alerts_v2?$filter=createdDateTime ge ${since}&$select=id`,
      token
    );
    raw.threats30d = alerts.length;
  } catch (err) {
    context.log.warn(`[${label}] threats failed:`, err.message);
  }

  return raw;
}

function aggregate(perTenant) {
  const metrics = {};

  const endpointsTotal = perTenant.reduce((sum, t) => sum + (t.endpoints || 0), 0);
  if (perTenant.some((t) => typeof t.endpoints === "number")) {
    metrics.endpoints_protected = { value: endpointsTotal, note: "Dispositivos Intune" };
  }

  const compliantTotal = perTenant.reduce((sum, t) => sum + (t.compliantDevices || 0), 0);
  const knownTotal = perTenant.reduce((sum, t) => sum + (t.knownDevices || 0), 0);
  if (knownTotal > 0) {
    metrics.compliance_rate = {
      value: round1((compliantTotal / knownTotal) * 100),
      suffix: "%",
      note: "Dispositivos Intune",
    };
  }

  const mfaTotal = perTenant.reduce((sum, t) => sum + (t.mfaUsers || 0), 0);
  const usersTotal = perTenant.reduce((sum, t) => sum + (t.totalUsers || 0), 0);
  if (usersTotal > 0) {
    metrics.mfa_coverage = {
      value: round1((mfaTotal / usersTotal) * 100),
      suffix: "%",
      note: "Microsoft Entra ID",
    };
  }

  const threatsTotal = perTenant.reduce((sum, t) => sum + (t.threats30d || 0), 0);
  if (perTenant.some((t) => typeof t.threats30d === "number")) {
    metrics.threats_blocked_30d = { value: threatsTotal, note: "Últimos 30 días" };
  }

  return metrics;
}

async function computeMetrics(context, debugTenants) {
  const perTenant = [];

  // Own tenant — application-only, exactly as in v1.
  try {
    const ownToken = (await getOwnTenantCredential().getToken(SCOPE)).token;
    const raw = await computeTenantRaw(context, ownToken, "stratos");
    perTenant.push(raw);
    if (debugTenants) debugTenants.push({ tenant: "stratos (own)", ...raw });
  } catch (err) {
    context.log.error("own tenant failed entirely:", err.message);
    if (debugTenants) debugTenants.push({ tenant: "stratos (own)", error: String(err.message).slice(0, 200) });
  }

  // Client tenants — delegated OBO token, one exchange per tenant.
  const clientTenantIds = getClientTenantIds();
  for (const tenantId of clientTenantIds) {
    try {
      const token = await getGraphTokenForTenant(context, tenantId);
      const raw = await computeTenantRaw(context, token, tenantId);
      perTenant.push(raw);
      if (debugTenants) debugTenants.push({ tenant: tenantId, ...raw });
    } catch (err) {
      // One client without consent/valid GDAP role shouldn't take down the
      // rest of the panel — skip it and keep going.
      context.log.warn(`client tenant ${tenantId} skipped:`, err.message);
      if (debugTenants) debugTenants.push({ tenant: tenantId, error: String(err.message).slice(0, 200) });
    }
  }

  return { metrics: aggregate(perTenant), tenantsRead: perTenant.length };
}

module.exports = async function (context, req) {
  const ttlMs = (parseInt(process.env.CACHE_TTL_SECONDS, 10) || 1800) * 1000;
  const now = Date.now();
  const debug = req && req.query && req.query.debug === "1";

  if (!debug && cache.data && cache.expiresAt > now) {
    context.res = { status: 200, headers: { "Content-Type": "application/json" }, body: cache.data };
    return;
  }

  try {
    const debugTenants = debug ? [] : null;
    const { metrics, tenantsRead } = await computeMetrics(context, debugTenants);

    if (Object.keys(metrics).length === 0) {
      throw new Error("No metric succeeded for any tenant — check app permissions/consent");
    }

    // tenants_read is a plain count, never per-tenant detail — safe to
    // expose, and useful for confirming client aggregation is happening
    // without needing working Application Insights.
    const payload = { updated_at: new Date().toISOString(), metrics, tenants_read: tenantsRead };
    if (debug) payload._debug_tenants = debugTenants;
    if (!debug) cache = { data: payload, expiresAt: now + ttlMs };
    context.log(`security-stats: aggregated ${tenantsRead} tenant(s)`);

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
