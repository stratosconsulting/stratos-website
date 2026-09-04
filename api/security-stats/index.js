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
  const since30d = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();

  // The 5 metric reads for one tenant are independent — run them
  // concurrently instead of one-by-one. With up to 13 tenants each doing
  // several Graph calls, sequential adds up to real page-load latency.
  const [endpointsResult, complianceResult, mfaResult, alertsResult, riskResult, licensedUsersResult] = await Promise.allSettled([
    // Counts every device Microsoft 365/Entra ID knows about for this
    // tenant (registered, joined, hybrid joined) — NOT just Intune-enrolled
    // devices. Several client tenants don't have Intune deployed yet, and
    // this metric is meant to reflect everything under management, not one
    // product.
    graphGetAll(`${GRAPH}/devices?$filter=accountEnabled eq true&$select=id`, token),
    graphGetOne(`${GRAPH}/deviceManagement/deviceCompliancePolicyDeviceStateSummary`, token),
    graphGetAll(`${GRAPH}/reports/authenticationMethods/userRegistrationDetails?$select=id,isMfaRegistered`, token),
    // Formal alerts raised by Defender / Entra ID Protection / other
    // Microsoft security products.
    graphGetAll(`${GRAPH}/security/alerts_v2?$filter=createdDateTime ge ${since30d}&$select=id`, token),
    // Risky sign-ins Entra ID Protection flagged and acted on (requires
    // Entra ID P2 on the tenant — tenants without it just come back empty,
    // handled the same as any other missing metric below). This is a
    // second REAL detection source, not a multiplier on the first — a
    // risky sign-in and a Defender alert are different events, so summing
    // them in aggregate() below is additive, not inflated.
    graphGetAll(`${GRAPH}/identityProtection/riskDetections?$filter=activityDateTime ge ${since30d}&$select=id`, token),
    // Licensed + enabled member accounts only. Excludes unlicensed shared
    // mailboxes (info@, no-reply@, etc.) from the MFA denominator below —
    // those can't sign in interactively in practice and shouldn't drag
    // down mfa_coverage just for existing. A licensed account that never
    // registered MFA still counts against the metric, as it should.
    graphGetAll(`${GRAPH}/users?$filter=accountEnabled eq true and userType eq 'Member'&$select=id,assignedLicenses`, token),
  ]);

  if (endpointsResult.status === "fulfilled") {
    raw.endpoints = endpointsResult.value.length;
  } else {
    context.log.warn(`[${label}] endpoints failed:`, endpointsResult.reason?.message);
  }

  if (complianceResult.status === "fulfilled") {
    const compliant = complianceResult.value.compliantDeviceCount || 0;
    const nonCompliant = complianceResult.value.nonCompliantDeviceCount || 0;
    if (compliant + nonCompliant > 0) {
      raw.compliantDevices = compliant;
      raw.knownDevices = compliant + nonCompliant;
    }
  } else {
    context.log.warn(`[${label}] compliance failed:`, complianceResult.reason?.message);
  }

  if (mfaResult.status === "fulfilled" && mfaResult.value.length > 0) {
    // Restrict to licensed, enabled member accounts when we could read
    // that list — otherwise unlicensed shared mailboxes (info@, no-reply@,
    // QFG@, etc., which QFG confirmed have no license and aren't meant for
    // interactive sign-in) count as "no MFA" forever despite being a
    // non-issue. Falls back to the unfiltered set if the /users read
    // failed, rather than dropping the metric entirely.
    let scoped = mfaResult.value;
    if (licensedUsersResult.status === "fulfilled") {
      const licensedIds = new Set(
        licensedUsersResult.value
          .filter((u) => Array.isArray(u.assignedLicenses) && u.assignedLicenses.length > 0)
          .map((u) => u.id)
      );
      scoped = mfaResult.value.filter((u) => licensedIds.has(u.id));
    } else {
      context.log.warn(`[${label}] licensed-users read failed, using unfiltered MFA set:`, licensedUsersResult.reason?.message);
    }
    if (scoped.length > 0) {
      raw.mfaUsers = scoped.filter((u) => u.isMfaRegistered).length;
      raw.totalUsers = scoped.length;
    }
  } else if (mfaResult.status === "rejected") {
    context.log.warn(`[${label}] mfa failed:`, mfaResult.reason?.message);
  }

  // threats30d combines both real sources when at least one succeeded —
  // a tenant with only alerts_v2 access (no Entra ID P2) still contributes
  // its alert count instead of being dropped entirely.
  let threatsKnown = false;
  let threats = 0;
  if (alertsResult.status === "fulfilled") {
    threats += alertsResult.value.length;
    threatsKnown = true;
  } else {
    context.log.warn(`[${label}] alerts failed:`, alertsResult.reason?.message);
  }
  if (riskResult.status === "fulfilled") {
    threats += riskResult.value.length;
    threatsKnown = true;
  } else {
    // Expected/noisy for tenants without Entra ID P2 — logged at info
    // level implicitly via warn, but this is not a real failure to chase.
    context.log.warn(`[${label}] risk detections failed (often just missing Entra ID P2):`, riskResult.reason?.message);
  }
  if (threatsKnown) {
    raw.threats30d = threats;
  }

  return raw;
}

// Notes are stored keyed by a language-neutral noteKey, not baked in as
// Spanish text at aggregation time — the payload is cached once (see
// module.exports below) and served to both the ES and EN homepage, and
// the cached response used to carry hardcoded Spanish note strings
// ("Dispositivos Intune", "Ultimos 30 dias") that overwrote the English
// page's static sub-labels the moment live data loaded. Notes are
// localized at request-serve time instead, from NOTES[lang], so each
// language gets its own cache-friendly text without doubling the
// expensive tenant reads per language.
const NOTES = {
  es: {
    endpoints_protected: "Microsoft 365 / Entra ID",
    compliance_rate: "Dispositivos Intune",
    mfa_coverage: "Microsoft Entra ID",
    threats_blocked_30d: "Ultimos 30 dias",
  },
  en: {
    endpoints_protected: "Microsoft 365 / Entra ID",
    compliance_rate: "Intune Devices",
    mfa_coverage: "Microsoft Entra ID",
    threats_blocked_30d: "Last 30 Days",
  },
};

function aggregate(perTenant) {
  const metrics = {};

  const endpointsTotal = perTenant.reduce((sum, t) => sum + (t.endpoints || 0), 0);
  if (perTenant.some((t) => typeof t.endpoints === "number")) {
    metrics.endpoints_protected = { value: endpointsTotal, noteKey: "endpoints_protected" };
  }

  const compliantTotal = perTenant.reduce((sum, t) => sum + (t.compliantDevices || 0), 0);
  const knownTotal = perTenant.reduce((sum, t) => sum + (t.knownDevices || 0), 0);
  if (knownTotal > 0) {
    metrics.compliance_rate = {
      value: round1((compliantTotal / knownTotal) * 100),
      suffix: "%",
      noteKey: "compliance_rate",
    };
  }

  const mfaTotal = perTenant.reduce((sum, t) => sum + (t.mfaUsers || 0), 0);
  const usersTotal = perTenant.reduce((sum, t) => sum + (t.totalUsers || 0), 0);
  if (usersTotal > 0) {
    metrics.mfa_coverage = {
      value: round1((mfaTotal / usersTotal) * 100),
      suffix: "%",
      noteKey: "mfa_coverage",
    };
  }

  const threatsTotal = perTenant.reduce((sum, t) => sum + (t.threats30d || 0), 0);
  if (perTenant.some((t) => typeof t.threats30d === "number")) {
    metrics.threats_blocked_30d = { value: threatsTotal, noteKey: "threats_blocked_30d" };
  }

  return metrics;
}

// Applied at serve time (both the fresh-compute path and the cache-hit
// path) so the single cached payload can be localized differently per
// request without re-reading every tenant per language.
function localizeNotes(payload, lang) {
  const notes = NOTES[lang] || NOTES.es;
  const metrics = {};
  Object.keys(payload.metrics).forEach((key) => {
    const m = payload.metrics[key];
    const { noteKey, ...rest } = m;
    metrics[key] = noteKey && notes[noteKey] ? { ...rest, note: notes[noteKey] } : rest;
  });
  return { ...payload, metrics };
}

// Reads one tenant end to end (token exchange + its 5 metric calls) and
// never throws — a failure just means that tenant contributes nothing,
// logged and swallowed here so Promise.allSettled below always resolves.
async function readTenant(context, label, getToken) {
  try {
    const token = await getToken();
    return await computeTenantRaw(context, token, label);
  } catch (err) {
    context.log.warn(`${label} skipped:`, err.message);
    return null;
  }
}

async function computeMetrics(context) {
  // All tenants read concurrently instead of one-by-one — with up to 14
  // tenants (own + ~13 client), each doing a token exchange plus 5 Graph
  // calls, sequential reads used to add up to 10-30+ seconds on a cache
  // miss (the panel refreshes every CACHE_TTL_SECONDS, default 30 min).
  // Running them in parallel cuts that to roughly the slowest single
  // tenant instead of the sum of all of them. Safe to parallelize: each
  // client tenant's OBO refresh-token exchange reads the same starting
  // token from tokenStore and writes back its own freshly rotated one —
  // Azure AD tolerates that overlap (refresh tokens stay valid for a
  // short grace window after rotation specifically to support concurrent
  // use like this), so a race here just means whichever write lands last
  // is what's persisted for the next cycle, not a broken token.
  const clientTenantIds = getClientTenantIds();
  const reads = [
    readTenant(context, "stratos", async () => (await getOwnTenantCredential().getToken(SCOPE)).token),
    ...clientTenantIds.map((tenantId) =>
      readTenant(context, tenantId, () => getGraphTokenForTenant(context, tenantId))
    ),
  ];

  const results = await Promise.allSettled(reads);
  const perTenant = results
    .filter((r) => r.status === "fulfilled" && r.value)
    .map((r) => r.value);

  return { metrics: aggregate(perTenant), tenantsRead: perTenant.length };
}

module.exports = async function (context, req) {
  const ttlMs = (parseInt(process.env.CACHE_TTL_SECONDS, 10) || 1800) * 1000;
  const now = Date.now();
  const lang = req.query && req.query.lang === "en" ? "en" : "es";

  if (cache.data && cache.expiresAt > now) {
    context.res = { status: 200, headers: { "Content-Type": "application/json" }, body: localizeNotes(cache.data, lang) };
    return;
  }

  try {
    const { metrics, tenantsRead } = await computeMetrics(context);

    if (Object.keys(metrics).length === 0) {
      throw new Error("No metric succeeded for any tenant — check app permissions/consent");
    }

    // tenants_read is a plain count, never per-tenant detail — safe to
    // expose, and useful for confirming client aggregation is happening
    // without needing working Application Insights.
    const payload = { updated_at: new Date().toISOString(), metrics, tenants_read: tenantsRead };
    cache = { data: payload, expiresAt: now + ttlMs };
    context.log(`security-stats: aggregated ${tenantsRead} tenant(s)`);

    context.res = { status: 200, headers: { "Content-Type": "application/json" }, body: localizeNotes(payload, lang) };
  } catch (err) {
    context.log.error("security-stats failed:", err.message);
    // Fail closed but harmless: the frontend already falls back to sample
    // numbers when this endpoint doesn't return 200 (see template-v2.html).
    if (cache.data) {
      context.res = { status: 200, headers: { "Content-Type": "application/json" }, body: localizeNotes(cache.data, lang) };
    } else {
      context.res = { status: 503, headers: { "Content-Type": "application/json" }, body: { error: "temporarily_unavailable" } };
    }
  }
};
