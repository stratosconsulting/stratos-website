// delegatedAuth.js
//
// Exchanges the stored OBO refresh token for a Graph access token scoped to
// one specific client tenant. This is the "Secure Application Model" flow
// GDAP requires — application-only (client-credentials) tokens are NOT
// accepted against a GDAP customer tenant, only delegated app+user tokens
// obtained this way. See scripts/README-obo-consent.md for the one-time
// setup this depends on.
//
// Every exchange returns a NEW refresh token (Azure AD rotates it), which
// we persist immediately via tokenStore so the next call — for the next
// tenant, or the next cache cycle — keeps working.

const { readRefreshToken, writeRefreshToken } = require("./tokenStore");

const GRAPH_SCOPE = "https://graph.microsoft.com/.default offline_access";

async function getGraphTokenForTenant(context, tenantId) {
  const refreshToken = await readRefreshToken(context);

  // No client_secret here on purpose. The app has "Allow public client
  // flows" enabled (required for the device-code login in
  // scripts/obo-delegated-setup.ps1), and once that's on, Azure AD treats
  // this specific flow as a public client and REJECTS a client_secret with
  // AADSTS700025. The client-only-mode ClientSecretCredential path (own
  // tenant, application permissions) is unaffected — that one authenticates
  // as a confidential client separately and still needs the secret.
  const body = new URLSearchParams({
    client_id: process.env.GRAPH_CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: GRAPH_SCOPE,
  });

  const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const json = await res.json();

  if (!res.ok) {
    // Common causes: this tenant hasn't gone through the app-consent step
    // yet (scripts/obo-delegated-setup.ps1), the GDAP relationship/role
    // lapsed, or the OBO refresh token itself is stale (>90 days unused).
    throw new Error(
      `token exchange for tenant ${tenantId} -> ${res.status} ${json.error || ""} ${json.error_description || ""}`.slice(0, 400)
    );
  }

  if (json.refresh_token) {
    await writeRefreshToken(context, json.refresh_token);
  }

  return json.access_token;
}

module.exports = { getGraphTokenForTenant };
