// tokenStore.js
//
// Holds the OBO (on-behalf-of) delegated refresh token used to read client
// tenants via GDAP. Azure AD rotates the refresh token every time it's
// exchanged, so we persist the latest one in Azure Table Storage (the same
// storage account the Function App already uses for AzureWebJobsStorage —
// no new resource to provision).
//
// Seeding: the very first value comes from the GRAPH_OBO_REFRESH_TOKEN
// Application Setting (pasted in once, after running
// scripts/obo-delegated-setup.ps1). After that, this module owns rotation
// and the Application Setting is never read again — updating it later has
// no effect unless the Table Storage row is manually cleared.

const { TableClient } = require("@azure/data-tables");

const TABLE_NAME = "SecurityPanelTokens";
const PARTITION_KEY = "obo";
const ROW_KEY = "refresh_token";

let tableClientPromise = null;

function getTableClient(context) {
  if (!tableClientPromise) {
    const conn = process.env.AzureWebJobsStorage;
    if (!conn) {
      throw new Error("AzureWebJobsStorage is not set — can't reach Table Storage for the OBO token.");
    }
    const client = TableClient.fromConnectionString(conn, TABLE_NAME, {
      allowInsecureConnection: true,
    });
    tableClientPromise = client
      .createTable()
      .catch((err) => {
        // 409 = table already exists, which is the common case after the first run.
        if (err.statusCode !== 409) throw err;
      })
      .then(() => client);
  }
  return tableClientPromise;
}

async function readRefreshToken(context) {
  const client = await getTableClient(context);
  try {
    const entity = await client.getEntity(PARTITION_KEY, ROW_KEY);
    return entity.value;
  } catch (err) {
    if (err.statusCode === 404) {
      // Nothing stored yet — fall back to the seed value from Application Settings.
      const seed = process.env.GRAPH_OBO_REFRESH_TOKEN;
      if (!seed) {
        throw new Error(
          "No OBO refresh token in Table Storage and GRAPH_OBO_REFRESH_TOKEN is not set. Run scripts/obo-delegated-setup.ps1 first."
        );
      }
      return seed;
    }
    throw err;
  }
}

async function writeRefreshToken(context, value) {
  const client = await getTableClient(context);
  await client.upsertEntity(
    { partitionKey: PARTITION_KEY, rowKey: ROW_KEY, value, updatedAt: new Date().toISOString() },
    "Replace"
  );
}

module.exports = { readRefreshToken, writeRefreshToken };
