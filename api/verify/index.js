const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const msal = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET
  }
});

async function getToken() {
  const res = await msal.acquireTokenByClientCredential({ scopes: ['https://graph.microsoft.com/.default'] });
  return res.accessToken;
}

module.exports = async function (context, req) {
  try {
    const email = (req.body?.email || '').trim().toLowerCase();
    const code  = (req.body?.code  || '').trim();
    if (!email || !/^\d{6}$/.test(code)) {
      context.res = { status: 400, body: { error: 'Invalid input' } }; return;
    }

    const token = await getToken();
    const siteHost = process.env.SITE_HOSTNAME;   // e.g. firstchoicechildcare.sharepoint.com
    const sitePath = process.env.SITE_PATH;       // e.g. /sites/FirstChoiceManagement
    const listId  = process.env.LIST_ID_VERIF;    // GUID of Submission Verifications list
    const nowIso  = new Date().toISOString();

    // 1) Find matching, waiting, not-expired verification row
    const url = `https://graph.microsoft.com/v1.0/sites/${siteHost}:/sites${sitePath}:/lists/${listId}/items` +
                `?$expand=fields&$filter=fields/Token eq '${code}' and fields/Status eq 'Waiting' and fields/ExpiresOn ge '${nowIso}'`;
    const { data } = await axios.get(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!data.value || data.value.length === 0) {
      context.res = { status: 400, body: { error: 'Invalid or expired code' } }; return;
    }

    const item = data.value[0];
    const targetEmail = (item.fields.TargetEmail || '').toLowerCase();
    if (email !== targetEmail) {
      await axios.patch(
        `https://graph.microsoft.com/v1.0/sites/${siteHost}:/sites${sitePath}:/lists/${listId}/items/${item.id}/fields`,
        { Status: 'Suspicious', Notes: `Typed: ${email}` },
        { headers: { Authorization: `Bearer ${token}` } }
      );
      context.res = { status: 403, body: { error: 'Email does not match' } }; return;
    }

    // 2) Mark Verified
    await axios.patch(
      `https://graph.microsoft.com/v1.0/sites/${siteHost}:/sites${sitePath}:/lists/${listId}/items/${item.id}/fields`,
      { Status: 'Verified', VerifiedOn: nowIso },
      { headers: { Authorization: `Bearer ${token}` } }
    );

    context.res = { status: 200, body: { ok: true } };
  } catch (e) {
    context.log.error(e);
    context.res = { status: 500, body: { error: 'Server error' } };
  }
};
