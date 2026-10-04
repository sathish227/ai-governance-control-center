import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';

import { tenantId, clientId } from './config.js';
const configured = Boolean(tenantId && clientId);
const scopes = ['Application.Read.All'];
const agentScopes = ['AgentIdentity.Read.All'];
const pendingAgentKey = 'governance-load-agent-identities';
const graphOrigin = 'https://graph.microsoft.com';
const redirectUri = new URL('./', window.location.href).href.split('#')[0];
const byId = id => document.getElementById(id);
const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));

const pca = configured ? new PublicClientApplication({
  auth: { clientId, authority: 'https://login.microsoftonline.com/' + tenantId, redirectUri, navigateToLoginRequestUrl: false },
  cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
}) : null;
let principals = [];
let agentIdentities = [];
byId('tenant-redirect-uri').textContent = redirectUri;
if (!configured) {
  byId('tenant-connect').disabled = true;
  byId('agent-load').disabled = true;
  setStatus('Sample mode. Set your tenant ID and client ID in src/config.js, then rebuild to enable Microsoft Graph.');
}

function setStatus(message, error = false) {
  byId('tenant-status').textContent = message;
  byId('tenant-status').classList.toggle('error', error);
}

async function token(requestedScopes = scopes) {
  const account = pca.getActiveAccount() || pca.getAllAccounts()[0];
  if (!account) throw new Error('Sign in to your test tenant first.');
  try {
    return (await pca.acquireTokenSilent({ scopes: requestedScopes, account })).accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      if (requestedScopes.includes('AgentIdentity.Read.All')) sessionStorage.setItem(pendingAgentKey, 'true');
      await pca.acquireTokenRedirect({ scopes: requestedScopes, account, redirectUri });
      return null;
    }
    throw error;
  }
}

async function graphGet(pathOrUrl, requestedScopes = scopes) {
  const url = new URL(pathOrUrl, graphOrigin);
  if (url.origin !== graphOrigin || !url.pathname.startsWith('/v1.0/')) throw new Error('Unexpected Graph continuation URL.');
  const accessToken = await token(requestedScopes);
  if (!accessToken) return null;
  const response = await fetch(url, { headers: { Authorization: 'Bearer ' + accessToken, Accept: 'application/json' } });
  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    const code = body.error?.code || 'Graph request failed';
    throw new Error(code + ' (' + response.status + '). Check Graph permission, admin consent, and your Entra role.');
  }
  return response.json();
}

async function collection(first, maxPages = 10, requestedScopes = scopes) {
  let next = first, pages = 0;
  const items = [];
  while (next && pages < maxPages) {
    const data = await graphGet(next, requestedScopes);
    if (!data) return { items: [], truncated: false, redirecting: true };
    items.push(...(data.value || []));
    next = data['@odata.nextLink'];
    pages++;
  }
  return { items, truncated: !!next };
}

function renderInventory() {
  const query = byId('tenant-search').value.trim().toLocaleLowerCase();
  const matches = principals.filter(p => (p.displayName || '').toLocaleLowerCase().includes(query) || (p.appId || '').includes(query));
  byId('tenant-table').innerHTML = matches.map(p => `<tr><td><strong>${escapeHtml(p.displayName || 'Unnamed application')}</strong><span class="sub">${escapeHtml(p.appId)}</span></td><td>${escapeHtml(p.servicePrincipalType || '—')}</td><td>${escapeHtml(p.publisherName || '—')}</td><td><button class="tenant-table-btn" data-sp-id="${escapeHtml(p.id)}">View grants →</button></td></tr>`).join('');
  byId('tenant-empty').hidden = matches.length > 0;
}

async function refresh() {
  setStatus('Loading enterprise applications from Microsoft Graph…');
  byId('tenant-refresh').disabled = true;
  try {
    const result = await collection('/v1.0/servicePrincipals?$select=id,appId,displayName,servicePrincipalType,publisherName&$top=100', 20);
    if (result.redirecting) return;
    principals = result.items.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
    byId('tenant-inventory').hidden = false;
    byId('tenant-detail').hidden = true;
    byId('tenant-heading').textContent = 'Enterprise applications';
    byId('tenant-connect').hidden = true;
    byId('tenant-refresh').hidden = false;
    renderInventory();
    setStatus(`${principals.length} enterprise applications loaded${result.truncated ? ' (partial list; Graph returned more pages)' : ''}. Select one to inspect granted application permissions.`);
  } catch (error) {
    setStatus(error.message || 'Unable to load applications.', true);
  } finally {
    byId('tenant-refresh').disabled = false;
  }
}

function setAgentStatus(message, error = false) {
  byId('agent-status').textContent = message;
  byId('agent-status').classList.toggle('error', error);
}

function renderAgents() {
  const query = byId('agent-search').value.trim().toLocaleLowerCase();
  const matches = agentIdentities.filter(a => [a.displayName, a.id, a.agentIdentityBlueprintId].some(value => String(value || '').toLocaleLowerCase().includes(query)));
  byId('agent-table').innerHTML = matches.map(a => `<tr><td><strong>${escapeHtml(a.displayName || 'Unnamed agent')}</strong></td><td>${a.accountEnabled === true ? 'Enabled' : a.accountEnabled === false ? 'Disabled' : 'Not returned'}</td><td>${escapeHtml(a.createdDateTime || 'Not returned')}</td><td>${escapeHtml(a.id)}</td><td>${escapeHtml(a.agentIdentityBlueprintId || 'Not returned')}</td></tr>`).join('');
  byId('agent-empty').hidden = matches.length > 0;
}

async function loadAgents() {
  byId('agent-results').hidden = true;
  if (!pca?.getActiveAccount()) {
    setAgentStatus('Connect Entra ID first, then load agent identities.', true);
    return;
  }
  byId('agent-load').disabled = true;
  setAgentStatus('Loading Entra agent identities from Microsoft Graph…');
  try {
    const result = await collection('/v1.0/servicePrincipals/microsoft.graph.agentIdentity?$select=id,displayName,accountEnabled,createdDateTime,agentIdentityBlueprintId,servicePrincipalType&$top=100', 20, agentScopes);
    if (result.redirecting) return;
    agentIdentities = result.items.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
    renderAgents();
    byId('agent-results').hidden = false;
    setAgentStatus(`${agentIdentities.length} Entra agent identities loaded${result.truncated ? ' (partial list; more pages are available)' : ''}. This inventory covers agents with Entra agent identities.`);
  } catch (error) {
    setAgentStatus('Agent inventory could not be loaded. ' + (error.message || '') + ' Check delegated AgentIdentity.Read.All, admin consent, and the signed-in account’s Agent ID access (Agent ID Administrator for nonowners).', true);
  } finally {
    byId('agent-load').disabled = false;
  }
}

async function showGrants(id) {
  const principal = principals.find(p => p.id === id);
  if (!principal) return;
  const detail = byId('tenant-detail');
  detail.hidden = false;
  detail.replaceChildren();
  const heading = document.createElement('h2');
  heading.textContent = principal.displayName || 'Unnamed application';
  const info = document.createElement('p');
  info.textContent = 'Application (client) ID: ' + (principal.appId || '—');
  const loading = document.createElement('p');
  loading.textContent = 'Loading granted application permissions…';
  detail.append(heading, info, loading);
  try {
    const result = await collection('/v1.0/servicePrincipals/' + encodeURIComponent(id) + '/appRoleAssignments?$top=100', 10);
    const resourceIds = [...new Set(result.items.map(a => a.resourceId).filter(Boolean))];
    const resources = await Promise.all(resourceIds.map(async resourceId => {
      try { return await graphGet('/v1.0/servicePrincipals/' + encodeURIComponent(resourceId) + '?$select=id,displayName,appRoles'); }
      catch { return null; }
    }));
    const resourceMap = new Map(resources.filter(Boolean).map(r => [r.id, r]));
    loading.remove();
    const note = document.createElement('div');
    note.className = 'tenant-callout';
    note.textContent = 'Granted application permissions are live Entra data. A grant does not prove that an app actively uses it or accesses sensitive content.';
    detail.append(note);
    if (!result.items.length) {
      const empty = document.createElement('p');
      empty.textContent = 'No granted application permissions returned for this service principal.';
      detail.append(empty);
    } else {
      const list = document.createElement('ul');
      for (const grant of result.items) {
        const resource = resourceMap.get(grant.resourceId);
        const role = resource?.appRoles?.find(r => r.id?.toLowerCase() === grant.appRoleId?.toLowerCase());
        const item = document.createElement('li');
        item.textContent = (resource?.displayName || grant.resourceDisplayName || 'API') + ' — ' + (role?.value || 'Role ID: ' + grant.appRoleId);
        list.append(item);
      }
      detail.append(list);
    }
    if (result.truncated) {
      const note = document.createElement('p');
      note.className = 'tenant-small';
      note.textContent = 'Only the first 1,000 grants are shown.';
      detail.append(note);
    }
    detail.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  } catch (error) {
    loading.textContent = error.message || 'Unable to load granted permissions.';
  }
}

byId('tenant-connect').addEventListener('click', async () => {
  if (!configured) return;
  setStatus('Redirecting to Microsoft sign-in…');
  try { await pca.loginRedirect({ scopes, redirectUri, prompt: 'select_account' }); }
  catch (error) { setStatus(error.message || 'Sign-in could not start.', true); }
});
byId('tenant-refresh').addEventListener('click', refresh);
byId('tenant-search').addEventListener('input', renderInventory);
byId('agent-load').addEventListener('click', loadAgents);
byId('agent-search').addEventListener('input', renderAgents);
byId('tenant-table').addEventListener('click', event => {
  const id = event.target.closest('button[data-sp-id]')?.dataset.spId;
  if (id) showGrants(id);
});

(async () => {
  try {
    if (!configured) {
      const requestedView = window.location.hash.slice(1);
      window.showGovernanceView(window.governanceViews.includes(requestedView) ? requestedView : 'dashboard');
      return;
    }
    await pca.initialize();
    const result = await pca.handleRedirectPromise();
    if (result?.account) {
      pca.setActiveAccount(result.account);
      window.showGovernanceView('tenant');
    } else if (pca.getAllAccounts().length) {
      pca.setActiveAccount(pca.getAllAccounts()[0]);
    }
    if (!result?.account) {
      const requestedView = window.location.hash.slice(1);
      window.showGovernanceView(window.governanceViews.includes(requestedView) ? requestedView : 'dashboard');
    }
    if (pca.getActiveAccount()) await refresh();
    if (sessionStorage.getItem(pendingAgentKey) === 'true') {
      sessionStorage.removeItem(pendingAgentKey);
      await loadAgents();
    }
  } catch (error) {
    sessionStorage.removeItem(pendingAgentKey);
    window.showGovernanceView('tenant');
    setStatus(error.message || 'Sign-in could not be completed.', true);
  }
})();
