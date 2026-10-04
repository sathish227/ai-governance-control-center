# AI Governance Control Center

A static, interactive prototype for exploring AI governance. The Overview, AI applications, Control assessment, and Recommendations sections use **illustrative records and scores**. They reset on refresh. The separate Test tenant section optionally reads enterprise application inventory and granted application permissions from Microsoft Graph.

## Run locally

Requires Node.js and npm.

```bash
npm ci
npm run build
npx serve docs
```

Open the local URL shown by the server. The sample dashboard works with the empty configuration in `src/config.js`. Use a local HTTP server rather than opening `docs/index.html` directly.

## Connect your own Microsoft Entra test tenant

1. Register a **single-tenant** application in Microsoft Entra ID. Add a **Single-page application (SPA)** redirect URI matching the base URL shown in the Test tenant view, including the trailing slash. For example, a site served at `http://localhost:3000/` needs that exact URI. If deployed under a path such as `https://example.github.io/repo/`, register that exact URI instead.
2. Add Microsoft Graph **delegated** permission `Application.Read.All`. Have an administrator review and grant consent as required by your tenant. Do not add a client secret to this browser app.
3. Put the tenant ID and application (client) ID in `src/config.js`; run `npm run build` again. The generated `docs/tenant.js` includes these public identifiers. An empty configuration keeps the live connection disabled.
4. Serve `docs/`, open the Test tenant view, and select **Connect Entra ID**.

The app uses MSAL Browser with authorization code and PKCE, stores authentication state in browser session storage, and calls Microsoft Graph directly from the browser. Graph access tokens remain in the browser session. It reads `/v1.0/servicePrincipals` and selected service principals' `/appRoleAssignments`, including resource role names. It does not call an OpenAI API or upload tenant inventory to an AI model.

## Publishing considerations

- Review the values in `src/config.js` before pushing. Tenant and client IDs are not secrets, but they identify your environment. This package deliberately leaves them empty. Never commit tokens, client secrets, or screenshots showing live tenant inventory.
- A public repository and public hosted page expose the interface to everyone. Entra consent and tenant access still govern Graph calls. Use a test tenant and review who can sign in to the registered enterprise application before enabling live access on a public page.
- A public sample-only demo can leave `src/config.js` empty. The built `docs/` folder is ready for a static host; the Entra redirect URI must match the final URL if the tenant connection is enabled.
- This repository is source code for a standalone browser app. It is **not** a ChatGPT Enterprise integration or a tenant-wide AI discovery product.
- The 4 AI applications, control states, risk labels, and recommendations are synthetic. Enterprise application counts are not AI agent counts. A granted permission is not evidence that it was used. The live view does not assess delegated OAuth grants, Purview data access, or actual AI safety controls.

## Entra agent identity inventory

The Test tenant view now has a separate **Load agent identities** action. It queries `GET /v1.0/servicePrincipals/microsoft.graph.agentIdentity`; it does not infer agents from application names. Add Microsoft Graph **delegated** `AgentIdentity.Read.All` to your own registration and grant admin consent. Microsoft documents **Agent ID Administrator** as the supported role for nonowner delegated calls. Activate the role if your tenant uses PIM. Sign in to the tenant, then load the agent inventory to acquire a token for the extra read scope.

The agent table shows display name, account enabled status, creation date, object ID, and blueprint ID, with local search. Authorization failures are shown as errors rather than a zero count. This covers Entra agent identities only; it is not a complete Agent 365 inventory, activity assessment, or live governance score.

Reference: https://learn.microsoft.com/en-us/graph/api/agentidentity-list?view=graph-rest-1.0
