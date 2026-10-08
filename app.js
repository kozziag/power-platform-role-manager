/*
 * Power Platform Role Manager -- client-side only, no backend.
 *
 * Auth: MSAL.js public-client SPA flow (Authorization Code + PKCE) against a
 * dedicated, multi-tenant Entra app registration (NOT the Azure CLI's app ID
 * -- that trick is fine for a personal desktop tool but inappropriate for a
 * publicly hosted app). Every visitor signs in with their own org's Entra ID;
 * tokens live only in their own browser (sessionStorage), never touch any
 * server this project controls.
 *
 * Confirmed live before building this: both Microsoft Graph and the Power
 * Platform Authorization API send Access-Control-Allow-Origin: * on their
 * preflight responses, so calling them directly from a static page works.
 *
 * APIs used (same ones verified in the desktop version):
 *   GET    https://graph.microsoft.com/v1.0/users?$search=...
 *   GET    https://graph.microsoft.com/v1.0/servicePrincipals?$filter=...
 *   GET    https://graph.microsoft.com/v1.0/organization
 *   GET    https://api.powerplatform.com/authorization/roleDefinitions
 *   GET    https://api.powerplatform.com/authorization/roleAssignments?$filter=principalObjectId eq <guid>
 *   POST   https://api.powerplatform.com/authorization/roleAssignments
 *   DELETE https://api.powerplatform.com/authorization/roleAssignments/{id}
 */

// Registered app (see README.md for how this was created / how to re-point it
// at your own registration). Multi-tenant: AzureADMultipleOrgs.
const CLIENT_ID = "2fad7640-e268-49b8-8a9c-0ea1b0bf0921";
const PP_AUTH_BASE = "https://api.powerplatform.com/authorization";
const API_VERSION = "1";

const GRAPH_SCOPES = ["User.Read.All"];
const PP_SCOPES = [
  "https://api.powerplatform.com/Authorization.RoleAssignments.Read",
  "https://api.powerplatform.com/Authorization.RoleAssignments.Write",
];

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: "https://login.microsoftonline.com/organizations",
    redirectUri: window.location.href.split("#")[0].split("?")[0],
  },
  cache: {
    cacheLocation: "sessionStorage", // cleared when the tab closes -- nothing persists to disk
    storeAuthStateInCookie: false,
  },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
let activeAccount = null;
let selectedPrincipal = null;
let roleDefs = [];
const activityLog = [];

function authHeader(token) {
  const scheme = "Bearer";
  return scheme + " " + token;
}

async function getGraphToken() {
  const req = { scopes: GRAPH_SCOPES, account: activeAccount };
  try {
    const r = await msalInstance.acquireTokenSilent(req);
    return r.accessToken;
  } catch (e) {
    const r = await msalInstance.acquireTokenPopup(req);
    return r.accessToken;
  }
}

async function getPpToken() {
  const req = { scopes: PP_SCOPES, account: activeAccount };
  try {
    const r = await msalInstance.acquireTokenSilent(req);
    return r.accessToken;
  } catch (e) {
    const r = await msalInstance.acquireTokenPopup(req);
    return r.accessToken;
  }
}

async function graphFetch(path, opts = {}) {
  const token = await getGraphToken();
  const headers = Object.assign({}, opts.headers, { Authorization: authHeader(token) });
  const r = await fetch("https://graph.microsoft.com/v1.0" + path, Object.assign({}, opts, { headers }));
  const data = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error((data.error && data.error.message) || r.statusText);
  return data;
}

async function ppFetch(path, opts = {}) {
  const token = await getPpToken();
  const headers = Object.assign(
    { Accept: "application/json", "Content-Type": "application/json" },
    opts.headers,
    { Authorization: authHeader(token) }
  );
  const r = await fetch(PP_AUTH_BASE + path, Object.assign({}, opts, { headers }));
  if (r.status === 204) return {};
  const data = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(typeof data === "string" ? data : JSON.stringify(data));
  return data;
}

function banner(msg, kind) {
  document.getElementById("banner").innerHTML = `<div class="banner ${kind}">${msg}</div>`;
}

function showConfirm(message, { title = "Confirm", confirmLabel = "OK", danger = false } = {}) {
  return new Promise((resolve) => {
    const overlay = document.getElementById("modalOverlay");
    const confirmBtn = document.getElementById("modalConfirmBtn");
    const cancelBtn = document.getElementById("modalCancelBtn");
    document.getElementById("modalTitle").textContent = title;
    document.getElementById("modalMessage").textContent = message;
    confirmBtn.textContent = confirmLabel;
    confirmBtn.className = danger ? "danger" : "";
    cancelBtn.classList.remove("hidden");
    overlay.classList.remove("hidden");
    const cleanup = (result) => {
      overlay.classList.add("hidden");
      confirmBtn.onclick = null;
      cancelBtn.onclick = null;
      resolve(result);
    };
    confirmBtn.onclick = () => cleanup(true);
    cancelBtn.onclick = () => cleanup(false);
  });
}

async function signIn() {
  const statusEl = document.getElementById("signInStatus");
  statusEl.textContent = "Opening sign-in…";
  document.getElementById("btnSignIn").disabled = true;
  try {
    // prompt: 'select_account' avoids silently reusing whichever Microsoft
    // account already has an SSO session in this browser.
    const result = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES, prompt: "select_account" });
    activeAccount = result.account;
    msalInstance.setActiveAccount(activeAccount);
    statusEl.textContent = "";
    await afterSignIn();
  } catch (e) {
    statusEl.textContent = "Sign-in failed: " + e.message;
  } finally {
    document.getElementById("btnSignIn").disabled = false;
  }
}

async function afterSignIn() {
  document.getElementById("signedOutView").classList.add("hidden");
  document.getElementById("signedInView").classList.remove("hidden");
  document.getElementById("btnSignOut").classList.remove("hidden");

  let tenantName = null;
  try {
    const org = await graphFetch("/organization?$select=displayName");
    tenantName = (org.value && org.value[0] && org.value[0].displayName) || null;
  } catch (e) { /* non-fatal -- tenant display name is a nice-to-have */ }

  const tenantLabel = tenantName
    ? `${tenantName} · tenant ${activeAccount.tenantId}`
    : `tenant ${activeAccount.tenantId}`;
  document.getElementById("whoami").textContent = `${activeAccount.name || activeAccount.username} · ${tenantLabel}`;

  loadRoles();
}

async function signOut() {
  document.getElementById("signedOutView").classList.remove("hidden");
  document.getElementById("signedInView").classList.add("hidden");
  document.getElementById("btnSignOut").classList.add("hidden");
  document.getElementById("whoami").textContent = "";
  selectedPrincipal = null;
  document.getElementById("userResults").innerHTML = "";
  document.getElementById("appResult").innerHTML = "";
  document.getElementById("selectedPrincipal").innerHTML = "";
  document.getElementById("userSearch").value = "";
  if (activeAccount) {
    try {
      await msalInstance.logoutPopup({ account: activeAccount, mainWindowRedirectUri: window.location.href });
    } catch (e) { /* popup closed/blocked -- local state is already cleared above */ }
  }
  activeAccount = null;
}

document.getElementById("btnSignIn").onclick = signIn;
document.getElementById("btnSignOut").onclick = signOut;

let searchTimer;
document.getElementById("userSearch").addEventListener("input", (e) => {
  clearTimeout(searchTimer);
  const q = e.target.value;
  searchTimer = setTimeout(async () => {
    if (q.trim().length < 2) { document.getElementById("userResults").innerHTML = ""; return; }
    try {
      const search = `"displayName:${q}" OR "mail:${q}" OR "userPrincipalName:${q}"`;
      const r = await graphFetch(
        `/users?$search=${encodeURIComponent(search)}&$count=true&$select=id,displayName,userPrincipalName,mail,jobTitle&$top=15`,
        { headers: { ConsistencyLevel: "eventual" } }
      );
      const el = document.getElementById("userResults");
      el.innerHTML = (r.value || []).map(u => `
        <div class="result-row" data-id="${u.id}" data-label="${u.displayName} (${u.userPrincipalName || u.mail || ""})">
          <b>${u.displayName}</b> <span class="muted">${u.userPrincipalName || u.mail || ""}</span>
          ${u.jobTitle ? `<div class="muted">${u.jobTitle}</div>` : ""}
        </div>`).join("") || '<p class="muted">No matches.</p>';
      el.querySelectorAll(".result-row").forEach(row => {
        row.onclick = () => selectPrincipal(row.dataset.id, "User", row.dataset.label, el);
      });
    } catch (e) { banner("User search failed: " + e.message, "error"); }
  }, 350);
});

document.getElementById("btnLookupApp").onclick = async () => {
  const appId = document.getElementById("appIdInput").value.trim();
  if (!appId) return;
  const el = document.getElementById("appResult");
  try {
    const r = await graphFetch(`/servicePrincipals?$filter=appId eq '${appId}'&$select=id,displayName,appId`);
    const values = r.value || [];
    if (!values.length) {
      el.innerHTML = '<p class="muted">No enterprise application found for that App (client) ID in this tenant.</p>';
      return;
    }
    const sp = values[0];
    el.innerHTML = `<div class="result-row" data-id="${sp.id}" data-label="${sp.displayName} (service principal)">
      <b>${sp.displayName}</b> <span class="muted">object ID: ${sp.id}</span></div>`;
    el.querySelector(".result-row").onclick = () => selectPrincipal(sp.id, "ApplicationUser", `${sp.displayName} (service principal)`, el);
  } catch (e) {
    el.innerHTML = `<p class="muted">${e.message}</p>`;
  }
};

function selectPrincipal(objectId, type, label, containerEl) {
  selectedPrincipal = { objectId, type, label };
  containerEl.querySelectorAll(".result-row").forEach(r => r.classList.remove("selected"));
  const match = [...containerEl.querySelectorAll(".result-row")].find(r => r.dataset.id === objectId);
  if (match) match.classList.add("selected");
  const pillClass = type === "User" ? "pill-user" : "pill-app";
  document.getElementById("selectedPrincipal").innerHTML =
    `Selected: <b>${label}</b> <span class="pill ${pillClass}">${type}</span>`;
  updateAssignButton();
  loadPrincipalAssignments();
}

async function loadRoles() {
  const sel = document.getElementById("roleSelect");
  try {
    const r = await ppFetch(`/roleDefinitions?api-version=${API_VERSION}`);
    roleDefs = r.value || [];
    sel.innerHTML = roleDefs.map(rd => `<option value="${rd.roleDefinitionId}">${rd.roleDefinitionName}</option>`).join("");
    updateRoleDescription();
    updateAssignButton();
  } catch (e) {
    sel.innerHTML = "<option>Failed to load roles</option>";
    banner("Could not load role definitions: " + e.message, "error");
  }
}

function updateRoleDescription() {
  const sel = document.getElementById("roleSelect");
  const rd = roleDefs.find(r => r.roleDefinitionId === sel.value);
  document.getElementById("roleDescription").textContent = rd ? rd.description : "";
}
document.getElementById("roleSelect").addEventListener("change", updateRoleDescription);

function updateAssignButton() {
  document.getElementById("btnAssign").disabled = !(selectedPrincipal && document.getElementById("roleSelect").value);
}

document.getElementById("btnAssign").onclick = async () => {
  const statusEl = document.getElementById("assignStatus");
  const roleDefinitionId = document.getElementById("roleSelect").value;
  const roleDef = roleDefs.find(r => r.roleDefinitionId === roleDefinitionId);
  statusEl.textContent = "Assigning…";
  try {
    await ppFetch(`/roleAssignments?api-version=${API_VERSION}`, {
      method: "POST",
      body: JSON.stringify({
        scope: `/tenants/${activeAccount.tenantId}`,
        principalType: selectedPrincipal.type,
        principalObjectId: selectedPrincipal.objectId,
        roleDefinitionId,
      }),
    });
    statusEl.textContent = "Assigned ✓";
    addActivity("Assigned", selectedPrincipal.label, roleDef ? roleDef.roleDefinitionName : roleDefinitionId);
    loadPrincipalAssignments();
  } catch (e) {
    statusEl.textContent = "";
    banner("Assignment failed: " + e.message, "error");
  }
};

async function loadPrincipalAssignments() {
  const card = document.getElementById("principalAssignmentsCard");
  const tbody = document.querySelector("#principalAssignmentsTable tbody");
  if (!selectedPrincipal) { card.classList.add("hidden"); return; }
  card.classList.remove("hidden");
  tbody.innerHTML = '<tr><td colspan="3" class="muted">Loading…</td></tr>';
  try {
    const filt = `principalObjectId eq ${selectedPrincipal.objectId}`;
    const r = await ppFetch(`/roleAssignments?api-version=${API_VERSION}&$filter=${encodeURIComponent(filt)}`);
    const roleDefsById = {};
    roleDefs.forEach(rd => { roleDefsById[rd.roleDefinitionId] = rd; });
    tbody.innerHTML = (r.value || []).map(a => {
      const rd = roleDefsById[a.roleDefinitionId];
      const roleLabel = rd ? rd.roleDefinitionName : `<code>${a.roleDefinitionId}</code>`;
      return `<tr>
        <td>${roleLabel}</td>
        <td class="muted">${(a.createdOn || "").slice(0, 10)}</td>
        <td><button class="danger" onclick="removeAssignment('${a.roleAssignmentId}')">Remove</button></td>
      </tr>`;
    }).join("") || '<tr><td colspan="3" class="muted">No existing role assignments for this principal.</td></tr>';
  } catch (e) {
    tbody.innerHTML = `<tr><td colspan="3" class="muted">Could not load: ${e.message}</td></tr>`;
  }
}

function addActivity(action, principalLabel, roleName) {
  activityLog.unshift({ when: new Date(), action, principalLabel, roleName });
  if (activityLog.length > 25) activityLog.length = 25;
  const tbody = document.querySelector("#activityTable tbody");
  tbody.innerHTML = activityLog.map(e => `
    <tr>
      <td class="muted">${e.when.toLocaleTimeString()}</td>
      <td><span class="pill ${e.action === "Assigned" ? "pill-user" : "pill-app"}">${e.action}</span></td>
      <td>${e.principalLabel}</td>
      <td>${e.roleName}</td>
    </tr>`).join("");
}

async function removeAssignment(id) {
  const ok = await showConfirm("Remove this role assignment? This cannot be undone.", {
    title: "Remove role assignment", confirmLabel: "Remove", danger: true,
  });
  if (!ok) return;
  try {
    const row = [...document.querySelectorAll("#principalAssignmentsTable tbody tr")]
      .find(tr => tr.innerHTML.includes(id));
    const roleName = row ? row.querySelector("td").textContent : id;
    await ppFetch(`/roleAssignments/${id}?api-version=${API_VERSION}`, { method: "DELETE" });
    addActivity("Removed", selectedPrincipal ? selectedPrincipal.label : "(unknown)", roleName);
    loadPrincipalAssignments();
  } catch (e) {
    banner("Remove failed: " + e.message, "error");
  }
}

// Boot: if MSAL already has a cached account (e.g. page refresh), re-use it silently.
(async function boot() {
  await msalInstance.initialize();
  await msalInstance.handleRedirectPromise().catch(() => {});
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    activeAccount = accounts[0];
    msalInstance.setActiveAccount(activeAccount);
    await afterSignIn();
  }
})();
