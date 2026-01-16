// Rename this file to config.js and fill values.
// Tip: keep this file OUT of source control if you add any sensitive info (you shouldn't need secrets).

window.COCKPIT_CONFIG = {
  appName: "Cockpit Central",

  // Microsoft Entra ID (Azure AD) App Registration
  // - Single-page application (SPA)
  // - Redirect URI: https://localhost:5173 (or your deployed URL)
  tenantId: "YOUR_TENANT_ID_OR_DOMAIN", // e.g. contoso.onmicrosoft.com or GUID
  clientId: "YOUR_APP_CLIENT_ID",
  redirectUri: "http://localhost:5173",

  // SharePoint list backend (Microsoft Lists)
  // You can set these after running the PowerShell scripts in /scripts
  siteId: "YOUR_SITE_ID",   // e.g. contoso.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy
  listId: "YOUR_LIST_ID",

  // Quick links
  links: {
    bienChezSoiApp: "https://example.com",
    evolumisApp: "https://example.com",
    persoApp: "https://example.com",

    // SharePoint folder (web URL). Use the library/folder you want.
    sharePointFolderUrl: "https://contoso.sharepoint.com/sites/SolutionsEvolumis/Shared%20Documents",

    // Optional: direct link to open the List in SharePoint
    listWebUrl: "",
  },

  // Pole labels (UI)
  poles: [
    { key: "BCS",   label: "Bien Chez Soi",  emoji: "🏡" },
    { key: "EVO",   label: "Evolumis",       emoji: "🚀" },
    { key: "PERSO", label: "Personnel",     emoji: "🧠" },
  ],

  // Status config (must match List internal values)
  statuses: [
    { key: "Backlog", label: "Backlog" },
    { key: "EnCours", label: "En cours" },
    { key: "EnAttente", label: "En attente" },
    { key: "Termine", label: "Terminé" },
  ],

  priorities: [
    { key: "P1", label: "P1" },
    { key: "P2", label: "P2" },
    { key: "P3", label: "P3" },
  ],
};
