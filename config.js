// Cockpit Central - Configuration (aucun secret requis)

window.COCKPIT_CONFIG = {
  appName: "Cockpit Central",

  // Microsoft Entra ID (Azure AD) App Registration (SPA)
  // Redirect URI: https://vercel.com/martins-projects-96b96d99/cockpit-central/4Sg5VYkdjuSpdMgyguKbrZb3GrfM
  tenantId: "bd2452f8-ef2b-4ddf-a972-a0bcc233768a",
  clientId: "edd4c435-217e-4f81-9d7a-1a1e53fe9de3",
  redirectUri: "https://vercel.com/martins-projects-96b96d99/cockpit-central/4Sg5VYkdjuSpdMgyguKbrZb3GrfM,

  // SharePoint list backend (Microsoft Lists)
  // SiteId (Graph): hostname,siteCollectionId,siteId
  siteId: "bcsjoliette.sharepoint.com,fa75f19a-a192-4177-b975-4a2d34e7f816,141af805-b71f-49e8-9671-63dde1bec8e6",
  listId: "51764fdd-96c4-4ad4-bab2-9dfa045d13fa",

  // Quick links
  links: {
    // Ajuste ces URLs si tu veux pointer vers des apps spécifiques
    bienChezSoiApp: "https://bcsjoliette.sharepoint.com",
    evolumisApp: "https://bcsjoliette.sharepoint.com/sites/SolutionsEvolumis",
    persoApp: "https://outlook.office.com/todo/",

    // SharePoint folder (web URL)
    sharePointFolderUrl: "https://bcsjoliette.sharepoint.com/sites/SolutionsEvolumis/Shared%20Documents",

    // Optional: direct link to open the List in SharePoint
    listWebUrl: "https://bcsjoliette.sharepoint.com/sites/SolutionsEvolumis/Lists/Cockpit%20%20Tasks/AllItems.aspx",
  },

  // Pôles (UI)
  poles: [
    { key: "BCS", label: "Bien Chez Soi", emoji: "🏡" },
    { key: "EVO", label: "Evolumis", emoji: "🚀" },
    { key: "PERSO", label: "Personnel", emoji: "🧠" },
  ],

  // Statuts (doivent matcher les valeurs de ta liste)
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
