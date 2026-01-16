# Cockpit Central

Un cockpit web (SPA) ultra simple:
- Accueil: 3 tuiles XXL (Bien Chez Soi / Evolumis / Personnel) + aperçu dynamique des tâches
- Pages pôle: Kanban, Calendrier, Table (édition inline + tri), Liste
- Backend: Microsoft Lists (SharePoint List) via Microsoft Graph
- Auth: Microsoft Entra ID (MSAL Browser)

## 1) Pré-requis
- Un site SharePoint (ex: https://contoso.sharepoint.com/sites/SolutionsEvolumis)
- Un compte M365 avec accès au site
- PowerShell (Windows) pour créer la liste automatiquement

## 2) App Registration (Entra ID)
1. Microsoft Entra admin center -> App registrations -> New registration
2. Type: Single-page application (SPA)
3. Redirect URI:
   - Dev local: http://localhost:5173
   - Prod: l'URL de ton hébergement
4. Permissions (Delegated):
   - User.Read
   - Sites.ReadWrite.All (rapide pour démarrer; tu pourras durcir ensuite)

## 3) Créer la List + colonnes (1 commande)
Dans PowerShell:
```powershell
cd .\scripts
.\01-create-cockpit-list.ps1 -SiteWebUrl "https://contoso.sharepoint.com/sites/SolutionsEvolumis" -ListName "Cockpit - Tasks"
```
Le script affiche: SiteId, ListId, et l'URL de la liste.

## 4) Configurer l'app
1. Ouvre `config.js`
2. Remplace:
   - tenantId
   - clientId
   - redirectUri
   - siteId
   - listId
   - liens (apps + dossier SharePoint)

## 5) Lancer en local
Depuis le dossier racine:
```powershell
python -m http.server 5173
```
Puis ouvre:
- http://localhost:5173

## 6) Déployer (gratuit)
Tu peux héberger ce dossier tel quel (statique) sur:
- GitHub Pages
- Cloudflare Pages
- Azure Static Web Apps
- Vercel (static)

## Modèle de données (colonnes)
- Title (texte)
- Pole (choix: BCS, EVO, PERSO)
- Status (choix: Backlog, EnCours, EnAttente, Termine)
- Priority (choix: P1, P2, P3)
- DueDate (date)
- Notes (texte multiligne)
- LinkUrl (hyperlien)
- SortOrder (nombre)

## Notes "Ouvrir dans l'Explorateur"
Les navigateurs ne peuvent pas forcer l'ouverture dans l'explorateur Windows. La méthode robuste est de synchroniser le dossier avec OneDrive, puis l'accès se fait via l'explorateur.
