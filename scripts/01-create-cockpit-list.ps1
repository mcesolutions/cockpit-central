<#
Creates the "Cockpit - Tasks" Microsoft List (SharePoint List) + required columns.

PREREQ:
  - PowerShell 7+ recommended
  - Microsoft.Graph PowerShell SDK

RUN:
  1) Open PowerShell
  2) cd to this folder
  3) .\01-create-cockpit-list.ps1 -SiteWebUrl "https://.../sites/YourSite" -ListName "Cockpit - Tasks"

OUTPUT:
  - SiteId, ListId, List webUrl
  - Copy/paste snippet for config.js
#>

param(
  [Parameter(Mandatory=$true)]
  [string]$SiteWebUrl,

  [Parameter(Mandatory=$false)]
  [string]$ListName = "Cockpit - Tasks"
)

$ErrorActionPreference = "Stop"

Write-Host "\n[1/4] Installing/Importing Microsoft.Graph..." -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
  Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Sites

Write-Host "\n[2/4] Connecting to Microsoft Graph (delegated)..." -ForegroundColor Cyan
# You can reduce privilege later (Sites.Selected). This is fastest for initial setup.
Connect-MgGraph -Scopes "User.Read","Sites.ReadWrite.All" -UseDeviceAuthentication | Out-Null

# Parse site URL -> hostname + server relative path
$u = [System.Uri]$SiteWebUrl
$spHost = $u.Host
$path = $u.AbsolutePath.TrimEnd('/')

Write-Host "\n[3/4] Resolving SiteId for $SiteWebUrl (Graph sites/hostname:/path) ..." -ForegroundColor Cyan
$site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/${spHost}:$path"
$siteId = $site.id
if (-not $siteId) { throw "Unable to resolve SiteId. Check SiteWebUrl." }

Write-Host "SiteId: $siteId" -ForegroundColor Green

Write-Host "\n[4/4] Creating List + Columns..." -ForegroundColor Cyan

# Create list
$createListBody = @{
  displayName = $ListName
  list = @{ template = "genericList" }
} | ConvertTo-Json -Depth 10

$list = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists" -Body $createListBody -ContentType "application/json"
$listId = $list.id
$listWebUrl = $list.webUrl

Write-Host "ListId: $listId" -ForegroundColor Green
Write-Host "List webUrl: $listWebUrl" -ForegroundColor Green

function Add-Column($params) {
  $json = ($params | ConvertTo-Json -Depth 20)
  Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/columns" -Body $json -ContentType "application/json" | Out-Null
}

# Pole (choice)
Add-Column @{
  name = "Pole"
  displayName = "Pôle"
  choice = @{
    allowTextEntry = $false
    choices = @("BCS","EVO","PERSO")
  }
}

# Status (choice)
Add-Column @{
  name = "Status"
  displayName = "Statut"
  choice = @{
    allowTextEntry = $false
    choices = @("Backlog","EnCours","EnAttente","Termine")
  }
}

# Priority (choice)
Add-Column @{
  name = "Priority"
  displayName = "Priorité"
  choice = @{
    allowTextEntry = $false
    choices = @("P1","P2","P3")
  }
}

# DueDate (dateTime)
Add-Column @{
  name = "DueDate"
  displayName = "Échéance"
  dateTime = @{ format = "dateOnly" }
}

# Notes (multiline text)
Add-Column @{
  name = "Notes"
  displayName = "Notes"
  text = @{ allowMultipleLines = $true; linesForEditing = 8 }
}

# LinkUrl (hyperlink)
Add-Column @{
  name = "LinkUrl"
  displayName = "Lien"
  hyperlinkOrPicture = @{ isPicture = $false }
}

# SortOrder (number)
Add-Column @{
  name = "SortOrder"
  displayName = "Ordre"
  number = @{ decimalPlaces = "none" }
}

Write-Host "\nDone." -ForegroundColor Green
Write-Host "\n=== COPY INTO config.js ===" -ForegroundColor Yellow
Write-Host "tenantId: \"YOUR_TENANT_ID_OR_DOMAIN\","
Write-Host "clientId: \"YOUR_APP_CLIENT_ID\","
Write-Host "redirectUri: \"http://localhost:5173\","
Write-Host "siteId: \"$siteId\","
Write-Host "listId: \"$listId\","
Write-Host "links: { listWebUrl: \"$listWebUrl\" }," 

