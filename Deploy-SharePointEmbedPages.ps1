<#
.SYNOPSIS
    Creates SharePoint pages with embedded GitHub Pages dashboards.
.DESCRIPTION
    Creates SharePoint modern pages with Embed web parts pointing to
    GitHub Pages hosted dashboards. This provides full custom UX while
    keeping SharePoint as the user-facing portal.

    NO Power Automate flows created/enabled
    NO emails or Teams messages sent
    Just creates SharePoint pages with iframe embeds
.PARAMETER SiteUrl
    SharePoint site URL
.PARAMETER GitHubPagesUrl
    Base URL for GitHub Pages (e.g., https://username.github.io/repo)
.EXAMPLE
    .\Deploy-SharePointEmbedPages.ps1 -GitHubPagesUrl "https://jevenson76.github.io/atl-dashboards"
#>

param(
    [string]$SiteUrl = "https://chamberlaingroup.sharepoint.com/sites/PrincipalGTMStrategy-InternalUseOnly-ATLIntegrationProject",
    [string]$GitHubPagesUrl = "https://jevenson76.github.io/atl-dashboards"
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " ATL Integration Hub - Embed Pages Deploy" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "SharePoint: $SiteUrl" -ForegroundColor Gray
Write-Host "GitHub Pages: $GitHubPagesUrl" -ForegroundColor Gray
Write-Host ""

# Page definitions mapping to GitHub Pages URLs
$Pages = @(
    @{ Name = "ATL-Dashboard"; Title = "ATL Integration Hub"; File = "index.html"; IsHome = $true },
    @{ Name = "ATL-Status-Hub"; Title = "Status Hub"; File = "ATL_Status_Hub.html" },
    @{ Name = "ATL-Project-Gantt"; Title = "Project Gantt"; File = "ATL_Project_Gantt.html" },
    @{ Name = "ATL-Milestone-Timeline"; Title = "Milestone Timeline"; File = "ATL_Milestone_Timeline.html" },
    @{ Name = "ATL-Task-Detail"; Title = "Task Detail"; File = "ATL_Task_Detail.html" },
    @{ Name = "ATL-My-Tasks"; Title = "My Tasks"; File = "ATL_My_Tasks.html" },
    @{ Name = "ATL-Team-Workload"; Title = "Team Workload"; File = "ATL_Team_Workload.html" },
    @{ Name = "ATL-Budget-Tracker"; Title = "Budget Tracker"; File = "ATL_Budget_Tracker.html" },
    @{ Name = "ATL-Blocker-Dashboard"; Title = "Blocker Dashboard"; File = "ATL_Blocker_Dashboard.html" },
    @{ Name = "ATL-Reports"; Title = "Reports"; File = "ATL_Reports.html" },
    @{ Name = "ATL-Admin"; Title = "Admin"; File = "ATL_Admin.html" }
)

# Connect to SharePoint
Write-Host "Connecting to SharePoint..." -ForegroundColor Yellow
$env:PNPLEGACYMESSAGE = 'false'
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue
Connect-PnPOnline -Url $SiteUrl -UseWebLogin
Write-Host "Connected!" -ForegroundColor Green
Write-Host ""

# Create each page
foreach ($page in $Pages) {
    Write-Host "Creating: $($page.Title)..." -NoNewline

    # Build iframe embed HTML
    $embedUrl = "$GitHubPagesUrl/$($page.File)"
    $iframeHtml = @"
<iframe src="$embedUrl" width="100%" height="900" frameborder="0" style="border: none; overflow: hidden;"></iframe>
"@

    # Remove existing page if present
    $existing = Get-PnPClientSidePage -Identity $page.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Remove-PnPClientSidePage -Identity $page.Name -Force
    }

    # Create new page
    try {
        $newPage = Add-PnPClientSidePage -Name $page.Name -LayoutType Article -PromoteAs None

        # Add embed web part with full-width iframe
        Add-PnPClientSidePageSection -Page $page.Name -SectionTemplate OneColumnFullWidth -Order 1

        # Add the embed web part
        Add-PnPClientSideWebPart -Page $page.Name -DefaultWebPartType ContentEmbed -Section 1 -Column 1 -WebPartProperties @{
            embedCode = $iframeHtml
        }

        # Publish the page
        Set-PnPClientSidePage -Identity $page.Name -Publish

        Write-Host " CREATED" -ForegroundColor Green
    } catch {
        Write-Host " ERROR: $_" -ForegroundColor Red
    }
}

Write-Host ""

# Set home page
Write-Host "Setting ATL-Dashboard as site home page..." -ForegroundColor Yellow
try {
    Set-PnPHomePage -RootFolderRelativeUrl "SitePages/ATL-Dashboard.aspx"
    Write-Host "Home page set!" -ForegroundColor Green
} catch {
    Write-Host "Could not set home page: $_" -ForegroundColor Yellow
}

Write-Host ""

# Configure navigation
Write-Host "Configuring site navigation..." -ForegroundColor Yellow

$navItems = @(
    @{ Title = "Home"; Url = "/sites/PrincipalGTMStrategy-InternalUseOnly-ATLIntegrationProject/SitePages/ATL-Dashboard.aspx" },
    @{ Title = "Status"; Url = "/sites/PrincipalGTMStrategy-InternalUseOnly-ATLIntegrationProject/SitePages/ATL-Status-Hub.aspx" },
    @{ Title = "Timeline"; Url = "/sites/PrincipalGTMStrategy-InternalUseOnly-ATLIntegrationProject/SitePages/ATL-Project-Gantt.aspx" },
    @{ Title = "Tasks"; Url = "/sites/PrincipalGTMStrategy-InternalUseOnly-ATLIntegrationProject/SitePages/ATL-Task-Detail.aspx" },
    @{ Title = "My Tasks"; Url = "/sites/PrincipalGTMStrategy-InternalUseOnly-ATLIntegrationProject/SitePages/ATL-My-Tasks.aspx" },
    @{ Title = "Team"; Url = "/sites/PrincipalGTMStrategy-InternalUseOnly-ATLIntegrationProject/SitePages/ATL-Team-Workload.aspx" },
    @{ Title = "Budget"; Url = "/sites/PrincipalGTMStrategy-InternalUseOnly-ATLIntegrationProject/SitePages/ATL-Budget-Tracker.aspx" },
    @{ Title = "Blockers"; Url = "/sites/PrincipalGTMStrategy-InternalUseOnly-ATLIntegrationProject/SitePages/ATL-Blocker-Dashboard.aspx" },
    @{ Title = "Reports"; Url = "/sites/PrincipalGTMStrategy-InternalUseOnly-ATLIntegrationProject/SitePages/ATL-Reports.aspx" }
)

$topNav = Get-PnPNavigationNode -Location TopNavigationBar

foreach ($item in $navItems) {
    $existing = $topNav | Where-Object { $_.Title -eq $item.Title }
    if (-not $existing) {
        Add-PnPNavigationNode -Location TopNavigationBar -Title $item.Title -Url $item.Url
        Write-Host "  Added: $($item.Title)" -ForegroundColor Gray
    } else {
        Write-Host "  Exists: $($item.Title)" -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host " DEPLOYMENT COMPLETE!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Dashboard URL:" -ForegroundColor Cyan
Write-Host "$SiteUrl/SitePages/ATL-Dashboard.aspx" -ForegroundColor White
Write-Host ""
Write-Host "GitHub Pages URL (direct):" -ForegroundColor Cyan
Write-Host "$GitHubPagesUrl" -ForegroundColor White
Write-Host ""
Write-Host "NEXT STEP: Create the Power Automate data proxy flow" -ForegroundColor Yellow
Write-Host "See SETUP.md for instructions" -ForegroundColor Gray
Write-Host ""

Disconnect-PnPOnline
