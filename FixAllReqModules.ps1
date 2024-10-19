$DesktopPath = [System.Environment]::GetFolderPath('Desktop')
$InstallAzModules = @("Az.Accounts","Az.KeyVault","Az.Resources")
$installEXOhModule = "ExchangeOnlineManagement"
$InstallTeamsModule = "MicrosoftTeams"
$InstallGraphPSModules = @("Microsoft.Graph.Authentication","Microsoft.Graph.Calendar","Microsoft.Graph.DirectoryObjects","Microsoft.Graph.Groups","Microsoft.Graph.Users","Microsoft.Graph.Users.Actions")
$installPnPModule = "PnP.PowerShell"
$ExchangeReqVersion = [version]"3.4.0"
$PnPReqVersion = [version]"1.12.0"
$GraphPSReqVersion = [version]"2.8.0"

#Teams Part
if ((!(Get-Module -ListAvailable $InstallTeamsModule)) -and (!(Get-InstalledModule $InstallTeamsModule))) {
    Write-Host "Installing $InstallTeamsModule..." -ForegroundColor Cyan
    Install-Module $InstallTeamsModule -AllowClobber -Force
}

#Exchange Online Part
$CheckEXOModules = Get-Module -ListAvailable $installEXOhModule
if (!$CheckEXOModules) {
    Install-Module $installEXOhModule -AllowClobber -Force
}
if ($CheckEXOModules) {
    foreach ($CheckEXOModule in $CheckEXOModules) {
        if ($CheckEXOModule.Version -lt $ExchangeReqVersion) {
            Write-Host "Removing old $($CheckEXOModule.name) module..." -ForegroundColor Cyan
            Uninstall-Module $CheckEXOModule.name -RequiredVersion $CheckEXOModule.Version -Force -ErrorAction SilentlyContinue
        }
    }
}
Start-Sleep -Seconds 5
Clear-Variable $CheckEXOModules -ErrorAction SilentlyContinue
$CheckEXOModules = Get-Module -ListAvailable $installEXOhModule | Sort-Object version -Descending
if ((!$CheckEXOModules) -or ($CheckEXOModules[0].Version -lt $ExchangeReqVersion)) {
    Write-Host "Installing $installEXOhModule..." -ForegroundColor Cyan
    Install-Module $installEXOhModule -AllowClobber -Force
}

#AZ Part
$installedAzModules = Get-Module -ListAvailable Az.*
$AzModulesToUninstall = @()
if ($installedAzModules) {
    foreach ($AzModule in $installedAzModules) {
        if ($InstallAzModules -notcontains $AzModule.Name) {
            $AzModulesToUninstall += $AzModule.Name
        }
    }
    foreach ($InstallAzModule in $InstallAzModules) {
        if (-not (Get-Module -ListAvailable -Name $InstallAzModule)) {
            Install-Module -Name $InstallAzModule -Force
            Write-Output "$InstallAzModule has been installed."
        } else {
            Write-Output "$InstallAzModule is already installed."
        }
    }
    if ($AzModulesToUninstall.Count -gt 0) {
        foreach ($AzModuleToUninstall in $AzModulesToUninstall) {
            Uninstall-Module -Name $AzModuleToUninstall -Force
            Write-Output "$AzModuleToUninstall has been uninstalled."
        }
    } else {
        Write-Output "No other Az modules to uninstall."
    }
} else {
    foreach ($InstallAzModule in $InstallAzModules) {
        Install-Module -Name $InstallAzModule -Force
        Write-Output "$InstallAzModule has been installed."
    }
}

#PnP Part
$checkPnPModule = Get-Module -ListAvailable $installPnPModule | Sort-Object Version -Descending
$PnPVersionExists = $false
$PnPVersionsToUninstall = @()

if ($checkPnPModule) {
    foreach ($PnPModule in $checkPnPModule) {
        if ($PnPModule.Version -eq $PnPReqVersion) {
            $PnPVersionExists = $true
        } else {
            $PnPVersionsToUninstall += $PnPModule.Version
        }
    }

    if (-not $PnPVersionExists) {
        Install-Module -Name $installPnPModule -RequiredVersion $PnPReqVersion -AllowClobber -Force
        Write-Output "$installPnPModule version $PnPReqVersion has been installed."
    } else {
        Write-Output "$installPnPModule version $PnPReqVersion is already installed."
    }

    if ($PnPVersionsToUninstall.Count -gt 0) {
        foreach ($PnPVersion in $PnPVersionsToUninstall) {
            Uninstall-Module -Name $installPnPModule -RequiredVersion $PnPVersion -Force
            Write-Output "$installPnPModule version $PnPVersion has been uninstalled."
        }
    } else {
        Write-Output "No other versions to uninstall."
    }

} else {
    Install-Module -Name $installPnPModulel -RequiredVersion $PnPReqVersion -AllowClobber -Force
    Write-Output "$installPnPModule version $PnPReqVersion has been installed."
}

#Graph Part
# List of required Microsoft Graph modules
$GraphPSReqModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Calendar", "Microsoft.Graph.DirectoryObjects", "Microsoft.Graph.Groups", "Microsoft.Graph.Users", "Microsoft.Graph.Users.Actions")

# Check for installed Microsoft Graph modules
$GraphPSModules = Get-Module -ListAvailable Microsoft.Grap* | Sort-Object Version -Descending
$GraphPSReqVersion = [version]"2.8.0"

if (!$GraphPSModules) {
    foreach ($GraphPSReqModule in $GraphPSReqModules) {
    Write-Host "installing $GraphPSReqModule..." -ForegroundColor Cyan
    Install-Module $GraphPSReqModule -Force
    }
}

if ($GraphPSModules) {

    foreach ($GraphPSModule in $GraphPSModules) {
    if (Get-Module -ListAvailable $GraphPSModule.name | Where-Object version -eq $GraphPSModule.version | Select-Object -ExpandProperty RequiredModules | where {$_.Name -ne "Microsoft.Graph.Authentication"}) {
        Write-Host "Removing Dependednt Module $($GraphPSModule.Name) $($GraphPSModule.Version)" -ForegroundColor Cyan
        Uninstall-Module $GraphPSModule.Name -RequiredVersion $GraphPSModule.Version -Force
        }
    }

    foreach ($GraphPSModule in $GraphPSModules) {
        if ($GraphPSModule -notin $GraphPSReqModules) {
            Write-Host "Removing not required module $($GraphPSModule.name) version $($GraphPSModule.version)" -ForegroundColor Cyan
            Uninstall-Module $GraphPSModule.name -RequiredVersion $GraphPSModule.version -Force
            }
          }    
    }

if ($GraphPSModules[0].Version -lt $GraphPSReqVersion) {
    foreach ($GraphPSReqModule in $GraphPSReqModules) {
    {
    Write-Host "Installing missing module $($GraphPSReqModule)" -ForegroundColor Cyan
    Install-Module $GraphPSReqModule -Force
      }
    }
}
if ($GraphPSModules[0].Version -ge $GraphPSReqVersion) {
    foreach ($GraphPSReqModule in $GraphPSReqModules)
        {
        $InstalledGraphPSReqModule = Get-Module -ListAvailable $GraphPSReqModule | Sort-Object Version -Descending | select -First 1
        if (($InstalledGraphPSReqModule.name -notin $GraphPSReqModules) -or ($InstalledGraphPSReqModule.version -lt $GraphPSModules[0].Version)) {
        Write-Host "Installing missing module $($GraphPSReqModule) version $($GraphPSModules[0].Version)" -ForegroundColor Cyan
        Install-Module $GraphPSReqModule -RequiredVersion $GraphPSModules[0].Version -Force
        }
        
      }

    }

$GraphPDuplicateModules = Get-Module -ListAvailable Microsoft.Graph.* | Sort-Object Version -Descending | Group-Object Name | Where-Object {$_.Count -gt 1} | Select-Object -ExpandProperty name
if ($GraphPDuplicateModules -gt 0) {
    Write-Host "Below Modules have multiple versions installed" -ForegroundColor Yellow
    $GraphPDuplicateModules | fl
    foreach ($GraphPDuplicateModule in $GraphPDuplicateModules) {
        $DuplicateModules = Get-Module -ListAvailable $GraphPDuplicateModule | Sort-Object version -Descending | select -Skip 1
        foreach ($DuplicateModule in $DuplicateModules) {
        Write-Host "removing $($DuplicateModule.name) Version $($DuplicateModule.Version)"
        Uninstall-Module $DuplicateModule.name -RequiredVersion $DuplicateModule.version -Force
            }
        }
    }
