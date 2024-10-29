# UTP + Info Monitoreo
# Luis Giraldez
Add-Type -AssemblyName UIAutomationClient
$logFile = Join-Path $env:TEMP "utp_monitor_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Level - $Message"
    Add-Content -Path $logFile -Value $logMessage
    
    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        default { Write-Host $logMessage }
    }
}

function Get-BrowserTabs {
    try {
        $automation = [System.Windows.Automation.AutomationElement]::RootElement
        $allTabs = @()
        
        $browserProcesses = @{
            'chrome' = 'Google Chrome'
            'firefox' = 'Mozilla Firefox'
            'opera' = 'Opera'
            'msedge' = 'Microsoft Edge'
            'brave' = 'Brave'
        }
        
        foreach ($browserName in $browserProcesses.Keys) {
            $processes = Get-Process -Name $browserName -ErrorAction SilentlyContinue
            
            foreach ($process in $processes) {
                try {
                    $condition = New-Object System.Windows.Automation.PropertyCondition(
                        [System.Windows.Automation.AutomationElement]::ProcessIdProperty, 
                        $process.Id
                    )
                    
                    $window = $automation.FindFirst(
                        [System.Windows.Automation.TreeScope]::Children, 
                        $condition
                    )
                    
                    if ($window) {
                        $tabCondition = New-Object System.Windows.Automation.PropertyCondition(
                            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                            [System.Windows.Automation.ControlType]::TabItem
                        )
                        
                        $tabs = $window.FindAll(
                            [System.Windows.Automation.TreeScope]::Descendants, 
                            $tabCondition
                        )
                        
                        foreach ($tab in $tabs) {
                            $title = $tab.Current.Name
                            if ($title) {
                                $allTabs += @{
                                    'Browser' = $browserProcesses[$browserName]
                                    'Title' = $title
                                    'ProcessName' = $browserName
                                    'ProcessId' = $process.Id
                                    'TabElement' = $tab
                                }
                            }
                        }
                        
                        # Fallback to main window title if no tabs found
                        if ($tabs.Count -eq 0 -and $process.MainWindowTitle) {
                            $allTabs += @{
                                'Browser' = $browserProcesses[$browserName]
                                'Title' = $process.MainWindowTitle
                                'ProcessName' = $browserName
                                'ProcessId' = $process.Id
                                'TabElement' = $null
                            }
                        }
                    }
                }
                catch {
                    Write-Log "Error al procesar ventana de navegador: $_" -Level "ERROR"
                    # Fallback method
                    if ($process.MainWindowTitle) {
                        $allTabs += @{
                            'Browser' = $browserProcesses[$browserName]
                            'Title' = $process.MainWindowTitle
                            'ProcessName' = $browserName
                            'ProcessId' = $process.Id
                            'TabElement' = $null
                        }
                    }
                    continue
                }
            }
        }
        
        return $allTabs
    }
    catch {
        Write-Log "Error al obtener pestañas: $_" -Level "ERROR"
        return @()
    }
}

function Get-TabSelectionPattern {
    param (
        [System.Windows.Automation.AutomationElement]$element
    )
    
    try {
        return [System.Windows.Automation.SelectionItemPattern]::Pattern
    }
    catch {
        return $null
    }
}

function Test-UTPOpen {
    $utpPatterns = @(
        'Info',
        'info\.utp\.edu\.pe'
        #'Universidad Tecnológica del Perú',
        #'Intranet UTP',
        #'Canvas UTP',
        #'Portal UTP'
    )
    
    $tabs = Get-BrowserTabs
    
    $utpTabs = $tabs | Where-Object { 
        $title = $_.Title
        $utpPatterns | Where-Object { $title -match $_ }
    }
    
    if ($utpTabs.Count -gt 0) {
        foreach ($tab in $utpTabs) {
            $isOpen = $true
            
            # Solo registramos que existe la pestaña, sin importar si está activa o no
            Write-Log "UTP+ Info detectado en $($tab.Browser): $($tab.Title)" -Level "INFO"
        }
        
        return @{
            'IsOpen' = $true
            'TabCount' = $utpTabs.Count
        }
    }
    
    Write-Log "No se detectó UTP+ Info en ninguna pestaña" -Level "INFO"
    return @{
        'IsOpen' = $false
        'TabCount' = 0
    }
}

function Start-UTPMonitor {
    param(
        [string]$url = "https://info.utp.edu.pe/",
        [int]$checkInterval = 3
    )
    
    $host.UI.RawUI.WindowTitle = "UTP+ Info Monitor"
    Write-Log "Iniciando monitoreo del portal UTP+ Info..."
    Write-Log "Monitorizando: Chrome, Firefox, Opera, Edge, Brave"
    
    $lastOpened = 0
    $cooldownPeriod = 10
    $dotCount = 0
    
    while ($true) {
        try {
            $currentTime = [int](Get-Date -UFormat %s)
            $utpStatus = Test-UTPOpen
            
            if (-not $utpStatus.IsOpen -and ($currentTime - $lastOpened -ge $cooldownPeriod)) {
                Start-Sleep -Milliseconds 500
                $utpStatus = Test-UTPOpen
                
                if (-not $utpStatus.IsOpen) {
                    Write-Host "`n"
                    Write-Log "UTP+ Info no detectado. Reabriendo..." -Level "WARNING"
                    
                    try {
                        Start-Process $url
                        $lastOpened = $currentTime
                        Write-Log "UTP+ Info reabierto exitosamente"
                    } catch {
                        Write-Log "Error al abrir UTP Portal: $_" -Level "ERROR"
                        
                        # Método alternativo de apertura
                        try {
                            $ie = New-Object -ComObject InternetExplorer.Application
                            $ie.Visible = $true
                            $ie.Navigate($url)
                            $lastOpened = $currentTime
                            Write-Log "UTP+ Info reabierto usando método alternativo"
                        } catch {
                            Write-Log "Error en método alternativo: $_" -Level "ERROR"
                        }
                    }
                }
            }
            
            $dotCount = ($dotCount + 1) % 4
            Write-Host "`rMonitoreando $('.' * $dotCount)$(' ' * (3 - $dotCount))" -NoNewline
            
            Start-Sleep -Seconds $checkInterval
            
        } catch {
            Write-Log "Error en el monitoreo: $_" -Level "ERROR"
            Start-Sleep -Seconds $checkInterval
        }
    }
}

function Main {
    try {
        Clear-Host
        Write-Host @"
===================================
     UTP + Info Monitor v2.5
===================================
"@ -ForegroundColor Cyan
        
        Write-Log "=== Iniciando UTP + Info Monitor v3.0 ==="
        Start-UTPMonitor
        
    } catch {
        if ($_.Exception.Message -match "OperationStopped") {
            Write-Log "Monitoreo detenido por el usuario" -Level "WARNING"
        } else {
            Write-Log "Error inesperado: $_" -Level "ERROR"
        }
    } finally {
        Write-Log "=== Monitoreo finalizado ==="
        Write-Host "`nPresiona Enter para cerrar..."
        $null = Read-Host
    }
}

# Iniciar el programa
Main