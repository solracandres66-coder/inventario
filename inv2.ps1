# ============================================
# INVENTARIO → TXT + Google Sheets (Auto Headers)
# ============================================

$ErrorActionPreference = "Continue"
$OutputDir = [Environment]::GetFolderPath("Desktop")
$fecha = Get-Date -Format "yyyy-MM-dd_HH-mm"
$computer = $env:COMPUTERNAME

$txtPath = Join-Path $OutputDir "Inventario_$computer_$fecha.txt"
$GoogleWebAppURL = "https://script.google.com/macros/s/AKfycbzTqMmna3JUlx4QEUhU8B-SopUzHkYUgYaP75uRZeSB649Y5C6pg_ZiB89pDKE-xnhQ/exec"

function titulo($text) {
    $line = "=" * 75
    Add-Content -Path $txtPath -Value "`n$line" -Encoding UTF8
    Add-Content -Path $txtPath -Value " $text " -Encoding UTF8
    Add-Content -Path $txtPath -Value "$line" -Encoding UTF8
}

"Inventario para Cotización - Generado: $(Get-Date)" | Out-File -FilePath $txtPath -Encoding UTF8 -Force

# ... (Mantengo tu TXT igual) ...
titulo "INFORMACION GENERAL"
try { Get-ComputerInfo | Format-List WindowsProductName, WindowsVersion, CsManufacturer, CsModel, CsTotalPhysicalMemory | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}

titulo "PROCESADOR"
try { Get-CimInstance Win32_Processor | Select-Object Name, Manufacturer, MaxClockSpeed, NumberOfCores | Format-List | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}

titulo "MEMORIA RAM"
try {
    $ramModules = Get-CimInstance Win32_PhysicalMemory
    foreach ($r in $ramModules) {
        $gb = [math]::Round($r.Capacity / 1GB, 2)
        Add-Content -Path $txtPath -Value "Slot $($r.DeviceLocator): $gb GB $($r.Speed)MHz $($r.PartNumber)" -Encoding UTF8
    }
} catch {}

titulo "ALMACENAMIENTO + GPU + PLACA"
try { Get-PhysicalDisk | Select FriendlyName, MediaType, @{Name="SizeGB";E={[math]::Round($_.Size/1GB,1)}} | Format-Table -AutoSize | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}
try { Get-CimInstance Win32_VideoController | Select Name | Format-List | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}
try { Get-CimInstance Win32_BaseBoard | Format-List Manufacturer, Product, SerialNumber | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}

titulo "MONITORES"
try {
    Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID | ForEach-Object {
        $n = ($_.UserFriendlyName | ? {$_ -ne 0} | % {[char]$_}) -join ""
        $s = ($_.SerialNumberID | ? {$_ -ne 0} | % {[char]$_}) -join ""
        Add-Content -Path $txtPath -Value "Monitor: $n | Serial: $s" -Encoding UTF8
    }
} catch {}

# ====================== DATOS PARA GOOGLE SHEETS ======================
$data = @{
    computerID       = $computer
    equipoAsociadoID = ""
    tipoEquipo       = "Desktop"
    marca            = (Get-ComputerInfo).CsManufacturer
    modelo           = (Get-ComputerInfo).CsModel
    serial           = ""
    so               = (Get-ComputerInfo).WindowsProductName
    usuario          = $env:USERNAME
    ubicacion        = ""
    estado           = "Operativo"
    observaciones    = ""
    procesador       = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name
    ram              = "$([math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 0)) GB"
    ramSlots         = (Get-CimInstance Win32_PhysicalMemory).Count
    almacenamiento   = (Get-PhysicalDisk | Where BusType -ne "USB" | ForEach-Object {"$($_.MediaType) $([math]::Round($_.Size/1GB,0))GB"}) -join " + "
    gpu              = (Get-CimInstance Win32_VideoController | ForEach-Object {$_.Name}) -join " + "
    placaBase        = "$((Get-CimInstance Win32_BaseBoard).Manufacturer) $((Get-CimInstance Win32_BaseBoard).Product)"
    fuentePoder      = ""
    monitor          = ""
    monitorSerial    = ""
    perifericos      = ""
    licenciaWindows  = ""
    licenciaOffice   = ""
}

# Licencias
try {
    $win = Get-CimInstance SoftwareLicensingProduct | Where-Object {$_.Name -like "*Windows*" -and $_.PartialProductKey} | Select-Object -First 1
    if ($win) { $data.licenciaWindows = $win.Name }
} catch {}

try {
    $off = Get-CimInstance SoftwareLicensingProduct | Where-Object {$_.Name -like "*Office*" -and $_.PartialProductKey} | Select-Object -First 1
    if ($off) { $data.licenciaOffice = $off.Name }
} catch {}

# Monitores
try {
    $mons = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID | ForEach-Object {
        $n = ($_.UserFriendlyName | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
        "$n"
    }
    $data.monitor = $mons -join " | "
} catch {}

# ====================== ENVIAR ======================
try {
    $json = $data | ConvertTo-Json
    Invoke-WebRequest -Uri $GoogleWebAppURL -Method Post -Body $json -ContentType "application/json" | Out-Null
    Write-Host "✅ Enviado correctamente a Google Sheets" -ForegroundColor Green
} catch {
    Write-Host "❌ Error al enviar: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "=======================================" -ForegroundColor Green
Write-Host "INVENTARIO GENERADO" -ForegroundColor Green
Write-Host "TXT + Google Sheets actualizado" -ForegroundColor Cyan
