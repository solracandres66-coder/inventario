# ============================================
# INVENTARIO DETALLADO → TXT + Google Sheets
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

# ====================== TXT (igual que antes) ======================
titulo "INFORMACION GENERAL"
try { Get-ComputerInfo | Format-List WindowsProductName, WindowsVersion, CsManufacturer, CsModel, CsTotalPhysicalMemory | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}

titulo "PROCESADOR"
try { Get-CimInstance Win32_Processor | Select-Object Name, Manufacturer, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors | Format-List | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}

titulo "MEMORIA RAM (Detalle por Slot)"
$ramDetail = ""
try {
    $ramModules = Get-CimInstance Win32_PhysicalMemory
    foreach ($r in $ramModules) {
        $gb = [math]::Round($r.Capacity / 1GB, 2)
        $line = "Slot $($r.DeviceLocator): $gb GB | $($r.Speed) MHz | $($r.Manufacturer) $($r.PartNumber)"
        Add-Content -Path $txtPath -Value $line -Encoding UTF8
        $ramDetail += "$line`n"
    }
} catch {}

titulo "ALMACENAMIENTO"
try { Get-PhysicalDisk | Select-Object FriendlyName, MediaType, @{Name="SizeGB";E={[math]::Round($_.Size/1GB,2)}} | Format-Table -AutoSize | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}

titulo "GPU"
try { Get-CimInstance Win32_VideoController | Select-Object Name, @{Name="VRAM_GB";E={[math]::Round($_.AdapterRAM/1GB,2)}} | Format-List | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}

titulo "PLACA BASE"
try { Get-CimInstance Win32_BaseBoard | Format-List Manufacturer, Product, SerialNumber | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}

titulo "MONITORES"
try {
    Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID | ForEach-Object {
        $name = ($_.UserFriendlyName | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
        $serial = ($_.SerialNumberID | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
        Add-Content -Path $txtPath -Value "Monitor: $name | Serial: $serial" -Encoding UTF8
    }
} catch {}

# ====================== DATOS PARA GOOGLE SHEETS ======================
$data = @{
    computerID       = $computer
    equipoAsociadoID = ""
    tipoEquipo       = "Desktop"
    marca            = ""
    modelo           = ""
    serial           = ""
    so               = ""
    usuario          = $env:USERNAME
    ubicacion        = ""
    estado           = "Operativo"
    observaciones    = ""
    procesador       = ""
    ram              = $ramDetail.Trim()          # ← Aquí va el detalle completo por slot
    almacenamiento   = ""
    gpu              = ""
    placaBase        = ""
    fuentePoder      = ""
    monitor          = ""
    monitorSerial    = ""
    perifericos      = ""
    licenciaWindows  = ""
    licenciaOffice   = ""
}

# Rellenado
try {
    $ci = Get-ComputerInfo
    $data.so     = $ci.WindowsProductName
    $data.marca  = $ci.CsManufacturer
    $data.modelo = $ci.CsModel
} catch {}

try {
    $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
    $data.procesador = $cpu.Name
} catch {}

try {
    $disks = Get-PhysicalDisk | Where-Object BusType -ne "USB"
    $data.almacenamiento = ($disks | ForEach-Object { "$($_.MediaType) $([math]::Round($_.Size/1GB,0))GB" }) -join " + "
} catch {}

try {
    $gpus = Get-CimInstance Win32_VideoController
    $data.gpu = ($gpus.Name) -join " + "
} catch {}

try {
    $mb = Get-CimInstance Win32_BaseBoard
    $data.placaBase = "$($mb.Manufacturer) $($mb.Product)"
} catch {}

try {
    $mons = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID | ForEach-Object {
        $name = ($_.UserFriendlyName | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
        $serial = ($_.SerialNumberID | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
        "$name (SN:$serial)"
    }
    $data.monitor = $mons -join " | "
} catch {}

# Licencias
try {
    $win = Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.Name -like "*Windows*" -and $_.PartialProductKey } | Select-Object -First 1
    if ($win) { $data.licenciaWindows = $win.Name }
} catch {}

try {
    $off = Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.Name -like "*Office*" -and $_.PartialProductKey } | Select-Object -First 1
    if ($off) { $data.licenciaOffice = $off.Name }
} catch {}

# ====================== ENVIAR ======================
try {
    $json = $data | ConvertTo-Json
    Invoke-WebRequest -Uri $GoogleWebAppURL -Method Post -Body $json -ContentType "application/json" | Out-Null
    Write-Host "✅ Enviado correctamente a Google Sheets" -ForegroundColor Green
} catch {
    Write-Host "❌ Error enviando a Sheets" -ForegroundColor Red
}

Write-Host "=======================================" -ForegroundColor Green
Write-Host " INVENTARIO GENERADO CORRECTAMENTE" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green
