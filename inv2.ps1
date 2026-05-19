# ============================================
# INVENTARIO DETALLADO → TXT + Google Sheets
# ============================================

$ErrorActionPreference = "Continue"
$OutputDir = [Environment]::GetFolderPath("Desktop")
$fecha = Get-Date -Format "yyyy-MM-dd_HH-mm"
$computer = $env:COMPUTERNAME

$txtPath = Join-Path $OutputDir "Inventario_$computer_$fecha.txt"
$GoogleWebAppURL = "https://script.google.com/macros/s/AKfycbyUhL7A6RmeJViyuJC7PdSpuUx_Ifmf9WUwXVmOYH9YWqpY3qxd82sImcTYEmxxUzDr/exec"

function titulo($text) {
    $line = "=" * 75
    Add-Content -Path $txtPath -Value "`n$line" -Encoding UTF8
    Add-Content -Path $txtPath -Value " $text " -Encoding UTF8
    Add-Content -Path $txtPath -Value "$line" -Encoding UTF8
}

"Inventario para Cotización - Generado: $(Get-Date)" | Out-File -FilePath $txtPath -Encoding UTF8 -Force

# ====================== TXT DETALLADO (mantengo tu estilo) ======================
titulo "INFORMACION GENERAL"
try { Get-ComputerInfo | Format-List WindowsProductName, WindowsVersion, CsManufacturer, CsModel, CsTotalPhysicalMemory | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}

titulo "PROCESADOR"
try { Get-CimInstance Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors | Format-List | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}

titulo "MEMORIA RAM (Slots)"
try {
    $ramModules = Get-CimInstance Win32_PhysicalMemory
    foreach ($r in $ramModules) {
        $gb = [math]::Round($r.Capacity / 1GB, 2)
        Add-Content -Path $txtPath -Value "Slot $($r.DeviceLocator): $gb GB | $($r.Speed) MHz | $($r.Manufacturer) $($r.PartNumber)" -Encoding UTF8
    }
} catch {}

titulo "ALMACENAMIENTO"
try { Get-PhysicalDisk | Select-Object FriendlyName, MediaType, @{Name="SizeGB";E={[math]::Round($_.Size/1GB,2)}} | Format-Table -AutoSize | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}

titulo "GPU - PLACA BASE - FUENTE"
try { Get-CimInstance Win32_VideoController | Select-Object Name | Format-List | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}
try { Get-CimInstance Win32_BaseBoard | Format-List Manufacturer, Product, SerialNumber | Out-File -FilePath $txtPath -Append -Encoding UTF8 } catch {}

titulo "MONITORES"
try {
    Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID | ForEach-Object {
        $name = ($_.UserFriendlyName | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
        $serial = ($_.SerialNumberID | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
        Add-Content -Path $txtPath -Value "Monitor: $name | Serial: $serial" -Encoding UTF8
    }
} catch {}

# ====================== DATOS PARA GOOGLE SHEETS (MÁS DETALLADO) ======================
$data = @{
    noID              = $computer
    tipoEquipo        = "Desktop"
    marca             = ""
    modelo            = ""
    serial            = ""
    so                = ""
    usuario           = $env:USERNAME
    ubicacion         = ""
    estado            = "Operativo"
    observaciones     = ""
    procesador        = ""
    ram               = ""
    ramSlots          = ""
    almacenamiento    = ""
    gpu               = ""
    placaBase         = ""
    fuentePoder       = ""
    monitor           = ""
    monitorSerial     = ""
    perifericos       = ""
    licenciaWindows   = ""
    licenciaOffice    = ""
}

# Rellenado automático mejorado
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
    $totalRAM = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 0)
    $data.ram = "$totalRAM GB"
    $data.ramSlots = (Get-CimInstance Win32_PhysicalMemory).Count
} catch {}

try {
    $disk = Get-PhysicalDisk | Where-Object BusType -ne "USB" | Select-Object -First 1
    $data.almacenamiento = "$($disk.MediaType) $([math]::Round($disk.Size/1GB,0)) GB"
} catch {}

try {
    $gpu = Get-CimInstance Win32_VideoController | Select-Object -First 1
    $data.gpu = $gpu.Name
} catch {}

try {
    $mb = Get-CimInstance Win32_BaseBoard
    $data.placaBase = "$($mb.Manufacturer) $($mb.Product) SN:$($mb.SerialNumber)"
} catch {}

# Monitores detallados
try {
    $monData = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID | ForEach-Object {
        $name = ($_.UserFriendlyName | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
        $serial = ($_.SerialNumberID | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
        "$name ($serial)"
    }
    $data.monitor = ($monData | Where-Object {$_}) -join " | "
} catch {}

# Licencias
try {
    $winLic = Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.Name -like "*Windows*" -and $_.PartialProductKey }
    if ($winLic) { $data.licenciaWindows = "$($winLic.Name) - $($winLic.LicenseStatus)" }
} catch {}

try {
    $officeLic = Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.Name -like "*Office*" -and $_.PartialProductKey }
    if ($officeLic) { $data.licenciaOffice = $officeLic.Name }
} catch {}

# ====================== ENVIAR A GOOGLE SHEETS ======================
try {
    $json = $data | ConvertTo-Json
    Invoke-WebRequest -Uri $GoogleWebAppURL -Method Post -Body $json -ContentType "application/json" -TimeoutSec 30 | Out-Null
    Write-Host "✅ Enviado correctamente a Google Sheets" -ForegroundColor Green
} catch {
    Write-Host "⚠️ No se pudo enviar a Google Sheets" -ForegroundColor Yellow
}

Write-Host "=======================================" -ForegroundColor Green
Write-Host " INVENTARIO GENERADO CORRECTAMENTE" -ForegroundColor Green
Write-Host "TXT en Escritorio + Google Sheets actualizado" -ForegroundColor Cyan
