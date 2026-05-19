# ============================================
# INVENTARIO COMPLETO → TXT + Google Sheets
# ============================================

$ErrorActionPreference = "Continue"
$OutputDir = [Environment]::GetFolderPath("Desktop")
$fecha = Get-Date -Format "yyyy-MM-dd_HH-mm"
$computer = $env:COMPUTERNAME

$txtPath = Join-Path $OutputDir "Inventario_$computer_$fecha.txt"

# ================== URL DE GOOGLE SHEETS ==================
$GoogleWebAppURL = "https://script.google.com/macros/s/AKfycbyUhL7A6RmeJViyuJC7PdSpuUx_Ifmf9WUwXVmOYH9YWqpY3qxd82sImcTYEmxxUzDr/exec"

# ====================== FUNCIÓN TÍTULO ======================
function titulo($text) {
    $line = "=" * 75
    Add-Content -Path $txtPath -Value "`n$line" -Encoding UTF8
    Add-Content -Path $txtPath -Value " $text " -Encoding UTF8
    Add-Content -Path $txtPath -Value "$line" -Encoding UTF8
}

# ====================== INICIO DEL TXT ======================
"Inventario para Cotización - Generado: $(Get-Date)" | Out-File -FilePath $txtPath -Encoding UTF8 -Force

titulo "INFORMACION GENERAL"
try {
    Get-ComputerInfo | Format-List WindowsProductName, WindowsVersion, WindowsBuildLabEx, OsArchitecture,
    CsManufacturer, CsModel, CsTotalPhysicalMemory, OsUptime | Out-File -FilePath $txtPath -Append -Encoding UTF8
} catch {
    Add-Content -Path $txtPath -Value "No se pudo obtener informacion general" -Encoding UTF8
}

titulo "PROCESADOR (CPU)"
try {
    Get-CimInstance Win32_Processor | Select-Object Name, Manufacturer, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors |
    Format-List | Out-File -FilePath $txtPath -Append -Encoding UTF8
} catch {}

titulo "MEMORIA RAM"
try {
    $ram = Get-CimInstance Win32_PhysicalMemory
    foreach ($r in $ram) {
        $gb = [math]::Round($r.Capacity / 1GB, 2)
        Add-Content -Path $txtPath -Value "Slot: $($r.DeviceLocator) | Capacidad: $gb GB | Velocidad: $($r.Speed) MHz | Parte: $($r.PartNumber)" -Encoding UTF8
    }
    $total = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
    Add-Content -Path $txtPath -Value "TOTAL RAM: $total GB" -Encoding UTF8
} catch {}

titulo "ALMACENAMIENTO"
try {
    Get-PhysicalDisk | Select-Object FriendlyName, MediaType, @{Name="SizeGB";Expression={[math]::Round($_.Size/1GB,2)}}, BusType |
    Format-Table -AutoSize | Out-File -FilePath $txtPath -Append -Encoding UTF8
} catch {}

titulo "GPU"
try {
    Get-CimInstance Win32_VideoController | Select-Object Name, @{Name="RAM_GB";E={[math]::Round($_.AdapterRAM/1GB,2)}} |
    Format-List | Out-File -FilePath $txtPath -Append -Encoding UTF8
} catch {}

titulo "PLACA BASE"
try {
    Get-CimInstance Win32_BaseBoard | Format-List Manufacturer, Product, SerialNumber | Out-File -FilePath $txtPath -Append -Encoding UTF8
} catch {}

titulo "MONITORES"
try {
    Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID | ForEach-Object {
        $name = ($_.UserFriendlyName | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
        Add-Content -Path $txtPath -Value "Monitor: $name" -Encoding UTF8
    }
} catch {}

titulo "INFORMACION FINAL"
Add-Content -Path $txtPath -Value "Equipo     : $computer" -Encoding UTF8
Add-Content -Path $txtPath -Value "Usuario    : $env:USERNAME" -Encoding UTF8
Add-Content -Path $txtPath -Value "Fecha      : $(Get-Date)" -Encoding UTF8

# ====================== DATOS PARA GOOGLE SHEETS ======================
$data = @{
    computerID       = $computer
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
    ram              = ""
    almacenamiento   = ""
    gpu              = ""
    placaBase        = ""
    monitor          = ""
    perifericos      = ""
    estadoPerifericos = ""
}

# Rellenar datos
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
} catch {}

try {
    $disk = Get-PhysicalDisk | Where-Object BusType -ne "USB" | Select-Object -First 1
    $size = [math]::Round($disk.Size/1GB, 0)
    $data.almacenamiento = "$($disk.MediaType) $size GB"
} catch {}

try {
    $gpu = Get-CimInstance Win32_VideoController | Select-Object -First 1
    $data.gpu = $gpu.Name
} catch {}

try {
    $mb = Get-CimInstance Win32_BaseBoard
    $data.placaBase = "$($mb.Manufacturer) $($mb.Product)"
} catch {}

# ====================== ENVIAR A GOOGLE SHEETS ======================
try {
    $json = $data | ConvertTo-Json
    Invoke-WebRequest -Uri $GoogleWebAppURL -Method Post -Body $json -ContentType "application/json" -TimeoutSec 20 | Out-Null
    Write-Host "✅ Datos enviados a Google Sheets" -ForegroundColor Green
} 
catch {
    Write-Host "⚠️  No se pudo enviar a Google Sheets (sin internet o error)" -ForegroundColor Yellow
}

# ====================== MENSAJE FINAL ======================
Write-Host "=======================================" -ForegroundColor Green
Write-Host " INVENTARIO GENERADO CORRECTAMENTE" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green
Write-Host "Archivo TXT guardado en Escritorio" -ForegroundColor Cyan
Write-Host "Google Sheets actualizado" -ForegroundColor Magenta