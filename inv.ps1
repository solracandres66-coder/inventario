# ============================================
# INVENTARIO COMPLETO PARA COTIZACIÓN
# Ejecución remota desde GitHub:
#   iex (irm https://raw.githubusercontent.com/solracandres66-coder/inventario/main/inv.ps1)
# Descarga directa:
#   irm https://raw.githubusercontent.com/solracandres66-coder/inventario/main/inv.ps1 -o inv.ps1
# ============================================

$ErrorActionPreference = "Continue"
$OutputDir = [Environment]::GetFolderPath("Desktop")
$fecha = Get-Date -Format "yyyy-MM-dd_HH-mm"
$output = Join-Path $OutputDir "Inventario_$env:COMPUTERNAME_$fecha.txt"

function titulo($text) {
    $line = "=" * 75
    Add-Content -Path $output -Value "`n$line" -Encoding UTF8
    Add-Content -Path $output -Value " $text " -Encoding UTF8
    Add-Content -Path $output -Value "$line" -Encoding UTF8
}

"Inventario para Cotización - Generado: $(Get-Date)" | Out-File -FilePath $output -Encoding UTF8 -Force

titulo "INFORMACION GENERAL"
try {
    Get-ComputerInfo | Format-List WindowsProductName, WindowsVersion, WindowsBuildLabEx, OsArchitecture,
    CsManufacturer, CsModel, CsTotalPhysicalMemory, OsUptime | Out-File -FilePath $output -Append -Encoding UTF8
} catch {
    Add-Content -Path $output -Value "No se pudo obtener informacion general: $($_.Exception.Message)" -Encoding UTF8
}

titulo "PROCESADOR (CPU)"
try {
    Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object Name, Manufacturer, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors, SocketDesignation | Format-List | Out-File -FilePath $output -Append -Encoding UTF8
} catch {
    Add-Content -Path $output -Value "No se pudo obtener informacion del procesador: $($_.Exception.Message)" -Encoding UTF8
}

titulo "MEMORIA RAM (Importante para upgrade)"
try {
    $ram = Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop
    foreach ($r in $ram) {
        $gb = [math]::Round($r.Capacity / 1GB, 2)
        Add-Content -Path $output -Value "Fabricante : $($r.Manufacturer)" -Encoding UTF8
        Add-Content -Path $output -Value "Capacidad : $gb GB" -Encoding UTF8
        Add-Content -Path $output -Value "Velocidad : $($r.Speed) MHz" -Encoding UTF8
        Add-Content -Path $output -Value "Slot : $($r.DeviceLocator)" -Encoding UTF8
        Add-Content -Path $output -Value "Parte Number: $($r.PartNumber)" -Encoding UTF8
        Add-Content -Path $output -Value "-----------------------------------" -Encoding UTF8
    }
    $totalRAM = [math]::Round((Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).TotalPhysicalMemory / 1GB, 2)
    Add-Content -Path $output -Value "RAM TOTAL INSTALADA: $totalRAM GB" -Encoding UTF8
} catch {
    Add-Content -Path $output -Value "No se pudo obtener informacion de RAM: $($_.Exception.Message)" -Encoding UTF8
}

titulo "DISCOS / ALMACENAMIENTO (Ideal para upgrade a SSD)"
try {
    Get-PhysicalDisk -ErrorAction Stop | Select-Object FriendlyName, MediaType, BusType,
    @{Name="SizeGB";Expression={[math]::Round($_.Size/1GB,2)}}, HealthStatus, SerialNumber |
    Format-Table -AutoSize | Out-File -FilePath $output -Append -Encoding UTF8
} catch {
    Add-Content -Path $output -Value "No se pudo obtener informacion de discos: $($_.Exception.Message)" -Encoding UTF8
}

titulo "PARTICIONES"
try {
    Get-Volume -ErrorAction Stop | Where-Object DriveLetter | Select-Object DriveLetter, FileSystemLabel, FileSystem,
    @{Name="TotalGB";E={[math]::Round($_.Size/1GB,2)}},
    @{Name="LibreGB";E={[math]::Round($_.SizeRemaining/1GB,2)}} |
    Format-Table -AutoSize | Out-File -FilePath $output -Append -Encoding UTF8
} catch {
    Add-Content -Path $output -Value "No se pudo obtener informacion de particiones: $($_.Exception.Message)" -Encoding UTF8
}

titulo "GPU / TARJETA GRAFICA"
try {
    Get-CimInstance Win32_VideoController -ErrorAction Stop | Select-Object Name,
    @{Name="RAM_GB";E={[math]::Round($_.AdapterRAM/1GB,2)}}, DriverVersion, VideoProcessor |
    Format-List | Out-File -FilePath $output -Append -Encoding UTF8
} catch {
    Add-Content -Path $output -Value "No se pudo obtener informacion de GPU: $($_.Exception.Message)" -Encoding UTF8
}

titulo "PLACA BASE"
try {
    Get-CimInstance Win32_BaseBoard -ErrorAction Stop | Format-List Manufacturer, Product, SerialNumber, Version | Out-File -FilePath $output -Append -Encoding UTF8
} catch {
    Add-Content -Path $output -Value "No se pudo obtener informacion de placa base: $($_.Exception.Message)" -Encoding UTF8
}

titulo "FUENTE DE PODER (PSU) - Limitado"
try {
    $psu = Get-CimInstance Win32_PowerSupply -ErrorAction Stop
    if ($psu) {
        $psu | Select-Object Name, Model, Manufacturer, Capacity, PowerSupplyType | Format-List | Out-File -FilePath $output -Append -Encoding UTF8
    } else {
        Add-Content -Path $output -Value "No se pudo detectar informacion detallada de la Fuente de Poder." -Encoding UTF8
        Add-Content -Path $output -Value "Recomendacion: Abrir el gabinete y ver la etiqueta fisica." -Encoding UTF8
    }
} catch {
    Add-Content -Path $output -Value "No se pudo obtener informacion de la fuente de poder. Abrir gabinete y ver etiqueta fisica." -Encoding UTF8
}

titulo "MONITORES"
try {
    Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ErrorAction Stop | ForEach-Object {
        $nombre = ($_.UserFriendlyName | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
        $serial = ($_.SerialNumberID | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
        Add-Content -Path $output -Value "Monitor : $nombre" -Encoding UTF8
        Add-Content -Path $output -Value "Serial : $serial" -Encoding UTF8
        Add-Content -Path $output -Value "----------------------------------------" -Encoding UTF8
    }
} catch {
    Add-Content -Path $output -Value "No se pudo obtener informacion de monitores: $($_.Exception.Message)" -Encoding UTF8
}

titulo "LICENCIAS MICROSOFT"
try {
    Get-CimInstance SoftwareLicensingProduct -ErrorAction Stop | Where-Object { ($_.Name -like "*Windows*" -or $_.Name -like "*Office*") -and $_.PartialProductKey } |
    Select-Object Name, LicenseStatus, PartialProductKey, Description | Format-List | Out-File -FilePath $output -Append -Encoding UTF8
} catch {
    Add-Content -Path $output -Value "No se pudo obtener informacion de licencias: $($_.Exception.Message)" -Encoding UTF8
}

titulo "SOFTWARE INSTALADO (especialmente Office)"
try {
    Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction Stop |
    Where-Object DisplayName |
    Select-Object DisplayName, DisplayVersion, Publisher |
    Sort-Object DisplayName |
    Format-Table -AutoSize | Out-File -FilePath $output -Append -Encoding UTF8
} catch {
    Add-Content -Path $output -Value "No se pudo obtener informacion de software: $($_.Exception.Message)" -Encoding UTF8
}

titulo "PERIFERICOS DETECTADOS"
try {
    Get-PnpDevice -ErrorAction Stop | Where-Object Status -eq "OK" | Select-Object Class, FriendlyName, Manufacturer |
    Sort-Object Class | Format-Table -AutoSize | Out-File -FilePath $output -Append -Encoding UTF8
} catch {
    Add-Content -Path $output -Value "No se pudo obtener informacion de perifericos: $($_.Exception.Message)" -Encoding UTF8
}

titulo "INFORMACION FINAL"
Add-Content -Path $output -Value "Equipo : $env:COMPUTERNAME" -Encoding UTF8
Add-Content -Path $output -Value "Usuario : $env:USERNAME" -Encoding UTF8
Add-Content -Path $output -Value "Archivo : $output" -Encoding UTF8

Write-Host "=======================================" -ForegroundColor Green
Write-Host " INVENTARIO GENERADO CORRECTAMENTE" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green
Write-Host "Guardado en: $output" -ForegroundColor Cyan
