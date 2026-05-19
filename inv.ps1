# ============================================
# INVENTARIO COMPLETO PARA COTIZACIÓN
# Ejecución remota desde GitHub:
#   iex (irm https://raw.githubusercontent.com/solracandres66-coder/inventario/main/inv.ps1)
# Descarga directa:
#   irm https://raw.githubusercontent.com/solracandres66-coder/inventario/main/inv.ps1 -o inv.ps1
# ============================================

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
Get-ComputerInfo | Format-List WindowsProductName, WindowsVersion, WindowsBuildLabEx, OsArchitecture, 
    CsManufacturer, CsModel, CsTotalPhysicalMemory, OsUptime | Out-File -FilePath $output -Append -Encoding UTF8

titulo "PROCESADOR (CPU)"
Get-CimInstance Win32_Processor | Select-Object Name, Manufacturer, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors, SocketDesignation | Format-List | Out-File -FilePath $output -Append -Encoding UTF8

titulo "MEMORIA RAM (Importante para upgrade)"
$ram = Get-CimInstance Win32_PhysicalMemory
foreach ($r in $ram) {
    $gb = [math]::Round($r.Capacity / 1GB, 2)
    Add-Content -Path $output -Value "Fabricante : $($r.Manufacturer)" -Encoding UTF8
    Add-Content -Path $output -Value "Capacidad   : $gb GB" -Encoding UTF8
    Add-Content -Path $output -Value "Velocidad   : $($r.Speed) MHz" -Encoding UTF8
    Add-Content -Path $output -Value "Slot        : $($r.DeviceLocator)" -Encoding UTF8
    Add-Content -Path $output -Value "Parte Number: $($r.PartNumber)" -Encoding UTF8
    Add-Content -Path $output -Value "-----------------------------------" -Encoding UTF8
}
$totalRAM = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
Add-Content -Path $output -Value "RAM TOTAL INSTALADA: $totalRAM GB" -Encoding UTF8

titulo "DISCOS / ALMACENAMIENTO (Ideal para upgrade a SSD)"
Get-PhysicalDisk | Select-Object FriendlyName, MediaType, BusType, 
    @{Name="SizeGB";Expression={[math]::Round($_.Size/1GB,2)}}, HealthStatus, SerialNumber | 
Format-Table -AutoSize | Out-File -FilePath $output -Append -Encoding UTF8

titulo "PARTICIONES"
Get-Volume | Where-Object DriveLetter | Select-Object DriveLetter, FileSystemLabel, FileSystem,
    @{Name="TotalGB";E={[math]::Round($_.Size/1GB,2)}}, 
    @{Name="LibreGB";E={[math]::Round($_.SizeRemaining/1GB,2)}} | 
Format-Table -AutoSize | Out-File -FilePath $output -Append -Encoding UTF8

titulo "GPU / TARJETA GRÁFICA"
Get-CimInstance Win32_VideoController | Select-Object Name, 
    @{Name="RAM_GB";E={[math]::Round($_.AdapterRAM/1GB,2)}}, DriverVersion, VideoProcessor | 
Format-List | Out-File -FilePath $output -Append -Encoding UTF8

titulo "PLACA BASE"
Get-CimInstance Win32_BaseBoard | Format-List Manufacturer, Product, SerialNumber, Version | Out-File -FilePath $output -Append -Encoding UTF8

titulo "FUENTE DE PODER (PSU) - Limitado"
$psu = Get-CimInstance Win32_PowerSupply -ErrorAction SilentlyContinue
if ($psu) {
    $psu | Select-Object Name, Model, Manufacturer, Capacity, PowerSupplyType | Format-List | Out-File -FilePath $output -Append -Encoding UTF8
} else {
    Add-Content -Path $output -Value "No se pudo detectar información detallada de la Fuente de Poder." -Encoding UTF8
    Add-Content -Path $output -Value "Recomendación: Abrir el gabinete y ver la etiqueta física." -Encoding UTF8
}

titulo "MONITORES"
Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ErrorAction SilentlyContinue | ForEach-Object {
    $nombre = ($_.UserFriendlyName | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
    $serial = ($_.SerialNumberID | Where-Object {$_ -ne 0} | ForEach-Object {[char]$_}) -join ""
    Add-Content -Path $output -Value "Monitor : $nombre" -Encoding UTF8
    Add-Content -Path $output -Value "Serial  : $serial" -Encoding UTF8
    Add-Content -Path $output -Value "----------------------------------------" -Encoding UTF8
}

titulo "LICENCIAS MICROSOFT"
Get-CimInstance SoftwareLicensingProduct | Where-Object { ($_.Name -like "*Windows*" -or $_.Name -like "*Office*") -and $_.PartialProductKey } |
Select-Object Name, LicenseStatus, PartialProductKey, Description | Format-List | Out-File -FilePath $output -Append -Encoding UTF8

titulo "SOFTWARE INSTALADO (especialmente Office)"
Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*", 
                 "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
Where-Object DisplayName | 
Select-Object DisplayName, DisplayVersion, Publisher | 
Sort-Object DisplayName | 
Format-Table -AutoSize | Out-File -FilePath $output -Append -Encoding UTF8

titulo "PERIFÉRICOS DETECTADOS"
Get-PnpDevice | Where-Object Status -eq "OK" | Select-Object Class, FriendlyName, Manufacturer | 
Sort-Object Class | Format-Table -AutoSize | Out-File -FilePath $output -Append -Encoding UTF8

titulo "INFORMACION FINAL"
Add-Content -Path $output -Value "Equipo     : $env:COMPUTERNAME" -Encoding UTF8
Add-Content -Path $output -Value "Usuario    : $env:USERNAME" -Encoding UTF8
Add-Content -Path $output -Value "Archivo    : $output" -Encoding UTF8

Write-Host "=======================================" -ForegroundColor Green
Write-Host "  INVENTARIO GENERADO CORRECTAMENTE" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green
Write-Host "Guardado en: $output" -ForegroundColor Cyan

Pause
