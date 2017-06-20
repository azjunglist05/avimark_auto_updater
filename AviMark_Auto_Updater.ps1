$ZIPUpdate = "\\10.252.70.3\Users\Public\Downloads\AVImark Beta 2016.4.5.zip"
$Destination = "C:\Temp\Avimark"
$AVIMarkZip = "C:\Temp\Avimark\AVImark Beta 2016.4.5.zip"

if ((Test-Path -Path $Destination) -eq $false) {
    New-Item -Path $Destination -ItemType Container -Force -Verbose
    Copy-Item -Path $ZIPUpdate -Destination $Destination -Force -Verbose
}

if ((Test-Path -Path $AVIMarkZip) -eq $false) {
    Copy-Item -Path $ZIPUpdate -Destination $Destination -Force -Verbose
}

$AVIMarkProcess = Get-Process | ? {$_.ProcessName -like "AVIM*"}

if (($AVIMarkProcess | Select-Object -ExpandProperty Name) -like "AVIM*") {
     
    foreach ($Process in ($AVIMarkProcess | Select-Object -ExpandProperty ProcessName)) {
        Stop-Process -Name $Process -Verbose -Force
    }
}

$MPSProcess = Get-Process | ? {$_.ProcessName -like "MPS*"}

if (($MPSProcess | Select-Object -ExpandProperty ProcessName) -like "MPS*") {

    foreach ($Process in ($MPSProcess | Select-Object -ExpandProperty Name)) {
        Stop-Process -Name $Process -Verbose -Force
    }
}

$AVIMarkServer = Get-Service | ? {$_.ServiceName -like "AVIM*"}

if (($AVIMarkServer | Select-Object -ExpandProperty Status) -eq "Running") {
    
    foreach ($Service in ($AVIMarkServer | Select-Object -ExpandProperty Name)) {
        Stop-Service -Name $Service -Verbose -Force -ErrorAction SilentlyContinue
    }
}

$IDEXXService = Get-Service | ? {$_.DisplayName} -like "IDEXX*"

if (($IDEXXService | Select-Object -ExpandProperty Status) -eq "Running" ) {
    
    foreach ($Service in ($IDEXXService | Select-Object -ExpandProperty Name)) {
        Stop-Service -Name $Service -Verbose -Force -ErrorAction SilentlyContinue
    }
}

$Vetstoria = Get-Service | ? {$_.ServiceName -like "Vets*"}

if (($Vetstoria | Select-Object -ExpandProperty Status) -eq "Running") {
    
    foreach ($Service in ($Vetstoria | Select-Object -ExpandProperty Name)) {
        Stop-Service -Name $Service -Verbose -Force -ErrorAction SilentlyContinue
    }
}

Copy-Item -Path "D:\Avimark" -Recurse -Destination "D:\Backup" -Verbose -Force

$BackCheck = "D:\Backup\Avimark"

if ((Test-Path -Path $BackCheck) -eq $true) {

    $ZipFile = (New-Object -COM Shell.Application).NameSpace($AVIMarkZip)
    $DestinationPath = "D:\Avimark"
    $Destination = (New-Object -COM Shell.Application).NameSpace($DestinationPath)

    if ((Test-Path -Path $AVIMarkZip) -eq $true) {
        $Destination.CopyHere($ZipFile.Items(), 16)
    } else {
        Write-Error "AVIMark update file did not transfer properly!"
    }

} else {
        Write-Error "AVIMark backup did not complete successfully. Please, manually backup to 'D:\Backup\Avimark' and re-run script."
}

$Continue = Read-Host "Please, continue manual update of Avimark, then type in 'Resume' to start all services"

if ($Continue -eq "Resume") {

    foreach ($Service in ($AVIMarkServer | Select-Object -ExpandProperty Name)) {
        Start-Service -Name $Service -Verbose
    }

    foreach ($Service in ($Vetstoria | Select-Object -ExpandProperty Name)) {
        Start-Service -Name $Service -Verbose
    }

    foreach ($Service in ($IDEXXService | Select-Object -ExpandProperty Name)) {
        Start-Service -Name $Service -Verbose
    }
} else {
    Return
}