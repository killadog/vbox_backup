<#
.EXAMPLE
.\vbox_backup.ps1 -VM TESTVM -Destination D:\BACKUP_VM
#>
[cmdletBinding()]
Param
(
    [Parameter(Mandatory = $true)][String]$VM,
    [Parameter(Mandatory = $true)][String]$Destination
)
function Get-RunningVirtualBox($VM) {
    $VBoxManage = 'C:\Program Files\Oracle\VirtualBox\VBoxManage.exe'
    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $VBoxManage
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = "list runningvms"
    #$pinfo.Arguments = "$query"
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null
    $p.WaitForExit()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $VM = '"' + $VM + '"'
    if ($stdout.Contains($VM)) {
        return $stdout
    }
}

$Date = Get-Date -UFormat "%Y%m%d-%H%M%S"
$VBoxManage = 'C:\Program Files\Oracle\VirtualBox\VBoxManage.exe'
$OVA = "$VM-$Date.ova"
Get-RunningVirtualBox($VM)
Write-Host "Testing if $Destination exists, if not then create it"
if (-Not(Test-Path $Destination)) {
    New-Item -Path $Destination -ItemType Directory
}
Write-Host "Checking if $OVA already exists and removing it before beginning"
if (Test-Path $Destination\$OVA) {
    Remove-Item $Destination\$OVA -Force -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent)
}
#Write-Host "acpipowerbutton and Stopping $VM"
#Start-Process $VBoxManage -ArgumentList "controlvm $VM acpipowerbutton" -Wait -WindowStyle Hidden
Write-Host "SaveState and Stopping $VM"
Start-Process $VBoxManage -ArgumentList "controlvm $VM savestate" -Wait -WindowStyle Hidden
Write-Host "Waiting for $VM to have stopped"
While (Get-RunningVirtualBox($VM)) {
    Start-Sleep -Seconds 1
}
Write-Host "Exporting the VM appliance of $VM as $OVA to $Destination"
Start-Process $VBoxManage -ArgumentList "export $VM -o $Destination\$OVA" -Wait -WindowStyle Hidden
Write-Host "Starting $VM"
Start-Process $VBoxManage -ArgumentList "startvm $VM" -Wait -WindowStyle Hidden
Write-Host "Completed the Backup"