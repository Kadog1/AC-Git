function Get-TimeStamp {
    
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    
}
Write-Host "$(Get-TimeStamp) Start Script: SwitchOn.ps1"
$psInstance = "Powershell"
$processes = Get-Process | Where-Object {$_.Name -eq $psInstance}
$counter = 0
foreach ($prcs in $processes)
{
    $counter++
}

if($counter -eq 1)
{
# "Ich" bin das einzige aktive Powershell-Objekt -> "ich" kann nicht noch einmal ausgeführt werden während eine Instanz von "mir" läuft
    $olFolderInbox = 6 
    $outlook = new-object -comobject outlook.application;
    $namespace = $outlook.GetNamespace(“MAPI”)
    $emailbox = "adressabgleich@de.ey.com"
    $recipient = $namespace.CreateRecipient($emailbox)
    $inbox = $namespace.GetSharedDefaultFolder($recipient, $olFolderinbox) 
    $switch = $inbox.Folders | where-object { $_.name -eq "Switch" }
    $off = $switch.Folders | Where-Object {$_.name -eq "Off"}
    $on = $switch.Folders | Where-Object {$_.name -eq "On"}
    $off.Items | foreach {$_.move($on)}
}


Write-Host "$(Get-TimeStamp) End Script: SwitchOn.ps1"