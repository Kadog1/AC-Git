function Get-TimeStamp {
    
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    
}

function Write-Log([string]$logtext)
{
    $logdate = get-date -format "yyyy-MM-dd HH:mm:ss"
    $logtext = $logtext
    $text = "["+$logdate+"] - " + $logtext
    Write-Host $text
    $text >> $logfile
}

# Testenvironment yes / no
[boolean]$productivemode = $True
if($productivemode -eq $True)
{
    $testenv=""
    $orderbook_test = ""
    $emailbox = "adressabgleich@de.ey.com"
}
Else 
{
    $testenv=" Testumgebung"
    $orderbook_test = "_TEST"
    $emailbox = "anykey@de.ey.com"
}

$pathLog = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\A Order entry\LogFiles\"
$date = get-date -format "yyyy-MM-dd"
$file = ("AC_Storno_Log_" + $date + "_" +$productivemode+".log")
$logfile = $pathLog + $file

Write-Host "$(Get-TimeStamp) Start Script: StornoAC.ps1"
Write-Log "Start StornoAC.ps1  (Productive = $productivemode)"

try {
    # Part 1: Integrate Storno JR mails into db.
    # Initialize Variables & Define paths
    $tolSize = 10000
    $dateStornoMailSent = Get-Date -Format G
    $status = "Storno"

    $filepath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\A Order entry\Storno"
    $processedfolder = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\A Order entry\Storno\Processed"

    # A1 Connect to shared Inbox *@de.ey.com and retrieve e-mails
    Write-host "Running Connect to Outlook..."
    $olFolderInbox = 6
    $outlook = new-object -comobject outlook.application;
    $namespace = $outlook.GetNamespace(“MAPI”);
    $recipient = $namespace.CreateRecipient("adressabgleich@de.ey.com")
    $inbox = $namespace.GetSharedDefaultFolder($recipient, $olFolderinbox)
    $storno = $inbox.Folders | where-object {$_.name -eq "Storno AC"}
    $processed = $inbox.Folders | where-object {$_.name -eq "Processed Order"}
    $ordercount=$storno.items.count #Anzahl der Mails im Mailfolder $new
    write-host "Anzahl Storno Objekte: "$ordercount # Infozwecke Anzahl Neue Ordermails

    # A2 Connect to Database ACCDB and retrieve order list
    Write-host "Running Retrieve Orderlist..."
    $adOpenStatic = 3
    $adLockOptimistic = 3
    $Connection = New-Object -com "ADODB.Connection"
    $RecordSet = New-Object -com "ADODB.Recordset"
    $Connection.Open("Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;")
    $RecordSet.Open("Select * From tAC_Orderbook$orderbook_test", $Connection,$adOpenStatic,$adLockOptimistic)

    $orderbcount = 0
    #Clear-variable orderlist
    while ($RecordSet.EOF -ne $True)
    {
        $OrderNo=$RecordSet.Fields.Item("OrderNo").Value
        $orderlist = $orderlist + ,"$OrderNo"
        $orderbcount=$orderbcount + 1
        $RecordSet.MoveNext()
    }
    $RecordSet.Close()
    $Connection.Close()

    if ($ordercount -gt 0){
    # B1 Download attachments *.txt for each storno mail only if order number exists in orderlist
    Write-host "Processing Attachments for E-Mail with Subject..." 
    $storno.Items | foreach {
        $subject = $_.subject
        $Order_subject = $subject.Substring(0, 12)
        # Handle Attachments
        $_.attachments | foreach {
            [boolean]$orderbookentry=$FALSE #resetten des boolean wertes
            $txtname = $_.filename; $size = $_.size
            Write-host "Process Attachment " $txtname"..."
            If ($txtname -like “AC000*" -and $txtname -like "*.txt” -and $size -lt $tolSize){
                Write-host "Attachment detected."
                $txtname2 = $txtname.Substring(0,$txtname.Length-24)
                $txtname3 = "*$txtname2*"
                if ($orderlist -like $txtname3){
                    write-host "Attachment download"
                    $_.saveasfile((Join-Path $filepath $txtname))
                } else {
                    write-host "No attachment download"
                }
            }
        }
    }

    # B2 Clear folder 'Storno' and move all e-mails to folder 'Processed Order'
    for ($i = $ordercount; $i -gt 0; $i--) {
            If($storno.items($i).attachments(1).filename -like “AC000*" -and $storno.items($i).attachments(1).filename -like "*.txt” -and $storno.items($i).attachments(1).size -lt $tolSize) {
                $storno.items($i).move($processed) | Out-Null
        }
    }
    Clear-variable -name i, ordercount
    }

    $adOpenStatic = 3
    $adLockOptimistic = 3
    $Connection = New-Object -com "ADODB.Connection"
    $RecordSet = New-Object -com "ADODB.Recordset"
    $Connection.Open("Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;")

    # C Read Order *.txt as CSV and adjust database for storno request.
    Get-ChildItem $filepath -Filter *.txt |
        ForEach-Object{
            $fileName = $_.Name
            $order = $fileName.Substring(0, 12)
            $stornoList = $stornoList + "$order,"
            Import-Csv $_.FullName -Delimiter "|" -Header @("StornoContact") | 
                foreach-object {
                    # Update Orderbook for each storno record
                    $RecordSet.Open("SELECT * FROM tAC_Orderbook$orderbook_test WHERE [OrderNo] = '$order'", $Connection, $adOpenStatic, $adLockOptimistic)
                    Write-Host "Set 'Storno' for order " $RecordSet.Fields.Item("OrderNo").Value
                    $RecordSet.Fields.Item("AC_Status").Value = "Storno"
                    $RecordSet.Fields.Item("StornoCntct").Value = $_.StornoContact # updating the Order book with the new status after e-mail has been sent
                    $RecordSet.Update()
        }       
    }
    $Connection.Close()

    # D Alle txt dateien, die ausgelesen wurden und die Infos im Orderbook abgespeichert wurden verschoben in den Ordern "Processed"
    Get-ChildItem -Path $filepath -Recurse -file *.txt | Move-Item -Destination $processedfolder -Force

    if ($stornoList.length -gt 0){
        $stornoList = $stornoList.Substring(0, $stornoList.length - 1)



    # Part 2: Send Storno Confirmation Mails
    $stornoList =$stornoList.Split(",")

    $adOpenStatic = 3
    $adLockOptimistic = 3
    $Connection = New-Object -com "ADODB.Connection"
    $RecordSet = New-Object -com "ADODB.Recordset"
    $Connection.Open("Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;")

    # search through the Order book and find the Orders with prescribed status ("Storno" - meaning that Storno Mail has yet to be sent)
    for ($i=0; $i -lt $stornoList.length; $i++){
        $order = $stornoList[$i]
        $RecordSet.Open("SELECT * FROM tAC_Orderbook$orderbook_test WHERE [OrderNo] = '$order'", $Connection, $adOpenStatic, $adLockOptimistic)
        If ($RecordSet.Fields.Item("AC_Status").Value -eq $status -and [string]::IsNullOrEmpty($RecordSet.Fields.Item("tsStornoSent").Value)) # search for status flag in the Status field
        {
            Write-host $RecordSet.Fields.Item("OrderNo").Value 
            #Auslesen der benötígten Informationen aus dem Orderbook
            $To1 = $RecordSet.Fields.Item("engcontact").Value
            $To2 = $RecordSet.Fields.Item("engmngr").Value
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_2").Value) -eq $false) {$To2 = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_2").Value)+','+$To2} 
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_3").Value) -eq $false) {$To2 = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_3").Value)+','+$To2} 
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_4").Value) -eq $false) {$To2 = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_4").Value)+','+$To2}
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_5").Value) -eq $false) {$To2 = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_5").Value)+','+$To2}
            [string[]]$To2 = $To2.Split(',')

            $Order = $RecordSet.Fields.Item("OrderNo").Value
            $Client = $RecordSet.Fields.Item("client").Value
            $Orderdate = $RecordSet.Fields.Item("OrderDate").Value
            $EngCode = $RecordSet.Fields.Item("engcode").Value
            $GISID = $RecordSet.Fields.Item("GISID").Value
            $Lang=$RecordSet.Fields.Item("LangProduct").Value
        
            #Erstellung des arrays aus der analysislist um später im Mail tempaltes [list of order analysis] zu ersetzen abhängig von Sprache
            If($Lang -eq "DE") 
            {
                $Confimailtemplate= "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich\B Mail Templates\German\700 Stornierungsbestätigung.htm"
                $subject="Stornierung CAD AC Bestellung $Order für $client"
             }

             If($Lang -eq "ENG")
             {
                $Confimailtemplate= "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich\B Mail Templates\English\700 Cancellation confirmation_EN.htm"
                $subject="Cancelation CAD AC order $Order for $client"
             }
            $utf8 = New-Object System.Text.utf8encoding

            $AttachmentBanner= "\\DEVIDVAPFL04.ey.net\04EM1015\H\Heartbeat-CAD Team\CAD_Brand_Communications\CAD_Signatur+Banner\190807_CAD_Signature-Banner_500px_mittel.png"
            $AttachmentLogo= "\\DEVIDVAPFL04.ey.net\04EM1015\H\Heartbeat-CAD Team\CAD_Brand_Communications\CAD_Logo+Image\EY_Logo_Beam_RGB.png"

            $confihtmlbody = Get-content $Confimailtemplate -Encoding UTF8
            [string]$confimessage=$confihtmlbody.Replace("[OrderNo]", $Order).Replace("[client]", $client).replace("[OrderDate]", $OrderDate).replace("[GISID]", $GISID)
            Send-MailMessage -Encoding $utf8 -FROM $emailbox -To $To1 -Cc $To2 -BCc $emailbox -Subject $subject -Body $confimessage -BodyAshtml -SmtpServer "mail-de.ey.net" -Attachments $AttachmentLogo, $AttachmentBanner

            # updating the Order book
            $RecordSet.Fields.Item("tsStornoSent").Value = Get-Date -Format G # updating the Order book with the date of e-mail sent

            #$OrderFolderexists = Test-Path "\\DEVIDVAPFL04.ey.net\04EM1015\H\HeartbeatAR\Main Folder$testenv\F Workplace\1 Data received\$Order"
            #if($RecordSet.Fields.Item("Status").Value -eq $status -and $OrderFolderexists -eq "True") {
            #    #$Order Ordner aus Data Received verschieben in Recycle Bin
            #    move-item -path "\\DEVIDVAPFL04.ey.net\04EM1015\H\Heartbeat-AR\Main Folder$testenv\F Workplace\1 Data received\$Order" -destination "\\DEVIDVAPFL04.ey.net\04EM1015\H\Heartbeat-AR\temp DO NOT DELETE\RecycleBin\1\"
            #    Write-host "Verschoben" 
            #    }
        
            #Clear Variable auf dieser Ebene, da erst gecleared werden muss, wenn Status leer war
            }

        $RecordSet.Update()
        $RecordSet.Close()
        }
    }
}
catch {
    Write-Warning "ERROR: $($_.Exception.Message)"
    Write-Log "ERROR ConfiMailAC.ps1  (Productive = $productivemode): $($_.Exception.Message)"
    break
}
#Clear-variable stornoList, Array
#Ende des Skripts
Write-Host "$(Get-TimeStamp) End Script: StornoAC.ps1"
Write-Log "End StornoAC.ps1  (Productive = $productivemode)"



