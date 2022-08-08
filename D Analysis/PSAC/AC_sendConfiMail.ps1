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

#Wenn Produktivmodues dann auf TRUE stellen, wenn Testumgebung auf FALSE
[boolean]$productivemode = $False

write-host $PSScriptRoot

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
########################################################################

write-host "Productive = $productivemode" 

$pathLog = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\A Order entry\LogFiles\"
$date = get-date -format "yyyy-MM-dd"
$file = ("AC_ConfiMail_Log_" + $date + "_" +$productivemode+".log")
$logfile = $pathLog + $file

Write-Host "$(Get-TimeStamp) Start Script: ConfiMailAC.ps1"
Write-Log "Start ConfiMailAC.ps1  (Productive = $productivemode)"

try {
    #Ansteuern des Shared Inbox AC@de.ey.com
    $olFolderInbox = 6 
    $outlook = new-object -comobject outlook.application;
    $namespace = $outlook.GetNamespace(“MAPI”);
    $recipient = $namespace.CreateRecipient($emailbox)
    $inbox = $namespace.GetSharedDefaultFolder($recipient, $olFolderinbox) 
    $new = $inbox.Folders | where-object { $_.name -eq "New Order" }
    $processed = $inbox.Folders | where-object {$_.name -eq "Processed Order"}
    $ordercount=$new.items.count #Anzahl der Mails im Mailfolder $new
    write-host "Neue Bestellungen: $ordercount" #Infozwecke Anzahl Neue Ordermails

    #Verzeichnisse Fileserver HB
    #$filepath = “$testenv\A Order entry\Orders\”
    #produktiv
    $filepath = “\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\A Order entry\Orders”

    #$ordercount = 1

    if($ordercount -gt 0){
        #Öffnen der Verbindung zum ACCDB
        $adOpenStatic = 3
        $adLockOptimistic = 3
        $Connection = New-Object -com "ADODB.Connection"
        $RecordSet = New-Object -com "ADODB.Recordset"
        $Connection.Open("Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;")
        $RecordSet.Open("Select * From AC_CON_Helper_getRSConApAr('AC', '%')", $Connection,$adOpenStatic,$adLockOptimistic)

        #Erstellen Array alle Order aus ACCDB und Anzahl der Order im Orderbook
        #Wird verwendet um neue Order Validität zu checken
        while ($RecordSet.EOF -ne $True) 
        {
            $OrderNo=$RecordSet.Fields.Item("OrderNo").Value
            $orderlist = $orderlist + ,"$OrderNo"
            $orderbcount=$orderbcount+1
            $RecordSet.MoveNext()
        } 

        #Speichert alle Attachments von JR in das Verzeichnis $filepath, wenn die Ordernummer nicht im Orderbook gefunden werden kann
        $new.Items | foreach {
            $_.attachments | foreach {
                [boolean]$orderbookentry=$FALSE
                $txtname = $_.filename
                $size = $_.size
                If ($txtname -like “AC000*" -and $txtname -like "*.txt” -and $size -lt 10000) 
                {
                    for($i=0; $i -lt $orderbcount; $i++) 
                    {
                        if ($txtname.Substring(0,$txtname.Length-4) -contains $orderlist[$i]) {
                            $orderbookentry=$TRUE
                        }
                    }
                    if($orderbookentry -eq $FALSE)
                    {
                       $_.saveasfile((Join-Path $filepath $txtname))
                    }
                }
            }
        }
        Clear-variable -name i, orderbcount, orderlist

        #Schließen der Verbindung zur ACCDB
        $RecordSet.Close()
        $Connection.Close()

        #Verschiebt alle Mails mit Order AC*.txt Attachment in den Outlook Folder $processed
        for ($i = $ordercount; $i -gt 0; $i--) 
            {
            If($new.items($i).attachments(1).filename -like “AC000*" -and $new.items($i).attachments(1).filename -like "*.txt” -and $new.items($i).attachments(1).size -lt 10000) 
                {
                $new.items($i).move($processed)
                }
            }
        Clear-variable -name i, ordercount

        # Auslesen der txt/CSV Dateien und import in ACCDB"
        $adOpenStatic = 3
        $adLockOptimistic = 3
        $Connection = New-Object -com "ADODB.Connection"
        $RecordSet = New-Object -com "ADODB.Recordset"
        $Connection.Open("Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;")
        #$RecordSet.Open("Select * From tAC_Orderbook$orderbook_test", $Connection,$adOpenStatic,$adLockOptimistic)

        $processedfolder = “\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\A Order entry\Orders\Processed”
        $newfolder = “\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\A Order entry\Orders”
        #ANPASSUNG AN AC BASIC_INFO#
        Get-ChildItem -path $newfolder -Filter *.txt | 
        ForEach-Object{
            Import-Csv $_.FullName -Delimiter "|" -Header @("OrderNo","GISID","client","engcode","engname","engprtnr","engmngr","engcontact","LangProduct","Periodend","OrderDate","ADDITIONAL_CONTACT_2","ADDITIONAL_CONTACT_3","ADDITIONAL_CONTACT_4","ADDITIONAL_CONTACT_5") | 
            foreach-object{
                $InsertIntoQuery = "INSERT INTO CAD.dbo.tAC_Orderbook$orderbook_test (OrderNo, GISID, client, engcode, engname, engprtnr, engmngr, engcontact, LangProduct, Periodend, OrderDate, 
                ADDITIONAL_CONTACT_2, ADDITIONAL_CONTACT_3, ADDITIONAL_CONTACT_4, ADDITIONAL_CONTACT_5, tsOrderImported) 
                VALUES ('$($_.OrderNo)', $($_.GISID.ToString().PadLeft(10,"0")), '$($_.client)', '$($_.engcode)', '$($_.engname)', '$($_.engprtnr)', '$($_.engmngr)', '$($_.engcontact)', '$($_.LangProduct)', '$([datetime]::parseexact($_.Periodend, 'dd.MM.yyyy', $null).ToString("yyyy-MM-dd"))',
                 '$([datetime]::parseexact($_.OrderDate, 'dd.MM.yyyy', $null).ToString("yyyy-MM-dd"))', '$($_.ADDITIONAL_CONTACT_2)', '$($_.ADDITIONAL_CONTACT_3)', '$($_.ADDITIONAL_CONTACT_4)', '$($_.ADDITIONAL_CONTACT_5)', '$(Get-Date -Format "yyyy-MM-dd HH:mm")')"
                $InsertIntoQuery = $InsertIntoQuery.replace(" '',",' NULL,')
                $RecordSet.Open($InsertIntoQuery, $Connection,$adOpenStatic,$adLockOptimistic)
                <#
                $RecordSet.AddNew()
                $RecordSet.Fields.Item("OrderNo").Value = $_.OrderNo
                $RecordSet.Fields.Item("GISID").Value = $_.GISID.ToString().PadLeft(10,"0")
                $RecordSet.Fields.Item("client").Value = $_.client
                $RecordSet.Fields.Item("engcode").Value = $_.engcode
                $RecordSet.Fields.Item("engname").Value = $_.engname
                $RecordSet.Fields.Item("engprtnr").Value = $_.engprtnr
                $RecordSet.Fields.Item("engmngr").Value = $_.engmngr
                $RecordSet.Fields.Item("engcontact").Value = $_.engcontact
                $RecordSet.Fields.Item("LangProduct").Value = $_.LangProduct
                $RecordSet.Fields.Item("Periodend").Value = [datetime]::parseexact($_.Periodend, 'dd.MM.yyyy', $null).ToString("yyyy-MM-dd")
                $RecordSet.Fields.Item("OrderDate").Value = [datetime]::parseexact($_.OrderDate, 'dd.MM.yyyy', $null).ToString("yyyy-MM-dd")
                $RecordSet.Fields.Item("ADDITIONAL_CONTACT_2").Value = $_.ADDITIONAL_CONTACT_2
                $RecordSet.Fields.Item("ADDITIONAL_CONTACT_3").Value = $_.ADDITIONAL_CONTACT_3
                $RecordSet.Fields.Item("ADDITIONAL_CONTACT_4").Value = $_.ADDITIONAL_CONTACT_4
                $RecordSet.Fields.Item("ADDITIONAL_CONTACT_5").Value = $_.ADDITIONAL_CONTACT_5
                $RecordSet.Fields.Item("tsOrderImported").Value = Get-Date -Format "yyyy-MM-dd HH:mm"
                $RecordSet.Update()
                #>
            }
        }
        #Alle txt dateien, die ausgelesen wurden und die Infos im Orderbook abgespeichert wurden verschoben in den Ordern "Processed"
        Get-ChildItem -Path $newfolder -file *.txt | Move-Item -Destination $processedfolder -Force

        #Schließen der Verbindung
        #$RecordSet.Close()
        $Connection.Close()
        }
    
    # Versenden der ConfiMail und Customer Data Folrder Erstellung"
    $adOpenStatic = 3
    $adLockOptimistic = 3
    $Connection = New-Object -com "ADODB.Connection"
    $RecordSet = New-Object -com "ADODB.Recordset"
    $RecordSetUpdate = New-Object -com "ADODB.Recordset"
    $Connection.Open("Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;")
    $RecordSet.Open("Select * From AC_CON_Helper_getRSConApAr('AC', '%')", $Connection,$adOpenStatic,$adLockOptimistic)

    $dateConfiMailSent = Get-Date -Format "yyyy-MM-dd hh:mm" #get the current date and time
    $status2 = "InputDataSent" #status after ( = conf email has been sent)

    $attLogo ="\\devidvapfl04.ey.net\04em1015\H\Heartbeat-CAD Team\CAD_Brand_Communications\CAD_Logo+Image\EY_Logo_Beam_RGB.png"
    $attBanner ="\\devidvapfl04.ey.net\04em1015\H\Heartbeat-CAD Team\CAD_Brand_Communications\CAD_Signatur+Banner\190807_CAD_Signature-Banner_500px_mittel.png"
    $RecordSetCount=0
    while ($RecordSet.EOF -ne $True) 
        {
            $RecordSetCount = $RecordSetCount+1
            $RecordSet.MoveNext()
        } 
    $RecordSet.Close() 
    $RecordSet.Open("Select * From CAD.dbo.AC_CON_Helper_getRSConApAr('AC', '%')", $Connection,$adOpenStatic,$adLockOptimistic)
    #search through the Order book and find the Orders with empty status (meaning that the confirmation e-mail has not been sent yet)
    While ($RecordSet.EOF -ne $True -And $RecordSetCount -gt 0) {
        If ([string]::IsNullOrEmpty($RecordSet.Fields.Item("AC_Status").Value)){

            #Write-host "Aktuelle OrderNo: "$RecordSet.Fields.Item("OrderNo").Value 
            $OrderNo = $RecordSet.Fields.Item("OrderNo").Value
            Write-host "OrderNo: $OrderNo"
            #Auslesen der benötígten Informationen aus dem Orderbook
            $To1 = $RecordSet.Fields.Item("engcontact").Value
            $To2 = $RecordSet.Fields.Item("engpartner").Value
            $To3 = $RecordSet.Fields.Item("engmanager").Value
    
            $Periodend = [datetime]::parseexact($RecordSet.Fields.Item("YearEnd").Value, 'yyyy-MM-dd', $null).ToString("dd.MM.yyyy")
            $Client = $RecordSet.Fields.Item("client").Value
            $Orderdate = [datetime]::parseexact($RecordSet.Fields.Item("OrderDate").Value, 'yyyy-MM-dd', $null).ToString("dd.MM.yyyy")
            $EngOtherContact = $RecordSet.Fields.Item("OtherContact").Value
            $GISID = $RecordSet.Fields.Item("GISID").Value
            $EngCode = $RecordSet.Fields.Item("engCode").Value
            $Lang = $RecordSet.Fields.Item("DocLang").Value

            $To2 = $($RecordSet.Fields.Item("engmanager").Value)+','+$To2
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_2").Value) -eq $false) {$To2 = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_2").Value)+','+$To2} 
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_3").Value) -eq $false) {$To2 = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_3").Value)+','+$To2} 
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_4").Value) -eq $false) {$To2 = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_4").Value)+','+$To2}
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_5").Value) -eq $false) {$To2 = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_5").Value)+','+$To2}
            [string[]]$To2 = $To2.Split(',')
            If ([string]::IsNullOrEmpty($RecordSet.Fields.Item("AC_Status").Value)) # search for status flag in the Status field
            {

            $Excel = new-object -comobject excel.application
            $Excel.Visible = $false

            $Excel.EnableEvents = $false 
            $Excel.DisplayAlerts = $false

            if($Lang -eq "DE") {
                    $objWorkbook = $Excel.Workbooks.Open("\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\E Documentation Templates\German\1_CAD-Adressabgleich Adressenabfrage Mandant.xlsm")            
            }
             if(($Lang -eq "EN") -or ($Lang -eq "ENG")) {
                    $objWorkbook = $Excel.Workbooks.Open("\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\E Documentation Templates\English\1_CAD-Adressabgleich Adressenabfrage Mandant_EN.xlsm")
            }

            # Auf Sheet 3 (Validation) befinden sich die unique value inputs
            $objWorksheet = $objWorkbook.Worksheets.Item("basic_info")
            $objWorksheet.cells.Item(1,2) = $OrderNo
            $objWorksheet.cells.Item(2,2) = $Periodend
            $objWorksheet.cells.Item(3,2) = $Client
            $objWorksheet.cells.Item(4,2) = $To1
            $objWorksheet.cells.Item(5,2) = $To2
            $objWorksheet.cells.Item(6,2) = $To3
            $objWorksheet.cells.Item(7,2) = $EngOtherContact
            $objWorksheet.cells.Item(8,2) = $GISID
            $objWorksheet.cells.Item(9,2) = "TBD"
            $objWorksheet.cells.Item(10,2) = $EngCode

            $ExcelSavePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\C Workplace"
            $OrderFolderexists = Test-Path "$ExcelSavePath\$OrderNo”
            If($OrderFolderexists -eq "True") 
            {
            Write-host "OrderFolder exists"
            }
            else 
            {
            New-Item "$ExcelSavePath\$OrderNo” -type directory
            }

            $fileName = $GISID.ToString().PadLeft(10,"0") + " 1_CAD-Adressabgleich Adressenabfrage Mandant " + [datetime]::parseexact($RecordSet.Fields.Item("YearEnd").Value, 'yyyy-MM-dd', $null).ToString("yyyyMMdd")
    
            $objWorkbook.SaveAs($ExcelSavePath + "\" + "$OrderNo" + "\" + $fileName,52) # http://msdn.microsoft.com/en-us/library/bb241279.aspx
            $objWorkbook.Saved = $true
            $objWorkbook.Close()
            Write-host $ExcelSavePath+$OrderNo\$fileName

            If($Lang -eq "DE")
            {
               $Confimailtemplate= "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\B Mail Templates\German\AC_ConfiMail.htm"
               $subject="Adressabgleich für $client/Bestellnummer: $OrderNo"
            }

             if(($Lang -eq "EN") -or ($Lang -eq "ENG"))
            {
               $Confimailtemplate= "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\B Mail Templates\English\AC_ConfiMail.htm"
               $subject="Address comparison for $client/Order number: $OrderNo"
            }

            $Attachmentdata = $ExcelSavePath + "\" + $OrderNo + "\" + $fileName + ".xlsm"

            $utf8 = New-Object System.Text.utf8encoding
            $confihtmlbody = Get-content $Confimailtemplate -Encoding UTF8

            [string]$confimessage=$confihtmlbody.Replace("[OrderNo]", $OrderNo).Replace("[Client]", $Client).Replace("[GISID]", $GISID).Replace("[YearEnd]", $PeriodEnd)

            Send-MailMessage -Encoding $utf8 -FROM $emailbox -To $To1 -Cc "$To2, $To3" -bcc $emailbox -Subject $subject -Body $confimessage -BodyAshtml -SmtpServer "mail-de.ey.net" -Attachments $Attachmentdata, $attLogo, $attBanner

            $RecordSetUpdate.Open("EXECUTE CAD.dbo.AcCon_UpdateOrderbook @OrderNo = '$($OrderNo)', @Fields = 'tsConfimailSent,tsInputDataSent,AC_Status', @Values = '$dateConfiMailSent,$dateConfiMailSent,$status2'", $Connection,$adOpenStatic,$adLockOptimistic)
            }   
        }
        $RecordSet.MoveNext()
    }

    #Schließen der Verbindung
    $RecordSet.Close()
    $Connection.Close()
}
catch {
    Write-Warning "ERROR: $($_.Exception.Message)"
    Write-Log "ERROR ConfiMailAC.ps1  (Productive = $productivemode): $($_.Exception.Message)"
    break
}
#Ende des Skripts
Write-Host "$(Get-TimeStamp) End Script: ConfiMailAC.ps1"
Write-Log "End ConfiMailAC.ps1  (Productive = $productivemode)"
