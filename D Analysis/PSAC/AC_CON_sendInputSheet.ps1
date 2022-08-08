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

[boolean]$productivemode = $True

write-host $PSScriptRoot

if($productivemode -eq $True) 
{
    $testenv =""
    $orderbook_test = ""
    $emailbox = "adressabgleich@de.ey.com"
}
Else 
{
    $testenv =" Testumgebung"
    $orderbook_test = "_TEST"
    $emailbox = "adressabgleich@de.ey.com"
}

write-host "Productive = $productivemode"
$pathLog = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\A Order entry\LogFiles\"
$date = get-date -format "yyyy-MM-dd"
$file = ("AC_CON_sendInputData_Log_" + $date + "_" +$productivemode+".log")
$logfile = $pathLog + $file

Write-Log "Start sendInputDataConAC.ps1  (Productive = $productivemode)"

try {
    # Versenden der ConfiMail und Customer Data Folder Erstellung"
    $adOpenStatic = 3
    $adLockOptimistic = 3
    $Connection = New-Object -com "ADODB.Connection"
    $RecordSet = New-Object -com "ADODB.Recordset"
    $Connection.Open("Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;")
    $RecordSet.Open("SELECT * FROM [CAD].[dbo].[tCON_Orderbook$orderbook_test] WHERE AC_Status IS NULL AND tsInputDataSent IS NULL AND Address_Validation = 1 AND (EYStand = 0 OR EYStand IS NULL) AND Storno IS NULL", $Connection,$adOpenStatic,$adLockOptimistic)

    $dateInputSheetSent = Get-Date -Format G #get the current date and time
    $nextStatus = "InputDataSent" #status after ( = conf email has been sent)

    $attLogo ="\\devidvapfl04.ey.net\04em1015\H\Heartbeat-CAD Team\CAD_Brand_Communications\CAD_Logo+Image\EY_Logo_Beam_RGB.png"
    $attBanner ="\\devidvapfl04.ey.net\04em1015\H\Heartbeat-CAD Team\CAD_Brand_Communications\CAD_Signatur+Banner\190807_CAD_Signature-Banner_500px_mittel.png"

    $RecordSetCount=0
    while ($RecordSet.EOF -ne $True)
        {
            $RecordSetCount = $RecordSetCount+1
            $RecordSet.MoveNext()
        } 
    $RecordSet.Close() 
    $RecordSet.Open("SELECT * FROM [CAD].[dbo].[tCON_Orderbook$orderbook_test] WHERE AC_Status IS NULL AND tsInputDataSent IS NULL AND Address_Validation = 1 AND (EYStand = 0 OR EYStand IS NULL) AND Storno IS NULL", $Connection,$adOpenStatic,$adLockOptimistic)
    # Query updaten: Bei Banken-eConPLUS nicht rausschicken, da es hier aus dem Bankverzeichnis.xlsm kommt, nicht aus InputData
    #search through the Order book and find the Orders with empty status (meaning that the confirmation e-mail has not been sent yet)
    While ($RecordSet.EOF -ne $True -And $RecordSetCount -gt 0) {
        If ([string]::IsNullOrEmpty($RecordSet.Fields.Item("AC_Status").Value)){
            $Order = $RecordSet.Fields.Item("OrderNo").Value
            Write-host "OrderNo: $Order"
            $To1 = $RecordSet.Fields.Item("EngContact").Value
            $To2 = $RecordSet.Fields.Item("EngPartner").Value
            $To2 = $($RecordSet.Fields.Item("EngManager").Value)+','+$To2
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_2").Value) -eq $false) {$To2 = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_2").Value)+','+$To2} 
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_3").Value) -eq $false) {$To2 = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_3").Value)+','+$To2} 
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_4").Value) -eq $false) {$To2 = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_4").Value)+','+$To2}
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_5").Value) -eq $false) {$To2 = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_5").Value)+','+$To2}
            [string[]]$To2 = $To2.Split(',')
            $To3 = $RecordSet.Fields.Item("EngManager").Value
            $Periodend = [datetime]::parseexact($RecordSet.Fields.Item("YearEnd").Value, 'yyyy-MM-dd', $null).ToString("dd.MM.yyyy")
            $Client = $RecordSet.Fields.Item("Client").Value
            $EngOtherContact = $RecordSet.Fields.Item("OtherContact").Value
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_2").Value) -eq $false) {$EngOtherContact = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_2").Value)+','+$EngOtherContact} 
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_3").Value) -eq $false) {$EngOtherContact = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_3").Value)+','+$EngOtherContact} 
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_4").Value) -eq $false) {$EngOtherContact = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_4").Value)+','+$EngOtherContact}
            if ([string]::IsNullOrEmpty($RecordSet.Fields.Item("ADDITIONAL_CONTACT_5").Value) -eq $false) {$EngOtherContact = $($RecordSet.Fields.Item("ADDITIONAL_CONTACT_5").Value)+','+$EngOtherContact}
            $GISID = $RecordSet.Fields.Item("GISID").Value
            $EngCode = $RecordSet.Fields.Item("EngCode").Value
            $Lang = $RecordSet.Fields.Item("DocLang").Value 
            $Confirmation = $RecordSet.Fields.Item("Confirmation").Value 

            # InputTemplate
            $Excel = new-object -comobject excel.application
            $Excel.Visible = $false
            $Excel.EnableEvents = $false 
            $Excel.DisplayAlerts = $false
            if($Lang -eq "DE") {
                    $objWorkbook = $Excel.Workbooks.Open("\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\1_CAD-Adressabgleich Adressenabfrage Mandant.xlsm")            
            }
             if(($Lang -eq "EN") -or ($Lang -eq "ENG")) {
                    $objWorkbook = $Excel.Workbooks.Open("\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\1_CAD-Adressabgleich Adressenabfrage Mandant_EN.xlsm")
            }

            $objWorksheet = $objWorkbook.Worksheets.Item("basic_info")
            $objWorksheet.cells.Item(1,2) = $Order
            $objWorksheet.cells.Item(2,2) = $Periodend
            $objWorksheet.cells.Item(3,2) = $Client
            $objWorksheet.cells.Item(4,2) = $To1
            $objWorksheet.cells.Item(5,2) = $To2
            $objWorksheet.cells.Item(6,2) = $To3
            $objWorksheet.cells.Item(7,2) = $EngOtherContact
            $objWorksheet.cells.Item(8,2) = $GISID
            $objWorksheet.cells.Item(9,2) = $Confirmation
            $objWorksheet.cells.Item(10,2) = $EngCode

            $ExcelSavePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace"
            $OrderFolderexists = Test-Path "$ExcelSavePath\$Order\1. Adressenabgleich”
            If($OrderFolderexists -eq "True") 
            {
                Write-host "Folder exists"
            }
            else 
            {
                New-Item "$ExcelSavePath\$Order\1. Adressenabgleich" -type directory
            }
            $fileName = $GISID.ToString().PadLeft(10,"0") + " 1_CAD-Adressabgleich Adressenabfrage Mandant " + [datetime]::parseexact($RecordSet.Fields.Item("YearEnd").Value, 'yyyy-MM-dd', $null).ToString("yyyyMMdd")
            $objWorkbook.SaveAs($ExcelSavePath + "\" + "$Order" + "\1. Adressenabgleich\" + $fileName, 52) # http://msdn.microsoft.com/en-us/library/bb241279.aspx
            $objWorkbook.Saved = $true
            $objWorkbook.Close()
            Write-host $ExcelSavePath+$Order\$fileName
            $Attachmentdata = $ExcelSavePath + "\" + $Order + "\1. Adressenabgleich\" + $fileName + ".xlsm"
            
            # Subject
            If($Lang -eq "DE")
            {  
               $subject="CAD Adressabgleich Bestellung: $Confirmation $Order für $client"
            }
            if(($Lang -eq "EN") -or ($Lang -eq "ENG"))
            {
               $subject="CAD Address comparison order: $Confirmation $Order for $client"
            }

            # Mail Template
            $Confimailtemplate= "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\$Lang\eCon_AC_ConfiMail.htm"
            $utf8 = New-Object System.Text.utf8encoding
            $confihtmlbody = Get-content $Confimailtemplate -Encoding UTF8
            [string]$confimessage=$confihtmlbody.Replace("[OrderNo]", $Order).Replace("[Client]", $Client).Replace("[GISID]", $GISID).Replace("[YearEnd]", $PeriodEnd).Replace("[Confirmation]", $Confirmation)   

            Send-MailMessage -Encoding $utf8 -FROM $emailbox -To $To1 -Cc "$To2, $To3, $EngOtherContact" -bcc $emailbox -Subject $subject -Body $confimessage -BodyAshtml -SmtpServer "mail-de.ey.net" -Attachments $Attachmentdata, $attLogo, $attBanner

            $RecordSet.Fields.Item("tsInputDataSent").Value = $dateInputSheetSent # updating the Order book with the date of e-mail sent
            $RecordSet.Fields.Item("AC_Status").Value = $nextStatus # updating the Order book with the new status after e-mail has been sent
           
        }
        $RecordSet.MoveNext()
    }

    #Schließen der Verbindung
    $RecordSet.Close() 
    $Connection.Close()
}
catch {
    Write-Warning "ERROR: $($_.Exception.Message)"
    Write-Log "ERROR sendInputDataConAC.ps1  (Productive = $productivemode): $($_.Exception.Message)"
    break
}

#Ende des Skripts
Write-Log "End sendInputDataConAC.ps1  (Productive = $productivemode)"
