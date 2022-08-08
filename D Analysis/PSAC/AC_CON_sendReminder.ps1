#clear

function Write-Log([string]$logtext)
{
    $logdate = get-date -format "yyyy-MM-dd HH:mm:ss"
    $logtext = $logtext
    $text = "["+$logdate+"] - " + $logtext
    Write-Host $text
    $text >> $logfile
}

function getReminderOrderNoList {
    param(
    [int]$ReminderBarrierDays
    )
    $datetimeNow = Get-Date -Format "yyyy-MM-dd HH:mm:00"
    $returnQuery = "SELECT OrderNo, POrderNo, AC_Status, AC_Preparer, DocLang, Confirmation, client, GISID, YearEnd, engCode, engcontact, engmanager, OtherContact
        FROM AC_SELECT('AC','%')
        WHERE (AC_Status = 'TeamApprovalSent' AND tsTeamApprovalSent IS NOT NULL AND tsReminderTeamApprovalSent IS NULL AND DATEDIFF(DAY, tsTeamApprovalSent, '$datetimeNow') > $ReminderBarrierDays) OR
        (AC_Status = 'TeamApprovalReceived' AND tsTeamApprovalReceived IS NOT NULL AND tsReminderTeamApprovalReceived IS NULL AND DATEDIFF(DAY, tsTeamApprovalReceived, '$datetimeNow') > $ReminderBarrierDays)
        UNION
        SELECT OrderNo, POrderNo, AC_Status, AC_Preparer, DocLang, Confirmation, client, GISID, YearEnd, engCode, engcontact, engmanager, OtherContact
        FROM AC_SELECT('CO','%')
        WHERE (AC_Status = 'TeamApprovalSent' AND tsTeamApprovalSent IS NOT NULL AND tsReminderTeamApprovalSent IS NULL AND DATEDIFF(DAY, tsTeamApprovalSent, '$datetimeNow') > $ReminderBarrierDays) OR
        (AC_Status = 'TeamApprovalReceived' AND tsTeamApprovalReceived IS NOT NULL AND tsReminderTeamApprovalReceived IS NULL AND DATEDIFF(DAY, tsTeamApprovalReceived, '$datetimeNow') > $ReminderBarrierDays)"
    Write-Output $returnQuery
}

[boolean]$productivemode = $True
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
$thisScriptPath = Get-Location
$date = get-date -format "yyyy-MM-dd"
$file = "AC_CON_sendReminder_Log_" + $date + "_" +$productivemode+".log"
$logfile = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich\A Order entry\LogFiles\AC_CON_sendReminder\$file"

$dateReminderSent = Get-Date -Format "yyyy-MM-dd HH:mm"
$attLogo ="\\devidvapfl04.ey.net\04em1015\H\Heartbeat-CAD Team\CAD_Brand_Communications\CAD_Logo+Image\EY_Logo_Beam_RGB.png"
$attBanner ="\\devidvapfl04.ey.net\04em1015\H\Heartbeat-CAD Team\CAD_Brand_Communications\CAD_Signatur+Banner\190807_CAD_Signature-Banner_500px_mittel.png"

$ReminderBarrierDays = 15

Write-Log "Start AC_CON_sendReminder.ps1 (Productive = $productivemode)"
try {
    $adOpenStatic = 3
    $adLockOptimistic = 3
    $Connection = New-Object -com "ADODB.Connection"
    $RecordSet = New-Object -com "ADODB.Recordset"
    $RecordSetPOrderNo = New-Object -com "ADODB.Recordset"
    $RecordSetUpdate = New-Object -com "ADODB.Recordset"
    $Connection.Open("Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;")
    $RecordSet.Open("$(getReminderOrderNoList($ReminderBarrierDays))", $Connection,$adOpenStatic,$adLockOptimistic)
    $RecordSetCount=0
    while ($RecordSet.EOF -ne $True){
            $RecordSetCount = $RecordSetCount+1
            $RecordSet.MoveNext()
    } 
    $RecordSet.Close() 
    $RecordSet.Open("$(getReminderOrderNoList($ReminderBarrierDays))", $Connection,$adOpenStatic,$adLockOptimistic)
    # Query updaten: Bei Banken-eConPLUS nicht rausschicken, da es hier aus dem Bankverzeichnis.xlsm kommt, nicht aus InputData
    #search through the Order book and find the Orders with empty status (meaning that the confirmation e-mail has not been sent yet)
    While ($RecordSet.EOF -ne $True -And $RecordSetCount -gt 0){
        $OrderNo = $RecordSet.Fields.Item("OrderNo").Value
        $AC_Status = $RecordSet.Fields.Item("AC_Status").Value
        Write-Host "Sending Reminder-Mail for $OrderNo - $AC_Status..."
        
        $attachments = @($attLogo, $attBanner)
        $POrderNo = $RecordSet.Fields.Item("POrderNo").Value
        $lang = $RecordSet.Fields.Item("DocLang").Value
        $langFolder = "German"
        If($lang -eq "EN"){$langFolder = "English"}
        $confirmation = $RecordSet.Fields.Item("Confirmation").Value 
        $client = $RecordSet.Fields.Item("client").Value
        $orderNo = $RecordSet.Fields.Item("OrderNo").Value
        $pOrderNo = $RecordSet.Fields.Item("POrderNo").Value
        $GISID = $RecordSet.Fields.Item("GISID").Value
        $yearEnd = $RecordSet.Fields.Item("YearEnd").Value
        $engCode = $RecordSet.Fields.Item("engCode").Value
        $ACPreparer = $RecordSet.Fields.Item("AC_Preparer").Value
        $ACStatus = $RecordSet.Fields.Item("AC_Status").Value

        # Contact details
        $engContact = $RecordSet.Fields.Item("EngContact").Value
        $engManager = $RecordSet.Fields.Item("EngManager").Value
        $otherContact = $RecordSet.Fields.Item("OtherContact").Value
        
        # For TeamApprovalReceived: Aggregate Confirmation, OrderNo per POrderNo
        If ($ACStatus -eq "TeamApprovalReceived"){
            If (-not [string]::IsNullOrEmpty($POrderNo)){
                $listOrderNo = ""
                $listConfirmation = ""
                $RecordSetPOrderNo.Open("SELECT OrderNo, Confirmation FROM ($(getReminderOrderNoList($ReminderBarrierDays))) AS T WHERE POrderNo = '$POrderNo' AND AC_Status = 'TeamApprovalReceived'", $Connection,$adOpenStatic,$adLockOptimistic)
                While ($RecordSetPOrderNo.EOF -ne $True) {
                    $listOrderNo = $listOrderNo + "$($RecordSetPOrderNo.Fields.Item("OrderNo").Value), "
                    $listConfirmation = $listConfirmation + "$($RecordSetPOrderNo.Fields.Item("Confirmation").Value), "
                    $RecordSetUpdate.Open("EXECUTE CAD.dbo.AcCon_UpdateOrderbook @OrderNo = '$($RecordSetPOrderNo.Fields.Item("OrderNo").Value)', @Fields = 'tsReminder$ACStatus', @Values = '$dateReminderSent'", $Connection,$adOpenStatic,$adLockOptimistic)
                    $RecordSetPOrderNo.MoveNext()
                }
                $RecordSetPOrderNo.Close()
                try{
                    $orderNo = $listOrderNo.SubString(0, $($listOrderNo.length - 2))
                    $confirmation = $listConfirmation.SubString(0, $($listConfirmation.length - 2))
                } catch {
                    $RecordSet.MoveNext()
                    continue
                }
            } else {
                $RecordSetUpdate.Open("EXECUTE CAD.dbo.AcCon_UpdateOrderbook @OrderNo = '$($RecordSet.Fields.Item("OrderNo").Value)', @Fields = 'tsReminder$ACStatus', @Values = '$dateReminderSent'", $Connection,$adOpenStatic,$adLockOptimistic)
            }
        # For TeamApprovalSent: Attach Workbook and 3_Team Approval
        } elseif ($ACStatus -eq "TeamApprovalSent"){
            $gisidCode = '{0:d10}' -f [int]$GISID
            $yearEnd = $yearEnd -replace '-',''
            $filenameTeamApproval = "$gisidCode 3_CAD-Adressabgleich Team Approval_Template $yearEnd.xlsm"
            $filenameWorkbook = "$gisidCode CAD Confirmations Workbook $yearEnd.xlsm"
            foreach ($filenameAttachment in @($filenameTeamApproval, $filenameWorkbook)){
                $fileFullPath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\$orderNo\2. CAD_Abgleich\$filenameAttachment"
                if (Test-Path $fileFullPath) {
                    $attachments = $attachments + $fileFullPath
                } else {
                    Write-Log "$filenameAttachment not found for $orderNo"
                }
            }
            $RecordSetUpdate.Open("EXECUTE CAD.dbo.AcCon_UpdateOrderbook @OrderNo = '$($RecordSet.Fields.Item("OrderNo").Value)', @Fields = 'tsReminder$ACStatus', @Values = '$dateReminderSent'", $Connection,$adOpenStatic,$adLockOptimistic)
        }
        If ($lang -eq "EN" -and -not [string]::IsNullOrEmpty($confirmation)){
            $confirmation = $confirmation.Replace("Debitoren", "Debtor").Replace("Kreditoren", "Creditor").Replace("Rechtsanwalt", "Law Advisor").Replace("Steuerberater", "Tax Advisor").Replace("Banken", "Bank")
        }
        # Subject
        If($lang -eq "DE" -and $ACStatus -eq "TeamApprovalSent"){ $subject = "REMINDER: Team Approval für den CAD Adressabgleich für $confirmation $orderNo für $client"}
        elseif($lang -eq "EN" -and $ACStatus -eq "TeamApprovalSent"){ $subject = "REMINDER: Team Approval for Address comparison service for $confirmation $orderNo for $client"}
        elseif($lang -eq "DE" -and $ACStatus -eq "TeamApprovalReceived"){ $subject = "REMINDER: Canvas Einladung für den CAD Adressabgleich für $confirmation $orderNo für $client ausstehend"}
        elseif($lang -eq "EN" -and $ACStatus -eq "TeamApprovalReceived"){ $subject = "REMINDER: Canvas invitation for Address comparison service for $confirmation $orderNo for $client"}
        # Body
        $mailTemplateReminder = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich\B Mail Templates\$langFolder\Reminder\$($ACStatus)_$($lang).htm"
        $utf8 = New-Object System.Text.utf8encoding
        $mailHTMLbody = Get-content $mailTemplateReminder -Encoding UTF8
        [string]$reminderMessage = $mailHTMLbody.Replace("[OrderNo]", $orderNo).Replace("[Client]", $client).Replace("[GISID]", $GISID).Replace("[Confirmation]", $confirmation).Replace("[CAD Preparer]", $ACPreparer)
        $subject = $subject.Replace("  ", " ")
        $reminderMessage = $reminderMessage.Replace("( ", "(")
        $ccContacts = @($engManager)
        If(-not [string]::IsNullOrEmpty($otherContact)){
            $ccContacts = @($otherContact, $engManager)
        }
        Send-MailMessage -Encoding $utf8 -FROM $emailbox -To $engContact -Cc $ccContacts -bcc $emailbox -Subject $subject -Body $reminderMessage -BodyAshtml -SmtpServer "mail-de.ey.net" -Attachments $attachments
        $RecordSet.MoveNext()
        Start-Sleep -s 5
    }
    $Connection.Close()
}
catch {
    Write-Warning "ERROR: $($_.Exception.Message)"
    Write-Log "ERROR einmaligerReminderFebruar.ps1  (Productive = $productivemode): $($_.Exception.Message)"
    break
}

#Ende des Skripts
Write-Log "End AC_CON_sendReminder.ps1  (Productive = $productivemode)"
