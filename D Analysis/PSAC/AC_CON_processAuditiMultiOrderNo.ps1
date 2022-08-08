### This script processes the multi order number subjects sent from Auditi to CAD Adressabgleich
### PFP 2021/12/14
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
    $emailbox = "adressabgleich@de.ey.com"
}

write-host "Productive = $productivemode"
$pathLog = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich$testenv\A Order entry\LogFiles\"
$date = get-date -format "yyyy-MM-dd"
$file = ("processMultiOrderNoAuditi_Log_" + $date + "_" +$productivemode+".log")
$logfile = $pathLog + $file

Write-Log "$(Get-TimeStamp) Start processMultiOrderNoAuditi.ps1  (Productive = $productivemode)"

try {
        #Ansteuern des Shared Inbox Adressabgleich@de.ey.com
        $olFolderInbox = 6
        $outlook = new-object -comobject outlook.application;
        $namespace = $outlook.GetNamespace(“MAPI”);
        $recipient = $namespace.CreateRecipient($emailbox)
        $inbox = $namespace.GetSharedDefaultFolder($recipient, $olFolderinbox) 
        $TeamReply = $inbox.Folders | where-object { $_.name -eq "Team Reply" }
        $TeamReplyProcessed = $inbox.Folders | where-object {$_.name -eq "Team Reply Processed"}
        $TeamReplyNachlieferung = $inbox.Folders | where-object {$_.name -eq "Team Reply Nachlieferung"}
        $ordercount=$TeamReply.items.count 
        write-host "Neue Bestellungen: $ordercount"

        # Connect to data base
        $adOpenStatic = 3
        $adLockOptimistic = 3
        $Connection = New-Object -com "ADODB.Connection"
        $RecordSet = New-Object -com "ADODB.Recordset"
        $Connection.Open("Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;")
        
        $DBconfirmation = @("Banken", "Steuerberater", "Rechtsanwalt", "Kreditoren", "Debitoren", "Sonstige")
        $AuditiConfirmation = @("Bank", "Steuerberater", "Rechtsanwalt", "Kreditor", "Debitor", "Sonstige")
        
        $ExcelSavePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace"

        if($ordercount -gt 0){
            # Search for multi order number emails
            $TeamReply.Items | foreach {
                $TeamReplyItem = $_
                $subject = $TeamReplyItem.subject
                If ($subject -notlike "*CAD Adressabgleich*" -or ($subject -notlike "*,CON00*")){
                    return
                }
                Write-Log $Subject
                $listOrderNo = @(@($subject -split " ")[2] -split ",")
                $nachlieferung = $false
                $listOrderNo|ForEach {
                    $OrderNo = $_
                    Write-Log "Processing OrderNo $listOrderNo - $OrderNo"

                    # Retrieve order information
                    $queryOrderNo = "SELECT * FROM [CAD].[dbo].[tCON_Orderbook] WHERE OrderNo = '$OrderNo'"
                    $RecordSet.Open($queryOrderNo, $Connection,$adOpenStatic,$adLockOptimistic)
                    $GISID = $RecordSet.Fields.Item("GISID").Value
                    $Client = $RecordSet.Fields.Item("Client").Value
                    $Confirmation = $RecordSet.Fields.Item("Confirmation").Value
                    $idxConfirmation = (0..($DBconfirmation.Count-1)) | where {$DBconfirmation[$_] -eq $Confirmation}
                    $ConfirmationAuditi = $AuditiConfirmation[$idxConfirmation]

                    # Open excel attachment
                    $savePath = (Join-Path $ExcelSavePath $OrderNo )
                    $OrderFolderexists = "$ExcelSavePath\$OrderNo\2. CAD_Abgleich”
                    If((Test-Path $OrderFolderexists) -eq $false) {
                        New-Item $OrderFolderexists -type directory| Out-Null
                    }
                    $TeamReplyItem.attachments(1).saveasfile((Join-Path (Join-Path $savePath "2. CAD_Abgleich") $TeamReplyItem.attachments(1).filename))
                    $Excel = new-object -comobject excel.application
                    $Excel.Visible = $false
                    $Excel.EnableEvents = $false 
                    $Excel.DisplayAlerts = $false
                    $attachmentWorkbook = $Excel.Workbooks.Open((Join-Path (Join-Path $savePath "2. CAD_Abgleich") $TeamReplyItem.attachments(1).filename))
                    # find lastrow
                    $firstRow = $attachmentWorkbook.WorkSheets.Item(1).Range("A1:Z500").Find("Mit der Adresse ver-bundene Dienstleistung").Row + 2
                    $lastrow = $firstRow
                    for ($i= $firstRow + 1; $i -le 200; $i++){
                        If(-Not [string]::IsNullOrEmpty($attachmentWorkbook.WorkSheets.Item(1).Cells.Range("C$i").value2)){
                            $lastrow = $i
                        }
                    }
                    $counter = 0; $deleteCount = 0
                    Write-Log $ConfirmationAuditi
                    for ($i = 17; $i -le $lastrow; $i++){
                        $counter++  
                        If (($attachmentWorkbook.WorkSheets.Item(1).Cells.Range("C$i").value2 -ne $ConfirmationAuditi -and -Not [string]::IsNullOrEmpty($attachmentWorkbook.WorkSheets.Item(1).Cells.Range("C$i").value2) -and `
                            $counter -le ($lastrow - $firstRow + 1)) -or ($attachmentWorkbook.WorkSheets.Item(1).Cells.Range("E$i").value2 -eq "Confirmation.com")){                                              
                            $SourceRangeCopy = $attachmentWorkbook.WorkSheets.Item(1).Range("C$($i + 1):L$($lastrow)")
                            [void]$SourceRangeCopy.Copy()
                            [void]$attachmentWorkbook.worksheets.item(1).Range("C$i").PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteValues)
                            
                            $SourceRangeLastRow = $attachmentWorkbook.WorkSheets.Item(1).Range("C$($lastrow - $deleteCount):L$($lastrow - $deleteCount)")
                            [void]$SourceRangeLastRow.Clear()
                            $r = 255; $g = 245; $b = 153;
                            $rgb = $r + ($g * 256) + ($b * 256 * 256)
                            $SourceRangeLastRow.Interior.Color = $rgb
                            $deleteCount++
                            $i = $i - 1
                        }
                    }
                    [void]$attachmentWorkbook.WorkSheets.Item(1).Cells.Range("C$($firstrow):C$($lastrow)").Replace("Rechtsanwalt","Rechtsberater") 
                    [void]$attachmentWorkbook.WorkSheets.Item(1).Columns("E:E").Insert()
                    # basic info
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
                    $EngOtherContact = $RecordSet.Fields.Item("OtherContact").Value
                    $EngCode = $RecordSet.Fields.Item("EngCode").Value
                    $Lang = $RecordSet.Fields.Item("DocLang").Value
                    If ($Lang -eq "EN"){
                        $attachmentWorkbook.WorkSheets.Item(1).cells.Item(15,5) = "Additional address information"
                        $attachmentWorkbook.WorkSheets.Item(1).Name = "Address check"
                    }

                    $attachmentWorkbook.Worksheets.Item("basic_info").cells.Item(1,2) = $OrderNo
                    $attachmentWorkbook.Worksheets.Item("basic_info").cells.Item(2,2) = $Periodend
                    $attachmentWorkbook.Worksheets.Item("basic_info").cells.Item(3,2) = $Client
                    $attachmentWorkbook.Worksheets.Item("basic_info").cells.Item(4,2) = $To1
                    $attachmentWorkbook.Worksheets.Item("basic_info").cells.Item(5,2) = $To2
                    $attachmentWorkbook.Worksheets.Item("basic_info").cells.Item(6,2) = $To3
                    $attachmentWorkbook.Worksheets.Item("basic_info").cells.Item(7,2) = $EngOtherContact
                    $attachmentWorkbook.Worksheets.Item("basic_info").cells.Item(8,2) = $GISID
                    $attachmentWorkbook.Worksheets.Item("basic_info").cells.Item(9,2) = $Confirmation
                    $attachmentWorkbook.Worksheets.Item("basic_info").cells.Item(10,2) = $EngCode
                    
                    $fileName = $GISID.ToString().PadLeft(10,"0") + " 1_CAD-Adressabgleich Adressenabfrage Mandant " + [datetime]::parseexact($RecordSet.Fields.Item("YearEnd").Value, 'yyyy-MM-dd', $null).ToString("yyyyMMdd")

                    # Keine Nachlieferung?
                    $AC_Status = $RecordSet.Fields.Item("AC_Status").Value
                    $nachlieferung = $false
                    If ($AC_Status -eq "" -Or $AC_Status -eq "WaitingForInputData"){
                        Write-Log "Keine Nachlieferung $OrderNo. AC Status: $AC_Status"
                        $attachmentWorkbook.SaveAs("$ExcelSavePath\$OrderNo\2. CAD_Abgleich\$fileName", 51) # http://msdn.microsoft.com/en-us/library/bb241279.aspx
                        $attachmentWorkbook.Saved = $true
                        $attachmentWorkbook.close($false)
                        $RecordSet.Fields.Item("AC_Status").Value = "InputDataAvailable"
                        $RecordSet.Fields.Item("tsInputDataAvailable").Value = Get-Date -Format G
                    } ElseIf (($AC_Status -eq "TeamApprovalSent" -Or $AC_Status -eq "TeamApprovalReceived" -Or $AC_Status -eq "CanvasDone") -and $RecordSet.Fields.Item("Tool").Value -eq "eConfirmations"){
                        $attachmentWorkbook.close($false)
                        Write-Log "Keine Speicherung $OrderNo. AC Status: $AC_Status"
                    } ElseIf ($AC_Status -eq "InputDataReceived" -Or $AC_Status -eq "InProgress" -Or $AC_Status -eq "ReadyForReview" -Or $AC_Status -eq "TeamApprovalSent" -Or $AC_Status -eq "TeamApprovalReceived" -Or $AC_Status -eq "CanvasDone"){
                        $newStatus = 'InputDataAvailable'
                        If($RecordSet.Fields.Item("Tool").Value = "eConfirmations"){
                            $newStatus = $AC_Status
                        }
                        # Open excel attachment
                        $folderPrevious = "$ExcelSavePath\$OrderNo\2. CAD_Abgleich\Previous InputDatenSheets”
                        If((Test-Path $folderPrevious) -eq $false) {
                            New-Item $folderPrevious -type directory| Out-Null
                        }
                        If((Test-Path "$ExcelSavePath\$OrderNo\2. CAD_Abgleich\$($fileName).xlsx") -eq $true) {
                            Move-Item -Path "$ExcelSavePath\$OrderNo\2. CAD_Abgleich\$($fileName).xlsx" -Destination $folderPrevious
                            $datetimeSubfix = ([datetime]::now).tostring("yyyymmdd_HHmm")
                            Rename-Item -Path "$folderPrevious\$($fileName).xlsx" -NewName "$($fileName)_$($datetimeSubfix).xlsx"
                        }
                        $attachmentWorkbook.SaveAs("$ExcelSavePath\$OrderNo\2. CAD_Abgleich\$fileName", 51) # http://msdn.microsoft.com/en-us/library/bb241279.aspx
                        $attachmentWorkbook.Saved = $true
                        $attachmentWorkbook.close($false)
                        Write-Log "Nachlieferung $OrderNo. AC Status: $AC_Status"
                        
                        If(-Not [string]::IsNullOrEmpty($RecordSet.Fields.Item("AC_Preparer").Value)){
                            $nachlieferung = $true
                            $Confimailtemplate= "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich\B Mail Templates\German\AC_StatusMail.htm"
                            $subject= "Neue Adressdaten: $OrderNo"
                            $utf8 = New-Object System.Text.utf8encoding
                            $confihtmlbody = Get-content $Confimailtemplate -Encoding UTF8
                            [string]$confimessage = $confihtmlbody.Replace("[orderNo]", $Order).Replace("[client]", $Client).Replace("[GISID]", $GISID)
                            $preparer = $($RecordSet.Fields.Item("AC_Preparer").Value).Replace("ä", "ae").Replace("ö", "oe").Replace("ü", "ue")
                            $preparerwoSpace = $preparer.Replace(" ", "")
                            If (($preparer.length - $preparerwoSpace.length) -gt 1){
					            $preparerEmail = "$(@($preparer -split " ")[0]).$(@($preparer -split " ")[1].Substring(0,1)).$(@($preparer -split " ")[2])@de.ey.com"
                            } Else {
                                $preparerEmail = "$(@($preparer -split " ")[0]).$(@($preparer -split " ")[1])@de.ey.com"
                            }
                            Write-Log $preparerEmail
                            $emailAttachment = "$ExcelSavePath\$OrderNo\2. CAD_Abgleich\$($fileName).xlsx"
                            Send-MailMessage -Encoding $utf8 -FROM $emailbox -To $preparerEmail -Subject $subject -Body $confimessage -BodyAshtml -SmtpServer "mail-de.ey.net" -Attachments $emailAttachment
                        }
                        $RecordSet.Fields.Item("AC_Status").Value = "InputDataAvailable"
                        
                    } Else {
                        $attachmentWorkbook.close($false)
                        Write-Log "Keine Speicherung $OrderNo. AC Status: $AC_Status"
                    }
                    [void]$Excel.Quit()
                    $RecordSet.Update()
                    $RecordSet.Close()
                }
                if($nachlieferung){
                    [void]$TeamReplyItem.move($TeamReplyNachlieferung)
                } Else {
                    [void]$TeamReplyItem.move($TeamReplyProcessed)
                }
            }        
            $Connection.Close()
        }
}
catch {
    Write-Warning "ERROR: $($_.Exception.Message)"
    Write-Log "ERROR processMultiOrderNoAuditi.ps1  (Productive = $productivemode): $($_.Exception.Message)"
    break
}
#Ende des Skripts
Write-Log "$(Get-TimeStamp) End processMultiOrderNoAuditi.ps1  (Productive = $productivemode)"
