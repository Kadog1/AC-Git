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
$file = ("processReadyforAC_Log_" + $date + "_" +$productivemode+".log")
$logfile = $pathLog + $file

Write-Host "$(Get-TimeStamp) Start Script: sendInputDataConAC.ps1"
Write-Log "Start sendInputDataConAC.ps1  (Productive = $productivemode)"

$queryTable = "SELECT * FROM [CAD].[dbo].[tCON_Orderbook$orderbook_test] WHERE SB_Status = 'Ready for AC' AND tsInputDataReceived IS NULL"

try {
    # Connect to CON_Orderbook and count orders of interest
    $adOpenStatic = 3
    $adLockOptimistic = 3
    $Connection = New-Object -com "ADODB.Connection"
    $RecordSet = New-Object -com "ADODB.Recordset"
    $Connection.Open("Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;")
    $RecordSet.Open($queryTable, $Connection,$adOpenStatic,$adLockOptimistic)

    $dateInputSheetReceived = Get-Date -Format G #get the current date and time
    $nextACStatus = "InputDataReceived" #status after ( = conf email has been sent)

    $attLogo ="\\devidvapfl04.ey.net\04em1015\H\Heartbeat-CAD Team\CAD_Brand_Communications\CAD_Logo+Image\EY_Logo_Beam_RGB.png"
    $attBanner ="\\devidvapfl04.ey.net\04em1015\H\Heartbeat-CAD Team\CAD_Brand_Communications\CAD_Signatur+Banner\190807_CAD_Signature-Banner_500px_mittel.png"

    $RecordSetCount=0
    while ($RecordSet.EOF -ne $True) 
    {
        $RecordSetCount = $RecordSetCount+1
        $RecordSet.MoveNext()
    } 
    $RecordSet.Close() 
    Write-Host "OrderCount: $RecordSetCount"


    $RecordSet.Open($queryTable, $Connection,$adOpenStatic,$adLockOptimistic)
    #search through the Order book and find the Orders with empty status (meaning that the confirmation e-mail has not been sent yet)
    While ($RecordSet.EOF -ne $True -And $RecordSetCount -gt 0) {
        If ([string]::IsNullOrEmpty($RecordSet.Fields.Item("tsInputDataReceived").Value)){ 
            # Define Variables
            Write-Host "Processing " $RecordSet.Fields.Item("OrderNo").Value "..."
            Write-Host "Defining Variables..."
            $Order = $RecordSet.Fields.Item("OrderNo").Value
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
            $GISID = $RecordSet.Fields.Item("GISID").Value
            $EngCode = $RecordSet.Fields.Item("EngCode").Value
            $Lang = $RecordSet.Fields.Item("DocLang").Value 
            $Confirmation = $RecordSet.Fields.Item("Confirmation").Value
        
            # Create InputDataSent
            Write-Host "Creating InputDataSent..."
            $Excel = new-object -comobject excel.application
            $Excel.Visible = $false
            $Excel.EnableEvents = $false 
            $Excel.DisplayAlerts = $false
            if($Lang -eq "DE") {
                    $objWorkbook = $Excel.Workbooks.Open("\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\1_CAD-Adressabgleich Adressenabfrage Mandant.xlsm")
                    $sourceFileName = "Bankverzeichnis eConfirmations.xlsm"
            }
                if(($Lang -eq "EN") -or ($Lang -eq "ENG")) {
                    $objWorkbook = $Excel.Workbooks.Open("\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\1_CAD-Adressabgleich Adressenabfrage Mandant_EN.xlsm")
                    $sourceFileName = "Bank directory eConfirmations.xlsm" 
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
            $objWorkbook.SaveAs($ExcelSavePath + "\" + "$Order" + "\1. Adressenabgleich\" + $fileName, 52)
            $objWorkbook.Saved = $true
            $objWorkbook.Close()

            # Fill InputDataSent
            Write-Host "Filling InputDataSent..."
            $SourceWorkBook=$Excel.Workbooks.Open("\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" + "$Order" + "\" + "$sourceFileName")
            $TargetWorkbook=$Excel.workBooks.Open($ExcelSavePath + "\" + "$Order" + "\1. Adressenabgleich\" + $fileName + ".xlsm")

            $SourceWorkBook.WorkSheets.Item(1).Activate()
            $startDisclaimer = $($SourceWorkBook.WorkSheets.Item(1).Range("A1:Z500").Find("Hier sind ALLE Kreditinstitute").Row, `
                                 $SourceWorkBook.WorkSheets.Item(1).Range("A1:Z500").Find("ALL credit institutions with which business relations of any kind").Row | measure -Maximum).Maximum -1
            $lastrow = 10
            for ($i=11; $i -le $startDisclaimer; $i++){
                If(-Not [string]::IsNullOrEmpty($SourceWorkBook.WorkSheets.Item(1).Cells.Range("A$i").value2)){
                    $lastrow = $i
                }
            }
            $rowIndex = 0 
            for ($i=10; $i -le $lastrow; $i++){
                Write-Host $SourceWorkBook.WorkSheets.Item(1).Range("C$i").value2
                If($SourceWorkBook.WorkSheets.Item(1).Range("C$i").value2 -eq "No" -or $SourceWorkBook.WorkSheets.Item(1).Range("C$i").value2 -eq "Nein"){
                    $SourceRange1=$SourceWorkBook.WorkSheets.Item(1).Range("A$($i):B$($i)")
                    $SourceRange2=$SourceWorkBook.WorkSheets.Item(1).Range("D$($i):I$($i)")
                    [void]$SourceRange1.Copy()
                    [void]$TargetWorkBook.worksheets.item("Bank").Range("E$(17 + $rowIndex)").PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteValues)
                    [void]$SourceRange2.Copy()
                    [void]$TargetWorkBook.worksheets.item("Bank").Range("H$(17 + $rowIndex)").PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteValues)
                    $TargetWorkBook.worksheets.item("Bank").cells.Item($(17 + $rowIndex), 12) = ([string]$SourceWorkBook.WorkSheets.Item(1).Range("H$i").value2).split("|")[0]
                    $rowIndex++
                }               
            }

            <#
            $SourceRange1=$SourceWorkBook.WorkSheets.Item(1).Range("A10:B$lastrow")
            $SourceRange2=$SourceWorkBook.WorkSheets.Item(1).Range("D10:I$lastrow")
            [void]$SourceRange1.Copy()
            [void]$TargetWorkBook.worksheets.item("Bank").Range("E17").PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteValues)
            [void]$SourceRange2.Copy()
            [void]$TargetWorkBook.worksheets.item("Bank").Range("H17").PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteValues)
            #>
          

            [void]$SourceWorkBook.Close()
            [void]$TargetWorkBook.SaveAs($ExcelSavePath + "\" + "$Order" + "\" + $fileName, 52)
            $TargetWorkBook.Saved = $true
            [void]$TargetWorkBook.Close()

            $Excel.DisplayAlerts = $true

            [void]$Excel.Quit()

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null

            $RecordSet.Fields.Item("tsInputDataReceived").Value = $dateInputSheetReceived # updating the Order book with the date of e-mail sent
            $RecordSet.Fields.Item("AC_Status").Value = $nextACStatus # updating the Order book with the new status after e-mail has been sent
        }
        $RecordSet.MoveNext()
    }
    #Schließen der Verbindung
    $RecordSet.Close() 
    $Connection.Close()
}
catch {
    Write-Warning "ERROR: $($_.Exception.Message)"
    Write-Log "ERROR processReadyforAC.ps1  (Productive = $productivemode): $($_.Exception.Message)"
    break
}

#Ende des Skripts
Write-Host "$(Get-TimeStamp) End Script: processReadyforAC.ps1"
Write-Log "End processReadyforAC.ps1  (Productive = $productivemode)"
