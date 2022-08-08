#CADStorno procedure  @KA 20.06.2022

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

[boolean]$productivemode = $False

if($productivemode -eq $True) 
{
    $testenv=""
    $orderbook_test = ""
}
Else 
{
    $testenv=" Testumgebung"
    $orderbook_test = "_TEST"
}

$datetimeNow = Get-Date -Format "yyyy-MM-dd HH:mm:00"
$Validperiod = 15  

# GetCADStornoOrderNoList


# Team approval TAC status update 

function GetCADStornoOrderNoListTAC{
    param(
    [int]$Validperiod
    )
    $query1 = "SELECT OrderNo, AC_Status, AC_Preparer,tsReminderTeamApprovalSent,tsReminderTeamApprovalReceived
        FROM tAC_Orderbook$orderbook_test WHERE AC_Status = 'TeamApprovalSent' AND tsReminderTeamApprovalSent IS NOT NULL AND DATEDIFF(day,tsReminderTeamApprovalSent,'$datetimeNow') > $Validperiod  OR 
		(AC_Status = 'TeamApprovalReceived'AND tsReminderTeamApprovalReceived IS NOT NULL AND DATEDIFF(day,tsReminderTeamApprovalReceived,'$datetimeNow') > $Validperiod)      
        Order by tsReminderTeamApprovalSent desc"
            
    Write-Output $query1
}

                                   
$instancelocal = "DEFRNVMPDWASQ04\INST02"

write-host "Productive = $productivemode"
$pathLog = "C:\TestLocal\AC adressabgliech\Update procedure\logfiles$testenv"
$date = get-date -format "yyyy-MM-dd"
$file = (" UPDATECADStrono " + $date + "_" +$productivemode+".log")
$logfile = $pathLog + $file

Write-Log "Started runTAsent.ps1 (Productive = $productivemode) OrderNolist:"   
$newStatus = "CADStorno"              

try{
    $adOpenStatic = 3
    $adLockOptimistic = 3
    $Connection = New-Object -com "ADODB.Connection"
    $RecordSet = New-Object -com "ADODB.Recordset"
    $RecordSetUpdate = New-Object -com "ADODB.Recordset"
    $Connection.Open("Provider=SQLNCLI11;Data Source=$instancelocal; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;")
    $RecordSet.Open("$(GetCADStornoOrderNoListTAC($Validperiod))", $Connection,$adOpenStatic,$adLockOptimistic)
    $CountRecordset=0
    while ($RecordSet.EOF -ne $True){
            $OrderNo = $RecordSet.Fields.Item("OrderNo").Value
            $AC_Status = $RecordSet.Fields.Item("AC_Status").Value
            $ACPreparer = $RecordSet.Fields.Item("AC_Preparer").Value
            $tsreminderTAsent = $RecordSet.Fields.Item("tsReminderTeamApprovalSent").Value
            # Count rowitems while Recordset is open
            $CountRecordset = $CountRecordset+1
            Write-Host $CountRecordset ($OrderNo)
            Write-Log ($OrderNo)
            $newStatus = "CADStorno"
            #updating status for table tAC_Orderbook
            Write-Log "updating for $OrderNo $AC_Status ($ACPreparer)"
            If ($AC_Status -like "TeamApprovalSent" -or "TeamApprovalReceived" -and $OrderNo -like "AC0000" -or "CON000"){                                      #"TeamApprovalReceived" -or,-and $OrderNo -like "AC0000" -or "CON000"#AND tsReminderTeamApprovalSent IS NOT NULL AND tsReminderTeamApprovalReceived IS NOT NULL
                    $newStatus = "CADStorno"
                    $RecordSetUpdate = "Update CAD.dbo.tAC_Orderbook$orderbook_test SET AC_Status = '$newStatus' WHERE OrderNo = '$OrderNo'"
                    Invoke-Sqlcmd -Query $RecordSetUpdate -ServerInstance "$instancelocal" 
                    Write-Log "updated $OrderNo to $newStatus"                   
            }
                     
            $RecordSet.MoveNext()     
            
    } 
               
    $RecordSet.Close()
    Write-Host $CountRecordset "Orders updated"
    Write-Log "$CountRecordset Orders were updated"
                                            
}

catch {
    Write-Warning "ERROR: $($_.Exception.Message)"
    Write-Log "ERROR CADstorno update (Productive = $productivemode): $($_.Exception.Message)"
    break
}

Write-Log "end of script 1"



# Team approval Tcon status update 

function GetCADStornoOrderNoListTcon {
    param(
    [int]$Validperiod
    )
    $query2 = "SELECT OrderNo, AC_Status, AC_Preparer,tsReminderTeamApprovalSent,tsReminderTeamApprovalReceived
        FROM tCON_Orderbook$orderbook_test WHERE AC_Status = 'TeamApprovalSent' AND tsReminderTeamApprovalSent IS NOT NULL AND DATEDIFF(day,tsReminderTeamApprovalSent,'$datetimeNow') > $Validperiod  OR 
		(AC_Status = 'TeamApprovalReceived'AND tsReminderTeamApprovalReceived IS NOT NULL AND DATEDIFF(day,tsReminderTeamApprovalReceived,'$datetimeNow') > $Validperiod)      
        Order by tsReminderTeamApprovalSent desc"
            
    Write-Output $query2
}

Write-Log "Started runTArecieved.ps2 (Productive = $productivemode) OrderNolist:"
# Invoke-Sqlcmd -Query $query2  -ServerInstance "$instancelocal" | Format-Table

try{
    
    $adOpenStatic = 3
    $adLockOptimistic = 3
    $Connection = New-Object -com "ADODB.Connection"
    $RecordSet = New-Object -com "ADODB.Recordset"
    $RecordSetUpdate2 = New-Object -com "ADODB.Recordset"
    $Connection.Open("Provider=SQLNCLI11;Data Source=$instancelocal; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;")
    $RecordSet.Open("$(GetCADStornoOrderNoListTcon($Validperiod))", $Connection,$adOpenStatic,$adLockOptimistic)
    $CountRecordset=0
    while ($RecordSet.EOF -ne $True){
            $OrderNo = $RecordSet.Fields.Item("OrderNo").Value
            $AC_Status = $RecordSet.Fields.Item("AC_Status").Value
            $ACPreparer = $RecordSet.Fields.Item("AC_Preparer").Value
            $tsreminderTAsent = $RecordSet.Fields.Item("tsReminderTeamApprovalSent").Value
            # Count rowitems while Recordset is open
            $CountRecordset = $CountRecordset+1
            Write-Host $CountRecordset ($OrderNo)
            Write-Log ($OrderNo)
            
            #updating status for table tCON_Orderbook
            Write-Log "updating for $OrderNo $AC_Status ($ACPreparer)"
            If ($AC_Status -like "TeamApprovalSent" -or "TeamApprovalReceived" -and $OrderNo -like "AC0000" -or "CON000"){                                      
                    $RecordSetUpdate2 = "Update CAD.dbo.tCON_Orderbook$orderbook_test SET AC_Status = '$newStatus' WHERE OrderNo = '$OrderNo'"
                    Invoke-Sqlcmd -Query $RecordSetUpdate2 -ServerInstance "$instancelocal"
                    Write-Log "updated $OrderNo to $newStatus"                        
            }                                    
            $RecordSet.MoveNext()         
    } 
             
    $RecordSet.Close()
    Write-Host $CountRecordset "Orders updated"
    Write-Log "$CountRecordset Orders were updated"
    $Connection.Close()                                         
}

catch {
    Write-Warning "ERROR: $($_.Exception.Message)"
    Write-Log "ERROR CADstorno update (Productive = $productivemode): $($_.Exception.Message)"
    break
}

Write-Log "end of script 2"


Write-Host "end of script"