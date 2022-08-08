# Version 1
$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $TRUE
            
$WorkBook = $objExcel.Workbooks.Open("$pathAC\processTeamApprovalReceived.xlsm")

$objExcel.run("RunMain")
