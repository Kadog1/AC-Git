Attribute VB_Name = "Documentation"
Sub process_TeamApproval()

    'This sub loads input data into Tab [Input_Adressdaten]

    'Set this workbook objects
    Dim wb As Workbook
    Dim ws_TeamApproval As Worksheet
    Dim ws_basic_info As Worksheet

    Set wb = ThisWorkbook
    Set ws_TeamApproval = wb.Worksheets("Team Approval Documentation")
    Set ws_basic_info = wb.Worksheets("basic_info")
    
    'Get path to input file
    Dim str_inputpath As String, str_fileName As String, savePath As String
    
    Dim LogDir As String
    LogDir = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & orderNo & "\3. Team Approval\"
    Set fso = CreateObject("Scripting.FileSystemObject")
                    
    Set Source = fso.GetFolder(LogDir)
    Dim FileName As String

    For Each File In Source.Files
        
        FileName = File.Name
        Dim DestinationPathFile As String
        DestinationPathFile = ""
        DestinationPathFile = LogDir & "\" & FileName
                        
        If InStr(UCase(DestinationPathFile), ".msg") > 0 Then
        
            str_inputpath = ""
    
            ws_TeamApproval.OLEObjects.Add FileName:=DestinationPathFile, Link:=False, DisplayAsIcon:=False, Left:=100, Top:=100, Width:=1000.3, Height:=10
    
            wb.Save
            
            Exit Sub
            
        End If
    Next
        
End Sub






