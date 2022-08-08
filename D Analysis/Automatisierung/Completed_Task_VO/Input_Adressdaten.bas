Attribute VB_Name = "Input_Adressdaten"
Option Explicit

Sub LoadInput_Click(str_inputpath As String, str_fileName As String)

    'This sub loads input data into Tab [Input_Adressdaten]
    'Parameter information will be loaded from basic_info 1_CAD-Adressabgleich Adressenabfrage Mandant workbook
    'ActiveSheet.Shapes.Item("Button 2").Visible = False Hide/Unhide original button

    'Set this workbook objects
    Dim wb As Workbook
    Dim ws_InputAdressdaten As Worksheet
    Dim ws_basic_info As Worksheet

    Set wb = ThisWorkbook
    Set ws_InputAdressdaten = wb.Worksheets("Input Adressdaten")
    Set ws_basic_info = wb.Worksheets("basic_info")
    
    'Get path to input file
    'Dim str_inputpath As String, str_fileName As String, savePath As String
    Dim savePath As String
    
    Dim arr_directory() As String
'    arr_directory = Split(returnDirectoryInformation, "|")
'    str_inputpath = arr_directory(0)
'    str_fileName = arr_directory(1)
'
    savePath = str_inputpath
    
    If check_WBopen(str_inputpath & str_fileName) Then
    
        'Open Inputfile
    
        Dim wb_input As Workbook: Set wb_input = Workbooks(str_fileName)
        Dim ws_Input_basic_info As Worksheet: Set ws_Input_basic_info = wb_input.Worksheets("basic_info")

        'Set basic_info as an array
        Dim binfo() As Variant
        'Orderdetails:
        '1 - OrderNo
        '2 - YearEnd
        '3 - Client,4 - EngCntct, 5 - EngPrtnr, 6 - EngMngr, 7 - OtherCntct
        '8 - GISID
        '9 - Confirmation
        '10 - EngCode

        binfo = ws_Input_basic_info.range("B1:B10").Value
        ' Retrieve CAD Bearbeiter and set CAD Bearbeitungsdatum
        Dim ACPreparer As String
        ACPreparer = getACPreparer(CStr(binfo(1, 1)))
        
        ws_basic_info.range("E1") = ACPreparer
        ws_basic_info.range("E2") = Date
        
        
        'Paste binfo information to master template
        ws_basic_info.range("B1:B10") = binfo
    
        'Prepare which import worksheet to adress input
        Call sourceTOinput(wb_input, ws_InputAdressdaten)
        
        Call UppercaseISO(ws_InputAdressdaten)
        
        'fill index number of newly imported items
        Call fillIndex(ws_InputAdressdaten)
        
        'delete button
        Call deleteButton(ws_InputAdressdaten)

        Dim newName As String
        newName = CStr(Format(binfo(8, 1), "0000000000")) & " CAD Confirmations Workbook " & Format(CStr(binfo(2, 1)), "yyyyMMdd") & ".xlsm"
    
        'rename file
        Call renameWB(wb, savePath & newName)
        
        ws_InputAdressdaten.Cells(1, 1).Activate
        
        ' Trigger next procedure -> index number to single worksheet
        Call createDetailWorksheet
        
        ws_InputAdressdaten.Activate
        
        'Unhinde Summary worksheet
        wb.Worksheets("Summary").Visible = True
        wb.Worksheets("TabTemplate").Visible = False
        
    Else
        MsgBox ("Es ist ein Fehler aufgetreten. Bitte überprüfen Sie den Pfad")
    End If
    
End Sub


Sub UppercaseISO(ws As Worksheet)
    'This returns every value in the ISO Code column in uppercase
    ws.Activate
    
    Dim lastrow As Long, cell As range
    
    'This function returns last entry in passed column position
    
    lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    'This function returns all the values starting from K14 to the last filled cell in the column
    
    For Each cell In range("K14" & ":K" & lastrow)
        If Len(cell) > 0 Then cell = UCase(cell)
    Next cell
    
End Sub



