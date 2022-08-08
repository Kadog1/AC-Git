Attribute VB_Name = "TeamApproval"
Sub Transfer(arr_summaryValidated As Variant)
                 
    'open template
    Dim pathTemp As String
    Dim tempname As String
    Dim wbtemp As Workbook

    Dim wb As Workbook
    
    Dim sum As Worksheet, t_wsInput As Worksheet
    Dim wb_sum As Worksheet, wb_input As Worksheet
    
    Dim cel As range

    Set wb = ActiveWorkbook
    Set wb_sum = wb.Worksheets("Summary")
    Set wb_input = wb.Worksheets("Input Adressdaten")

    pathTemp = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\3_CAD-Adressabgleich Team Approval_Template.xlsm"
    'pathTemp = "C:\Users\devoh001\Desktop\DEV\eConfirmation\D Dokumentation Templates\3_CAD-Adressabgleich Team Approval_Template.xlsm"
 
    Set wbtemp = Workbooks.Open(pathTemp)
    Set sum = wbtemp.Worksheets("Summary")
    Set t_wsInput = wbtemp.Worksheets("Input Adressdaten")
    
    Dim lastrowInput As Long: lastrowInput = getlastentry(wb_input, 2)
    
    'Import Input Adressdaten
    Call copyContent(wb_input, "B14:L" & lastrowInput, t_wsInput, "B14:L" & lastrowInput, 3)
    
    t_wsInput.Activate
    t_wsInput.range("A1").Select
    'Import Basic_info
    
    Dim binfo() As Variant
    binfo = wb.Worksheets("basic_Info").range("B1:B10").Value
    wb.Worksheets("Versandliste").Copy After:=wbtemp.Sheets(wbtemp.Sheets.Count)
    wbtemp.Worksheets("Versandliste").Visible = xlSheetHidden

        
    'Paste binfo information to master template
    wbtemp.Worksheets("basic_info").range("B1:B10") = binfo
    
    'Check which Table to fill
    
    Dim b_FIS As Boolean: b_FIS = False
    Dim b_X As Boolean: b_X = False
    Dim b_ok As Boolean: b_ok = False
    
    Dim addRows As Long: addRows = 0
    
    wb.Save
    wb.Saved = True
    wb.Activate
    
    If arr_summaryValidated(2) <> "" Then
    
        'Copy FIS
        
        Call copyContent(wbtemp.Worksheets("basic_info"), "A22", sum, "B27", 3)
        sum.Cells(27, 2) = Replace(sum.Cells(27, 2), "[FIS]", "1)")

        Call copyContent(wb.Worksheets("TF_FIS"), "A1:P" & getlastentry(wb.Worksheets("TF_FIS"), 4), sum, "B28", 3)
        'Call copyContent(wb.Worksheets("TF_FIS"), "A1:A" & getlastentry(wb.Worksheets("TF_FIS"), 1), sum, "B28", 3)

        Call createFIScopy(wb)
        

    End If
    
    Dim str_Range As String
    
    
    If arr_summaryValidated(3) <> "" Or arr_summaryValidated(4) <> "" Then
    
        addRows = getlastentry(sum, 4) + 2
        ' Create Header
        If arr_summaryValidated(2) <> "" Then
            str_Range = sum.Cells(addRows, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            Call copyContent(wbtemp.Worksheets("basic_info"), "A24:A25", sum, str_Range, 3)
            sum.Cells(addRows, 2) = Replace(sum.Cells(addRows, 2), "[X]", "2)")
        Else
            Call copyContent(wbtemp.Worksheets("basic_info"), "A24:A25", sum, "B27", 3)
            sum.Cells(27, 2) = Replace(sum.Cells(27, 2), "[X]", "1)")
        End If
        
        
        ' Add X
        If arr_summaryValidated(3) <> "" Then
            If addRows = 3 Then
                'Copy X
                Call copyContent(wb.Worksheets("TF_X"), "A1:P" & getlastentry(wb.Worksheets("TF_X"), 4), sum, "B29", 3)
                Call createDropdown("B28:C" & CStr(getlastentry(sum, 2)), sum)
            Else
                'Copy X
                Call copyContent(wb.Worksheets("TF_X"), "A1:P" & getlastentry(wb.Worksheets("TF_X"), 4), sum, "B" & addRows + 3, 3)
                'Call copyContent(wb.Worksheets("TF_X"), "A1:A" & getlastentry(wb.Worksheets("TF_X"), 1), sum, "B" & addRows + 2, 3)
                Call createDropdown("B" & CStr(addRows + 2) & ":C" & CStr(getlastentry(sum, 4)), sum)
            End If
        End If
        
        ' Add ok
        If arr_summaryValidated(4) <> "" Then
            addRows = getlastentry(sum, 4) + 2
            If addRows = 3 Then
                'Copy TF_ok
                Call copyContent(wb.Worksheets("TF_ok"), "A1:P" & getlastentry(wb.Worksheets("TF_ok"), 4), sum, "B29", 3)
                Call createDropdown("B28:C" & CStr(getlastentry(sum, 4)), sum)
            ElseIf arr_summaryValidated(3) <> "" Then
                Call copyContent(wb.Worksheets("TF_ok"), "A3:P" & getlastentry(wb.Worksheets("TF_ok"), 4), sum, "B" & addRows - 1, 3)
                Call createDropdown("B" & CStr(addRows - 1) & ":C" & CStr(getlastentry(sum, 4)), sum)
            Else
                'Copy TF_ok
                Call copyContent(wb.Worksheets("TF_ok"), "A1:P" & getlastentry(wb.Worksheets("TF_ok"), 4), sum, "B" & addRows + 3, 3)
                Call createDropdown("B" & CStr(addRows + 2) & ":C" & CStr(getlastentry(sum, 4)), sum)
            End If
        End If
    End If
       
    'Copy Legend information
    addRows = getlastentry(sum, 2) + 2
    Call copyContent(wbtemp.Worksheets("basic_info"), "A12:B20", sum, "B" & addRows + 2, 3)
    
    'Insert Template Name in disclaimer
    Dim disclaimer As String: disclaimer = CStr(sum.Cells(getlastentry(sum, 3), 3))
    disclaimer = Replace(disclaimer, "[NameTemplate]", wb.Name)
    sum.Cells(getlastentry(sum, 3), 3) = disclaimer

    sum.Activate
    sum.range("A1").Select
    
    wbtemp.Worksheets(1).Activate

    'Get path to input file
    Dim str_fileName As String, savePath As String

    'savePath = wb.Path
    savePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & CStr(binfo(1, 1)) & "\2. CAD_Abgleich"
    
    'Open
    Dim newName As String, docuPath As String
    newName = CStr(Format(binfo(8, 1), "0000000000")) & " 3_CAD-Adressabgleich Team Approval_Template " & Format(CStr(binfo(2, 1)), "yyyyMMdd") & ".xlsm"
    docuPath = savePath & "\" & wb.Name
    
    'Protect Worksheet summary
    sum.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True

    'rename file
    Call renameWB(wbtemp, savePath & "\" & newName)
    
    Dim subject As String
    subject = "Action required: Team Approval für den Adressabgleich notwendig / CAD Bestellung [Confirmation] Bestätigungsaktion für [Client] / Bestellnummer: [OrderNo]"
    subject = Replace(subject, "[OrderNo]", CStr(binfo(1, 1)))
    subject = Replace(subject, "[Client]", CStr(binfo(3, 1)))
    subject = Replace(subject, "[Client]", CStr(binfo(3, 1)))
    subject = Replace(subject, "[Confirmation]", CStr(binfo(9, 1)))
    
    Dim templatepath As String
    templatepath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\DE\4_Team_Approval_Template_DE.htm"
    
    Dim body As String, ACPreparer As String
    ACPreparer = getACPreparer(CStr(binfo(1, 1)))
    
    'Vorgefertigter Text
    body = parseBody(templatepath)
    
    body = Replace(body, "[OrderNo]", binfo(1, 1))
    body = Replace(body, "[GISID]", CStr(Format(binfo(8, 1), "0000000000")))
    body = Replace(body, "[Client]", binfo(3, 1))
    body = Replace(body, "[YearEnd]", CStr(binfo(2, 1)))
    body = Replace(body, "[AC Preparer]", ACPreparer)
 
    Dim engCntct As String, engCC As String
    
    engCntct = binfo(4, 1)
    engCC = binfo(6, 1) & ";" & binfo(7, 1)

    'Close Team Aproval Template
    wbtemp.Close
    Set wbtemp = Nothing
    
    'create Email Draft
    Call createEmailDraft(savePath & "\" & newName, body, subject, engCntct, engCC, docuPath)
    

    'Create Separate FIS Document
    If arr_summaryValidated(2) <> "" Then
        
        Dim newSubject As String
    
        'fisSubject = "Action required: Forensic Adressabgleich notwendig / CAD Bestellung [Confirmation] Bestätigungsaktion für [Client] / Bestellnummer: [OrderNo]"
        newSubject = "FIS involvement required: CAD Adressabgleich Bestellung [Confirmation] für [Client] / Bestellnummer: [OrderNo]"
        newSubject = Replace(newSubject, "[Confirmation]", CStr(binfo(9, 1)))
        newSubject = Replace(newSubject, "[Client]", CStr(binfo(3, 1)))
        newSubject = Replace(newSubject, "[OrderNo]", CStr(binfo(1, 1)))
        
        Dim savedTemplate As String
        Dim templatep As String
        Dim engMgr As String
        Dim Engcnt As String
        
        savedTemplate = CStr(Format(binfo(8, 1), "0000000000")) & " CAD-Adressabgleich Forensics " & Format(CStr(binfo(2, 1)), "yyyyMMdd") & ".xlsm"
        
        'templatep = "C:\TestLocal\AC adressabgliech\FIS maill project\4_FIS Template.htm"
        templatep = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\DE\4_FIS Template.htm"
        
        body = parseBody(templatep)
    
        body = Replace(body, "[OrderNo]", binfo(1, 1))
        body = Replace(body, "[GISID]", CStr(Format(binfo(8, 1), "0000000000")))
        body = Replace(body, "[Client]", binfo(3, 1))
        body = Replace(body, "[YearEnd]", CStr(binfo(2, 1)))
        
        Engcnt = binfo(4, 1) & ";"
        engMgr = binfo(6, 1)
        
        Call createEmailDraft(savePath & "\" & savedTemplate, body, newSubject, "fis.adressabgleich@de.ey.com", Engcnt & engMgr, docuPath)
       
            
    End If

End Sub

Sub createFIScopy(wb As Workbook)

    'open template
    Dim pathTemp As String
    Dim tempname As String
    Dim wbtemp As Workbook

    pathTemp = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\EX_CAD_Adressabgleich Forensics_Template.xlsm"
    
    Dim sum As Worksheet
    Dim t_wsInput As Worksheet
    
    Set wbtemp = Workbooks.Open(pathTemp)
    Set sum = wbtemp.Worksheets("Input Adressdaten")

    Dim lastrowInput As Long: lastrowInput = getlastentry(wb.Worksheets("Input Adressdaten"), 2)
    
    'Import Basic_info
    
    Dim binfo() As Variant
    binfo = wb.Worksheets("basic_Info").range("B1:B10").Value
    
    'Paste binfo information to master template
    wbtemp.Worksheets("basic_info").range("B1:B10") = binfo
    
    Call copyContent(wb.Worksheets("TF_FIS"), "A1:P" & getlastentry(wb.Worksheets("TF_FIS"), 3), sum, "B19", 3)
    'Call copyContent(wb.Worksheets("TF_FIS"), "A1:A" & getlastentry(wb.Worksheets("TF_FIS"), 1), sum, "B19", 3)

    sum.Activate
    sum.range("A1").Select
    
    wbtemp.Worksheets(1).Activate

    'Get path to input file
    Dim str_fileName As String, savePath As String

    savePath = wb.Path
    
    'Open
    Dim newName As String
    newName = CStr(Format(binfo(8, 1), "0000000000")) & " CAD-Adressabgleich Forensics " & Format(CStr(binfo(2, 1)), "yyyyMMdd") & ".xlsm"
    
    'rename file
    Call renameWB(wbtemp, savePath & "\" & newName)
    
    wbtemp.Close
    Set wbtemp = Nothing

End Sub


Sub createDropdown(range As String, ws As Worksheet)

    Application.CutCopyMode = False
    With ws.range(range).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$R$2:$R$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    ws.range(range).Locked = False

End Sub




