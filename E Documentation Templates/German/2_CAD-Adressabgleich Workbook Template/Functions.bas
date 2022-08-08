Attribute VB_Name = "Functions"
Function getInputpath() As String

    'This function asks User for input path via msgbox
    
    Dim Message As String, Title As String

    Title = "Input Abfrage"    ' Set title.
    
    Message = "Bitte geben sie den Pfad zur Quelldatei ein: "    ' Set prompt.
    
    ' Display message, title, and default value.
    getInputpath = InputBox(Message, Title)
    

End Function

Function checkPathValidity(inputpath) As Boolean

    'This function checks inputpath

    If Len(Dir(inputpath, vbDirectory)) <> 0 Then
        checkPathValidity = True
    Else
        checkPathValidity = False
    End If

End Function

Function getInputWB(inputpath) As Workbook

    Workbooks.Open inputpath
    Set getInputWB = ActiveWorkbook

End Function


Function getWS(keyword As String, wb As Workbook) As Worksheet

    Select Case keyword

        Case "Rechtsanwalt"

            Set getWS = wb.Worksheets("Rechts-_Steuerberater")

        Case Else

            Set getWS = wb.Worksheets("Adresscheck")

    End Select

End Function

Function getlastentry(ws As Worksheet, columnposition As Long) As Long

    'This function returns last entry in passed column position

    getlastentry = ws.Cells(Rows.Count, columnposition).End(xlUp).row

End Function


Sub renameWB(wb As Workbook, newName As String)

    'rename Workbook
    Application.DisplayAlerts = False
    
    wb.Activate
    wb.SaveAs FileName:=newName
    wb.Saved = True

End Sub
Sub copyContent(wsSource As Worksheet, sourceRange As String, wsOutput As Worksheet, outputRange As String, Optional modus As Integer)

    wsSource.range(sourceRange).Copy
    If modus = 1 Then
    
        'xlPasteAllExceptBorders
        wsOutput.range(outputRange).PasteSpecial (xlPasteAllExceptBorders)
    
    ElseIf modus = 2 Then
    
    
        wsOutput.range(outputRange).PasteSpecial Paste:=xlPasteAllExceptBorders, Transpose:=True
        
    ElseIf modus = 3 Then
    
    
    wsOutput.range(outputRange).PasteSpecial (xlPasteAll)
    
    
    Else
    
        wsOutput.range(outputRange).PasteSpecial (xlPasteValues)
    
    End If

End Sub

Sub deleteButton(ws As Worksheet)

    'This method deletes a Button based on worksheet name

    ws.Activate
    
    With ws
    
        Select Case ws.Name
    
            Case "Input Adressdaten"
                .Shapes("Button 2").Delete

            Case Else
                Dim cmd As Object

                For Each cmd In ActiveSheet.Shapes
            
                    If InStr(1, cmd.Name, "Button", vbTextCompare) <> 0 Then
                            
                            cmd.Delete
                            
                    End If
                
                Next

        End Select
    
    End With

End Sub

Function returnDirectoryInformation() As String

'This function returns selecte file path and file name delimited with "|"

    Dim lngCount As Long
    Dim cl As range
    
    Dim filePathAbsolute As String
    Dim FileName As String

    Set cl = ActiveCell
    ' Open the file dialog
    
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Show
        ' Display paths of each file selected
        
        For lngCount = 1 To .SelectedItems.Count
        
            filePathAbsolute = .SelectedItems(lngCount)
            
            FileName = getFileName(filePathAbsolute)

        Next lngCount
        
    End With
    
    returnDirectoryInformation = filePathAbsolute & "|" & FileName
    
End Function

Function getFileName(filePathAbsolute As String) As String

'This function returns only the file name

    Dim arr_directoryPath() As String

    arr_directoryPath = Split(filePathAbsolute, "\")

    getFileName = arr_directoryPath(UBound(arr_directoryPath))

End Function

Function check_WBopen(filePathAbsolute As String) As Boolean

'This function openes a workbook and checks if it is actually open

    Dim wb As Workbook
    On Error GoTo FileNotOpened
    Set wb = getInputWB(filePathAbsolute)
    
    check_WBopen = True
    
    Exit Function
    
FileNotOpened:
    
    check_WBopen = False

End Function

Sub sourceTOinput(wbSource As Workbook, wsInputAdressdaten As Worksheet)

    'This method copy and pastes each information from ws to inputsheet

    Dim t_ws As Worksheet, wsVersandliste As Worksheet
    
    Dim lastrowImport As Long, lastrowInput As Long, i As Long
    Dim sourceRange As String, outputRange As String
    
    Dim arrayUnique As Variant, arrayUniqueRange As range, arrayDuplicated As Variant, arrayParentTab As Variant
    Dim countIter As Integer, lastrow As Integer, inputRange As String, parentTabCount As Integer
    
    ThisWorkbook.Sheets.Add.Name = "Versandliste"
    Set wsVersandliste = ThisWorkbook.Sheets("Versandliste")
    
    wsVersandliste.range("A1") = "Parent Tab"
    wsVersandliste.range("B1") = "Versand"
    wsVersandliste.range("C1") = "Worksheet"
    wsVersandliste.range("D1") = "Art der Dienstleistung"
    
    For Each t_ws In wbSource.Worksheets
    
        If t_ws.Name <> "Summary" And t_ws.Name <> "Inhalte" And t_ws.Name <> "ISO" And t_ws.Name <> "basic_info" Then
    
            With t_ws
    
                .Activate
            
                'Minimum rausfinden, wenn es keine entries gibt, dann in das nächste worksheet
                lastrowImport = 0
                lastrowInput = 0
                If t_ws.Name = "Bank" Then lastrowImport = getlastentry(t_ws, 5) Else lastrowImport = getlastentry(t_ws, 3)
                lastrowInput = getlastentry(wsInputAdressdaten, 9)
            
                If lastrowImport > 16 Then
    
                    Select Case t_ws.Name
                
                        Case "Debitor_Kreditor_Sonst"
                        
                            ' Define arrayUnique
                            inputRange = "C17:N" & lastrowImport
                            .range(inputRange).Copy
                            firstrow = WorksheetFunction.Max(wsVersandliste.Cells(Rows.Count, 4).End(xlUp).row + 1, wsVersandliste.Cells(Rows.Count, 6).End(xlUp).row + 1)
                            wsVersandliste.range("D" & firstrow).PasteSpecial Paste:=xlPasteValues
                            wsVersandliste.range("D" & firstrow & ":N" & firstrow - 1 + lastrowImport - 16).RemoveDuplicates Columns:=Array(1, 3, 5, 6, 7, 8, 9, 10, 11)
                            lastrow = wsVersandliste.Cells(Rows.Count, 4).End(xlUp).row
                            arrayUnique = wsVersandliste.range("D" & firstrow & ":N" & lastrow)
                            
                            ' Define arrayDuplicated
                            arrayDuplicated = .range("C17:M" & lastrowImport).Value
                            
                            ReDim arrayParentTab(1 To lastrowImport - 16)
                            parentTabCount = Application.WorksheetFunction.Max(wsVersandliste.range("A1:A" & wsVersandliste.Cells(Rows.Count, 1).End(xlUp).row))
                            ' Find Parent Tab
                            For i = 1 To UBound(arrayDuplicated)
                                For j = 1 To UBound(arrayUnique)
                                    If arrayDuplicated(i, 1) = arrayUnique(j, 1) And arrayDuplicated(i, 3) = arrayUnique(j, 3) And arrayDuplicated(i, 5) = arrayUnique(j, 5) And arrayDuplicated(i, 6) = arrayUnique(j, 6) And arrayDuplicated(i, 7) = arrayUnique(j, 7) And arrayDuplicated(i, 8) = arrayUnique(j, 8) And arrayDuplicated(i, 9) = arrayUnique(j, 9) And arrayDuplicated(i, 10) = arrayUnique(j, 10) And arrayDuplicated(i, 11) = arrayUnique(j, 11) Then
                                        Debug.Print "Parent Tab found:" & parentTabCount + j
                                        arrayParentTab(i) = parentTabCount + j
                                    End If
                                Next j
                            Next i
                                                         
                            'Import Unique from Versandliste
                            lastRowInpAdr = wsInputAdressdaten.Cells(Rows.Count, 3).End(xlUp).row
                            wsVersandliste.range("F" & firstrow & ":N" & lastrow).Copy
                            wsInputAdressdaten.range("D" & lastRowInpAdr + 1).PasteSpecial Paste:=xlPasteValues
                            ' Art der Dienstleistung
                            wsVersandliste.range("D" & firstrow & ":D" & lastrow).Copy
                            wsInputAdressdaten.range("C" & lastRowInpAdr + 1).PasteSpecial Paste:=xlPasteValues
                            
                            ' Save Input Adressdaten in Versandliste
                            .range(inputRange).Copy
                            wsVersandliste.range("D" & firstrow).PasteSpecial Paste:=xlPasteValues
                            
                            wsVersandliste.range("A" & firstrow & ":A" & firstrow - 1 + lastrowImport - 16) = WorksheetFunction.Transpose(arrayParentTab)
                            wsVersandliste.range("B" & firstrow & ":B" & firstrow - 1 + lastrowImport - 16) = "Nein"
                            wsVersandliste.range("C" & firstrow & ":C" & firstrow - 1 + lastrowImport - 16) = t_ws.Name
                            
                        Case "Bank"
            
                            'Inputargumente: Case, lastrowImport
                                                        
                            ' Define arrayUnique
                            inputRange = "C17:N" & lastrowImport
                            .range(inputRange).Copy
                            firstrow = WorksheetFunction.Max(wsVersandliste.Cells(Rows.Count, 4).End(xlUp).row + 1, wsVersandliste.Cells(Rows.Count, 6).End(xlUp).row + 1)
                            wsVersandliste.range("E" & firstrow).PasteSpecial Paste:=xlPasteValues
                            wsVersandliste.range("E" & firstrow & ":O" & firstrow - 1 + lastrowImport - 16).RemoveDuplicates Columns:=Array(3, 5, 6, 7, 8, 9, 10, 11)
                            lastrow = wsVersandliste.Cells(Rows.Count, 7).End(xlUp).row
                            arrayUnique = wsVersandliste.range("E" & firstrow & ":O" & lastrow)
                            
                            ' Define arrayDuplicated
                            arrayDuplicated = .range("C17:M" & lastrowImport).Value
                            
                            ReDim arrayParentTab(1 To lastrowImport - 16)
                            parentTabCount = Application.WorksheetFunction.Max(wsVersandliste.range("A1:A" & wsVersandliste.Cells(Rows.Count, 1).End(xlUp).row))
                            ' Find Parent Tab
                            For i = 1 To UBound(arrayDuplicated)
                                For j = 1 To UBound(arrayUnique)
                                    If arrayDuplicated(i, 3) = arrayUnique(j, 3) And arrayDuplicated(i, 5) = arrayUnique(j, 5) And arrayDuplicated(i, 6) = arrayUnique(j, 6) And arrayDuplicated(i, 7) = arrayUnique(j, 7) And arrayDuplicated(i, 8) = arrayUnique(j, 8) And arrayDuplicated(i, 9) = arrayUnique(j, 9) And arrayDuplicated(i, 10) = arrayUnique(j, 10) And arrayDuplicated(i, 11) = arrayUnique(j, 11) Then
                                        Debug.Print "Parent Tab found:" & parentTabCount + j
                                        arrayParentTab(i) = parentTabCount + j
                                    End If
                                Next j
                            Next i
                                                         
                            'Import Unique from Versandliste
                            lastRowInpAdr = wsInputAdressdaten.Cells(Rows.Count, 3).End(xlUp).row
                            wsVersandliste.range("G" & firstrow & ":O" & lastrow).Copy
                            wsInputAdressdaten.range("D" & lastRowInpAdr + 1).PasteSpecial Paste:=xlPasteValues
                            ' Art der Dienstleistung
                            wsInputAdressdaten.range("C" & lastRowInpAdr + 1 & ":C" & lastRowInpAdr + UBound(arrayUnique)) = "Bank"
                            
                            ' Save Input Adressdaten in Versandliste
                            t_ws.range(inputRange).Copy
                            wsVersandliste.range("E" & firstrow).PasteSpecial Paste:=xlPasteValues
                            
                            wsVersandliste.range("A" & firstrow & ":A" & firstrow - 1 + lastrowImport - 16) = WorksheetFunction.Transpose(arrayParentTab)
                            wsVersandliste.range("B" & firstrow & ":B" & firstrow - 1 + lastrowImport - 16) = "Nein"
                            wsVersandliste.range("C" & firstrow & ":C" & firstrow - 1 + lastrowImport - 16) = t_ws.Name ' Sheet
                            wsVersandliste.range("D" & firstrow & ":D" & firstrow - 1 + lastrowImport - 16) = "Bank"

                        Case "Adresscheck", "Rechts-_Steuerberater"
                            'Inputargumente: Case, lastrowImport
                                                        
                            ' Define arrayUnique
                            inputRange = "C17:N" & lastrowImport
                            t_ws.range(inputRange).Copy
                            firstrow = WorksheetFunction.Max(wsVersandliste.Cells(Rows.Count, 4).End(xlUp).row + 1, wsVersandliste.Cells(Rows.Count, 6).End(xlUp).row + 1)
                            wsVersandliste.range("D" & firstrow).PasteSpecial Paste:=xlPasteValues
                            wsVersandliste.range("D" & firstrow & ":N" & firstrow - 1 + lastrowImport - 16).RemoveDuplicates Columns:=Array(2, 4, 5, 6, 7, 8, 9, 10)
                            lastrow = wsVersandliste.Cells(Rows.Count, 4).End(xlUp).row
                            arrayUnique = wsVersandliste.range("D" & firstrow & ":N" & lastrow)
                            
                            ' Define arrayDuplicated
                            arrayDuplicated = .range("C17:M" & lastrowImport).Value
                            
                            ReDim arrayParentTab(1 To lastrowImport - 16)
                            parentTabCount = Application.WorksheetFunction.Max(wsVersandliste.range("A1:A" & wsVersandliste.Cells(Rows.Count, 1).End(xlUp).row))
                            ' Find Parent Tab
                            For i = 1 To UBound(arrayDuplicated)
                                For j = 1 To UBound(arrayUnique)
                                    If arrayDuplicated(i, 2) = arrayUnique(j, 2) And arrayDuplicated(i, 4) = arrayUnique(j, 4) And arrayDuplicated(i, 5) = arrayUnique(j, 5) And arrayDuplicated(i, 6) = arrayUnique(j, 6) And arrayDuplicated(i, 7) = arrayUnique(j, 7) And arrayDuplicated(i, 8) = arrayUnique(j, 8) And arrayDuplicated(i, 9) = arrayUnique(j, 9) And arrayDuplicated(i, 10) = arrayUnique(j, 10) Then
                                        arrayParentTab(i) = parentTabCount + j
                                    End If
                                Next j
                            Next i
                                                         
                            'Import Unique from input file / Versandliste
                            wsVersandliste.range("D" & firstrow & ":M" & lastrow).Copy
                            lastrow = wsInputAdressdaten.Cells(Rows.Count, 3).End(xlUp).row
                            wsInputAdressdaten.range("C" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
                            
                            ' Save Input Adressdaten in Versandliste
                            .range(inputRange).Copy
                            wsVersandliste.range("D" & firstrow).PasteSpecial Paste:=xlPasteValues
                            
                            wsVersandliste.range("A" & firstrow & ":A" & firstrow - 1 + lastrowImport - 16) = WorksheetFunction.Transpose(arrayParentTab)
                            wsVersandliste.range("B" & firstrow & ":B" & firstrow - 1 + lastrowImport - 16) = "Nein"
                            wsVersandliste.range("C" & firstrow & ":C" & firstrow - 1 + lastrowImport - 16) = t_ws.Name
                            
                    End Select
            
                End If
    
            End With
        
        End If
    Next
    
    Dim rngCopy As range, rngPaste As range

    Set rngCopy = wsInputAdressdaten.range("B15:L15")
    lastrow = wsInputAdressdaten.Cells(Rows.Count, 3).End(xlUp).row
    Set rngPaste = wsInputAdressdaten.range("B15:L" & lastrow)

    rngCopy.Copy
    rngPaste.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    wbSource.Worksheets(1).Activate

End Sub


Sub importAdresses(wb As Workbook, t_ws As Worksheet, rangeAdressses As String, rangeVersandliste As String, ArtderDienstleistung As String, lastrowImport As Integer)
    Dim arrayUnique As Variant, arrayUniqueRange As range, arrayDuplicated As Variant, arrayParentTab As Variant, countIter As Integer, lastrow As Integer
    
    ' Define arrayUnique
    t_ws.range(rangeAddresses).Copy
    wb.Sheets("Versandliste").range("C1").PasteSpecial Paste:=xlPasteValues
    wb.Sheets("Versandliste").range(rangeVersandliste).RemoveDuplicates Columns:=Array(2, 4, 5, 6, 7, 8, 9, 10)
    lastrow = wb.Sheets("Versandliste").Cells(Rows.Count, 3).End(xlUp).row
    arrayUnique = wb.Sheets("Versandliste").range("C1:L" & lastrow)
    
    ' Define arrayDuplicated
    arrayDuplicated = t_ws.range("C17:L" & lastrowImport).Value
    
    ReDim arrayParentTab(1 To lastrowImport - 16)
    ' Find Parent Tab
    For i = 1 To UBound(arrayDuplicated)
        For j = 1 To UBound(arrayUnique)
            If arrayDuplicated(i, 2) = arrayUnique(j, 2) And arrayDuplicated(i, 3) = arrayUnique(j, 3) And arrayDuplicated(i, 4) = arrayUnique(j, 4) And arrayDuplicated(i, 5) = arrayUnique(j, 5) And arrayDuplicated(i, 6) = arrayUnique(j, 6) And arrayDuplicated(i, 7) = arrayUnique(j, 7) And arrayDuplicated(i, 8) = arrayUnique(j, 8) And arrayDuplicated(i, 9) = arrayUnique(j, 9) And arrayDuplicated(i, 10) = arrayUnique(j, 10) Then
                Debug.Print "Parent Tab found:" & j
                arrayParentTab(i) = j
            End If
        Next j
    Next i
                                 
    'Import Unique from input file / Versandliste
    wb.Sheets("Versandliste").range("C1:L" & lastrow).Copy
    lastrow = wb.Sheets("Input Adressdaten").Cells(Rows.Count, 2).End(xlUp).row
    wb.Sheets("Input Adressdaten").range("C" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
    
    ' Save Input Adressdaten in Versandliste
    t_ws.range("C17:L" & lastrowImport).Copy
    wb.Sheets("Versandliste").range("B1").PasteSpecial Paste:=xlPasteValues
    lastrow = wb.Sheets("Versandliste").Cells(Rows.Count, 3).End(xlUp).row
    wb.Sheets("Versandliste").range("A" & lastrow + 1 & ":A" & lastrow + lastrowImport - 16) = WorksheetFunction.Transpose(arrayParentTab)
End Sub


Function splitEmail(str_input As String) As String

    Dim arr_input() As String
    
    arr_input = Split(str_input, "@")
    
    splitEmail = arr_input(1)
    
End Function



Sub fillIndex(ws As Worksheet)

    'this function autofills index of newly imported data

    Dim lastrow As Long: lastrow = getlastentry(ws, 3)

    Dim i As Long

    For i = 1 To lastrow - 13

        ws.Cells(13 + i, 2) = i
    
    Next i
    
    ws.Rows(lastrow + 1).EntireRow.Delete

End Sub

Sub Test()

Call fillIndex(ActiveSheet)


End Sub



Function SheetExist(shtName As String)


    Dim sht As Worksheet
    Dim exists As Boolean
    exists = False
    For Each sht In ActiveWorkbook.Worksheets

        If sht.Name = shtName Then
            exists = True
            SheetExist = exists
            Exit Function

        End If

    Next sht

    SheetExist = exists

 
End Function
Function validateEntry(keyword As String, comparison As String, Optional trigger As String) As Long

    'This method validates original entry provided by client and colors the cell accordingly
    
    'Validation is based on following checks
    
    '1st check: empty - coloring = grey
    '2nd check: keyword matches exactly = white
    '3rd check: keyword matches partially Instr() = yellow
    '4th check: keyword does not match = red
    
    If comparison = "" Then
    
        validateEntry = 1
        
        Exit Function
    
    ElseIf StrComp(keyword, comparison) = 0 Then
    
        validateEntry = 2
        
        Exit Function
        
    Else
    
        validateEntry = 4
        
        
    End If
    
'    Dim arr_string() As String
'
'    Select Case trigger
'
'        Case "email"
'
'            arr_string = Split(comparison, "@")
'
'            If UBound(arr_string) = 0 Then
'
'                If InStr(1, keyword, arr_string(0)) <> 0 Then
'
'                    validateEntry = 2
'
'                Else
'
'                    validateEntry = 4
'
'                End If
'
'            Else
'
'                If InStr(1, keyword, arr_string(1)) <> 0 Then
'
'                    validateEntry = 2
'
'                Else
'
'                    validateEntry = 4
'
'                End If
'
'            End If
'
'        Case "soft", "adresse"
'
'            arr_string = Split(comparison, " ")
'
'            Dim i As Long
'
'            For i = 0 To UBound(arr_string)
'
'                If InStr(1, keyword, arr_string(i), vbTextCompare) <> 0 Then
'                    validateEntry = 3
'                    If trigger = "soft" Then Exit Function
'                Else
'
'                    validateEntry = 4
'
'                End If
'
'            Next i
'
'    End Select
        
    
End Function

Sub getDataProvider(wsRegister As Worksheet, wsOutput As Worksheet, keyword As String, country As String, lastrow As Long, lastcol As Long, arr_CPI As Variant)


    'This sub checks Register Worksheet for matching criteria: country and keyword
    'if entries is found. Hyperlink is copied to wsOutput
    
    Dim i As Long, j As Long
    
    
    Dim str_Range As String
    
    ' Check for Datenanbieter
    For i = 4 To 11
    
        If UCase(wsRegister.Cells(1, i)) = UCase(keyword) Then
        
            'Get range as string to use copyContent function
            str_Range = wsRegister.Cells(lastrow, i).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        
            Call copyContent(wsRegister, str_Range, wsOutput, "F23", 1)
            
            Exit For

        End If
    
    Next i

    ' Check for CPI Score
    'Wendings x = û, ok = ü
    
    For i = 1 To UBound(arr_CPI)
    
        If UCase(country) = UCase(arr_CPI(i, 2)) Or UCase(country) = UCase(arr_CPI(i, 3)) Then
        
            If arr_CPI(i, 5) = 1 Then
            
                wsOutput.range("F24") = "ü"
                wsOutput.range("F24").Font.Color = RGB(0, 176, 81)
                
            
            Else
            
                wsOutput.range("F24") = "û"
                wsOutput.range("F24").Font.Color = RGB(255, 51, 0)
            
            End If
            
            Exit Sub

        End If
    Next i

End Sub

Sub getRegister(wsRegister As Worksheet, wsOutput As Worksheet, keyword As String, country As String, lastrow As Long, lastcol As Long)

    'This sub checks Register Worksheet for matching criteria: country and keyword
    'if entries is found. Hyperlink is copied to wsOutput
    
    Dim i As Long, j As Long
    
    Dim str_Range As String

    ' Check for Register Art
    For i = 2 To lastrow
    
        If UCase(country) = UCase(wsRegister.Cells(i, 3)) Or UCase(country) = UCase(wsRegister.Cells(i, 12)) Then
        
            For j = 4 To lastcol
            
                If UCase(keyword) = UCase(wsRegister.Cells(1, j)) Then
            
                    'E23 Register
                    str_Range = wsRegister.Cells(i, j).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    Call copyContent(wsRegister, str_Range, wsOutput, "E23", 1)
                    
                    If wsOutput.range("E23") <> "" Then
                        wsOutput.range("E24") = "ü"
                        wsOutput.range("E24").Font.Color = RGB(0, 176, 81)
                    End If
                    
                    Exit Sub

                End If
        
        
            Next j

        End If
    Next i
    
End Sub

Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    Dim element As Variant
    On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
    Exit Function
IsInArrayError:
    On Error GoTo 0
    IsInArray = False
End Function

Function checkTabs(arr_ws As Variant) As Variant


    Dim arr_temporary() As Variant
    ReDim arr_temporary(1 To 5)

    Dim t_Sheets As String
    Dim str_Sheets_errorFIS As String, str_Sheets_errorX As String, str_Sheets_okay As String
    Dim c_tabs As Long: c_tabs = 0


    Dim sht As Worksheet
    Dim c_entries As Long
    
    For Each sht In ActiveWorkbook.Worksheets
    
        If IsInArray(sht.Name, arr_ws) = False Then
            Debug.Print sht.Name
            Dim cmd As Object
            
            c_tabs = c_tabs + 1
 
            'Check if Button exists
            For Each cmd In sht.Shapes
            
                'Truncate sheet name to inform user if button is still existent
                If InStr(1, cmd.Name, "Button") <> 0 Then
                    t_Sheets = sht.Name & "|" & t_Sheets
                    
                    Exit For
                
                End If
                
            Next
            
            If sht.Cells(23, 8) = "ûFIS" Then str_Sheets_errorFIS = sht.Name & "|" & str_Sheets_errorFIS
            If sht.Cells(23, 8) = "û" Or sht.Cells(23, 8) = "" Then str_Sheets_errorX = sht.Name & "|" & str_Sheets_errorX
            If sht.Cells(23, 8) = "ü" Then str_Sheets_okay = sht.Name & "|" & str_Sheets_okay


        End If

    Next sht
    
    'Button exists - check
    If t_Sheets <> "" Then arr_temporary(1) = Left(t_Sheets, Len(t_Sheets) - 1)
    If str_Sheets_errorFIS <> "" Then arr_temporary(2) = Left(str_Sheets_errorFIS, Len(str_Sheets_errorFIS) - 1)
    If str_Sheets_errorX <> "" Then arr_temporary(3) = Left(str_Sheets_errorX, Len(str_Sheets_errorX) - 1)
    If str_Sheets_okay <> "" Then arr_temporary(4) = Left(str_Sheets_okay, Len(str_Sheets_okay) - 1)
    
    arr_temporary(5) = c_tabs
    
    checkTabs = arr_temporary


End Function

Sub colorConclusion(wsRange As String, colorScheme As Integer)

    'This method ensures that conclusion tickmarks are formatted correctly
    
    ActiveSheet.range(wsRange).Select

    Select Case colorScheme
    
        Case 1, 2
            With ActiveCell.Characters(start:=1, Length:=1).Font
                .Name = "Wingdings"
                .FontStyle = "Standard Bold"
                .Size = 16
                .Color = RGB(255, 51, 0)
            End With
    
            If colorScheme = 1 Then
                With ActiveSheet.range(wsRange).Characters(start:=2, Length:=3).Font
                    .Name = "Calibri Light"
                    .FontStyle = "Bold"
                    .Size = 10
                    .Color = RGB(0, 0, 0)
                End With
            
            End If
            
        Case 3
        
            With ActiveCell.Characters(start:=1, Length:=1).Font
                .Name = "Wingdings"
                .FontStyle = "Standard Bold"
                .Size = 16
                .Color = RGB(0, 176, 80)
            End With

    End Select
    
End Sub

Sub OutprintFindings(ws As Worksheet, start_pos As Long, col_start As Long, modus As Integer, Optional lastrow As Long)
    Dim i As Long, j As Long, r_end As Long, c_end As Long, start_row As Long
    Select Case modus
    
        Case 0
            'Check I Column
            r_end = 33
            c_end = 9
            start_row = 15
        Case 1
            'Check L Column
            r_end = lastrow
            c_end = 14
            start_row = 14
    
    End Select
    
    Dim arr_ColorsFound() As String: ReDim arr_ColorsFound(1 To 4)
    
    For i = start_pos To r_end
    
        For j = col_start To c_end
            
            Select Case modus
            
                Case 0
                    If ws.Cells(i, j).Interior.Color = RGB(180, 198, 231) Then arr_ColorsFound(1) = "Zu den blau unterlegten Adressbestandteilen haben wir unerhebliche Abweichungen festgestellt. Die Abweichungen wurden durch uns entsprechend den Abgleichquellen geändert."
                    If ws.Cells(i, j).Interior.Color = RGB(248, 203, 173) Then arr_ColorsFound(2) = "Zu den rot unterlegten  Adressbestandteilen haben wir erhebliche Abweichungen festgestellt. Die Abweichungen wurden durch uns entsprechend den Abgleichquellen geändert."
                    If ws.Cells(i, j).Interior.Color = RGB(217, 217, 217) Then arr_ColorsFound(3) = "Zu den grau unterlegten Adressbestandteilen haben wir keine Abgleichquellen gefunden oder es lagen uns keine Mandantendaten vor."
            
                Case 1
                
                    If ws.Cells(i, j).Interior.Color = RGB(180, 198, 231) Then arr_ColorsFound(1) = "Zu den blau unterlegten Adressbestandteilen haben wir unerhebliche Abweichungen festgestellt. Die Abweichungen wurden durch uns entsprechend den Abgleichquellen geändert."
                    If ws.Cells(i, j).Interior.Color = RGB(248, 203, 173) Then arr_ColorsFound(2) = "Zu den rot unterlegten  Adressbestandteilen haben wir erhebliche Abweichungen festgestellt. Die Abweichungen wurden durch uns entsprechend den Abgleichquellen geändert."
                    If ws.Cells(i, j).Interior.Color = RGB(217, 217, 217) Then arr_ColorsFound(3) = "Zu den grau unterlegten Adressbestandteilen haben wir keine Abgleichquellen gefunden oder es lagen uns keine Mandantendaten vor."
                    
            End Select
        Next j
    Next i

    Dim c_found As Long: c_found = 0
    
    For i = 1 To 3
    
        If arr_ColorsFound(i) <> "" Then
    
            ws.Cells(start_row + c_found, 2) = arr_ColorsFound(i)
            c_found = c_found + 1
    
        End If

    Next i
 
    
    If c_found = 0 Then ws.Cells(start_row, 2) = "Der Abgleich ergab keine Abweichungen."

End Sub

Sub protectWS(ws As Worksheet, range As String)


    ws.range(range).Locked = True

    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True

    'ws.EnableSelection = xlNoRestrictions


End Sub


Sub createEmailDraft(attachmentPath As String, body As String, subject As String, recipient As String, ccRecipient As String, docuName As String)

    Dim OutMail As Object
    Dim OutApp As Object
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    'Email Erstellung

    With OutMail

        .Attachments.Add attachmentPath
        .Attachments.Add docuName
        .Attachments.Add "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\EY_Logo_Beam_RGB.png", olByValue, 0
        .Attachments.Add "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\190807_CAD_Signature-Banner_500px_mittel.png", olByValue, 0
        
        .subject = subject
        .To = recipient
        .SentOnBehalfOfName = "adressabgleich@de.ey.com"
        .CC = ccRecipient
        .HTMLBody = body
        
        ' Funktioniert nur in Outlook - Bitte live schalten
        '.Recipients.ResolveAll

        'Outlook Draft Fenster wird geöffnet
        .Display
        
        'Simulierte "STRG" + ENTER Eingabe, um die Nachricht zu versenden
        'Live schalten wenn Sonja Weinmann schon von Anfang an als ACPreparer mit Alexandra Staicu ersetzt wurde
        'SendKeys "^{ENTER}"

      
    End With

    'Schließe Outlook Object
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub


Public Function parseBody(t_template As String) As String

    '##############################################
    '##### Import .HTM Content into Mail Body #####
    '##############################################
        
    Dim body As String
    body = ""

    ' Read htm content into body
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (t_template)
    
    body = objStream.ReadText()
    
    parseBody = body
    
    ' Clean up
    objStream.Close
    Set objStream = Nothing

End Function

Sub updateVersandliste(setToYes As range, wb As Workbook)
    ' Sub updated die Versandliste und setzt alle Adressen auf Versand Ja, wenn der Parent Tab auf "ok" gesetzt ist.
    Dim parTabsVersandliste As Variant
    lastrow = wb.Sheets("Versandliste").Cells(Rows.Count, 4).End(xlUp).row
    parTabsVersandliste = wb.Sheets("Versandliste").range("A2:A" & lastrow).Value
    
    ' Update Versandliste / Setze Versand auf Ja wenn ParentTab IN TF_ok
    For i = 1 To lastrow - 1
        If IsArray(parTabsVersandliste) Then
            If WorksheetFunction.CountIf(setToYes, parTabsVersandliste(i, 1)) > 0 Then ' setToYes = ThisWorkbook.Sheets("TF_ok").range("C3:C14")
                wb.Sheets("Versandliste").range("B" & i + 1) = "Ja"
            End If
        Else
            If WorksheetFunction.CountIf(setToYes, parTabsVersandliste) > 0 Then ' setToYes = ThisWorkbook.Sheets("TF_ok").range("C3:C14")
                wb.Sheets("Versandliste").range("B" & i + 1) = "Ja"
            End If
        End If
    Next i
End Sub
Sub updateVersandlisteNotInOk()
    ' Sub updated die Versandliste und setzt alle Adressen auf Versand Ja, wenn der Parent Tab nicht in TF_FIS, TF_X oder TF_ok auftaucht.
    Dim wb As Workbook, TF_FIS As range, TF_X As range, TF_ok As range
    Set wb = ThisWorkbook
    
    lastrow = wb.Sheets("TF_FIS").Cells(Rows.Count, 3).End(xlUp).row + 1
    Set TF_FIS = wb.Sheets("TF_FIS").range("C3:C" & lastrow)
    lastrow = wb.Sheets("TF_X").Cells(Rows.Count, 3).End(xlUp).row + 1
    Set TF_X = wb.Sheets("TF_X").range("C3:C" & lastrow)
    lastrow = wb.Sheets("TF_ok").Cells(Rows.Count, 3).End(xlUp).row + 1
    Set TF_ok = wb.Sheets("TF_ok").range("C3:C" & lastrow)

    
    ' Sub updated die Versandliste und setzt alle Adressen auf Versand Ja, wenn der Parent Tab auf "ok" gesetzt ist.
    Dim parTabsVersandliste As Variant
    lastrow = wb.Sheets("Versandliste").Cells(Rows.Count, 4).End(xlUp).row
    parTabsVersandliste = wb.Sheets("Versandliste").range("A2:A" & lastrow).Value
    
    ' Update Versandliste / Setze Versand auf Ja wenn ParentTab IN TF_ok
    For i = 1 To lastrow - 1
        If IsArray(parTabsVersandliste) Then
            If WorksheetFunction.CountIf(TF_FIS, parTabsVersandliste(i, 1)) = 0 Then
                If WorksheetFunction.CountIf(TF_X, parTabsVersandliste(i, 1)) = 0 Then
                    If WorksheetFunction.CountIf(TF_ok, parTabsVersandliste(i, 1)) = 0 Then
                        wb.Sheets("Versandliste").range("B" & i + 1) = "Ja"
                    End If
                End If
            End If
        Else
            If WorksheetFunction.CountIf(TF_FIS, parTabsVersandliste) = 0 Then
                If WorksheetFunction.CountIf(TF_X, parTabsVersandliste) = 0 Then
                    If WorksheetFunction.CountIf(TF_ok, parTabsVersandliste) = 0 Then
                        wb.Sheets("Versandliste").range("B" & i + 1) = "Ja"
                    End If
                End If
            End If
        End If
    Next i
End Sub

Sub createVersandlisteFile(wb As Workbook)
    ' Open Versandliste
    Dim wbVersandliste As Workbook, saveName As String, orderNo As String, orderbook As String, strFilePath As String
    Set wbVersandliste = Application.Workbooks.Open("\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\5_CAD-Adressabgleich Adressen für externe Bestätigungen_Template.xlsx")
    
    ' Fill Versandliste
    Call fillVersandlisteFile(wbVersandliste, wb)
    
    ' Save Versandliste
    orderNo = wb.Sheets("basic_info").range("B1").Value
    If Left(orderNo, 3) = "CON" Then
        orderbook = "tCON_Orderbook"
        strFilePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & orderNo & "\5. Versandliste"
    Else
        orderbook = "tAC_Orderbook"
        strFilePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich\C Workplace\" & orderNo & "\5. Versandliste"
    End If
    Dim fso As New FileSystemObject
    If Not fso.FolderExists(strFilePath) Then
        fso.CreateFolder strFilePath
    End If
    saveName = strFilePath & "\" & CStr(Format(wb.Sheets("basic_info").range("B8").Value, "0000000000")) & " 5_CAD-Adressabgleich Adressen für externe Bestätigungen " & Format(CStr(wb.Sheets("basic_info").range("B2").Value), "yyyyMMdd") & ".xlsx"
    ActiveWorkbook.SaveAs FileName:=saveName, FileFormat:=51
    ActiveWorkbook.Close
End Sub

Sub fillVersandlisteFile(wbVersandliste As Workbook, wb As Workbook)
    Dim arrayVersandliste As Variant
    lastrow = wb.Sheets("Versandliste").Cells(Rows.Count, 4).End(xlUp).row
    arrayVersandliste = wb.Sheets("Versandliste").range("A2" & ":O" & lastrow)
    For i = 1 To UBound(arrayVersandliste)
        If arrayVersandliste(i, 2) = "Ja" Then
            If arrayVersandliste(i, 3) = "Debitor_Kreditor_Sonst" Then
                If arrayVersandliste(i, 4) = "Debitor" Then
                    If wbVersandliste.Sheets("Debitoren").Cells(Rows.Count, 5).End(xlUp).row = 26 Then lastrow = 27 Else lastrow = wbVersandliste.Sheets("Debitoren").Cells(Rows.Count, 5).End(xlUp).row
                    wbVersandliste.Sheets("Debitoren").range("C" & lastrow + 1 & ":E" & lastrow + 1).Value = wb.Sheets("Versandliste").range("D" & i + 1 & ":F" & i + 1).Value
                    wbVersandliste.Sheets("Debitoren").range("F" & lastrow + 1 & ":M" & lastrow + 1).Value = wb.Sheets("Versandliste").range("H" & i + 1 & ":O" & i + 1).Value
                ElseIf arrayVersandliste(i, 4) = "Kreditor" Then
                    If wbVersandliste.Sheets("Kreditoren").Cells(Rows.Count, 5).End(xlUp).row = 26 Then lastrow = 27 Else lastrow = wbVersandliste.Sheets("Kreditoren").Cells(Rows.Count, 5).End(xlUp).row
                    wbVersandliste.Sheets("Kreditoren").range("C" & lastrow + 1 & ":E" & lastrow + 1).Value = wb.Sheets("Versandliste").range("D" & i + 1 & ":F" & i + 1).Value
                    wbVersandliste.Sheets("Kreditoren").range("F" & lastrow + 1 & ":N" & lastrow + 1).Value = wb.Sheets("Versandliste").range("H" & i + 1 & ":O" & i + 1).Value
                ElseIf arrayVersandliste(i, 4) = "Sonstige" Then
                    If wbVersandliste.Sheets("Sonstige").Cells(Rows.Count, 5).End(xlUp).row = 26 Then lastrow = 27 Else lastrow = wbVersandliste.Sheets("Sonstige").Cells(Rows.Count, 5).End(xlUp).row
                    wbVersandliste.Sheets("Sonstige").range("C" & lastrow + 1 & ":E" & lastrow + 1).Value = wb.Sheets("Versandliste").range("D" & i + 1 & ":F" & i + 1).Value
                    wbVersandliste.Sheets("Sonstige").range("F" & lastrow + 1 & ":N" & lastrow + 1).Value = wb.Sheets("Versandliste").range("H" & i + 1 & ":O" & i + 1).Value
                End If
            ElseIf arrayVersandliste(i, 3) = "Bank" Then
                If wbVersandliste.Sheets("Bank").Cells(Rows.Count, 3).End(xlUp).row = 27 Then lastrow = 28 Else lastrow = wbVersandliste.Sheets("Bank").Cells(Rows.Count, 3).End(xlUp).row
                wbVersandliste.Sheets("Bank").range("C" & lastrow + 1 & ":C" & lastrow + 1).Value = wb.Sheets("Versandliste").range("G" & i + 1 & ":G" & i + 1).Value
                wbVersandliste.Sheets("Bank").range("D" & lastrow + 1 & ":E" & lastrow + 1).Value = wb.Sheets("Versandliste").range("E" & i + 1 & ":F" & i + 1).Value
                wbVersandliste.Sheets("Bank").range("F" & lastrow + 1 & ":L" & lastrow + 1).Value = wb.Sheets("Versandliste").range("I" & i + 1 & ":O" & i + 1).Value
            ElseIf arrayVersandliste(i, 3) = "Rechts-_Steuerberater" Then
                If wbVersandliste.Sheets(arrayVersandliste(i, 4)).Cells(Rows.Count, 4).End(xlUp).row = 27 Then lastrow = 28 Else lastrow = wbVersandliste.Sheets(arrayVersandliste(i, 4)).Cells(Rows.Count, 4).End(xlUp).row
                wbVersandliste.Sheets(arrayVersandliste(i, 4)).range("C" & lastrow + 1 & ":D" & lastrow + 1).Value = wb.Sheets("Versandliste").range("D" & i + 1 & ":E" & i + 1).Value
                wbVersandliste.Sheets(arrayVersandliste(i, 4)).range("E" & lastrow + 1 & ":L" & lastrow + 1).Value = wb.Sheets("Versandliste").range("G" & i + 1 & ":N" & i + 1).Value
            ElseIf arrayVersandliste(i, 3) = "Adresscheck" Then
                Select Case arrayVersandliste(i, 4)
                    Case "Kreditor", "Debitor", "Sonstige"
                        Dim sheetName As String
                        If arrayVersandliste(i, 4) = "Kreditor" Then
                            sheetName = "Kreditoren"
                        ElseIf arrayVersandliste(i, 4) = "Debitor" Then
                            sheetName = "Debitoren"
                        Else
                            sheetName = "Sonstige"
                        End If
                        If wbVersandliste.Sheets(sheetName).Cells(Rows.Count, 5).End(xlUp).row = 26 Then lastrow = 27 Else lastrow = wbVersandliste.Sheets(sheetName).Cells(Rows.Count, 5).End(xlUp).row
                        wbVersandliste.Sheets(sheetName).range("C" & lastrow + 1 & ":C" & lastrow + 1).Value = wb.Sheets("Versandliste").range("D" & i + 1 & ":D" & i + 1).Value
                        wbVersandliste.Sheets(sheetName).range("E" & lastrow + 1 & ":E" & lastrow + 1).Value = wb.Sheets("Versandliste").range("E" & i + 1 & ":E" & i + 1).Value
                        wbVersandliste.Sheets(sheetName).range("F" & lastrow + 1 & ":M" & lastrow + 1).Value = wb.Sheets("Versandliste").range("G" & i + 1 & ":N" & i + 1).Value
                    Case "Bank"
                        If wbVersandliste.Sheets(arrayVersandliste(i, 4)).Cells(Rows.Count, 3).End(xlUp).row = 27 Then lastrow = 28 Else lastrow = wbVersandliste.Sheets(arrayVersandliste(i, 4)).Cells(Rows.Count, 3).End(xlUp).row
                        wbVersandliste.Sheets(arrayVersandliste(i, 4)).range("C" & lastrow + 1 & ":C" & lastrow + 1).Value = wb.Sheets("Versandliste").range("E" & i + 1 & ":E" & i + 1).Value
                        wbVersandliste.Sheets(arrayVersandliste(i, 4)).range("F" & lastrow + 1 & ":M" & lastrow + 1).Value = wb.Sheets("Versandliste").range("G" & i + 1 & ":N" & i + 1).Value
                    Case "Steuerberater", "Rechtsberater", "Wirtschaftsprüfer", "Sonstige Berater"
                        If wbVersandliste.Sheets(arrayVersandliste(i, 4)).Cells(Rows.Count, 4).End(xlUp).row = 27 Then lastrow = 28 Else lastrow = wbVersandliste.Sheets(arrayVersandliste(i, 4)).Cells(Rows.Count, 4).End(xlUp).row
                        wbVersandliste.Sheets(arrayVersandliste(i, 4)).range("C" & lastrow + 1 & ":D" & lastrow + 1).Value = wb.Sheets("Versandliste").range("D" & i + 1 & ":E" & i + 1).Value
                        wbVersandliste.Sheets(arrayVersandliste(i, 4)).range("E" & lastrow + 1 & ":L" & lastrow + 1).Value = wb.Sheets("Versandliste").range("G" & i + 1 & ":N" & i + 1).Value
                End Select
            End If
        End If
    Next i
End Sub
Sub sendNoTeamApprovalMail(binfo As Variant)
    
    Dim subject As String, body As String, templatePath As String
    If Left(binfo(1, 1), 3) = "CON" Then
        templatePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\DE\4_ConAC_NoApprovalMail.htm"
    Else
        templatePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\DE\4_AC_NoApprovalMail.htm"
    End If
    
    subject = "Action required: CAD Adressabgleich Bestellung [Confirmation] [OrderNo] für [Client]"
    subject = Replace(subject, "[OrderNo]", CStr(binfo(1, 1)))
    subject = Replace(subject, "[Client]", CStr(binfo(3, 1)))
    subject = Replace(subject, "[Confirmation]", CStr(binfo(9, 1)))
    
    Dim ACPreparer As String
    ACPreparer = getACPreparer(CStr(binfo(1, 1)))
    
    'Vorgefertigter Text
    body = parseBody(templatePath)
    body = Replace(body, "[OrderNo]", binfo(1, 1))
    body = Replace(body, "[GISID]", CStr(Format(binfo(8, 1), "0000000000")))
    body = Replace(body, "[Client]", binfo(3, 1))
    body = Replace(body, "[YearEnd]", CStr(binfo(2, 1)))
    body = Replace(body, "[AC Preparer]", ACPreparer)
    body = Replace(body, "[Confirmation]", CStr(binfo(9, 1)))
 
    Dim engCntct As String, engCC As String
    engCntct = binfo(4, 1)
    engCC = binfo(6, 1) & ";" & binfo(7, 1)
    
    'create Email Draft
    Dim OutMail As Object
    Dim OutApp As Object
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    'Email Erstellung
    With OutMail

        .Attachments.Add "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\EY_Logo_Beam_RGB.png", olByValue, 0
        .Attachments.Add "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\190807_CAD_Signature-Banner_500px_mittel.png", olByValue, 0
        
        .subject = subject
        .To = engCntct
        .CC = engCC
        .HTMLBody = body

        .Display

      
    End With

    'Schließe Outlook Object
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Public Function updateRS(SQL As String) As ADODB.Recordset
    'FP20210826
    'The function sets connection to SQL db and pulls a recordset acc to provided SQL-querry
    'Input: SQL querry as String
    'Output: ADO Recordset
    
    Dim sConnString As String
    Dim conn As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset

    'Create connection string
    sConnString = "Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;"
    'Open connection to SQL db
    conn.Open sConnString
    'Create RS
    myRecordSet.CursorLocation = adUseClient
    myRecordSet.ActiveConnection = conn

    myRecordSet.Open SQL
    Set myRecordSet = Nothing

End Function

Public Function getRS(SQL As String) As ADODB.Recordset
          'The function sets connection to SQL db and pulls a recordset acc to provided SQL-querry
          'Input: SQL querry as String
          'Output: ADO Recordset
          
          Dim sConnString As String
          Dim conn As New ADODB.Connection
          Dim myRecordSet As New ADODB.Recordset

          'Create connection string
          '#BM-1/2019-08-08 10:49
1         sConnString = "Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;"
          '#BM-1/2019-08-08 10:49 End
          'Open connection to SQL db
2         conn.Open sConnString

          'Create RS
3         myRecordSet.CursorLocation = adUseClient
4         myRecordSet.ActiveConnection = conn

          Debug.Print SQL
5         myRecordSet.Open SQL
          
6         Set getRS = myRecordSet.Clone
          
          'Clean up
7         myRecordSet.Close
8         Set myRecordSet = Nothing

End Function

Public Function getACPreparer(orderNo As String) As String
    ' Function to get AC Preparer for AC / CON and AP / AR Case
    Dim SQL As String, rsOrderbook As Object, ACPreparer As String
    If Left(orderNo, 2) = "AC" Or Left(orderNo, 3) = "CON" Then ' AC / CON Case
        SQL = "SELECT * FROM (SELECT OrderNo, AC_Preparer FROM [CAD].[dbo].[tCON_Orderbook] " & vbCrLf
        SQL = SQL & "UNION "
        SQL = SQL & "SELECT OrderNo, AC_Preparer FROM [CAD].[dbo].[tAC_Orderbook]) AS U" & vbCrLf
        SQL = SQL & " WHERE U.OrderNo = '" & orderNo & "'"
        Set rsOrderbook = getRS(SQL)
        If IsNull(rsOrderbook.Fields("AC_Preparer").Value) Then
            ACPreparer = "Chantal Berg"
        Else
            ACPreparer = rsOrderbook.Fields("AC_Preparer").Value
        End If
    Else ' AP / AR Case
        ACPreparer = "Chantal Berg"
    End If
    getACPreparer = ACPreparer
End Function
