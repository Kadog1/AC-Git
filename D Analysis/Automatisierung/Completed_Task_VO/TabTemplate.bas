Attribute VB_Name = "TabTemplate"
Sub createDetailWorksheet()
    
    'This procedure creates each detail worksheet based on WS("Input_Adressdaten")

    Dim lastrow As Long
    lastrow = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    Dim arr_pbc As Variant
    ReDim arr_pbc(1 To lastrow - 13, 1 To 11)
    arr_pbc = ActiveSheet.range("B14:L" & lastrow)

    ActiveWorkbook.Worksheets("TabTemplate").Visible = True

    Dim copyafter As String
    copyafter = "Summary"
    
    For i = 1 To lastrow - 13

        'create tabs nach der Vorlage

        'wenn das worksheet von exisitiert, nächstes
        If SheetExist(CStr(i)) = False Then

            ActiveWorkbook.Worksheets("TabTemplate").Copy After:=Worksheets(copyafter)
            ActiveWorkbook.Worksheets("TabTemplate (2)").Name = i
            copyafter = i
    
            ActiveWorkbook.Worksheets(CStr(i)).range("D19").Value = arr_pbc(i, 1) 'laufende Nummer
            ActiveWorkbook.Worksheets(CStr(i)).range("D20").Value = arr_pbc(i, 2) 'Art der Dienstleistung

            ActiveWorkbook.Worksheets(CStr(i)).range("D25").Value = arr_pbc(i, 3) 'Firma/Bank/Kanzlei
            ActiveWorkbook.Worksheets(CStr(i)).range("D26").Value = arr_pbc(i, 4) 'Adresszusatz
            ActiveWorkbook.Worksheets(CStr(i)).range("D27").Value = arr_pbc(i, 5) 'Vorname
            ActiveWorkbook.Worksheets(CStr(i)).range("D28").Value = arr_pbc(i, 6) 'Nachname
            ActiveWorkbook.Worksheets(CStr(i)).range("D29").Value = arr_pbc(i, 7) 'Straße/Postfach
            ActiveWorkbook.Worksheets(CStr(i)).range("D30").Value = arr_pbc(i, 8) 'PLZ
            ActiveWorkbook.Worksheets(CStr(i)).range("D31").Value = arr_pbc(i, 9) 'Stadt
            ActiveWorkbook.Worksheets(CStr(i)).range("D32").Value = arr_pbc(i, 10) 'Land
            
            ActiveWorkbook.Worksheets(CStr(i)).range("D33").Value = arr_pbc(i, 11) 'EmailAdresse
            If InStr(CStr(arr_pbc(i, 11)), "@") > 0 Then ActiveWorkbook.Worksheets(CStr(i)).range("D34").Value = splitEmail(CStr(arr_pbc(i, 11))) 'EmailDomain
            
            'load CPI Score Information
        
            Dim arr_CPI() As Variant
            Dim range_CPI As range
        
            Dim wsCPI As Worksheet: Set wsCPI = ActiveWorkbook.Worksheets("CPI Score")

            arr_CPI = wsCPI.range("A1:E181")
        
            'prefill newly created tabs
            
            Call getDataProvider(ActiveWorkbook.Worksheets("Register"), ActiveWorkbook.Worksheets(CStr(i)), ActiveWorkbook.Worksheets(CStr(i)).Cells(20, 4), ActiveWorkbook.Worksheets(CStr(i)).Cells(32, 4), 57, 9, arr_CPI)
            
            'wsRegister As Worksheet, wsOutput As Worksheet, keyword As String, country As String, lastrow As Long, lastcol As Long
            Call getRegister(ActiveWorkbook.Worksheets("Register"), ActiveWorkbook.Worksheets(CStr(i)), ActiveWorkbook.Worksheets(CStr(i)).Cells(20, 4), ActiveWorkbook.Worksheets(CStr(i)).Cells(32, 4), 56, 9)

            
            ActiveWorkbook.Worksheets(CStr(i)).Cells(1, 1).Select
            
        End If
    
    Next i


    ActiveWorkbook.Worksheets("TabTemplate").Visible = False

End Sub

Sub PreFill_Click()

    'This procedure prefills conclusion decision based on entries

    
    Dim lastrowEntries As Long: lastrowEntries = getlastentry(ActiveWorkbook.ActiveSheet, 5)
    Dim i As Long, j As Long
    Dim colorScheme As Long
    Dim trigger As String
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet
    
    If ws.Cells(19, 4) <> "" Then
    
        For i = 25 To 34
        
            Select Case i
            
                Case 28, 30, 31, 32, 33
            
                    trigger = "adresse"
            
                Case 34
                
                    trigger = "email"
            
                Case Else
            
                    trigger = "soft"

            End Select
            
            'Color cell grey if it is empty
            If ws.Cells(i, 4) = "" Then ws.Cells(i, 4).Interior.Color = RGB(217, 217, 217)
    
            For j = 5 To 7

                colorScheme = validateEntry(ws.Cells(i, 4), ws.Cells(i, j), trigger)

                Select Case colorScheme
                 
                    Case 1
                    
                        ws.Cells(i, j).Interior.Color = RGB(217, 217, 217)
                                       
                    Case 2
                    
                        If i <> 34 Then ws.Cells(i, 8) = "Keine Abweichung"
                
                    Case 3
                    
                        ws.Cells(i, j).Interior.Color = RGB(180, 198, 231)
                        If i <> 34 Then ws.Cells(i, 8) = "Unerhebliche Abweichung"
                 
                    Case 4
                    
                        If i <> 34 Then ws.Cells(i, j).Interior.Color = RGB(248, 203, 173)
                        If i <> 34 Then ws.Cells(i, 8) = "Erhebliche Abweichung"
                 
                End Select
                 
            Next j

        Next i
        
        'Delete Button
        ActiveWorkbook.ActiveSheet.Shapes("Button 2").Delete
        
        'Cleaning
        Call CleaningFormatting
        Call CleaningBorder
        
        Call visibilityCheckBox(1)
        
        ' Copy Tickbox Legend
        Call copyContent(ActiveWorkbook.Worksheets("basic_info"), "A19:C26", ActiveSheet, "K22:M29", 1)
        
    Else
    
        MsgBox ("Es wurden kein Input gefunden")
    
    End If
    

End Sub

Sub EinzelValidierung_Click()

    AnswerYes = MsgBox("Möchten Sie den Einzelabgleich fertigstellen? Alle Buttons werden entfernt", vbQuestion + vbYesNo, "User Repsonse")

    If AnswerYes = vbYes Then
    
        'Validate if entries are ready to be finalized
        
        If ActiveWorkbook.ActiveSheet.Cells(23, 8) = "" Then
        
            ' Check if Email Domain was entered. If not -> msgbox: Conclusion is mandatory
            '
            ' FP Improvement 20210830 Start If ActiveWorkbook.ActiveSheet.Cells(34, 8) = "" Then
            '   MsgBox ("Es wurde keine Conclusion hinterlegt. Die Daten können nicht weiterverarbeitet werden")
            '   Exit Sub
            'Else
            
            ActiveWorkbook.ActiveSheet.Cells(23, 8).Interior.Color = RGB(217, 217, 217)
            
            ' FP Improvement 20210830 End End If
            
        End If
        
        ' Check if conclusion is correctly formatted
        If ActiveWorkbook.ActiveSheet.Cells(23, 8) <> "" Then
            
            If ActiveWorkbook.ActiveSheet.Cells(23, 8) <> "ü" And ActiveWorkbook.ActiveSheet.Cells(23, 8) <> "û" And ActiveWorkbook.ActiveSheet.Cells(23, 8) <> "ûFIS" Then
                MsgBox ("Das Format der Conclusion ist ungültig.")
                Exit Sub
            End If
            
        End If

        Dim b_entryFound As Boolean, b_colorMatch As Boolean
        Dim str_Range As String, str_goal As String
        Dim i As Long, j As Long
        Dim c_email_found As Long: c_email_found = 0
        
        For j = 5 To 7
        
            If ActiveWorkbook.ActiveSheet.Cells(34, j) <> "" And ActiveWorkbook.ActiveSheet.Cells(34, 8) = "" Then
    
                MsgBox ("Es wurde keine Conclusion (Email) hinterlegt. Die Daten können nicht weiterverarbeitet werden")
                Exit Sub
            End If
            
            If ActiveWorkbook.ActiveSheet.Cells(34, j) = "" And ActiveWorkbook.ActiveSheet.Cells(34, 8) <> "" Then
                c_email_found = c_email_found + 1
            End If
        
            
            If ActiveWorkbook.ActiveSheet.Cells(34, 8) <> "ü" And ActiveWorkbook.ActiveSheet.Cells(34, 8) = "û" And ActiveWorkbook.ActiveSheet.Cells(34, 8) = "ûFIS" Then
                MsgBox ("Das Format der Conclusion (Email) ist ungültig.")
                Exit Sub
            End If
            
        Next j
        
        If c_email_found = 3 Then
        
            MsgBox ("Es wurde Conclusion (Email Domain) angewählt jedoch keine Eingabe getätigt.")
            Exit Sub
        
        End If

        If ActiveWorkbook.ActiveSheet.Cells(19, 4) <> "" Then
    
            For i = 25 To 34
            
                b_entryFound = False
                b_colorMatch = False
                

                For j = 5 To 7
                
                    If ActiveWorkbook.ActiveSheet.Cells(i, j) <> "" Then
                    
                        'Validation: Check if conclusion matches Color
                        b_colorMatch = validateConclusion(ActiveSheet, i, j)
                        If b_colorMatch = False Then
                
                            MsgBox ("Conclusion und Farbton stimmen nicht überein!")
                            Exit Sub
                    
                        End If
                                        
                        str_goal = ActiveWorkbook.ActiveSheet.Cells(i, 9).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                        str_Range = ActiveWorkbook.ActiveSheet.Cells(i, j).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                        
                        Call copyContent(ActiveWorkbook.ActiveSheet, str_Range, ActiveWorkbook.ActiveSheet, str_goal, 1)
                        
                        b_entryFound = True
                        Exit For
                    End If

                Next j
                
                If b_entryFound = False Then
                    ActiveWorkbook.ActiveSheet.Cells(i, 9) = ActiveWorkbook.ActiveSheet.Cells(i, 4)
                    ActiveWorkbook.ActiveSheet.Cells(i, 9).Interior.Color = RGB(217, 217, 217)
                    ActiveWorkbook.ActiveSheet.Cells(i, 8).Interior.Color = RGB(217, 217, 217)
                End If
                
                If ActiveSheet.Cells(34, 8) = "" Then
                
                    ActiveWorkbook.ActiveSheet.Cells(33, 8).Interior.Color = RGB(217, 217, 217)
                
                End If


            Next i
 
            'delete button
            Call deleteButton(ActiveWorkbook.ActiveSheet)
            
            'Color Conclusion accordingly - Conclusion adress
            Select Case ActiveSheet.Cells(23, 8).Value
            
                Case "ûFIS"
            
                    Call colorConclusion("H23", 1)
            
                Case "û"
                
                    Call colorConclusion("H23", 2)
            
                Case "ü"
                
                    Call colorConclusion("H23", 3)
    
            End Select
            
            'Color Conclusion accordingly - Conclusion email - adress
            
            Select Case ActiveSheet.Cells(34, 8).Value
            
                Case "ûFIS"
            
                    Call colorConclusion("H34", 1)
            
                Case "û"
                
                    Call colorConclusion("H34", 2)
            
                Case "ü"
                
                    Call colorConclusion("H34", 3)
    
            End Select
            
            'Clear color / markup tips
            ActiveSheet.range("K22:M29").Clear
            ActiveSheet.range("A1").Select
            
            'Print Findings
            Call OutprintFindings(ActiveSheet, 25, 4, 0)
            
            'Delete red square
            'ActiveSheet.Shapes("Rechteck 15").Delete
            
            'Color unnecessary information grey if Col E is fully entered
            Call adjustRegister
            
            'Cleaning
            Call CleaningFormatting
            Call CleaningBorder
            
            'Protect Cells
            Call protectWS(ActiveSheet, "D25:I34")
            
            'Hide CheckBoxes
            Call visibilityCheckBox(0)
            
            

        Else
    
            MsgBox ("Es wurde kein Input gefunden")
    
        End If

    End If
    

End Sub

Sub adjustRegister()

    ' if col E is filled clear F + G 23/24

    Dim i As Long, c_found As Long
    c_found = 0

    Dim b_found As Boolean: b_found = False

    For i = 25 To 33

        If ActiveSheet.Cells(i, 5) <> "" Then
            c_found = c_found + 1
    
        End If
    
    Next i

    If ActiveSheet.Cells(34, 4) <> "" Then

        If ActiveSheet.Cells(34, 5) <> "" Then
            c_found = c_found + 1
            b_found = True
        End If
    End If

    Select Case b_found

        Case True

            If c_found = 10 Then
            
                ActiveSheet.Cells(23, 6) = ""
                ActiveSheet.Cells(23, 6).Interior.Color = RGB(217, 217, 217)
                ActiveSheet.Cells(24, 6) = ""
                ActiveSheet.Cells(24, 6).Interior.Color = RGB(217, 217, 217)
                
                ActiveSheet.Cells(23, 7) = ""
                ActiveSheet.Cells(23, 7).Interior.Color = RGB(217, 217, 217)
                ActiveSheet.Cells(24, 7) = ""
                ActiveSheet.Cells(24, 7).Interior.Color = RGB(217, 217, 217)

            Else
                Exit Sub

            End If


        Case False
        
            If c_found = 9 Then
            
                ActiveSheet.Cells(23, 6) = ""
                ActiveSheet.Cells(23, 6).Interior.Color = RGB(217, 217, 217)
                ActiveSheet.Cells(24, 6) = ""
                ActiveSheet.Cells(24, 6).Interior.Color = RGB(217, 217, 217)
                
                ActiveSheet.Cells(23, 7) = ""
                ActiveSheet.Cells(23, 7).Interior.Color = RGB(217, 217, 217)
                ActiveSheet.Cells(24, 7) = ""
                ActiveSheet.Cells(24, 7).Interior.Color = RGB(217, 217, 217)
        
            Else
        
                Exit Sub
        
            End If

    End Select



End Sub


Function validateConclusion(ws As Worksheet, i As Long, j As Long) As Boolean
    validateConclusion = True
    If i <> 34 Then
        
        If ws.Cells(i, j).Interior.Color = RGB(0, 0, 0) Then
            If ws.Cells(i, 8) <> "Keine Abweichung" Then validateConclusion = False
        End If

        If ws.Cells(i, j).Interior.Color = RGB(180, 198, 231) Then
            If ws.Cells(i, 8) <> "Unerhebliche Abweichung" Then validateConclusion = False
        End If

        If ws.Cells(i, j).Interior.Color = RGB(248, 203, 173) Then
            If ws.Cells(i, 8) <> "Erhebliche Abweichung" Then validateConclusion = False
        End If
    
    End If

End Function

Sub visibilityCheckBox(modus As Integer)

'Modus 0 - Hides Objects
'Modus 1 - Shows Objects
    
    Dim cmd As Object
    
    With ActiveSheet
    
        For Each cmd In ActiveSheet.Shapes
            
            If InStr(1, cmd.Name, "Check Box", vbTextCompare) <> 0 Then
                            
                If modus = 0 Then
                cmd.Visible = False
                Else
                cmd.Visible = True
                End If
                            
            End If
                
        Next

    End With



End Sub
' This sub goes from DetailWorksheet to DetailWorksheet and retrieves the corresponding screenshot if available



Sub CleaningFormatting()
    'Cleaning-Code: Hyperlinks, Alignment, Font/Size, WraptText

    Dim ws As Worksheet: Set ws = ActiveWorkbook.ActiveSheet

    Dim rngWorkp As range: Set rngWorkp = ws.range("D25:G34")
    Dim rngSource As range: Set rngSource = ws.range("E23:G24")
    Dim rngWhole As range: Set rngWhole = ws.range("D23:G44")

    'Hyperlink entfernen
    rngWorkp.Hyperlinks.Delete
    
    'LeftAlignment
    rngWorkp.HorizontalAlignment = xlLeft
    
    'CenterAlignment Vertical
    rngWorkp.VerticalAlignment = xlCenter
    
    'Font
    rngWorkp.Font.Size = 10
    rngWorkp.Font.Name = "Calibri"

    'WrapText
    rngWhole.WrapText = True

End Sub

Sub CleaningBorder()
    'Cleaning-Code: Borders

    Dim ws As Worksheet: Set ws = ActiveWorkbook.ActiveSheet

    Dim rngSource As range: Set rngSource = ws.range("E23:G24")

    Dim colReg As range: Set colReg = ws.range("E23:E34")
    Dim colDatap As range: Set colDatap = ws.range("F23:F34")
    Dim colHP As range: Set colHP = ws.range("G23:G34")

    'Border - colReg
    colReg.Borders.LineStyle = xlContinous
    colReg.Borders.Weight = xlHairline
        
    colReg.Borders(xlEdgeLeft).LineStyle = xlContinous
    colReg.Borders(xlEdgeLeft).Weight = xlMedium
    
    'Border - colDatap
    colDatap.Borders.LineStyle = xlContinous
    colDatap.Borders.Weight = xlHairline
        
    colDatap.Borders(xlEdgeLeft).LineStyle = xlContinous
    colDatap.Borders(xlEdgeLeft).Weight = xlThin
        
    colDatap.Borders(xlEdgeRight).LineStyle = xlContinous
    colDatap.Borders(xlEdgeRight).Weight = xlThin
    
    'Border - colHP
    colHP.Borders.LineStyle = xlContinous
    colHP.Borders.Weight = xlHairline
        
    colHP.Borders(xlEdgeRight).LineStyle = xlContinous
    colHP.Borders(xlEdgeRight).Weight = xlMedium
        
    'Border - DataSource
    rngSource.Borders(xlInsideHorizontal).LineStyle = xlNone
        
    rngSource.Borders(xlEdgeTop).LineStyle = xlContinous
    rngSource.Borders(xlEdgeTop).Weight = xlThin
        
    rngSource.Borders(xlEdgeBottom).LineStyle = xlContinous
    rngSource.Borders(xlEdgeBottom).Weight = xlThin

End Sub


Sub retrieveScreenshot()
    
    Dim idxAddress As String, OrderNo As String, tsScreenshotCreated As String, locationPath As String, FileName As String, pathPNG As String
    Dim binfo() As Variant
    Dim wb As Workbook
    Dim ws_basic_info As Worksheet
    Dim SQL As String, rsOrderbook As Object
    
    ' get OrderNo, tsScreenShotCreated
    Set wb = ThisWorkbook
    Set ws_basic_info = wb.Worksheets("basic_info")
    binfo = ws_basic_info.range("B1:B11").Value
    
    OrderNo = CStr(binfo(1, 1))
    tsScreenshotCreated = CStr(binfo(11, 1))
    'OrderNo = "CON0000023789"
    'tsScreenshotCreated = "2022-02-28 12:07:00.000"
    
    listSkipSheets = Array("Start", "Summary", "Input Adressdaten", "Input Beurteilung", "TabTemplate", "Team Approval Documentation", _
        "TF_FIS", "TF_ok", "TF_X", "Register", "CPI Score", "basic_info")
    
    For Each t_ws In wb.Worksheets
        
        If IsError(Application.Match(t_ws.Name, listSkipSheets, 0)) Then  ' not case sensitive
            idxAddress = t_ws.range("D19").Value
            'idxAddress = 3
            
            ' get RS for OrderNo, idx & tsScreenshotCreated
            SQL = "SELECT TOP (1) p.Company, p.OrderNo, p.idxAddress, p.tsScreenshotCreated, s.locationPath, s.pngFileName FROM [CAD].[dbo].[tAC_ProdScreenshots_TEST] p" & vbCrLf
            SQL = SQL & "LEFT JOIN [CAD].[dbo].[tAC_Screenshots_TEST] s ON p.Company = s.Company " & vbCrLf
            SQL = SQL & "WHERE p.OrderNo = '" & OrderNo & "'" & vbCrLf
            SQL = SQL & "AND p.idxAddress = '" & idxAddress & "'" & vbCrLf
            SQL = SQL & "AND p.tsScreenshotCreated = '" & tsScreenshotCreated & "'"
            
            Set rsSQL = getRS(SQL)
            
            If rsSQL.RecordCount > 0 Then ' Get Screenshot
                locationPath = rsSQL.Fields("locationPath")
                'locationPath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich Testumgebung\F Screenshots\M\Microsoft Deutschland GmbH"
                FileName = rsSQL.Fields("pngFileName")
                'FileName = "Microsoft Deutschland GmbH_Muenchen_20220302_1820.png"
                pathPNG = locationPath & "\" & FileName
                wb.Worksheets("TabTemplate").Shapes.AddPicture pathPNG, True, True, 40, 650, 960, 540
            End If
        End If
    Next

End Sub

