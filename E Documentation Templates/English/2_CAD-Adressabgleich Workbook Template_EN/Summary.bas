Attribute VB_Name = "Summary"
Sub createSummary_Click()

    'verarbeitet die Information aus den Tabs zu Summary
    'loop alle die nicht Start, Summary, Summary (2), TabTemplate, Input Adressdaten, Input Beurteilung heißen
    ' fühlt Information + Farbe in ein array

    Dim AnswerYes As String


    AnswerYes = MsgBox("Summary is created from all tabs. Please make sure all tabs are ready", vbQuestion + vbYesNo, "User Repsonse")

    If AnswerYes = vbYes Then
                                   
        Dim arr_transferReady() As Variant
        
        Dim arr_ws As Variant
        ReDim arr_ws(1 To 14)
        arr_ws(1) = "Start"
        arr_ws(2) = "Summary"
        arr_ws(3) = "Summary (2)"
        arr_ws(4) = "TabTemplate"
        arr_ws(5) = "Input Address data"
        arr_ws(6) = "Input evaluation"
        arr_ws(7) = "basic_info"
        arr_ws(8) = "Register"
        arr_ws(9) = "CPI Score"
        arr_ws(10) = "TF_FIS"
        arr_ws(11) = "TF_X"
        arr_ws(12) = "TF_ok"
        arr_ws(13) = "Team Approval Documentation"
        arr_ws(14) = "Versandliste"

        'Check if Button in detail sheets exist
        arr_transferReady = checkTabs(arr_ws)
        
        If arr_transferReady(5) <> getlastentry(ActiveWorkbook.Worksheets("Input Address data"), 2) - 13 Then
            MsgBox ("The number of detail sheets does not match the number of input address data!")
            Exit Sub
        End If

        If arr_transferReady(1) = "" Then
        
            'If there are already entries. Clear content
            If ActiveSheet.Cells(22, 1) <> "" Then Call resetSummarySheet(getlastentry(ActiveSheet, 1))
            
            'conclusion
            
            'import xFIS tabs to summary
            If arr_transferReady(4) <> "" Then Call import_data(CStr(arr_transferReady(4)))
            
            'import x tabs to summary
            If arr_transferReady(3) <> "" Then Call import_data(CStr(arr_transferReady(3)))

            'import green tabs to summary
            If arr_transferReady(2) <> "" Then Call import_data(CStr(arr_transferReady(2)))
                      
            ActiveSheet.Rows(22).EntireRow.Delete
            ActiveSheet.Rows(22).EntireRow.Delete
            
            Call OutprintFindings(ActiveSheet, 22, 3, 1, getlastentry(ActiveSheet, 1))

        Else
            MsgBox ("Please make sure that the following tabs have been finalized: " & arr_transferReady(1))
        End If

    Else

    End If
        
End Sub
Sub finalizeSummary_Click()

    Dim arr_summaryValidated() As Variant
    Dim lastrow As Long: lastrow = getlastentry(ActiveSheet, 1)

    arr_summaryValidated = validationSummary(lastrow)
        
        
    If lastrow = 21 Then
        
        MsgBox ("Fehlender Input im Summary Sheet.")
            
        Exit Sub
    
    End If
    
    ' Update Versandliste Vesand Yes / No according to sheet TF_ok.
    lastrow = ThisWorkbook.Sheets("TF_ok").Cells(Rows.Count, 3).End(xlUp).Row + 1
    Call updateVersandliste(ThisWorkbook.Sheets("TF_ok").range("C3:C" & lastrow), ThisWorkbook)
    Call updateVersandlisteNotInOk
    
    Dim orderNo As String, orderbook As String
    orderNo = ThisWorkbook.Sheets("basic_info").range("B1").Value
    If Left(orderNo, 3) = "CON" Then
        orderbook = "tCON_Orderbook"
    Else
        orderbook = "tAC_Orderbook"
    End If

    If arr_summaryValidated(1) = True Then

        'MsgBox ("Team approval benötigt. " & vbCrLf & arr_summaryValidated(2) & vbCrLf & arr_summaryValidated(3) & vbCrLf & arr_summaryValidated(4))
        
        'Delete all Buttons on Summary Sheet
        Call deleteButton(ActiveSheet)
    
        'Create Document for Team approval
        Call Transfer(arr_summaryValidated)
        Dim timeNow As String
        timeNow = Format(Now(), "yyyy-MM-dd hh:mm:ss")
        Call updateRS("UPDATE " & orderbook & " Set AC_Status = 'TeamApprovalSent', tsTeamApprovalSent = '" & timeNow & "' WHERE OrderNo = '" & orderNo & "'")
        
        Exit Sub
        
    Else
        MsgBox ("Finales Dokument kann erstellt werden.")
        
        Call createVersandlisteFile(ThisWorkbook)
        
        Call updateRS("UPDATE " & orderbook & " Set AC_Status = 'TeamApprovalReceived' WHERE OrderNo = '" & orderNo & "'")
        Call deleteButton(ActiveSheet)
        Dim binfo As Variant
        binfo = ThisWorkbook.Worksheets("basic_Info").range("B1:B10").Value
        Call sendNoTeamApprovalMail(binfo)
        
    End If

'Schließt Workbook
ThisWorkbook.Close

End Sub

Function validationSummary(lastrow As Long) As Variant

    'Validation and preparation of Email Draft - Create Table in temporary Worksheets

    Dim arr_entries() As Variant: ReDim arr_entries(1 To 4)

    Dim b_teamApproval As Boolean: b_teamApproval = False
    Dim b_okFalse As Boolean
    
    Dim str_goal As String, str_Range As String

    Dim i As Long

    For i = 22 To lastrow
    
        Select Case ActiveSheet.Cells(i, 11)
        
            Case "ûFIS"
            
                arr_entries(2) = "Es ist ein Versand an Forensic nötig. Siehe Tab: " & ActiveSheet.Cells(i, 1) & vbCrLf & arr_entries(2)
                
                
                'Copy Content to tempSheet FIS
                
                'Clean up Range
                If ActiveWorkbook.Worksheets("TF_FIS").Cells(3, 1) <> "" Then
                    ActiveWorkbook.Worksheets("TF_FIS").range("A3:P" & getlastentry(ActiveWorkbook.Worksheets("TF_FIS"), 1) + 2).Clear
                    'ActiveWorkbook.Worksheets("TF_FIS").range("A3:A" & getlastentry(ActiveWorkbook.Worksheets("TF_FIS"), 1) + 2).Clear
                    'ActiveWorkbook.Worksheets("TF_FIS").range("M3:M" & getlastentry(ActiveWorkbook.Worksheets("TF_FIS"), 1) + 2).Clear
                End If

                ActiveWorkbook.Worksheets("TF_FIS").Rows(4).EntireRow.Insert
                Call copyContent(ActiveWorkbook.ActiveSheet, "A" & i & ":N" & i, ActiveWorkbook.Worksheets("TF_FIS"), "C4:P4", 1)
                'Call copyContent(ActiveWorkbook.ActiveSheet, "K" & i, ActiveWorkbook.Worksheets("TF_FIS"), "A4", 1) 'Verlässlichkeitsgrad Adresse
                'Call copyContent(ActiveWorkbook.ActiveSheet, "L" & i & ":M" & i, ActiveWorkbook.Worksheets("TF_FIS"), "N4", 1)  'Verlässlichkeitgrad (Domain)
                
                b_teamApproval = True
    
            Case "û"
                
                arr_entries(3) = "Der Verlässlichkeitsgrad für Tab: " & ActiveSheet.Cells(i, 1) & " ist nicht ausreichend." & vbCrLf & arr_entries(3)
                
                'Copy Content to tempSheet X
                
                'Clean up Range
                If ActiveWorkbook.Worksheets("TF_X").Cells(3, 1) <> "" Then
                    ActiveWorkbook.Worksheets("TF_X").range("A3:P" & getlastentry(ActiveWorkbook.Worksheets("TF_X"), 1) + 2).Clear
                    'ActiveWorkbook.Worksheets("TF_X").range("A3:A" & getlastentry(ActiveWorkbook.Worksheets("TF_X"), 1) + 2).Clear
                    'ActiveWorkbook.Worksheets("TF_X").range("M3:M" & getlastentry(ActiveWorkbook.Worksheets("TF_X"), 1) + 2).Clear
                End If

                ActiveWorkbook.Worksheets("TF_X").Rows(4).EntireRow.Insert
                Call copyContent(ActiveWorkbook.ActiveSheet, "A" & i & ":N" & i, ActiveWorkbook.Worksheets("TF_X"), "C4:P4", 1)
                'Call copyContent(ActiveWorkbook.ActiveSheet, "K" & i, ActiveWorkbook.Worksheets("TF_X"), "A4", 1)
                'Call copyContent(ActiveWorkbook.ActiveSheet, "L" & i & ":M" & i, ActiveWorkbook.Worksheets("TF_X"), "N4", 1)
                ActiveWorkbook.Worksheets("TF_X").range("A4") = "No"
                ActiveWorkbook.Worksheets("TF_X").range("B4") = "Yes"
                
                b_teamApproval = True
                
            Case "ü"
                If ActiveSheet.Cells(i, 14) = "û" Then
                    arr_entries(3) = "Der Verlässlichkeitsgrad für Tab: " & ActiveSheet.Cells(i, 1) & " ist nicht ausreichend." & vbCrLf & arr_entries(3)
            
                    'Copy Content to tempSheet X
            
                    'Clean up Range
                    If ActiveWorkbook.Worksheets("TF_X").Cells(3, 1) <> "" Then
                        ActiveWorkbook.Worksheets("TF_X").range("A3:P" & getlastentry(ActiveWorkbook.Worksheets("TF_X"), 1) + 2).Clear
                        'ActiveWorkbook.Worksheets("TF_X").range("A3:A" & getlastentry(ActiveWorkbook.Worksheets("TF_X"), 1) + 2).Clear
                        'ActiveWorkbook.Worksheets("TF_X").range("M3:M" & getlastentry(ActiveWorkbook.Worksheets("TF_X"), 1) + 2).Clear
                    End If

                    ActiveWorkbook.Worksheets("TF_X").Rows(4).EntireRow.Insert
                    Call copyContent(ActiveWorkbook.ActiveSheet, "A" & i & ":N" & i, ActiveWorkbook.Worksheets("TF_X"), "C4:P4", 1)
                    'Call copyContent(ActiveWorkbook.ActiveSheet, "K" & i, ActiveWorkbook.Worksheets("TF_X"), "A4", 1)
                    'Call copyContent(ActiveWorkbook.ActiveSheet, "L" & i & ":M" & i, ActiveWorkbook.Worksheets("TF_X"), "N4", 1)
                    ActiveWorkbook.Worksheets("TF_X").range("A4") = "No"
                    ActiveWorkbook.Worksheets("TF_X").range("B4") = "Yes"
            
                    b_teamApproval = True
                Else
                    Dim j As Long
                    
                    b_okFalse = False
                    
                    For j = 3 To 12
                    
                        If ActiveSheet.Cells(i, j).Interior.Color = RGB(248, 203, 173) Then
                            arr_entries(4) = "Der Verlässlichkeitsgrad für Tab: " & ActiveSheet.Cells(i, 1) & " ist ausreichend. Jedoch wurden erhebliche Abweichung festgestellt. " & vbCrLf & arr_entries(4)
                            b_teamApproval = True
                            b_okFalse = True
    
                            Exit For
                        End If
                    
                    Next j
                    
                    If b_okFalse = False Then
                    
                        If ActiveSheet.Cells(i, 14) = "û" Then
                    
                            b_okFalse = True
                    
                    
                        End If
                    
                    End If
                    
                    'Copy Content to tempSheet ok
                    If b_okFalse Then
                        
                        'Clean up Range
                        If ActiveWorkbook.Worksheets("TF_ok").Cells(3, 1) <> "" Then
                            ActiveWorkbook.Worksheets("TF_ok").range("A3:P" & getlastentry(ActiveWorkbook.Worksheets("TF_ok"), 1) + 2).Clear
                            'ActiveWorkbook.Worksheets("TF_ok").range("A3:A" & getlastentry(ActiveWorkbook.Worksheets("TF_ok"), 1) + 2).Clear
                            'ActiveWorkbook.Worksheets("TF_ok").range("M3:M" & getlastentry(ActiveWorkbook.Worksheets("TF_ok"), 1) + 2).Clear
                        End If
                        
                        ActiveWorkbook.Worksheets("TF_ok").Rows(4).EntireRow.Insert
                        Call copyContent(ActiveWorkbook.ActiveSheet, "A" & i & ":N" & i, ActiveWorkbook.Worksheets("TF_ok"), "C4:P4", 1)
                        'Call copyContent(ActiveWorkbook.ActiveSheet, "K" & i, ActiveWorkbook.Worksheets("TF_ok"), "A4", 1)
                        'Call copyContent(ActiveWorkbook.ActiveSheet, "L" & i & ":M" & i, ActiveWorkbook.Worksheets("TF_ok"), "N4", 1)
                        ActiveWorkbook.Worksheets("TF_ok").range("A4") = "Yes"
                    End If
                End If
                
                
            Case ""
            
                Select Case ActiveSheet.Cells(i, 14)
                    Case "û"
                
                        arr_entries(3) = "Der Verlässlichkeitsgrad für Tab: " & ActiveSheet.Cells(i, 1) & " ist nicht ausreichend." & vbCrLf & arr_entries(3)
                
                        'Copy Content to tempSheet X
                
                        'Clean up Range
                        If ActiveWorkbook.Worksheets("TF_X").Cells(3, 1) <> "" Then
                            ActiveWorkbook.Worksheets("TF_X").range("A3:P" & getlastentry(ActiveWorkbook.Worksheets("TF_X"), 1) + 2).Clear
                            'ActiveWorkbook.Worksheets("TF_X").range("A3:A" & getlastentry(ActiveWorkbook.Worksheets("TF_X"), 1) + 2).Clear
                            'ActiveWorkbook.Worksheets("TF_X").range("M3:M" & getlastentry(ActiveWorkbook.Worksheets("TF_X"), 1) + 2).Clear
                        End If

                        ActiveWorkbook.Worksheets("TF_X").Rows(4).EntireRow.Insert
                        Call copyContent(ActiveWorkbook.ActiveSheet, "A" & i & ":N" & i, ActiveWorkbook.Worksheets("TF_X"), "C4:P4", 1)
                        'Call copyContent(ActiveWorkbook.ActiveSheet, "K" & i, ActiveWorkbook.Worksheets("TF_X"), "A4", 1)
                        'Call copyContent(ActiveWorkbook.ActiveSheet, "L" & i & ":M" & i, ActiveWorkbook.Worksheets("TF_X"), "N4", 1)
                        ActiveWorkbook.Worksheets("TF_X").range("A4") = "No"
                        ActiveWorkbook.Worksheets("TF_X").range("B4") = "Yes"
                
                        b_teamApproval = True
                
                    Case "ü"
                
                        Dim k As Long
                
                        b_okFalse = False
                
                        For k = 3 To 12
                
                            If ActiveSheet.Cells(i, k).Interior.Color = RGB(248, 203, 173) Then
                                arr_entries(4) = "Der Verlässlichkeitsgrad für Tab: " & ActiveSheet.Cells(i, 1) & " ist ausreichend. Jedoch wurden erhebliche Abweichung festgestellt. " & vbCrLf & arr_entries(4)
                                b_teamApproval = True
                                b_okFalse = True

                                Exit For
                            End If
                
                        Next k
                
                        'Copy Content to tempSheet ok
                        If b_okFalse Then
                    
                            'Clean up Range
                            If ActiveWorkbook.Worksheets("TF_ok").Cells(3, 1) <> "" Then
                                ActiveWorkbook.Worksheets("TF_ok").range("A3:P" & getlastentry(ActiveWorkbook.Worksheets("TF_ok"), 1) + 2).Clear
                                'ActiveWorkbook.Worksheets("TF_ok").range("A3:A" & getlastentry(ActiveWorkbook.Worksheets("TF_ok"), 1) + 2).Clear
                                'ActiveWorkbook.Worksheets("TF_ok").range("M3:M" & getlastentry(ActiveWorkbook.Worksheets("TF_ok"), 1) + 2).Clear
                            End If
                    
                            ActiveWorkbook.Worksheets("TF_ok").Rows(4).EntireRow.Insert
                            Call copyContent(ActiveWorkbook.ActiveSheet, "A" & i & ":N" & i, ActiveWorkbook.Worksheets("TF_ok"), "C4:P4", 1)
                            'Call copyContent(ActiveWorkbook.ActiveSheet, "K" & i, ActiveWorkbook.Worksheets("TF_ok"), "A4", 1)
                            'Call copyContent(ActiveWorkbook.ActiveSheet, "L" & i & ":M" & i, ActiveWorkbook.Worksheets("TF_ok"), "N4", 1)
                            ActiveWorkbook.Worksheets("TF_ok").range("A4") = "Yes"
                        End If
                
                End Select
        End Select

    Next i
    
    If arr_entries(2) <> "" Then ActiveWorkbook.Worksheets("TF_FIS").Rows(3).EntireRow.Delete
    If arr_entries(3) <> "" Then ActiveWorkbook.Worksheets("TF_X").Rows(3).EntireRow.Delete
    If arr_entries(4) <> "" Then ActiveWorkbook.Worksheets("TF_ok").Rows(3).EntireRow.Delete
    
    arr_entries(1) = b_teamApproval
    
    validationSummary = arr_entries

End Function
Sub import_data(arr_input As String)

    Dim arr_content() As String
    
    If InStr(arr_input, "|") <> 0 Then
    
        arr_content = Split(arr_input, "|")
    
    Else
    
        ReDim arr_content(0 To 0)
        arr_content(0) = arr_input
    
    End If
    
    Dim i As Long
        
    For i = 0 To UBound(arr_content)
    
        ActiveWorkbook.Worksheets(arr_content(i)).Unprotect
        
        ActiveWorkbook.Worksheets("Summary").Activate
        
        ActiveWorkbook.ActiveSheet.Cells(23, 1) = arr_content(i)
        Call copyContent(ActiveWorkbook.Worksheets(arr_content(i)), "D20", ActiveWorkbook.ActiveSheet, "B23", 1)

        Call copyContent(ActiveWorkbook.Worksheets(arr_content(i)), "I25:I32", ActiveWorkbook.ActiveSheet, "C23:J23", 2)
        Call copyContent(ActiveWorkbook.Worksheets(arr_content(i)), "H23", ActiveWorkbook.ActiveSheet, "K23", 1)
        Call copyContent(ActiveWorkbook.Worksheets(arr_content(i)), "I33", ActiveWorkbook.ActiveSheet, "L23", 1)
        Call copyContent(ActiveWorkbook.Worksheets(arr_content(i)), "I34", ActiveWorkbook.ActiveSheet, "M23", 1) 'EmailAdresse
        Call copyContent(ActiveWorkbook.Worksheets(arr_content(i)), "H34", ActiveWorkbook.ActiveSheet, "N23", 1) 'EmailDomain
        ActiveSheet.Hyperlinks.Add ActiveSheet.Cells(23, 15), "", ActiveWorkbook.Worksheets(arr_content(i)).Name & "!A1", TextToDisplay:="Tab " & ActiveWorkbook.Worksheets(arr_content(i)).Name
        ActiveSheet.Cells(23, 15).Font.Color = RGB(255, 0, 0)
        ActiveSheet.Rows(23).EntireRow.AutoFit
        ActiveSheet.Rows(23).EntireRow.Insert
        
        Call protectWS(ActiveWorkbook.Worksheets(arr_content(i)), "D25:I34")
        
        ActiveWorkbook.Worksheets("Summary").Activate
        
    Next i
    
    

End Sub

Sub resetSummarySheet(lastrow As Long)

    ActiveSheet.Rows(22).EntireRow.Insert
    ActiveSheet.Rows(lastrow + 2).EntireRow.Insert
    ActiveSheet.range("A23:L" & lastrow + 1) = ""

    ActiveSheet.Rows("23:" & lastrow + 1).EntireRow.Delete


    Call copyContent(ActiveWorkbook.Worksheets("basic_info"), "A14:L15", ActiveSheet, "A22:L23", 3)


End Sub

