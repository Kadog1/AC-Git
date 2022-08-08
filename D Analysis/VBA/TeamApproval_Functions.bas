Attribute VB_Name = "TeamApproval_Functions"
Sub processTeamApprovalReceived_CON()
    Dim rs As Object, strSQL As String, strSQLCount As String, counter As Integer, i As Integer, j As Integer, arrayRS() As Variant
        
    ' Load orders with AC_Status = 'TeamApprovalReceived'
    strSQLCount = "SELECT COUNT (*) AC_Status FROM [CAD].[dbo].[tCON_Orderbook] WHERE AC_Status = 'TeamApprovalReceived'"
    strSQL = "SELECT * FROM [CAD].[dbo].[tCON_Orderbook] WHERE AC_Status = 'TeamApprovalReceived'"
    
    Set rs = getRS(strSQLCount)
    counter = rs.Fields(0)
    If counter = 0 Then
        Exit Sub
    End If
    Set rs = getRS(strSQL)
    ReDim arrayRS(1 To counter, 1 To 73)
    
    'Alle gefundene Datensätze in ein Array laden
    For i = 1 To counter
        For j = 1 To 73
            arrayRS(i, j) = rs.Fields(j - 1)
        Next j
        rs.MoveNext
    Next i
    
    Dim orderNo As String
    For i = 1 To counter
        orderNo = arrayRS(i, 2)
        Call processTeamApproval(orderNo)
    Next i

End Sub

Sub processTeamApprovalReceived_AC()
    Dim rs As Object, strSQL As String, strSQLCount As String, counter As Integer, i As Integer, j As Integer, arrayRS() As Variant
        
    ' Load orders with AC_Status = 'TeamApprovalReceived'
    strSQLCount = "SELECT COUNT (*) AC_Status FROM [CAD].[dbo].[tAC_Orderbook] WHERE AC_Status = 'TeamApprovalReceived'"
    strSQL = "SELECT * FROM [CAD].[dbo].[tAC_Orderbook] WHERE AC_Status = 'TeamApprovalReceived'"
    Set rs = getRS(strSQLCount)
    counter = rs.Fields(0)
    If counter = 0 Then
        Exit Sub
    End If
    
    Set rs = getRS(strSQL)
    ReDim arrayRS(1 To counter, 1 To 22)
    
    'Alle gefundene Datensätze in ein Array laden
    For i = 1 To counter
        For j = 1 To 22
            arrayRS(i, j) = rs.Fields(j - 1)
        Next j
        rs.MoveNext
    Next i
    
    Dim orderNo As String
    For i = 1 To counter
        orderNo = arrayRS(i, 1)
        Call processTeamApproval(orderNo)
    Next i
    
End Sub

Sub processTeamApproval(orderNo As String)
    Debug.Print "Processing Order " & orderNo & "..."
    Dim strFilePath As String, strAttPath As String, orderbook As String
    ' Define paths
    If Left(orderNo, 3) = "CON" Then
        orderbook = "tCON_Orderbook"
        strFilePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & orderNo & "\3. Team Approval\"
    Else
        orderbook = "tAC_Orderbook"
        strFilePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich\C Workplace\" & orderNo & "\3. Team Approval\"
    End If
    
    strFile = Dir(strFilePath & "*3_CAD-Adressabgleich Team Approval_Template*.xls*")
    If Len(strFile) > 0 Then
        Dim wbTeamApproval As Workbook
        Set wbTeamApproval = Workbooks.Open(strFilePath & strFile)

        If WorksheetExists("Versandliste", wbTeamApproval) = False Then
           wbTeamApproval.Close
           Call updateRS("UPDATE " & orderbook & " Set AC_Status = 'VersandlisteE' WHERE OrderNo = '" & orderNo & "'")
           Exit Sub
        End If
        lastrow = wbTeamApproval.Sheets("Summary").Cells(Rows.Count, 4).End(xlUp).row
        For j = 1 To lastrow - 29
           Dim row As Integer
           row = 29 + j
           If wbTeamApproval.Sheets("Summary").range("B" & row).Value = "Ja" Then
               Call updateVersandliste(wbTeamApproval.Sheets("Summary").range("D" & row & ":D" & row), wbTeamApproval)
           End If
        Next j
        
        Call createVersandlisteFile(orderNo, wbTeamApproval)
        Call updateRS("UPDATE " & orderbook & " Set AC_Status = 'VersandlisteDone' WHERE OrderNo = '" & orderNo & "'")
        strFile = Dir
        wbTeamApproval.Close
    Else
        Call updateRS("UPDATE " & orderbook & " Set AC_Status = 'VersandlisteE' WHERE OrderNo = '" & orderNo & "'")
    End If
End Sub

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

Sub createVersandlisteFile(orderNo As String, wb As Workbook)
    ' Open Versandliste
    Dim wbVersandliste As Workbook, saveName As String, orderbook As String, strFilePath As String
    Set wbVersandliste = Application.Workbooks.Open("C:\Users\DEPPOPP1\OneDrive - EY\Desktop\Task for Viktor\Adressabgleich\Feature Versandliste Duplicates\5_CAD-Adressabgleich Adressen für externe Bestätigungen_Template.xlsx")
    
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

Public Function getRS(SQL As String) As ADODB.RecordSet
    'FP20210826
    'The function sets connection to SQL db and pulls a recordset acc to provided SQL-querry
    'Input: SQL querry as String
    'Output: ADO Recordset
    
    Dim sConnString As String
    Dim conn As New ADODB.Connection
    Dim myRecordSet As New ADODB.RecordSet

    'Create connection string
    sConnString = "Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;"
    'Open connection to SQL db
    conn.Open sConnString
    'Create RS
    myRecordSet.CursorLocation = adUseClient
    myRecordSet.ActiveConnection = conn

    myRecordSet.Open SQL
        
    Set getRS = myRecordSet.Clone
    
    'Clean up
    myRecordSet.Close
    Set myRecordSet = Nothing

End Function
Public Function updateRS(SQL As String) As ADODB.RecordSet
    'FP20210826
    'The function sets connection to SQL db and pulls a recordset acc to provided SQL-querry
    'Input: SQL querry as String
    'Output: ADO Recordset
    
    Dim sConnString As String
    Dim conn As New ADODB.Connection
    Dim myRecordSet As New ADODB.RecordSet

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

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

