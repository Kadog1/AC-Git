Attribute VB_Name = "TeamApprovalReceived"
Sub processTeamApproval()
    Dim rs As Object, strSQL As String, strSQLCount As String, counter As Integer, i As Integer, j As Integer, arrayRS() As Variant
        
    ' Load orders with AC_Status = 'TeamApprovalReceived'
    strSQLCount = "SELECT COUNT (*) AC_Status FROM [CAD].[dbo].[tCON_Orderbook] WHERE AC_Status = 'TeamApprovalReceived'"
    strSQL = "SELECT * FROM [CAD].[dbo].[tCON_Orderbook] WHERE AC_Status = 'TeamApprovalReceived'"
    Set rs = getRS(strSQLCount)
    counter = rs.Fields(0)
    Set rs = getRS(strSQL)
    ReDim arrayRS(1 To counter, 1 To 73)
    
    'Alle gefundene Datensätze in ein Array laden
    For i = 1 To counter
        For j = 1 To 73
            arrayRS(i, j) = rs.Fields(j - 1)
        Next j
        rs.MoveNext
    Next i
    
    ' Save TeamApproval from .msg, open Team Approval, updateVersandliste, createVersandlisteFile
    Dim msg As Outlook.MailItem, att As Outlook.Attachment, strFilePath As String, strAttPath As String, orderNo As String
    
    'path for creating
    For i = 1 To UBound(arrayRS)
        orderNo = arrayRS(i, 2)
        strFilePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & orderNo & "\3. Team Approval\"
        strAttPath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & orderNo & "\3. Team Approval\"
        strFile = Dir(strFilePath & "*.msg")
        Do While Len(strFile) > 0
            Set olApp = New Outlook.Application
            Set msg = olApp.CreateItemFromTemplate(strFilePath & strFile)
            If msg.Attachments.Count > 0 Then
                 For Each att In msg.Attachments
                     ' Update Versandliste
                     att.SaveAsFile strAttPath & att.FileName
                     Dim wbTeamApproval As Workbook
                     Set wbTeamApproval = Workbooks.Open(strAttPath & att.FileName)
                     lastrow = wbTeamApproval.Sheets("Summary").Cells(Rows.Count, 4).End(xlUp).row
                     For j = 1 To lastrow - 29
                        Dim row As Integer
                        row = 29 + j
                        If wbTeamApproval.Sheets("Summary").range("B" & row).Value = "Ja" Then
                            Call updateVersandliste(wbTeamApproval.Sheets("Summary").range("D" & row & ":D" & row), wbTeamApproval)
                        End If
                     Next j

                     Call createVersandlisteFile(wbTeamApproval)
                     
                 Next
            End If
            strFile = Dir
        Loop
    Next i
    

    
    ' Create Versandliste

End Sub
