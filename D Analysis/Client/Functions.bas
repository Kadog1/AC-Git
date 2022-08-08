Attribute VB_Name = "Functions"
Option Compare Database
Public AlltasksQuery As String
Public strTestTable As String

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Public Sub hideNavPRib()

    'NavigationPane
    DoCmd.NavigateTo ("acNavigationCategoryObjectType")
    DoCmd.RunCommand (acCmdWindowHide)
    'Ribbon Menu
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    'hide Statusbar
    Application.SetOption "Show Status Bar", False
        
End Sub
    
Sub unhideNavPRib()
    
    'NavigationPane
    DoCmd.SelectObject acTable, , True
    'Ribbon Menu
    DoCmd.ShowToolbar "Ribbon", acToolbarYes
    'Statusbar
    Application.SetOption "Show Status Bar", False
    
End Sub

Public Function getRS(sql As String) As ADODB.RecordSet

          '    Dim cnxn As ADODB.Connection
          '    Set cnxn = CurrentProject.AccessConnection
          '    Dim myRecordSet As New ADODB.Recordset

          Dim sConnString As String
          Dim conn As New ADODB.Connection
          Dim myRecordSet As New ADODB.RecordSet

            Call isTesting

          'Create connection string
          '#BM-1/2019-08-08 10:49
1         sConnString = "Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;"
          '#BM-1/2019-08-08 10:49 End
          'Open connection to SQL db
2         conn.Open sConnString

    
          'Create RS
3         myRecordSet.CursorLocation = adUseClient
4         myRecordSet.ActiveConnection = conn

          Debug.Print sql
5         myRecordSet.Open sql
          
6         Set getRS = myRecordSet.Clone
        
          'Clean up
          myRecordSet.Close
          Set myRecordSet = Nothing
          'conn.Close

End Function

Public Sub updateSQL(strSQL As String)

 

    Dim RecordSet As ADODB.RecordSet
    Dim Connection As ADODB.Connection

    Call isTesting

    Set Connection = CreateObject("ADODB.Connection")

 

    Connection.ConnectionString = "Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;"

 

    Connection.Open
    Debug.Print strSQL

 

    Set RecordSet = Connection.Execute(strSQL)

 

    Connection.Close
End Sub

Private Sub testOrderField()

    Dim sql As String, t_case As String
    
    sql = "SELECT [2nd_Request_gewünscht] From tCon_Orderbook" & strTestTable & ";"

    t_case = "ExportAll"
    
    'Dim cnxn As ADODB.Connection
    'Set cnxn = CurrentProject.AccessConnection
    Dim myRecordSet As New ADODB.RecordSet
    
    ' Execute SQL Query
    'myRecordSet.Open SQL, cnxn, adOpenStatic, adLockOptimistic
    Set myRecordSet = getRS(sql)
    
    Debug.Print myRecordSet.RecordCount

    'Call ExportRecordset(myRecordSet, t_case)
    
    cnxn.Close
    
    Set cnxn = Nothing

    Dim oExcel            As Excel.Application
    
    Dim oExcelWrkBk       As Excel.Workbook
    Dim oExcelWrkSht      As Excel.Worksheet
    Dim oExcelWrkShtFinal As Excel.Worksheet

    'Start Excel
    Set oExcel = CreateObject("Excel.Application")
            oExcel.DisplayStatusBar = True
            oExcel.EnableEvents = True
            oExcel.ScreenUpdating = True
            oExcel.Visible = True
            oExcel.displayalerts = True


End Sub
Public Function getTableValue(OrderNo As String, Column As String) As String

    Dim rs As ADODB.RecordSet
    
    Dim arrTemp() As Variant

    Set rs = getRS("SELECT " & Column & " From tCon_Orderbook" & strTestTable & " WHERE OrderNo='" & OrderNo & "'")

    arrTemp = rs.GetRows(rs.RecordCount)
    If IsNull(arrTemp(0, 0)) Then
        getTableValue = ""
    Else
        getTableValue = arrTemp(0, 0)
    End If

    rs.Close
    Set rs = Nothing

End Function

Public Sub RowsSelected(NewUser As String, orderbHeader As String)
   ' On Error GoTo Errormsg
    
    Call isTesting
    
    DoCmd.SetWarnings False
    
    Dim ctlList As Control, varItem As Variant, sql As String, selOrderNo As String, DBtable As String
    If orderbHeader = "Mandant_Anr" Or orderbHeader = "Mandant_Name" Or orderbHeader = "Mandant_Email" Then
        DBtable = "tCON_ClientContact" & strTestTable
    ElseIf orderbHeader = "Name_Bank" Or orderbHeader = "Portal_Bank" Then
        DBtable = "tCON_BankConfirmation" & strTestTable
    Else
        DBtable = "tCON_Orderbook" & strTestTable
    End If

    Set ctlList = Forms![Update_Confi Client].Alltasks
    

    For Each varItem In ctlList.ItemsSelected
    
        selOrderNo = ctlList.Column(0, varItem)
        
        sql = "UPDATE " & DBtable & " SET " & orderbHeader & "=" & handleSQLValues(NewUser) & " WHERE OrderNo='" & selOrderNo & "'"
        updateSQL (sql)
        If NewUser = "Storno" And orderbHeader <> "Adressen_Status" And orderbHeader <> "Forensics_Status" Then
            sql = "UPDATE " & DBtable & " SET Storno='manual storno' WHERE OrderNo='" & selOrderNo & "'"
            updateSQL (sql)
        End If
    Next varItem
    
    DoCmd.SetWarnings True
    
    On Error GoTo 0
    Exit Sub
    
Errormsg:
    If CStr(Err.Number) = "3044" Then
        MsgBox ("No connection to the server. Please check the connection and try again")
    Else
        MsgBox ("An error is occured. Please restart the client. Error Info: " & Err.Description)
    End If
    
End Sub

Sub testRS2XLSX()

    Dim sql As String
    sql = "SELECT *  From tCon_Orderbook" & strTestTable & " LEFT OUTER JOIN tCON_ClientContact" & strTestTable & " ON tCON_Orderbook" & strTestTable & ".OrderNo = tCON_ClientContact" & strTestTable & ".OrderNo;"
    
    'Dim cnxn As ADODB.Connection
    'Set cnxn = CurrentProject.AccessConnection
    Dim myRecordSet As New ADODB.RecordSet
    
    ' Execute SQL Query
    'myRecordSet.Open SQL, cnxn, adOpenStatic, adLockOptimistic
    myRecordSet = getRS(sql)

    Call ExportRecordset(myRecordSet)

End Sub



Sub RearrangeColumns(modus As Integer, oExcel As Object)

    Dim oExcelWrkBk       As Object
    Dim oExcelWrkSht      As Object
    Dim oExcelWrkShtFinal      As Object
    
    ' Set Objects
    Set oExcelWrkBk = oExcel.ActiveWorkbook
    Set oExcelWrkSht = oExcelWrkBk.Worksheets("t_Export all data")
    
    ' Start a new worksheet
    oExcelWrkBk.Worksheets.Add().name = "Export all data"
    Set oExcelWrkShtFinal = oExcelWrkBk.Sheets("Export all data")

    Dim header() As Variant
    
    'header = Array(1, 2, 3, 4, 5, 15, 16, 6, 13, 14, 17, 31, 20, 19, 22, 21, 18, 7, 8, 23, 9, 10, 12, 11, 25, 27, 55, 28, 29, 30, 59, 60, 61, 36, 37, 39, 24, 40, 41, 38, 44, 56, 57, 42, 43, 26, 58, 35, 45, 46, 47, 48, 49, 52, 53, 54, 32, 33, 34, 50, 51)
    header = Array(1, 2, 3, 4, 5, 15, 16, 6, 13, 14, 17, 31, 20, 19, 22, 21, 18, 7, 8, 23, 9, 10, 12, 11, 25, 27, 55, 28, 29, 30, 36, 37, 39, 24, 40, 41, 38, 44, 56, 57, 42, 43, 26, 58, 35, 45, 46, 47, 48, 49, 52, 53, 54, 32, 33, 34, 50, 51)
    
    Dim i As Integer
    
    For i = 0 To UBound(header)
 
    ' Cut Column
    oExcelWrkSht.Columns(header(i)).Cut

    ' Paste Column
    oExcelWrkShtFinal.Columns(i + 1).Insert

    Next i

    
End Sub

Public Function SetColumnWidths(ctrList As Control) As String
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim colNum As Integer
    Dim rowNum As Integer
    Dim aryLen() As Single
    Dim ctrValue As String
    Dim colRatio As Single
    Dim colWidths As String

    '--- Find and store the length of the largest piece of text in each column ----

    colNum = ctrList.ColumnCount - 1
    rowNum = ctrList.ListCount - 1

    ReDim aryLen(colNum) 'make the array's slots equal to the number of columns

    For X = 0 To colNum ' for every column in list box
        For Y = 0 To rowNum ' for every row in the column (including heading)
            ctrValue = ctrList.Column(X, Y)
            If Len(ctrValue) > aryLen(X) Then 'if the length of current record is larger than already stored
                aryLen(X) = Len(ctrValue) 'store the largest value length
            End If
        Next Y
    Next X



    '--- Set the column widths ---

    colRatio = 0.2

    For i = 0 To colNum 'For each stored maximum lenght
        If i = colNum Then
            colWidths = colWidths & Round((aryLen(i) * colRatio), 0) & " cm"
        Else
            colWidths = colWidths & Round((aryLen(i) * colRatio), 0) & " cm;"
            
        End If
    Next i


    '---Return the calculated value---

    SetColumnWidths = colWidths
    Debug.Print SetColumndWidths

End Function

Private Sub Alltasks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer

    Dim RSpreparer As ADODB.RecordSet

    Dim Spl As String
    Dim ColumnID As Integer, anzCol As Integer
    Dim width As Double, stepWidth As Double
    Dim columnWidths() As Double


    Select Case Me.Label76.Caption
    
        Case "Start"
            ReDim columnWidths(1 To 11)
            Alltasks.ColumnCount = 11
            t_width = 11
        Case "Versand"
            ReDim columnWidths(1 To 26)
            Alltasks.ColumnCount = 26
            t_width = 26
        Case "Auftragsende"
            ReDim columnWidths(1 To 25)
            Alltasks.ColumnCount = 25
            t_width = 25
        Case "Rest"
            ReDim columnWidths(1 To 29)
            Alltasks.ColumnCount = 29
            t_width = 29
    End Select
    
    width = Alltasks.width
    anzCol = Alltasks.ColumnCount
    
    stepWidth = Round(width / 45, 3)
    ' Gleichtverteilung der Spaltenbreiten wäre 3*stepWidth pro Spalte
    ' Wenn alles addiert 45 * stepWidth ergibt, stimmt die Breite
    columnWidths(1) = 3 * stepWidth
    columnWidths(2) = 4 * stepWidth
    columnWidths(3) = 4 * stepWidth
    columnWidths(4) = 3 * stepWidth
    columnWidths(5) = 3 * stepWidth
    columnWidths(6) = 3 * stepWidth
    columnWidths(7) = 2 * stepWidth
    columnWidths(8) = 2 * stepWidth
    columnWidths(9) = 2 * stepWidth
    columnWidths(10) = 5 * stepWidth
    columnWidths(11) = 3 * stepWidth
    columnWidths(12) = 3 * stepWidth
    
    If Me.Label76.Caption = "Rest" Then
        columnWidths(9) = 5 * stepWidth
        columnWidths(11) = 5 * stepWidth
        columnWidths(12) = 5 * stepWidth
    End If
    columnWidths(13) = 4 * stepWidth
    columnWidths(14) = 3 * stepWidth
    columnWidths(15) = 3 * stepWidth
    columnWidths(16) = 2 * stepWidth
    columnWidths(17) = 2 * stepWidth
    columnWidths(18) = 2 * stepWidth
    columnWidths(19) = 2 * stepWidth
    If Me.Label76.Caption = "Start" Then columnWidths(20) = width - (42.5 * stepWidth)
    
    If Me.Label76.Caption <> "Start" Then
        columnWidths(20) = 2 * stepWidth
        columnWidths(21) = 2 * stepWidth
        columnWidths(22) = 2 * stepWidth
        columnWidths(23) = 2 * stepWidth
        columnWidths(24) = 2 * stepWidth
    End If
    
    If Me.Label76.Caption = "Auftragsende" Then columnWidths(25) = width - (42.5 * stepWidth)
    
    If Me.Label76.Caption = "Versand" Then columnWidths(26) = width - (42.5 * stepWidth)
    
    If Me.Label76.Caption = "Rest" Then
        columnWidths(25) = 2 * stepWidth
        columnWidths(26) = 2 * stepWidth
        columnWidths(27) = 2 * stepWidth
        columnWidths(28) = width - (42.5 * stepWidth)  '2.5 * stepWidth
    End If
    
    '#JK-6/2020-12-17 START
    If Me.Label76.Caption = "Adressen" Then
        columnWidths(1) = 3 * stepWidth
        columnWidths(2) = 4 * stepWidth
        columnWidths(3) = 4 * stepWidth
        columnWidths(4) = 2 * stepWidth
        columnWidths(5) = 3.5 * stepWidth
        columnWidths(6) = 1.5 * stepWidth
        columnWidths(7) = 3 * stepWidth
        columnWidths(8) = 1.5 * stepWidth
        columnWidths(9) = 1.5 * stepWidth
        columnWidths(10) = 2 * stepWidth
        columnWidths(11) = 2 * stepWidth
        columnWidths(12) = 3.5 * stepWidth
        columnWidths(13) = 2 * stepWidth
        columnWidths(14) = 5 * stepWidth
        columnWidths(15) = 3 * stepWidth
        columnWidths(16) = width - (42.5 * stepWidth)
    End If
    '#JK-6/2020-12-17
    
    Dim sum As Double
    Dim str_colWidths As String
    For i = 1 To anzCol
        sum = sum + columnWidths(i)
        str_colWidths = str_colWidths & CStr(columnWidths(i)) & ";"
    Next i
    
    str_colWidths = Left(str_colWidths, Len(str_colWidths) - 1)
    
    'Alltasks.columnWidths = str_colWidths

    Dim leftBound As Double, rightBound As Double
    leftBound = 0
    rightBound = 0

    For i = 0 To anzCol - 1
        If i + 1 > anzCol Then Exit For
        rightBound = rightBound + columnWidths(i + 1)
        If i > 0 Then leftBound = leftBound + columnWidths(i)
        If leftBound < X And X < rightBound Then
        

            Select Case Me.Label76.Caption
            
                Case "Start":
                
                    Select Case Me.Label93.Caption:
                        Case "Start1":
                            Select Case i:
                        
                                Case 0: Spl = "OrderNo"
                                Case 1: Spl = "NameOfCompanyGroup"
                                Case 2: Spl = "Client"
                                Case 3: Spl = "Tool"
                                Case 4: Spl = "Confirmation"
                                Case 5: Spl = "TypeConfi"
                                Case 6: Spl = "DateConfi"
                                Case 7: Spl = "YearEnd"
                                Case 8: Spl = "EngCode"
                                Case 9: Spl = "EngContact"
                                Case 10: Spl = "DeliveryType"
                            
                            End Select
                        Case "Start2":
                            Select Case i:
                                Case 0: Spl = "OrderNo"
                                Case 1: Spl = "Client"
                                Case 2: Spl = "Job_Status"
                                Case 3: Spl = "CAD_Bearbeiter"
                                Case 4: Spl = "CAD_Bearbeiter_vorherig"
                                Case 5: Spl = "Datum_Versand_EMail_Erstkontakt"
                                Case 6: Spl = "Informationen_zum_Auftrag"
                                Case 7: Spl = "Versand_durch"
                                Case 8: Spl = "PlannedDelivery"
                                Case 9: Spl = "First_Upload_Gewünscht"
                                Case 10: Spl = "DatumAuftragsende"
                            End Select
                    End Select
                
                '#JK-7/2020-12-09 START
                Case "Adressen":
                    Select Case i:
                        Case 0: Spl = "OrderNo"
                        Case 1: Spl = "NameOfCompanyGroup"
                        Case 2: Spl = "Client"
                        Case 3: Spl = "Tool"
                        Case 4: Spl = "Confirmation"
                        Case 5: Spl = "PlannedDelivery"
                        Case 6: Spl = "TypeConfi"
                        Case 7: Spl = "DeliveryType"
                        Case 8: Spl = "DateConfi"
                        Case 9: Spl = "YearEnd"
                        Case 10: Spl = "EngCode"
                        Case 11: Spl = "EngContact"
                        Case 12: Spl = "Country"
                        Case 13: Spl = "Adressen_Bearbeiter"
                        Case 14: Spl = "Adressen_Status"
                        Case 15: Spl = "Forensics_Status"
                        
                    End Select
                '#JK-7/2020-12-09 START
                
                Case "Versand":
                
                    Select Case i
                        Case 0: Spl = "OrderNo"
                        Case 1: Spl = "NameOfCompanyGroup"
                        Case 2: Spl = "Client"
                        Case 3: Spl = "Tool"
                        Case 4: Spl = "Confirmation"
                        Case 5: Spl = "TypeConfi"
                        Case 6: Spl = "DateConfi"
                        Case 7: Spl = "YearEnd"
                        Case 8: Spl = "EngCode"
                        Case 9: Spl = "EngContact"
                        Case 10: Spl = "DeliveryType"
                        Case 11: Spl = "Datum_Versand_EMail_Erstkontakt"
                        Case 12: Spl = "Informationen_zum_Auftrag"
                        Case 13: Spl = "Versand_durch"
                        Case 14: Spl = "Job_Status"
                        Case 15: Spl = "CAD_Bearbeiter"
                        Case 16: Spl = "CAD_Bearbeiter_vorherig"
                        Case 17: Spl = "PlannedDelivery"
                        Case 18: Spl = "Eingangsdatum_CAD"
                        Case 19: Spl = "Versanddatum_CAD"
                        Case 20: Spl = "Anzahl_Bestätigungen"
                        Case 21: Spl = "Antwort_Deadline"
                        Case 22: Spl = "Kommentar_Versand"
                        Case 23: Spl = "2nd_Request_gewünscht"
                        Case 24: Spl = "Datum_2nd_Request"
                        Case 25: Spl = "Date2ndReq"
                    End Select
            
                Case "Auftragsende:"
                
                    Select Case i
                        Case 0: Spl = "OrderNo"
                        Case 1: Spl = "NameOfCompanyGroup"
                        Case 2: Spl = "Client"
                        Case 3: Spl = "Tool"
                        Case 4: Spl = "Confirmation"
                        Case 5: Spl = "TypeConfi"
                        Case 6: Spl = "DateConfi"
                        Case 7: Spl = "YearEnd"
                        Case 8: Spl = "EngCode"
                        Case 9: Spl = "EngContact"
                        Case 10: Spl = "DeliveryType"
                        Case 11: Spl = "Informationen_zum_Auftrag"
                        Case 12: Spl = "Job_Status"
                        Case 13: Spl = "CAD_Bearbeiter"
                        Case 14: Spl = "CAD_Bearbeiter_vorherig"
                        Case 15: Spl = "DatumAuftragsende"
                        Case 16: Spl = "Anzahl_erhaltene_Salden"
                        Case 17: Spl = "Final_Senior_Review_durch"
                        Case 18: Spl = "Datum_final_Review"
                        Case 19: Spl = "Versanddatum_Originale"
                        Case 20: Spl = "Empfänger_Originale"
                        Case 21: Spl = "Empfänger_Nachläufer"
                        Case 22: Spl = "Datum_Feedbackanfrage"
                        Case 23: Spl = "Datum_Reminder"
                        Case 24: Spl = "Feedback_erhalten"
                    End Select
                
                Case "Rest":
                
                    Select Case i
                        Case 0: Spl = "OrderNo"
                        Case 1: Spl = "NameOfCompanyGroup"
                        Case 2: Spl = "Client"
                        Case 3: Spl = "OrderNo"
                        Case 4: Spl = "Tool"
                        Case 5: Spl = "Confirmation"
                        Case 6: Spl = "TypeConfi"
                        Case 7: Spl = "YearEnd"
                        Case 8: Spl = "EngCode"
                        Case 9: Spl = "EngPartner"
                        Case 10: Spl = "EngManager"
                        Case 11: Spl = "EngContact"
                        Case 12: Spl = "OtherContact"
                        Case 13: Spl = "ID"
                        Case 14: Spl = "POrderNo"
                        Case 15: Spl = "Storno"
                        Case 16: Spl = "GISID"
                        Case 17: Spl = "BelongingToCompanyGroup"
                        Case 18: Spl = "BusinessEstablishmentOAreAvailable"
                        Case 19: Spl = "BusinessEstablishmentOAre"
                        Case 20: Spl = "TypeFinancials"
                        Case 21: Spl = "EYStand"
                        Case 22: Spl = "EngName"
                        Case 23: Spl = "Lang"
                        Case 24: Spl = "OtherTypeConfi"
                        Case 25: Spl = "ConfiFlag"
                        Case 26: Spl = "ConfiMailSent"
                        Case 27: Spl = "TemplateSent"
                        Case 28: Spl = "AuditiSent"
                    End Select

            End Select
            ColumnID = i
            Me!LblHeader.Caption = Spl
            If Spl = "Tool" Then
                Me!TxtEntry.Visible = False
                Me!cbEntry.Visible = True
                Do While Me!cbEntry.ListCount > 0
                    Me!cbEntry.RemoveItem (0)
                Loop
                Me!cbEntry.AddItem "eConfirmations"
                Me!cbEntry.AddItem "CAD"
                Me!cbEntry.value = Me!Alltasks.Column(ColumnID, Me.Alltasks.ListIndex + 1)
            ElseIf Spl = "1_CAD_Preparer" Or Spl = "2_CAD_Preparer" Then
                Me!TxtEntry.Visible = False
                Me!cbEntry.Visible = True
                Do While Me!cbEntry.ListCount > 0
                    Me!cbEntry.RemoveItem (0)
                Loop
                Set RSpreparer = getRS("SELECT * FROM tUser")
                Do While RSpreparer.EOF = False
                    Me!cbEntry.AddItem RSpreparer("PreparerID")
                    RSpreparer.MoveNext
                Loop
                RSpreparer.Close
                Me!cbEntry.value = Me!Alltasks.Column(ColumnID, Me.Alltasks.ListIndex + 1)
            ElseIf Spl = "Job_Status" Then
                Me!TxtEntry.Visible = False
                Me!cbEntry.Visible = True
                Do While Me!cbEntry.ListCount > 0
                    Me!cbEntry.RemoveItem (0)
                Loop
                Me!cbEntry.AddItem "Not Started"
                Me!cbEntry.AddItem "In Progress"
                Me!cbEntry.AddItem "Closed"
                Me!cbEntry.AddItem "Storno"
                Me!cbEntry.value = Me!Alltasks.Column(ColumnID, Me.Alltasks.ListIndex + 1)
                
            '#JK-8/2020-12-09 START
            ElseIf Spl = "Adressen_Status" Then
                Me!TxtEntry.Visible = False
                Me!cbEntry.Visible = True
                Do While Me!cbEntry.ListCount > 0
                    Me!cbEntry.RemoveItem (0)
                Loop
                Me!cbEntry.AddItem "n/a"
                Me!cbEntry.AddItem "Not Started"
                Me!cbEntry.AddItem "In Progress"
                Me!cbEntry.AddItem "ET review"
                Me!cbEntry.AddItem "ReadyForOutput"
                Me!cbEntry.AddItem "ET completed"
                Me!cbEntry.AddItem "ET-Nachtrag"
                Me!cbEntry.AddItem "ET-Nachtrag Review"
                Me!cbEntry.AddItem "ET Nachtrag completed"
                Me!cbEntry.AddItem "Storno"
                Me!cbEntry.AddItem "Closed"
                Me!cbEntry.value = Me!Alltasks.Column(ColumnID, Me.Alltasks.ListIndex + 1)
            ElseIf Spl = "Forensics_Status" Then
                Me!TxtEntry.Visible = False
                Me!cbEntry.Visible = True
                Do While Me!cbEntry.ListCount > 0
                    Me!cbEntry.RemoveItem (0)
                Loop
                Me!cbEntry.AddItem "n/a"
                Me!cbEntry.AddItem "Forensics-Optional"
                Me!cbEntry.AddItem "Forensics-Pflicht"
                Me!cbEntry.AddItem "ET completed"
                Me!cbEntry.AddItem "Storno"
                Me!cbEntry.AddItem "Closed"
                Me!cbEntry.value = Me!Alltasks.Column(ColumnID, Me.Alltasks.ListIndex + 1)
            '#JK-8/2020-12-09 END
            
            Else
                Me!TxtEntry.Visible = True
                Me!cbEntry.Visible = False
                Do While Me!cbEntry.ListCount > 0
                    Me!cbEntry.RemoveItem (0)
                Loop
                Me!TxtEntry = Me!Alltasks.Column(ColumnID, Me.Alltasks.ListIndex + 1)
            End If
        End If
    Next i

End Sub


'#JK-14/2020-12-16 START
Public Sub isTesting()

    Dim isTest As Boolean
    
    isTest = False
    
    If isTest = True Then
        strTestTable = "_TEST"
    
    Else
        strTestTable = ""
    
    End If
End Sub
'#JK-14/2020-12-16 END



'#JK-2/2020-12-17 START

Public Function handleSQLValues(value As String, Optional mode As Integer = 1) As String

    'Modi:
    '   1: String
    '   2: Boolean
    
    If Not (mode = 1 Or mode = 2) Then
        Debug.Print ("Für die Function handleSQLValues würde ein ungültiger mode eingegeben")
        handleSQLValues = value
        Exit Function
        
    End If

    If IsEmpty(value) Then
        handleSQLValues = "NULL"
    
    Else
        If mode = 1 Then
            value = Replace(value, "'", """")
        
        ElseIf mode = 2 Then
            If UCase(value) = "TRUE" Or UCase(value) = "WAHR" Then
                value = "1"
                
            ElseIf UCase(value) = "FALSE" Or UCase(value) = "FALSCH" Then
                value = "0"
            
            End If
                
        End If
        
        handleSQLValues = "'" & value & "'"
        
    End If
        

End Function

'#JK-2/2020-12-17 END
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
    Debug.Print t_template
    objStream.LoadFromFile (t_template)
    
    body = objStream.ReadText()
    
    parseBody = body
    
    ' Clean up
    objStream.Close
    Set objStream = Nothing

End Function

Sub createEmailDraft(attachmentPath As String, body As String, subject As String, recipient As String, ccRecipient As String)

    Dim OutMail As Object
    Dim OutApp As Object
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    'Email Erstellung

    With OutMail

        .Attachments.Add attachmentPath
        .Attachments.Add "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\EY_Logo_Beam_RGB.png", olByValue, 0
        .Attachments.Add "\\devidvapfl04.ey.net\04em1015\H\Heartbeat-CAD Team\CAD_Brand_Communications\CAD_Signatur+Banner\190807_CAD_Signature-Banner_500px_mittel.png", olByValue, 0
        '.Attachments.Add "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\190807_CAD_Signature-Banner_500px_mittel.png", olByValue, 0
        
        .subject = subject
        .To = recipient
        .CC = ccRecipient
        .HTMLBody = body
        
        ' Funktioniert nur in Outlook - Bitte live schalten
        '.Recipients.ResolveAll

        'Outlook Draft Fenster wird geöffnet
        .Display

      
    End With

    'Schließe Outlook Object
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub


Public Function getQueryPOrderNo(strOrderNo As String) As String

    Dim queryPOrderNo As String, strPOrderNo As String, fieldPOrderNo As String, strSelectColumns As String, strOrderTable As String
    Dim rsPOrderNo As ADODB.RecordSet
    strPOrderNo = "POrderNo"
        
    ' Get column names and table name for each solution
    strSelectColumns = "OrderNo, Confirmation, Client, PlannedDelivery, YearEnd, EngContact, EngManager, OtherContact"
    strOrderTable = "[CAD].[dbo].[tCON_Orderbook" & strTestTable & "]"
    If Left(strOrderNo, 2) = "AC" Then
        strSelectColumns = "OrderNo, NULL As Confirmation, client, NULL As PlannedDelivery, Periodend, engcontact, engmngr, EngOtherContact"
        strOrderTable = "[CAD].[dbo].[tAC_Orderbook" & strTestTable & "]"
    ElseIf Left(strOrderNo, 2) = "AP" Then
        strSelectColumns = "OrderNo, 'Creditor' As Confirmation, GISName As client, NULL As PlannedDelivery, oh.PeriodEnd As Periodend, EngContact_oi AS engcontact, EngManager_oi AS engmngr, ADD_Contact1_oi AS EngOtherContact"
        strOrderTable = "[CAD].[dbo].[tT_CADDB_CAD_OrderedItems" & strTestTable & "] oi LEFT JOIN [CAD].[dbo].[tT_CADDB_CAD_OrderHeader] oh ON oi.CADOrderNo = oh.CADOrderNo"
        strPOrderNo = "oi.CADOrderNo"
    ElseIf Left(strOrderNo, 2) = "AR" Then
        strSelectColumns = "oi.OrderNo, 'Debtitor' As Confirmation, oi.GISName As client, NULL As PlannedDelivery, oh.PeriodEnd As Periodend, oi.EngContact_oi AS engcontact, oi.EngManager_oi AS engmngr, oi.ADD_Contact1_oi AS EngOtherContact"
        strOrderTable = "[CAD].[dbo].[tT_CADDB_CAD_OrderedItems" & strTestTable & "] oi LEFT JOIN [CAD].[dbo].[tT_CADDB_CAD_OrderHeader] oh ON oi.CADOrderNo = oh.CADOrderNo"
        strPOrderNo = "oi.CADOrderNo"
    End If
    
    ' Get POrderNo / CADOrderNo
    If Left(strOrderNo, 2) = "AP" Or Left(strOrderNo, 2) = "AR" Then
        Set rsPOrderNo = getRS("Select oi.CADOrderNo " & _
        "FROM [CAD].[dbo].[tT_CADDB_CAD_OrderedItems" & strTestTable & "] AS oi " & _
        "WHERE OrderNo = '" & strOrderNo & "';")
        fieldPOrderNo = rsPOrderNo.Fields("CADOrderNo")
    Else
        Set rsPOrderNo = getRS("Select POrderNo " & _
        "FROM [CAD].[dbo].[tCON_Orderbook" & strTestTable & "] AS oi " & _
        "WHERE OrderNo = '" & strOrderNo & "';")
        fieldPOrderNo = rsPOrderNo.Fields(strPOrderNo)
    End If
    
    ' Get POrderNo / CADOrderNo recordset
    queryPOrderNo = "SELECT * FROM ("
    queryPOrderNo = queryPOrderNo & "SELECT " & strSelectColumns & ", 1 AS FILTER FROM " & strOrderTable & vbCrLf
    queryPOrderNo = queryPOrderNo & "WHERE OrderNo = '" & strOrderNo & "'" & vbCrLf
    If Left(strOrderNo, 2) <> "AC" Then
        queryPOrderNo = queryPOrderNo & " UNION " & vbCrLf
        queryPOrderNo = queryPOrderNo & "SELECT " & strSelectColumns & ", 2 AS FILTER FROM " & strOrderTable & vbCrLf
        queryPOrderNo = queryPOrderNo & "WHERE " & strPOrderNo & " = '" & fieldPOrderNo & "' AND OrderNo <> '" & strOrderNo & "'" & vbCrLf
    End If
    queryPOrderNo = queryPOrderNo & ") t ORDER BY FILTER, OrderNo DESC" & vbCrLf
    getQueryPOrderNo = queryPOrderNo
    
End Function
