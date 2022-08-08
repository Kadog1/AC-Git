Attribute VB_Name = "DEV"
Option Compare Database

Sub testLabelcontent()

    Dim oExcel            As Object
    Dim oExcelWrkBk       As Object
    Dim oExcelWrkSht      As Object
    Dim oExcelWrkShtFinal As Object



End Sub

Sub test()


    Dim sql As String, t_case As String
    
    sql = "SELECT OrderNo, LGISID, GAAP, PeriodEnd FROM Orderbook WHERE OrderNo = 'ADD0000008020' Or OrderNo = 'ADD0000006030';"
    
    Dim cnxn As ADODB.Connection
    Set cnxn = CurrentProject.AccessConnection
    Dim myRecordSet As New ADODB.RecordSet
    
    myRecordSet.Open sql, cnxn, adOpenStatic, adLockOptimistic
    
    Dim X() As Variant
    
    X = myRecordSet.GetRows
    
    Debug.Print X(1, 1)
    Debug.Print X(1, 2)
    Debug.Print X(2, 1)
    Debug.Print X(2, 2)
    
    
End Sub

Private Function buildSQL(t_case As String, t_econfiBtn As String, Optional modus As Integer, Optional searchText As String, Optional filterCol As String) As String

    Dim strSQL As String
    Dim b_buildQuery As Boolean
    b_buildQuery = False

    ' Improve SQL Query by selecting only necessary information
    If modus = 0 Then b_buildQuery = True
 
    ' Start Query
    str_SQL = "SELECT " & vbCrLf

    ' Select basic_info
    str_SQL = str_SQL & "OrderNo ," & vbCrLf
    str_SQL = str_SQL & "NameOfCompanyGroup as GroupName," & vbCrLf
    str_SQL = str_SQL & "Client ," & vbCrLf
    str_SQL = str_SQL & "Tool ," & vbCrLf
    str_SQL = str_SQL & "Confirmation ," & vbCrLf
    str_SQL = str_SQL & "TypeConfi ," & vbCrLf
    
    If t_case = "Start" Or t_case = "Versand" Or t_case = "Auftragsende" Then str_SQL = str_SQL & "DateConfi ," & vbCrLf
    
    str_SQL = str_SQL & "YearEnd ," & vbCrLf
    str_SQL = str_SQL & "EngCode ," & vbCrLf
    
    If t_case = "Start" Or t_case = "Versand" Or t_case = "Auftragsende" Then str_SQL = str_SQL & "EngContact ," & vbCrLf
    If t_case = "Start" Or t_case = "Versand" Or t_case = "Auftragsende" Then str_SQL = str_SQL & "DeliveryType ," & vbCrLf
    
    If t_case = "Start" Or t_case = "Versand" Then str_SQL = str_SQL & "[Datum_Versand_EMail_Erstkontakt] ," & vbCrLf
    If t_case = "Start" Or t_case = "Versand" Or t_case = "Auftragsende" Then str_SQL = str_SQL & "Informationen_zum_Auftrag ," & vbCrLf
    If t_case = "Start" Or t_case = "Versand" Then str_SQL = str_SQL & "Versand_durch ," & vbCrLf
    
    If t_case = "Start" Then
        str_SQL = str_SQL & "PlannedDelivery ," & vbCrLf
        str_SQL = str_SQL & "First_Upload_Gewünscht ," & vbCrLf
        str_SQL = str_SQL & "DatumAuftragsende ," & vbCrLf
    End If
    
    If t_case = "Start" Or t_case = "Versand" Or t_case = "Auftragsende" Then
        str_SQL = str_SQL & "[SB_Status] As Job_Status, " & vbCrLf
        str_SQL = str_SQL & "[CAD_Bearbeiter] As 1_CAD_Preparer, " & vbCrLf
    End If
    If t_case = "Start" Then str_SQL = str_SQL & "[CAD_Bearbeiter_vorherig] As 2_CAD_Preparer " & vbCrLf
    If t_case = "Versand" Or t_case = "Auftragsende" Then
        str_SQL = str_SQL & "[CAD_Bearbeiter_vorherig] As 2_CAD_Preparer ," & vbCrLf
    End If
    If t_case = "Versand" Then
    
        str_SQL = str_SQL & "[PlannedDelivery] ," & vbCrLf
        str_SQL = str_SQL & "[Eingangsdatum_CAD] ," & vbCrLf
        str_SQL = str_SQL & "[Versanddatum_CAD] ," & vbCrLf
        str_SQL = str_SQL & "[Anzahl_Bestätigungen] as Anzahl_Anschreiben," & vbCrLf
        str_SQL = str_SQL & "[Antwort_Deadline] ," & vbCrLf
        str_SQL = str_SQL & "[Kommentar_Versand] ," & vbCrLf
        str_SQL = str_SQL & "[2nd_Request_gewünscht] ," & vbCrLf
        str_SQL = str_SQL & "[Datum_2nd_Request] ," & vbCrLf
        str_SQL = str_SQL & "[Date2ndReq] " & vbCrLf

    End If
    
    If t_case = "Auftragsende" Then
    
        str_SQL = str_SQL & "DatumAuftragsende ," & vbCrLf
        str_SQL = str_SQL & "Anzahl_erhaltene_Salden ," & vbCrLf
        str_SQL = str_SQL & "Final_Senior_Review_durch ," & vbCrLf
        str_SQL = str_SQL & "Datum_final_Review ," & vbCrLf
        str_SQL = str_SQL & "Versanddatum_Originale ," & vbCrLf
        str_SQL = str_SQL & "[Empfänger_Originale] ," & vbCrLf
        str_SQL = str_SQL & "[Empfänger_Nachläufer] ," & vbCrLf
        str_SQL = str_SQL & "Datum_Feedbackanfrage ," & vbCrLf
        str_SQL = str_SQL & "Datum_Reminder ," & vbCrLf
        str_SQL = str_SQL & "Feedback_erhalten " & vbCrLf
    
    End If
    
    
    If t_case = "Rest" Then
    
        str_SQL = str_SQL & "EngPartner ," & vbCrLf
        str_SQL = str_SQL & "EngManager ," & vbCrLf
        str_SQL = str_SQL & "EngContact ," & vbCrLf
        str_SQL = str_SQL & "OtherContact ," & vbCrLf
        str_SQL = str_SQL & "ID ," & vbCrLf
        str_SQL = str_SQL & "POrderNo ," & vbCrLf
        str_SQL = str_SQL & "Storno ," & vbCrLf
        str_SQL = str_SQL & "GISID ," & vbCrLf
        str_SQL = str_SQL & "BelongingToCompanyGroup ," & vbCrLf
        str_SQL = str_SQL & "BusinessEstablishmentOAreAvailable ," & vbCrLf
        str_SQL = str_SQL & "BusinessEstablishmentOAre ," & vbCrLf
        str_SQL = str_SQL & "TypeFinancials ," & vbCrLf
        str_SQL = str_SQL & "EYStand ," & vbCrLf
        str_SQL = str_SQL & "EngName ," & vbCrLf
        str_SQL = str_SQL & "Lang ," & vbCrLf
        str_SQL = str_SQL & "OtherTypeConfi ," & vbCrLf
        str_SQL = str_SQL & "ConfiFlag ," & vbCrLf
        str_SQL = str_SQL & "ConfiMailSent ," & vbCrLf
        str_SQL = str_SQL & "TemplateSent ," & vbCrLf
        str_SQL = str_SQL & "AuditiSent " & vbCrLf
        
    End If
 
    ' Join table Orderbook x ConfiClientContact
    str_SQL = str_SQL & "FROM Orderbook " & vbCrLf
    
    If b_buildQuery = True Then
    
        If t_econfiBtn = "einblenden" Then
            str_SQL = str_SQL & "WHERE Tool = 'CAD' AND (Storno is Null) Order By OrderNo ASC;"
        Else
            str_SQL = str_SQL & "WHERE (Storno is Null) Order By OrderNo ASC;"
        End If
    
    End If
    
    Select Case modus
    
        Case 1

            If t_econfiBtn = "einblenden" Then
                str_SQL = str_SQL & "WHERE " & filterCol & " LIKE '%" & searchText & "%' AND Storno is null AND Tool = 'CAD'"
            Else
                str_SQL = str_SQL & "WHERE " & filterCol & " LIKE '%" & searchText & "%' AND Storno is null"
            End If
            
            
        Case 2
    
            If t_econfiBtn = "einblenden" Then
                str_SQL = str_SQL & " WHERE (OrderNo LIKE '%" & searchText & "%' OR GISID LIKE '%" & searchText & "%' OR Client LIKE '%" & searchText & "%' OR EngCode LIKE '%" & searchText & "%' OR EngName LIKE '%" & searchText & "%' OR EngPartner LIKE '%" & searchText & "%' OR EngManager LIKE '%" & searchText & "%' OR EngContact LIKE '%" & searchText & "%' OR OtherContact LIKE '%" & searchText & "%' OR BusinessEstablishmentOAre LIKE '%" & searchText & "%' OR NameOfCompanyGroup LIKE '%" & searchText & "%' OR TypeFinancials LIKE '%" & searchText & "%' OR Confirmation LIKE '%" & searchText & "%' OR TypeConfi LIKE '%" & searchText & "%'" _
                          & " OR Lang LIKE '%" & searchText & "%' OR PlannedDelivery LIKE '%" & searchText & "%' OR DeliveryType LIKE '%" & searchText & "%' OR OtherTypeConfi LIKE '%" & searchText & "%' OR Tool LIKE '%" & searchText & "%' OR CAD_Bearbeiter LIKE '%" & searchText & "%' OR CAD_Bearbeiter_vorherig LIKE '%" & searchText & "%' OR CAD_Reviewer LIKE '%" & searchText & "%' OR SB_Status LIKE '%" & searchText & "%' OR Anzahl_erhaltene_Salden LIKE '%" & searchText & "%' OR Informationen_Zum_Auftrag LIKE '%" & searchText & "%' OR Versand_durch LIKE '%" & searchText & "%' OR Kommentar_Versand LIKE '%" & searchText & "%' OR Final_Senior_Review_durch LIKE '%" & searchText & "%' OR Empfänger_Originale LIKE '%" & searchText & "%' OR Empfänger_Nachläufer LIKE '%" & searchText & "%') AND Storno is null AND Tool = 'CAD'"
            Else
                str_SQL = str_SQL & " WHERE (OrderNo LIKE '%" & searchText & "%' OR GISID LIKE '%" & searchText & "%' OR Client LIKE '%" & searchText & "%' OR EngCode LIKE '%" & searchText & "%' OR EngName LIKE '%" & searchText & "%' OR EngPartner LIKE '%" & searchText & "%' OR EngManager LIKE '%" & searchText & "%' OR EngContact LIKE '%" & searchText & "%' OR OtherContact LIKE '%" & searchText & "%' OR BusinessEstablishmentOAre LIKE '%" & searchText & "%' OR NameOfCompanyGroup LIKE '%" & searchText & "%' OR TypeFinancials LIKE '%" & searchText & "%' OR Confirmation LIKE '%" & searchText & "%' OR TypeConfi LIKE '%" & searchText & "%'" _
                          & " OR Lang LIKE '%" & searchText & "%' OR PlannedDelivery LIKE '%" & searchText & "%' OR DeliveryType LIKE '%" & searchText & "%' OR OtherTypeConfi LIKE '%" & searchText & "%' OR Tool LIKE '%" & searchText & "%' OR CAD_Bearbeiter LIKE '%" & searchText & "%' OR CAD_Bearbeiter_vorherig LIKE '%" & searchText & "%' OR CAD_Reviewer LIKE '%" & searchText & "%' OR SB_Status LIKE '%" & searchText & "%' OR Anzahl_erhaltene_Salden LIKE '%" & searchText & "%' OR Informationen_Zum_Auftrag LIKE '%" & searchText & "%' OR Versand_durch LIKE '%" & searchText & "%' OR Kommentar_Versand LIKE '%" & searchText & "%' OR Final_Senior_Review_durch LIKE '%" & searchText & "%' OR Empfänger_Originale LIKE '%" & searchText & "%' OR Empfänger_Nachläufer LIKE '%" & searchText & "%') AND Storno is null"
            End If
    
    End Select
    
    modus = Empty
    
    buildSQL = str_SQL


End Function

Private Sub adjustColumns(str_input As String)

    Dim t_width As Integer
    Dim width As Double, stepWidth As Double
    
    Dim columnWidths() As Double
    width = Alltasks.width
    
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
    If Me.Label76.Caption = "Start" Then columnWidths(10) = width - (42.5 * stepWidth)
    
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
    
    Dim sum As Double
    Dim str_colWidths As String
    For i = 1 To t_width
        sum = sum + columnWidths(i)
        str_colWidths = str_colWidths & CStr(columnWidths(i)) & ";"
    Next i

    str_colWidths = Left(str_colWidths, Len(str_colWidths) - 1)
    
    'Alltasks.ColumnWidth = sum
    Alltasks.columnWidths = str_colWidths
    
    
    Debug.Print Alltasks.ColumnCount


End Sub

