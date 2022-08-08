VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_OrderDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnDashboardClose_Click()
DoCmd.Close acForm, Me.name
End Sub

Private Sub Form_Load()
    '#DL 2020-10-06 Start (ganzer Sub)
    On Error GoTo Errormsg
    
    Call isTesting
    
    Dim strInfo1 As String, strInfo2 As String, strInfo3 As String, strInfo4 As String
    Dim i As Integer
    Dim headers1 As Variant
    Dim headers2 As Variant
    Dim OrderNo As String
    Dim rs As ADODB.RecordSet, RS2 As ADODB.RecordSet
    Dim ValHeader2 As Variant
    Dim args() As String
    Dim t_case As String, t_econfiBtn As String
    
    txtCopyInfo = "Bitte wählen sie eine Reihe aus!"
    
    'OrderNo = Screen.ActiveForm![Alltasks].Column(0)

    OrderNo = Forms!OrderDetail.OpenArgs
'    args = Split(Forms!OrderDetail.OpenArgs, "|")
'
'    orderNo = args(0)
'    t_case = args(1)
'    t_econfiBtn = args(2)
    
    If Len(OrderNo) > 0 Then
    

        'Set Headers
'        headers = Array( _
'        "OrderNo:", "GroupName:", "Client:", "Tool:", "Confirmation:", "PlannedDelivery:", "TypeConfi:", "DeliveryType:", "DateConfi:", "YearEnd:", "EngCode:" _
'        , "EngContact:", "Country:", "Adressen-Bearbeiter:", "Adressen-Status:", "Forensics-Status:", "Job_Status:", "1_CAD_Preparer:" _
'        , "2_CAD_Preparer:", "Adressen-Status:", "Forensics-Status:", "DeliveryType:", "Datum_Versand_EMail_Erstkontakt:", "Versand_durch:" _
'        , "First_Upload_Gewünscht:", "DatumAuftragsende:", "Informationen_zum_Auftrag:", "GroupName:", "PlannedDelivery:", "Versand_durch:" _
'        , "Eingangsdatum_CAD:", "Versanddatum_CAD:", "Anzahl_Anschr:", "Antwort_Deadline:", "Date2ndReq:", "Kommentar_Versand:" _
'        , "DatumAuftragsende:", "Final_Senior_Review_durch:", "Datum_final_Review:", "Anzahl_erhaltene_Salden:", "Versanddatum_Originale:" _
'        , "Empfänger_Originale:", "Empfänger_Nachläufer:", "EngPartner:", "EngManager:", "OtherContact:", "TypeFinancials:", "Lang:", "EYStand:" _
'        , "BelongingToCompanyGroup:", "BusinessEstablishmentOAreAvailable:", "BusinessEstablishmentOAre:", "GISID:", "EngName:", "OtherTypeConfi:" _
'        , "ConfiMailSent:", "TemplateSent:", "AuditiSent:", "Storno:", "ConfiFlag:" _
'        )

        'Set Values
        Dim strSQL As String
        
        '!!!!!!
        strSQL = buildSQLOrderDetail(OrderNo)
        
        Debug.Print strSQL
        Set rs = getRS(strSQL)
        ' Get Values for 2. Column

        Dim strTab As String: strTab = "   "


        For i = 0 To 28
            strInfo3 = strInfo3 & rs.Fields(i).name & ":" & strTab & rs.Fields(i).value & vbCrLf
        Next i
        
        For i = 29 To 55

            strInfo4 = strInfo4 & rs.Fields(i).name & ":" & strTab & rs.Fields(i) & vbCrLf
        Next i


        'Debug.Print strInfo3
        Me.Values1.SetFocus
        Me.Values1.Text = strInfo3
        Me.Values2.SetFocus
        Me.Values2.Text = strInfo4
        'Me.Header1.value = strInfo1
        'Me.Header2.value = strInfo2

        Exit Sub
    
    Else
    
        MsgBox ("Please select an order")
    
    End If
Errormsg:
    If CStr(Err.Number) = "3044" Then
        MsgBox ("No connection to the server. Please check the connection and try again")
    ElseIf CStr(Err.Number) = "94" Then
        MsgBox ("Please select an order")
        DoCmd.Close acForm, Me.name
    Else
        MsgBox ("An error is occured. Please restart the client. Error Info: " & Err.Description)
    End If

End Sub

Private Function buildSQLOrderDetail(OrderNo As String) As String

    Dim strSQL As String
    Dim b_buildQuery As Boolean
    b_buildQuery = False

    ' Improve SQL Query by selecting only necessary information
    If modus = 0 Then b_buildQuery = True
 
    ' Start Query
    str_SQL = "SELECT " & vbCrLf

    str_SQL = str_SQL & "OrderNo ," & vbCrLf
    str_SQL = str_SQL & "NameOfCompanyGroup as GroupName," & vbCrLf
    str_SQL = str_SQL & "Client ," & vbCrLf
    str_SQL = str_SQL & "Tool ," & vbCrLf
    str_SQL = str_SQL & "Confirmation ," & vbCrLf
    str_SQL = str_SQL & "PlannedDelivery ," & vbCrLf
    str_SQL = str_SQL & "TypeConfi ," & vbCrLf
    str_SQL = str_SQL & "DeliveryType ," & vbCrLf
    str_SQL = str_SQL & "DateConfi ," & vbCrLf
    str_SQL = str_SQL & "YearEnd ," & vbCrLf
    str_SQL = str_SQL & "EngCode ," & vbCrLf
    str_SQL = str_SQL & "EngContact ," & vbCrLf
    str_SQL = str_SQL & "Country ," & vbCrLf
    str_SQL = str_SQL & "Adressen_Bearbeiter as [Adressen-Bearbeiter], " & vbCrLf
    str_SQL = str_SQL & "Adressen_Status as [Adressen-Status], " & vbCrLf
    str_SQL = str_SQL & "Forensics_Status as [Forensics-Status], " & vbCrLf
    str_SQL = str_SQL & "[SB_Status] As Job_Status, " & vbCrLf
    str_SQL = str_SQL & "[CAD_Bearbeiter] As [1_CAD_Preparer], " & vbCrLf
    str_SQL = str_SQL & "[CAD_Bearbeiter_vorherig] As [2_CAD_Preparer], " & vbCrLf
    str_SQL = str_SQL & "Address_Validation as [Adressen-Status], " & vbCrLf
    str_SQL = str_SQL & "Forensics_Status as [Forensics-Status], " & vbCrLf
    str_SQL = str_SQL & "DeliveryType " & vbCrLf
    str_SQL = str_SQL & "[Datum_Versand_EMail_Erstkontakt] ," & vbCrLf
    str_SQL = str_SQL & "[Versand_durch] ," & vbCrLf
    str_SQL = str_SQL & "First_Upload_Gewünscht ," & vbCrLf
    str_SQL = str_SQL & "DatumAuftragsende, " & vbCrLf
    str_SQL = str_SQL & "[Informationen_zum_Auftrag], " & vbCrLf
    str_SQL = str_SQL & "NameOfCompanyGroup as GroupName ," & vbCrLf
    str_SQL = str_SQL & "[PlannedDelivery] ," & vbCrLf
    str_SQL = str_SQL & "Versand_durch ," & vbCrLf
    str_SQL = str_SQL & "[Eingangsdatum_CAD] ," & vbCrLf
    str_SQL = str_SQL & "[Versanddatum_CAD] ," & vbCrLf
    str_SQL = str_SQL & "[Anzahl_Bestätigungen] as Anzahl_Anschr ," & vbCrLf
    str_SQL = str_SQL & "[Antwort_Deadline], " & vbCrLf
    str_SQL = str_SQL & "[Date2ndReq], " & vbCrLf
    str_SQL = str_SQL & "[Kommentar_Versand] " & vbCrLf
    str_SQL = str_SQL & "DatumAuftragsende ," & vbCrLf
    str_SQL = str_SQL & "Final_Senior_Review_durch ," & vbCrLf
    str_SQL = str_SQL & "Datum_final_Review ," & vbCrLf
    str_SQL = str_SQL & "Anzahl_erhaltene_Salden ," & vbCrLf
    str_SQL = str_SQL & "Versanddatum_Originale ," & vbCrLf
    str_SQL = str_SQL & "[Empfänger_Originale] ," & vbCrLf
    str_SQL = str_SQL & "[Empfänger_Nachläufer] " & vbCrLf
    str_SQL = str_SQL & "EngPartner ," & vbCrLf
    str_SQL = str_SQL & "EngManager ," & vbCrLf
    str_SQL = str_SQL & "OtherContact ," & vbCrLf
    str_SQL = str_SQL & "TypeFinancials ," & vbCrLf
    str_SQL = str_SQL & "Lang ," & vbCrLf
    str_SQL = str_SQL & "EYStand ," & vbCrLf
    str_SQL = str_SQL & "BelongingToCompanyGroup " & vbCrLf
    str_SQL = str_SQL & "BusinessEstablishmentOAreAvailable ," & vbCrLf
    str_SQL = str_SQL & "BusinessEstablishmentOAre ," & vbCrLf
    str_SQL = str_SQL & "GISID ," & vbCrLf
    str_SQL = str_SQL & "EngName ," & vbCrLf
    str_SQL = str_SQL & "OtherTypeConfi ," & vbCrLf
    str_SQL = str_SQL & "ConfiMailSent ," & vbCrLf
    str_SQL = str_SQL & "TemplateSent ," & vbCrLf
    str_SQL = str_SQL & "AuditiSent ," & vbCrLf
    str_SQL = str_SQL & "Storno ," & vbCrLf
    str_SQL = str_SQL & "ConfiFlag " & vbCrLf

    str_SQL = str_SQL & "From tCon_Orderbook" & strTestTable & vbCrLf
    
    str_SQL = str_SQL & "WHERE OrderNo = '" & OrderNo & "'" & vbCrLf
    
    
    modus = Empty
    
    buildSQLOrderDetail = str_SQL


End Function



