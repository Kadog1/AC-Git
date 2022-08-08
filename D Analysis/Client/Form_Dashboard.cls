VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Refresh_Click()

    On Error GoTo Errormsg
    
    'Storno  DeliveryType    Tool    SB_Status   Versand_durch

    Dim DashboardQuery As String
    DashboardQuery = "SELECT Storno, Tool, SB_Status, Confirmation, TypeConfi,DeliveryType, Versand_durch, EYStand, Anzahl_Bestätigungen FROM tCON_Orderbook" & strTestTable
    Me.ListJobStatus.RowSource = ""

    Dim i As Long, RSFiltered As ADODB.RecordSet, RecordCount As Long
    Dim rs As ADODB.RecordSet
    Dim arrAlltasks() As Variant, arrResult() As Variant, arrOrder() As Variant, arrPostal() As Variant, arrEmail() As Variant, arrCAD() As Variant, arrOther() As Variant, arrAnzahl() As Variant
    Dim arrAnschreiben() As Variant
    Dim arrEConfResult() As Variant, arrEConfOrder() As Variant, arrEConfPostal() As Variant, arrEConfEmail() As Variant, arrEConf() As Variant
    Set rs = getRS(DashboardQuery)
 
    arrAlltasks = rs.GetRows
    
    ' ##### CAD output #####
    ReDim arrResult(1, 0 To 7)
    ReDim arrOrder(1, 0 To 23)
    ReDim arrPostal(1, 0 To 11)
    ReDim arrEmail(1, 0 To 11)
    ReDim arrAnschreiben(1, 0 To 11)
    
    ReDim arrCAD(1, 0 To 11)
    ReDim arrOther(1, 0 To 11)
    
    ' ##### eConfi output #####
    ReDim arrEConfResult(1, 0 To 3)
    ReDim arrEConfOrder(1, 0 To 9)
    ReDim arrEConfPostal(1, 0 To 4)
    ReDim arrEConfEmail(1, 0 To 4)
    ReDim arrEConf(1, 0 To 3)
    
    '0 Storno,1 Tool,2 SB_Status,3 Confirmation,4 TypeConfi,5 DeliveryType,6 Versand_durch, 7 EYStand

    For i = 0 To UBound(arrAlltasks, 2)
    
        ' Overall overview information
        
        Select Case arrAlltasks(1, i)
        
            Case "CAD"
                If IsNull(arrAlltasks(0, i)) Then
                    ' SB_Status
                    If arrAlltasks(2, i) = "Not Started" Then
                        arrResult(1, 1) = arrResult(1, 1) + 1
                    ElseIf arrAlltasks(2, i) = "In Progress" Then
                        arrResult(1, 3) = arrResult(1, 3) + 1
                    ElseIf arrAlltasks(2, i) = "Closed" Then
                        arrResult(1, 5) = arrResult(1, 5) + 1
                    End If
                End If
                If IsNull(arrAlltasks(0, i)) = False And arrAlltasks(1, i) = "CAD" Then
                    If arrAlltasks(0, i) <> "" Then arrResult(1, 7) = arrResult(1, 7) + 1
                End If
                
                ' Order
                If IsNull(arrAlltasks(0, i)) And arrAlltasks(2, i) <> "Storno" Then
        
                    If arrAlltasks(3, i) = "Banken" Then
                        arrOrder(1, 1) = arrOrder(1, 1) + 1
                    ElseIf arrAlltasks(3, i) = "Rechtsanwalt" Then
                        arrOrder(1, 3) = arrOrder(1, 3) + 1
                    ElseIf arrAlltasks(3, i) = "Steuerberater" Then
                        arrOrder(1, 5) = arrOrder(1, 5) + 1
                    ElseIf arrAlltasks(3, i) = "BankenPortal" Then
                        arrOrder(1, 7) = arrOrder(1, 7) + 1
                    ElseIf arrAlltasks(3, i) = "Sonstige" Then
                        arrOrder(1, 9) = arrOrder(1, 9) + 1
                    ElseIf arrAlltasks(3, i) = "VorräteDritte" Then
                        arrOrder(1, 11) = arrOrder(1, 11) + 1
                    ElseIf arrAlltasks(3, i) = "VorräteFremdlagerung" Then
                        arrOrder(1, 13) = arrOrder(1, 13) + 1
                    ElseIf arrAlltasks(3, i) = "Kreditoren" Then
                        arrOrder(1, 15) = arrOrder(1, 15) + 1
                        
                    ElseIf arrAlltasks(3, i) = "Debitoren" Then
        
                        arrOrder(1, 17) = arrOrder(1, 17) + 1
            
                        If arrAlltasks(4, i) = "Einzelposten" Then
                            arrOrder(1, 19) = arrOrder(1, 19) + 1
                        ElseIf arrAlltasks(4, i) = "Salden" Then
                            arrOrder(1, 21) = arrOrder(1, 21) + 1
                        End If
                    End If
        
                End If
        
                ' POSTAL
                If IsNull(arrAlltasks(0, i)) And arrAlltasks(2, i) <> "Storno" Then
        
                    If (arrAlltasks(3, i) = "Banken") And InStr(arrAlltasks(5, i), "post") > 0 Then
                        arrPostal(1, 0) = arrPostal(1, 0) + 1
                    ElseIf arrAlltasks(3, i) = "Rechtsanwalt" And InStr(arrAlltasks(5, i), "post") > 0 Then
                        arrPostal(1, 1) = arrPostal(1, 1) + 1
                    ElseIf arrAlltasks(3, i) = "Steuerberater" And InStr(arrAlltasks(5, i), "post") > 0 Then
                        arrPostal(1, 2) = arrPostal(1, 2) + 1
                    ElseIf arrAlltasks(3, i) = "BankenPortal" And InStr(arrAlltasks(5, i), "post") > 0 Then
                        arrPostal(1, 3) = arrPostal(1, 3) + 1
                    ElseIf arrAlltasks(3, i) = "Sonstige" And InStr(arrAlltasks(5, i), "post") > 0 Then
                        arrPostal(1, 4) = arrPostal(1, 4) + 1
                    ElseIf arrAlltasks(3, i) = "VorräteDritte" And InStr(arrAlltasks(5, i), "post") > 0 Then
                        arrPostal(1, 5) = arrPostal(1, 5) + 1
                    ElseIf arrAlltasks(3, i) = "VorräteFremdlagerung" And InStr(arrAlltasks(5, i), "post") > 0 Then
                        arrPostal(1, 6) = arrPostal(1, 6) + 1
                        

                    ElseIf arrAlltasks(3, i) = "Kreditoren" And InStr(arrAlltasks(5, i), "post") > 0 Then
                        arrPostal(1, 7) = arrPostal(1, 7) + 1
                    ElseIf arrAlltasks(3, i) = "Debitoren" And InStr(arrAlltasks(5, i), "post") > 0 Then
        
                        arrPostal(1, 8) = arrPostal(1, 8) + 1
            
                        If arrAlltasks(4, i) = "Einzelposten" Then
                            arrPostal(1, 9) = arrPostal(1, 9) + 1
                        ElseIf arrAlltasks(4, i) = "Salden" Then
                            arrPostal(1, 10) = arrPostal(1, 10) + 1
                        End If
                    End If
        
                End If
        
                ' EMAIL
                If IsNull(arrAlltasks(0, i)) And arrAlltasks(2, i) <> "Storno" Then
        
                    If (arrAlltasks(3, i) = "Banken") And InStr(arrAlltasks(5, i), "mail") > 0 Then
                        arrEmail(1, 0) = arrEmail(1, 0) + 1
                    ElseIf arrAlltasks(3, i) = "Rechtsanwalt" And InStr(arrAlltasks(5, i), "mail") > 0 Then
                        arrEmail(1, 1) = arrEmail(1, 1) + 1
                    ElseIf arrAlltasks(3, i) = "Steuerberater" And InStr(arrAlltasks(5, i), "mail") > 0 Then
                        arrEmail(1, 2) = arrEmail(1, 2) + 1
                    ElseIf arrAlltasks(3, i) = "BankenPortal" And InStr(arrAlltasks(5, i), "mail") > 0 Then
                        arrEmail(1, 3) = arrEmail(1, 3) + 1
                    ElseIf arrAlltasks(3, i) = "Sonstige" And InStr(arrAlltasks(5, i), "mail") > 0 Then
                        arrEmail(1, 4) = arrEmail(1, 4) + 1
                    ElseIf arrAlltasks(3, i) = "VorräteDritte" And InStr(arrAlltasks(5, i), "mail") > 0 Then
                        arrEmail(1, 5) = arrEmail(1, 5) + 1
                    ElseIf arrAlltasks(3, i) = "VorräteFremdlagerung" And InStr(arrAlltasks(5, i), "mail") > 0 Then
                        arrEmail(1, 6) = arrEmail(1, 6) + 1

                    ElseIf arrAlltasks(3, i) = "Kreditoren" And InStr(arrAlltasks(5, i), "mail") > 0 Then
                        arrEmail(1, 7) = arrEmail(1, 7) + 1
                    ElseIf arrAlltasks(3, i) = "Debitoren" And InStr(arrAlltasks(5, i), "mail") > 0 Then
        
                        arrEmail(1, 8) = arrEmail(1, 8) + 1
            
                        If arrAlltasks(4, i) = "Einzelposten" Then
                            arrEmail(1, 9) = arrEmail(1, 9) + 1
                        ElseIf arrAlltasks(4, i) = "Salden" Then
                            arrEmail(1, 10) = arrEmail(1, 10) + 1
                        End If
                    End If
                End If
        
                ' CAD + Eng Team
        
                If IsNull(arrAlltasks(0, i)) And arrAlltasks(2, i) <> "Storno" Then

                    If arrAlltasks(6, i) = "CAD" Then
                        If (arrAlltasks(3, i) = "Banken") Then
                            arrCAD(1, 0) = arrCAD(1, 0) + 1
                        ElseIf arrAlltasks(3, i) = "Rechtsanwalt" Then
                            arrCAD(1, 1) = arrCAD(1, 1) + 1
                        ElseIf arrAlltasks(3, i) = "Steuerberater" Then
                            arrCAD(1, 2) = arrCAD(1, 2) + 1
                        ElseIf arrAlltasks(3, i) = "BankenPortal" Then
                            arrCAD(1, 3) = arrCAD(1, 3) + 1
                        ElseIf arrAlltasks(3, i) = "Sonstige" Then
                            arrCAD(1, 4) = arrCAD(1, 4) + 1
                        ElseIf arrAlltasks(3, i) = "VorräteDritte" Then
                            arrCAD(1, 5) = arrCAD(1, 5) + 1
                        ElseIf arrAlltasks(3, i) = "VorräteFremdlagerung" Then
                            arrCAD(1, 6) = arrCAD(1, 6) + 1
 
                        ElseIf arrAlltasks(3, i) = "Kreditoren" Then
                            arrCAD(1, 7) = arrCAD(1, 7) + 1
                        ElseIf arrAlltasks(3, i) = "Debitoren" Then

                            arrCAD(1, 8) = arrCAD(1, 8) + 1

                            If arrAlltasks(4, i) = "Einzelposten" Then
                                arrCAD(1, 9) = arrCAD(1, 9) + 1
                            ElseIf arrAlltasks(4, i) = "Salden" Then
                                arrCAD(1, 10) = arrCAD(1, 10) + 1
                            End If
                        End If
                    End If

                    If InStr(arrAlltasks(6, i), "Team") > 0 And arrAlltasks(2, i) <> "Storno" Then
                        If (arrAlltasks(3, i) = "Banken" Or arrAlltasks(3, i) = "BankenPortal") Then
                            arrOther(1, 0) = arrOther(1, 0) + 1
                        ElseIf arrAlltasks(3, i) = "Rechtsanwalt" Then
                            arrOther(1, 1) = arrOther(1, 1) + 1
                        ElseIf arrAlltasks(3, i) = "Steuerberater" Then
                            arrOther(1, 2) = arrOther(1, 2) + 1
                        ElseIf arrAlltasks(3, i) = "BankenPortal" Then
                            arrOther(1, 3) = arrOther(1, 3) + 1
                        ElseIf arrAlltasks(3, i) = "Sonstige" Then
                            arrOther(1, 4) = arrOther(1, 4) + 1
                        ElseIf arrAlltasks(3, i) = "VorräteDritte" Then
                            arrOther(1, 5) = arrOther(1, 5) + 1
                        ElseIf arrAlltasks(3, i) = "VorräteFremdlagerung" Then
                            arrOther(1, 6) = arrOther(1, 6) + 1

                        ElseIf arrAlltasks(3, i) = "Kreditoren" Then
                            arrOther(1, 7) = arrOther(1, 7) + 1
                        ElseIf arrAlltasks(3, i) = "Debitoren" Then

                            arrOther(1, 8) = arrOther(1, 8) + 1

                            If arrAlltasks(4, i) = "Einzelposten" Then
                                arrOther(1, 9) = arrOther(1, 9) + 1
                            ElseIf arrAlltasks(4, i) = "Salden" Then
                                arrOther(1, 10) = arrOther(1, 10) + 1
                            End If
                        End If
                    End If
                End If
                
                ' AnzahlAnschreiben
                If IsNull(arrAlltasks(0, i)) And arrAlltasks(2, i) <> "Storno" And Not IsNull(arrAlltasks(8, i)) Then
        
                    If (arrAlltasks(3, i) = "Banken") Then
                        arrAnschreiben(1, 0) = arrAnschreiben(1, 0) + arrAlltasks(8, i)
                    ElseIf arrAlltasks(3, i) = "Rechtsanwalt" Then
                        arrAnschreiben(1, 1) = arrAnschreiben(1, 1) + arrAlltasks(8, i)
                    ElseIf arrAlltasks(3, i) = "Steuerberater" Then
                        arrAnschreiben(1, 2) = arrAnschreiben(1, 2) + arrAlltasks(8, i)
                    ElseIf arrAlltasks(3, i) = "BankenPortal" Then
                        arrAnschreiben(1, 3) = arrAnschreiben(1, 3) + arrAlltasks(8, i)
                    ElseIf arrAlltasks(3, i) = "Sonstige" Then
                        arrAnschreiben(1, 4) = arrAnschreiben(1, 4) + arrAlltasks(8, i)
                    ElseIf arrAlltasks(3, i) = "VorräteDritte" Then
                        arrAnschreiben(1, 5) = arrAnschreiben(1, 5) + arrAlltasks(8, i)
                    ElseIf arrAlltasks(3, i) = "VorräteFremdlagerung" Then
                        arrAnschreiben(1, 6) = arrAnschreiben(1, 6) + arrAlltasks(8, i)

                    ElseIf arrAlltasks(3, i) = "Kreditoren" Then
                        arrAnschreiben(1, 7) = arrAnschreiben(1, 7) + arrAlltasks(8, i)
                    ElseIf arrAlltasks(3, i) = "Debitoren" Then
        
                        arrAnschreiben(1, 8) = arrAnschreiben(1, 8) + arrAlltasks(8, i)
            
                        If arrAlltasks(4, i) = "Einzelposten" Then
                            arrAnschreiben(1, 9) = arrAnschreiben(1, 9) + arrAlltasks(8, i)
                        ElseIf arrAlltasks(4, i) = "Salden" Then
                            arrAnschreiben(1, 10) = arrAnschreiben(1, 10) + arrAlltasks(8, i)
                        End If
                    End If
                End If
                
                
        
            Case "eConfirmations"

                ' Listbox Aufträge + Storno
                If IsNull(arrAlltasks(0, i)) Then
                    arrEConfResult(1, 1) = arrEConfResult(1, 1) + 1
                End If
                If IsNull(arrAlltasks(0, i)) = False Then
                    If arrAlltasks(0, i) <> "" Then arrEConfResult(1, 3) = arrEConfResult(1, 3) + 1
                End If

                ' Listbox Kreditoren | Debitoren + Einzelposten + Salden TOTAL
                If IsNull(arrAlltasks(0, i)) Then

                    If arrAlltasks(3, i) = "Kreditoren" Then
                        arrEConfOrder(1, 1) = arrEConfOrder(1, 1) + 1

                    ElseIf arrAlltasks(3, i) = "Debitoren" Then

                        arrEConfOrder(1, 3) = arrEConfOrder(1, 3) + 1

                        If arrAlltasks(4, i) = "Einzelposten" Then
                            arrEConfOrder(1, 5) = arrEConfOrder(1, 5) + 1
                        ElseIf arrAlltasks(4, i) = "Salden" Then
                            arrEConfOrder(1, 7) = arrEConfOrder(1, 7) + 1
                        End If
                    End If

                End If

                ' Listbox Post + Mail
                If IsNull(arrAlltasks(0, i)) Then

                    If arrAlltasks(3, i) = "Kreditoren" And InStr(arrAlltasks(5, i), "post") > 0 Then
                        arrEConfPostal(1, 0) = arrEConfPostal(1, 0) + 1
                    ElseIf arrAlltasks(3, i) = "Kreditoren" And InStr(arrAlltasks(5, i), "mail") > 0 Then
                        arrEConfEmail(1, 0) = arrEConfEmail(1, 0) + 1

                    ElseIf arrAlltasks(3, i) = "Debitoren" And InStr(arrAlltasks(5, i), "post") > 0 Then

                        arrEConfPostal(1, 1) = arrEConfPostal(1, 1) + 1

                        If arrAlltasks(4, i) = "Einzelposten" Then
                            arrEConfPostal(1, 2) = arrEConfPostal(1, 2) + 1
                        ElseIf arrAlltasks(4, i) = "Salden" Then
                            arrEConfPostal(1, 3) = arrEConfPostal(1, 3) + 1
                        End If

                    ElseIf arrAlltasks(3, i) = "Debitoren" And InStr(arrAlltasks(5, i), "mail") > 0 Then

                        arrEConfEmail(1, 1) = arrEConfEmail(1, 1) + 1

                        If arrAlltasks(4, i) = "Einzelposten" Then
                            arrEConfEmail(1, 2) = arrEConfEmail(1, 2) + 1
                        ElseIf arrAlltasks(4, i) = "Salden" Then
                            arrEConfEmail(1, 3) = arrEConfEmail(1, 3) + 1
                        End If
                    End If
                End If
                
        End Select

    Next i
    
    ' ###### CAD #######
    ' SB_Status
    arrResult(1, 0) = "Not Started"
    arrResult(1, 2) = "In Progress"
    arrResult(1, 4) = "Closed"
    arrResult(1, 6) = "Storno"
    
    ' Order
    arrOrder(1, 0) = "Banken"
    arrOrder(1, 2) = "RA"
    arrOrder(1, 4) = "StB"
    arrOrder(1, 6) = "BankenPortal"
    arrOrder(1, 8) = "Sonstige"
    arrOrder(1, 10) = "VorräteDritte"
    arrOrder(1, 12) = "VorräteFremdlagerung"
    arrOrder(1, 14) = "Kreditoren"
    arrOrder(1, 16) = "Debitoren"
    
    arrOrder(1, 18) = "   Einzelposten"
    arrOrder(1, 20) = "   Salden"
    arrOrder(1, 22) = "GESAMT"
    arrOrder(1, 23) = arrOrder(1, 1) + arrOrder(1, 3) + arrOrder(1, 5) + arrOrder(1, 7) + arrOrder(1, 9) + arrOrder(1, 11) + arrOrder(1, 13) + arrOrder(1, 15) + arrOrder(1, 17)
    
    ' Postal
    arrPostal(1, 11) = arrPostal(1, 0) + arrPostal(1, 1) + arrPostal(1, 2) + arrPostal(1, 3) + arrPostal(1, 4) + arrPostal(1, 5) + arrPostal(1, 6) + arrPostal(1, 7) + arrPostal(1, 8)
    
    ' Email
    arrEmail(1, 11) = arrEmail(1, 0) + arrEmail(1, 1) + arrEmail(1, 2) + arrEmail(1, 3) + arrEmail(1, 4) + arrEmail(1, 5) + arrEmail(1, 6) + arrEmail(1, 7) + arrEmail(1, 8)
    
    ' Total - CAD
    arrCAD(1, 11) = arrCAD(1, 0) + arrCAD(1, 1) + arrCAD(1, 2) + arrCAD(1, 3) + arrCAD(1, 4) + arrCAD(1, 5) + arrCAD(1, 6) + arrCAD(1, 7) + arrCAD(1, 8)
    
    ' Total - Other
    arrOther(1, 11) = arrOther(1, 0) + arrOther(1, 1) + arrOther(1, 2) + arrOther(1, 3) + arrOther(1, 4) + arrOther(1, 5) + arrOther(1, 6) + arrOther(1, 7) + arrOther(1, 8)
    
    ' Total - Anschreiben
    arrAnschreiben(1, 11) = arrAnschreiben(1, 0) + arrAnschreiben(1, 1) + arrAnschreiben(1, 2) + arrAnschreiben(1, 3) + arrAnschreiben(1, 4) + arrAnschreiben(1, 5) + arrAnschreiben(1, 6) + arrAnschreiben(1, 7) + arrAnschreiben(1, 8)
    
    ' Output into ListBox
    Dim str_tempList As String
    
    str_tempList = ""
    str_tempList = createString(7, arrResult)
    Me.ListJobStatus.RowSource = str_tempList
    Me.ListJobStatus.columnWidths = "2.3cm;1.0cm"
    
    str_tempList = ""
    str_tempList = createString(23, arrOrder)
    Me.ListOrderStatus.RowSource = str_tempList
    Me.ListOrderStatus.columnWidths = "4.0cm;1.0cm"
    
    str_tempList = ""
    str_tempList = createString(11, arrPostal)
    Me.ListPostal.RowSource = str_tempList
    Me.ListPostal.columnWidths = "1.0cm"
    
    str_tempList = ""
    str_tempList = createString(11, arrEmail)
    Me.ListEmail.RowSource = str_tempList
    Me.ListEmail.columnWidths = "1.0cm"
    
    str_tempList = ""
    strListCAD = createString(11, arrCAD)
    Me.ListCAD.RowSource = strListCAD
    Me.ListCAD.columnWidths = "1.0cm"
    
    str_tempList = ""
    str_tempList = createString(11, arrOther)
    Me.ListOther.RowSource = str_tempList
    Me.ListOther.columnWidths = "1.0cm"
    
    str_tempList = ""
    str_tempList = createString(11, arrAnschreiben)
    Me.List7.RowSource = str_tempList
    Me.List7.columnWidths = "1.0cm"
   
    Erase arrResult, arrOrder, arrPostal, arrEmail, arrCAD, arrOther, arrAnschreiben
    
    ' ###### eConfirmations #######
    ' SB_Status
    arrEConfResult(1, 0) = "Anzahl Aufträge"
    arrEConfResult(1, 2) = "Storno"

    ' Order
    arrEConfOrder(1, 0) = "Kreditoren"
    arrEConfOrder(1, 2) = "Debitoren"
    arrEConfOrder(1, 4) = "   Einzelposten"
    arrEConfOrder(1, 6) = "   Salden"
    arrEConfOrder(1, 8) = "GESAMT"
    arrEConfOrder(1, 9) = arrEConfOrder(1, 1) + arrEConfOrder(1, 3)
    
    ' Postal
    arrEConfPostal(1, 4) = arrEConfPostal(1, 1) + arrEConfPostal(1, 3)
    
    ' Email
    arrEConfEmail(1, 4) = arrEConfEmail(1, 1) + arrEConfEmail(1, 3)
    
    Dim t_query As String
    t_query = "Select tCON_Orderbook" & strTestTable & ".OrderNo From tCon_Orderbook" & strTestTable & " Where Exists (Select 1 From tCON_ClientContact" & strTestTable & " WHERE tCON_Orderbook" & strTestTable & ".OrderNo = tCON_ClientContact" & strTestTable & ".OrderNo AND Storno is Null AND Tool = 'eConfirmations')"
    Dim rsCount As ADODB.RecordSet
    Set rsCount = getRS(t_query)

    arrEConf(1, 1) = rsCount.RecordCount
    
    Dim t_total As Integer, t_count As Integer, t_econfi As Integer
    
    t_econfi = arrEConfResult(1, 1)
    t_count = rsCount.RecordCount
                        
    arrEConf(1, 3) = t_econfi - t_count

    ' MandantenKontakt
    arrEConf(1, 0) = "Mandantenkontakt"
    arrEConf(1, 2) = "Ohne Mandantenkontakt"
    
    str_tempList = ""
    str_tempList = createString(3, arrEConfResult)
    Me.List8.RowSource = str_tempList
    Me.List8.columnWidths = "4.0cm;1.0cm"
    
    str_tempList = ""
    str_tempList = createString(9, arrEConfOrder)
    Me.List9.RowSource = str_tempList
    Me.List9.columnWidths = "2.3cm;1.0cm"
    
    str_tempList = ""
    str_tempList = createString(4, arrEConfPostal)
    Me.List10.RowSource = str_tempList
    Me.List10.columnWidths = "1.0cm"
    
    str_tempList = ""
    str_tempList = createString(4, arrEConfEmail)
    Me.List11.RowSource = str_tempList
    Me.List11.columnWidths = "1.0cm"
    
    str_tempList = ""
    str_tempList = createString(3, arrEConf)
    Me.List12.RowSource = str_tempList
    Me.List12.columnWidths = "5.0cm;2.0cm"
    
    Erase arrEConfResult, arrEConfOrder, arrEConfPostal, arrEConfEmail, arrEConf
    
    rs.Close
    rsCount.Close
    Set rs = Nothing
    Set rsCount = Nothing
    
    On Error GoTo 0
    Exit Sub

Errormsg:
    If CStr(Err.Number) = "3044" Then
        MsgBox ("No connection to the server. Please check the connection and try again")
    Else
        MsgBox ("An error is occured. Please restart the client. Error info: " & Err.Description)
    End If

End Sub

Private Function createString(t_width As Integer, arr_input() As Variant) As String

    Dim i As Integer

    For i = 0 To t_width
        If arr_input(1, i) = "" Then arr_input(1, i) = 0
        
        t_output = t_output & Format(arr_input(1, i), "#,##0") & ";"
    Next i
    
    createString = t_output

End Function

Private Sub Form_Load()
    Call isTesting
    Call Refresh_Click
End Sub

