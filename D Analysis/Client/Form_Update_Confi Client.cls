VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Update_Confi Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AV_Click()


    'Beim Klicken: erstellt die Folder & öffnet das mailtemaplte

    Dim path_2 As String
    Dim path_3 As String
    Dim tempname As String
    Dim OrderNo As String
    Dim period As String

    'On Error GoTo Errormsg

    OrderNo = Me.Alltasks.Column(0)
    period = convertDateToString(Me.Alltasks.Column(11))





    If Len(Dir("\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & OrderNo, vbDirectory)) = 0 Then
        MkDir "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & OrderNo
    End If



    path_2 = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & OrderNo & "\2. CAD_Abgleich\"
    If Len(Dir(path_2, vbDirectory)) = 0 Then
        MkDir path_2
    End If


    path_3 = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & OrderNo & "\3. Team Approval\"
    If Len(Dir(path_3, vbDirectory)) = 0 Then
        MkDir path_3
    End If


    'copy template
    tempname = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\2_CAD-Adressenabgleich Workbook_Template.xlsm"
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile tempname, path_2
    'rename
    Dim name_wb As String
    Dim name_ta As String
    name_wb = path_2 & period & " CAD Confirmations Workbook " & OrderNo & ".xlsm"

    On Error Resume Next
    Name path_2 & "2_CAD-Adressenabgleich Workbook_Template.xlsm" As name_wb


    tempname = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\3_CAD-Adressenabgleich Team Approval_Template.xlsm"
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile tempname, path_3

    name_ta = path_3 & period & " CAD Confirmations Team Approval Nachtrag" & OrderNo & ".xlsm"
    'rename

    On Error Resume Next
    Name path_3 & "3_CAD-Adressenabgleich Team Approval_Template.xlsm" As name_ta


    'open & load Basic Info
    Dim XL As Excel.Application
    Dim wb As Excel.Workbook



    Set XL = CreateObject("Excel.Application")
    Set wb = XL.Workbooks.Open(name_wb)
    XL.Visible = True

    wb.Worksheets("Basic Info").Range("B1").value = OrderNo
    wb.Worksheets("Basic Info").Range("B2").value = Me.Alltasks.Column(11)
    wb.Worksheets("Basic Info").Range("B3").value = Me.Alltasks.Column(2)
    wb.Worksheets("Basic Info").Range("B4").value = Me.Alltasks.Column(13)

    'query für EngPrnt & EngMngr
    Dim sql As String
    sql = "Select EngPartner, EngManager, OtherContact, CONVERT(varchar(50),FORMAT(GISID, '0000000000')) as GISID, lang  from tCON_Orderbook_TEST Where OrderNo='" & OrderNo & "'"
    Dim rs As ADODB.RecordSet
    Set rs = getRS(sql)
    
    wb.Worksheets("Basic Info").Range("B5").value = rs("EngPartner")
    wb.Worksheets("Basic Info").Range("B6").value = rs("EngManager")
    wb.Worksheets("Basic Info").Range("B7").value = rs("OtherContact")

    'gisid wird später für den Mailversand gebraucht
    wb.Worksheets("Basic Info").Range("B8").value = rs("GISID")
    ' Lang für die Auswahl des Mailtemps
    wb.Worksheets("Basic Info").Range("B9").value = rs("lang")



    rs.Close


    Set XL = Nothing

    'xl.Quit
    'Set xl = Nothing
    'Set wb = Nothing


Errormsg:
    If CStr(Err.Number) = "3044" Then
        MsgBox ("No connection to the server. Please check the connection and try again")
    ElseIf CStr(Err.Number) = "94" Then
        MsgBox ("Please select an order")
    End If


End Sub

Private Sub AC_View_Click()
DoCmd.OpenForm "AC_Confi Client"
DoCmd.Close acForm, "Update_Confi Client", acSaveYes
End Sub

Private Sub Form_Load()

    On Error GoTo Errormsg
    
    Call isTesting
    
    Application.SetOption "Auto Compact", True
    Application.SetOption "Default record locking", 2
    Application.SetOption "Refresh Interval (sec)", 9000
    Application.SetOption "ODBC Refresh Interval (sec)", 9000
    Application.SetOption "Number of Update Retries", 2
    Application.SetOption "Update Retry Interval (msec)", 250
    Application.SetOption "Enable DDE Refresh", False
    Application.SetOption "Ignore DDE Requests", True
    
    Debug.Print "Auto Compact " & Application.GetOption("Auto Compact")
    Debug.Print "Default record locking " & Application.GetOption("Default record locking")
    Debug.Print "Refresh Interval (sec) " & Application.GetOption("Refresh Interval (sec)")
    Debug.Print "ODBC Refresh Interval (sec) " & Application.GetOption("ODBC Refresh Interval (sec)")
    Debug.Print "Number of Update Retries " & Application.GetOption("Number of Update Retries")
    Debug.Print "Update Retry Interval (msec) " & Application.GetOption("Update Retry Interval (msec)")
    Debug.Print "Enable DDE Refresh " & Application.GetOption("Enable DDE Refresh")
    Debug.Print "Ignore DDE Requests " & Application.GetOption("Ignore DDE Requests")

    Call hideNavPRib
    
    Me.Label76.Caption = "Start"
    Me.Next_Page.Visible = True
    Me.Previous_Page.Visible = True

    If Me!btneConfi.Caption = "eConfirmations einblenden" Then
        ' Start
        If Me.Label76.Caption = "Start" Then
            Me.Label87.Caption = "Start1"
            AlltasksQuery = buildSQLStart(Me.Label87.Caption, "einblenden")
        End If
    Else
        If Me.Label76.Caption = "Start" Then
            AlltasksQuery = buildSQLStart(Me.Label87.Caption, "ausblenden")
        End If
    End If

    Set Me.Alltasks.RecordSet = getRS(AlltasksQuery)

    Call adjustStart(CStr(Me.Label87.Caption))
    
    Me.ShowStorno.Caption = "Storno/Closed ausblenden"

    'Call Refresh_Click
    
    On Error GoTo 0
    Exit Sub

Errormsg:
    If CStr(Err.Number) = "3044" Then
        MsgBox ("No connection to the server. Please check the connection and try again")
    Else
        MsgBox ("An error is occured. Please restart the client. Error info: " & Err.Description)
    End If

End Sub

'#JK-2/2020-12-16 START
Private Sub btn_OrderDetail_Click()

    On Error GoTo Errormsg:

'    Dim args As String
'    args = Me.Alltasks.Column(0) & "|" & Me.Label87.Caption
'
'    If Me!btneConfi.Caption = "eConfirmations einblenden" Then
'        args = args & "|einblenden"
'
'    Else
'        args = args & "|ausblenden"
'
'    End If

    DoCmd.OpenForm "OrderDetail", OpenArgs:=Me.Alltasks.Column(0)
    Forms![OrderDetail].SetFocus
3         Exit Sub
Errormsg:
4         MsgBox ("Please select an OrderNo in the Mainlist and try again.")


End Sub
'#JK-2/2020-12-16 END


' ######### BUTTON START ###########

Private Sub btn_ImportOrder_Click()
    
    'Passwort Abfrage
    Dim Passwort As String
    Passwort = InputBox("Bitte geben Sie das Passwort ein", "StarTrek")
    
    If Passwort = "StarTrek" Then
    
        Dim oExcel As Object
    
        Set oExcel = CreateObject("Excel.Application")
        Set oWorkbook = oExcel.Workbooks.Open("C:\eConfirmation\eConfirmations_import.xlsm")
    
        oExcel.Visible = True
        Set oExcel = Nothing
        
    End If
End Sub

Private Sub btn_Start_Click()
    Me.Label76.Caption = "Start"
    Me.Label87.Caption = "Start1"
    Me.Next_Page.Visible = True
    Me.Previous_Page.Visible = True
    Call Refresh_Click
End Sub

'#JK-3/2020-12-16 START
Private Sub btn_Adressen_Click()
    Me.Label76.Caption = "Adressen"
    Me.Label87.Caption = "Adressen1"
    Me.Next_Page.Visible = True
    Me.Previous_Page.Visible = True
    Call Refresh_Click
End Sub
'#JK-3/2020-12-16 END

Private Sub btn_Versand_Click()
    Me.Label76.Caption = "Versand"
    Me.Label87.Caption = "Versand1"
    'Me.Label87.Caption = "Start2"
    
    Me.Next_Page.Visible = False
    Me.Previous_Page.Visible = False
    Call Refresh_Click
End Sub
Private Sub btn_Auftrag_Click()
    Me.Label76.Caption = "Auftragsende"
    Me.Label87.Caption = "Auftragsende1"
        
    Me.Next_Page.Visible = False
    Me.Previous_Page.Visible = False
    Call Refresh_Click
End Sub

Private Sub btn_rest_Click()
    Me.Label76.Caption = "Rest"
    Me.Label87.Caption = "Rest1"
    Me.Next_Page.Visible = True
    Me.Previous_Page.Visible = True
    Call Refresh_Click
End Sub

Private Sub Next_Page_Click()

    Select Case Label76.Caption

        Case "Start"

            Me.Label87.Caption = "Start2"
            If Me.Label87.Caption = "Start2" Then
                Call Refresh_Click
            Else
                Exit Sub
            End If

'        Case "Versand"
'
'            Me.Label87.Caption = "Versand2"
'            If Me.Label87.Caption = "Versand2" Then
'                Call Refresh_Click
'            Else
'                Exit Sub
'            End If
'
'        Case "Auftragsende"
'            Me.Label87.Caption = "Auftragsende2"
'            If Me.Label87.Caption = "Auftragsende2" Then
'                Call Refresh_Click
'            Else
'                Exit Sub
'            End If
            
        Case "Rest"
            Me.Label87.Caption = "Rest2"
            If Me.Label87.Caption = "Rest2" Then
                Call Refresh_Click
            Else
                Exit Sub
            End If

    End Select




End Sub

Private Sub Previous_Page_Click()

    Select Case Label76.Caption

        Case "Start"
            Me.Label87.Caption = "Start1"
            If Me.Label87.Caption = "Start1" Then
                Call Refresh_Click
            Else
                Exit Sub
            End If
            
        Case "Versand"
            Me.Label87.Caption = "Versand1"
            If Me.Label87.Caption = "Versand1" Then
                Call Refresh_Click
            Else
                Exit Sub
            End If

        Case "Auftragsende"
            Me.Label87.Caption = "Auftragsende1"
            If Me.Label87.Caption = "Auftragsende1" Then
                Call Refresh_Click
            Else
                Exit Sub
            End If

        Case "Rest"
            Me.Label87.Caption = "Rest1"
            If Me.Label87.Caption = "Rest1" Then
                Call Refresh_Click
            Else
                Exit Sub
            End If

    End Select

End Sub

Private Sub openDashboard_Click()
DoCmd.OpenForm "Dashboard"
End Sub


Private Sub BtnChange_Click()
    Dim header As String
    
    If Me!TxtEntry.Visible = True Then
        If IsNull(Me!TxtEntry.value) Then
            If MsgBox("Sind Sie sich sicher, dass Sie für die ausgewählte(n) Zeile(n) das Datenfeld " & Me!LblHeader.Caption & " löschen möchten?", vbYesNo) = vbYes Then
                If Me!LblHeader.Caption = "GroupName" Then
                    header = "NameOfCompanyGroup"
                ElseIf Me!LblHeader.Caption = "Job_Status" Then
                    header = "SB_Status"
                ElseIf Me!LblHeader.Caption = "eConfirmationsSent" Then
                    header = "AuditiSent"
                ElseIf Me!LblHeader.Caption = "Sprache" Then
                    header = "Lang"
                ElseIf Me!LblHeader.Caption = "Anzahl_Anschr" Then
                    header = "Anzahl_Bestätigungen"
                '#JK-4/2020-12-17 START
                ElseIf Me!LblHeader.Caption = "Adressen_Status" Then
                    header = "Adressen_Status"
                ElseIf Me!LblHeader.Caption = "Forensics_Status" Then
                    header = "Forensics_Status"
                ElseIf Me!LblHeader.Caption = "Adressen_Bearbeiter" Then
                    header = "Adressen_Bearbeiter"
                '#JK-4/2020-12-17 START
                Else
                    header = Me!LblHeader.Caption
                End If

                Call RowsSelected("", header)
            End If
        Else
            If Me!LblHeader.Caption = "DateConfi" Or Me!LblHeader.Caption = "YearEnd" Or Me!LblHeader.Caption = "First_Upload_Gewünscht" Or Me!LblHeader.Caption = "DatumAuftragsende" Or Me!LblHeader.Caption = "Eingangsdatum_CAD" Or Me!LblHeader.Caption = "Versanddatum_CAD" Or Me!LblHeader.Caption = "Antwort_Deadline" Or Me!LblHeader.Caption = "Datum_final_Review" Or Me!LblHeader.Caption = "Versanddatum_Originale" Then
                If validateDateFormat(Me!TxtEntry.value) = False Then
                    MsgBox ("Bitte das Datumsformat beachten! Das Datum muss im Format YYYY-MM-DD sein also z.B. 2020-12-31")
                    Exit Sub
                End If
                
                
            ElseIf Me!LblHeader.Caption = "OrderNo" Then
                If isManualOrderNo() = False Then
                    MsgBox ("Das OrderNo Feld lässt sich nur für manuell aufgenommene Bestellungen bearbeiten!")
                    Exit Sub
                End If
            End If
            
            If MsgBox("Sind Sie sich sicher, dass Sie für die ausgewählte(n) Zeile(n) das Datenfeld " & Me!LblHeader.Caption & " zu " & Me!TxtEntry.value & " ändern möchten?", vbYesNo) = vbYes Then
                If Me!LblHeader.Caption = "GroupName" Then
                    header = "NameOfCompanyGroup"
                ElseIf Me!LblHeader.Caption = "Job_Status" Then
                    header = "SB_Status"
                ElseIf Me!LblHeader.Caption = "eConfirmationsSent" Then
                    header = "AuditiSent"
                ElseIf Me!LblHeader.Caption = "Sprache" Then
                    header = "Lang"
                ElseIf Me!LblHeader.Caption = "Anzahl_Anschr" Then
                    header = "Anzahl_Bestätigungen"
                '#JK-5/2020-12-17 START
                ElseIf Me!LblHeader.Caption = "Adressen_Status" Then
                    header = "Adressen_Status"
                ElseIf Me!LblHeader.Caption = "Forensics_Status" Then
                    header = "Forensics_Status"
                ElseIf Me!LblHeader.Caption = "Adressen_Bearbeiter" Then
                    header = "Adressen_Bearbeiter"
                '#JK-5/2020-12-17 START
                Else
                    header = Me!LblHeader.Caption
                End If

                Call RowsSelected(Me!TxtEntry.value, header)
            End If
        End If
    ElseIf Me!cbEntry.Visible = True Then
        If IsNull(Me!cbEntry.value) Then
            If MsgBox("Sind Sie sich sicher, dass Sie für die ausgewählte(n) Zeile(n) das Datenfeld " & Me!LblHeader.Caption & " löschen möchten?", vbYesNo) = vbYes Then
                If Me!LblHeader.Caption = "[1_CAD_Preparer]" Then
                    header = "CAD_Bearbeiter"
                ElseIf Me!LblHeader.Caption = "[2_CAD_Preparer]" Then
                    header = "CAD_Bearbeiter_vorherig"
                ElseIf Me!LblHeader.Caption = "Job_Status" Then
                    header = "SB_Status"
                ElseIf Me!LblHeader.Caption = "Anzahl_Anschr" Then
                    header = "Anzahl_Bestätigungen"
                '#JK-6/2020-12-17 START
                ElseIf Me!LblHeader.Caption = "Adressen_Status" Then
                    header = "Adressen_Status"
                ElseIf Me!LblHeader.Caption = "Forensics_Status" Then
                    header = "Forensics_Status"
                ElseIf Me!LblHeader.Caption = "Adressen_Bearbeiter" Then
                    header = "Adressen_Bearbeiter"
                '#JK-6/2020-12-17 START
                Else
                    header = Me!LblHeader.Caption
                End If
                Call RowsSelected("", header)
            End If
        Else
            If MsgBox("Sind Sie sich sicher, dass Sie für die ausgewählte(n) Zeile(n) das Datenfeld " & Me!LblHeader.Caption & " zu " & Me!cbEntry.value & " ändern möchten?", vbYesNo) = vbYes Then
                If Me!LblHeader.Caption = "[1_CAD_Preparer]" Then
                    header = "CAD_Bearbeiter"
                ElseIf Me!LblHeader.Caption = "[2_CAD_Preparer]" Then
                    header = "CAD_Bearbeiter_vorherig"
                ElseIf Me!LblHeader.Caption = "Job_Status" Then
                    header = "SB_Status"
                ElseIf Me!LblHeader.Caption = "Anzahl_Anschr" Then
                    header = "Anzahl_Bestätigungen"
                '#JK-4/2020-12-09 START
                ElseIf Me!LblHeader.Caption = "Adressen_Status" Then
                    header = "Adressen_Status"
                ElseIf Me!LblHeader.Caption = "Forensics_Status" Then
                    header = "Forensics_Status"
                ElseIf Me!LblHeader.Caption = "Adressen_Bearbeiter" Then
                    header = "Adressen_Bearbeiter"
                '#JK-4/2020-12-09 START
                Else
                    header = Me!LblHeader.Caption
                End If
                Call RowsSelected(Me!cbEntry.value, header)
            End If
        End If
    End If
    
    Call Refresh_Click

End Sub

Private Function isManualOrderNo()

    Dim ctlList As Control
    Set ctlList = Forms![Update_Confi Client].Alltasks
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "^CONM"
    
    isManualOrderNo = True
    
    For Each varItem In ctlList.ItemsSelected
        If regex.test(ctlList.Column(0, varItem)) = False Then
            isManualOrderNo = False
            Exit For
        End If
    Next varItem

End Function

Private Function validateDateFormat(strDate)

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "\d\d\d\d-\d\d-\d\d"
    validateDateFormat = regex.test(strDate)

End Function

Private Sub Refresh_Click()

    'On Error GoTo ErrorMsg
    
    If Me!btneConfi.Caption = "eConfirmations einblenden" Then
        ' Start
        If Me.Label76.Caption = "Start" Then
            AlltasksQuery = buildSQLStart(Me.Label87.Caption, "einblenden")
        End If
        '#JK-4/2020-12-16 START
        If Me.Label76.Caption = "Adressen" Then
            AlltasksQuery = buildSQLAdressen(Me.Label87.Caption, "einblenden")
        End If
        '#JK-4/2020-12-16 END
        If Me.Label76.Caption = "Versand" Then
            AlltasksQuery = buildSQLVersand(Me.Label87.Caption, "einblenden")
        End If
        If Me.Label76.Caption = "Auftragsende" Then
            AlltasksQuery = buildSQLAuftragsende(Me.Label87.Caption, "einblenden")
        End If
        If Me.Label76.Caption = "Rest" Then
            AlltasksQuery = buildSQLRest(Me.Label87.Caption, "einblenden")
        End If
    Else
        If Me.Label76.Caption = "Start" Then
            AlltasksQuery = buildSQLStart(Me.Label87.Caption, "ausblenden")
        End If
        '#JK-5/2020-12-16 START
        If Me.Label76.Caption = "Adressen" Then
            AlltasksQuery = buildSQLAdressen(Me.Label87.Caption, "ausblenden")
        End If
        '#JK-5/2020-12-16 END
        If Me.Label76.Caption = "Versand" Then
            AlltasksQuery = buildSQLVersand(Me.Label87.Caption, "ausblenden")
        End If
        If Me.Label76.Caption = "Auftragsende" Then
            AlltasksQuery = buildSQLAuftragsende(Me.Label87.Caption, "ausblenden")
        End If
        If Me.Label76.Caption = "Rest" Then
            AlltasksQuery = buildSQLRest(Me.Label87.Caption, "ausblenden")
        End If
        
    End If
    
   ' Debug.Print AlltasksQuery
    
    Set Me.Alltasks.RecordSet = getRS(AlltasksQuery)

    
    Select Case Me.Label76.Caption
    
        Case "Start"
            Call adjustStart(CStr(Me.Label87.Caption))
            
        '#JK-6/2020-12-16 START
        Case "Adressen"
            Call adjustAdressen(CStr(Me.Label87.Caption))
        '#JK-6/2020-12-16 START
           
        Case "Versand"
            Call adjustVersand(CStr(Me.Label87.Caption))
            
        Case "Auftragsende"
            Call adjustAuftragsende(CStr(Me.Label87.Caption))
            
        Case "Rest"
            Call adjustRest(CStr(Me.Label87.Caption))
    
    End Select
    
    If Me.txtSearch <> "" Then Call btnSearch_Click
    
    
'
    Dim RSFiltered As ADODB.RecordSet

    Set RSFiltered = Me.Alltasks.RecordSet
    RSFiltered.Filter = Me.Alltasks.RecordSet.Filter

    If InStr(Me.ShowStorno.Caption, "einblenden") Then

        RSFiltered.Filter = ("([Job_Status] <> 'Closed' AND [Job_Status] <> 'Storno') OR [Job_Status] = Null")
        Set Me.Alltasks.RecordSet = RSFiltered

    End If
    
    'Set Me.ListStammdaten.RecordSet = Nothing
'
    On Error GoTo 0
    Exit Sub

Errormsg:
    If CStr(Err.Number) = "3044" Then
        MsgBox ("No connection to the server. Please check the connection and try again")
    Else
        MsgBox ("An error is occured. Please restart the client. Error info: " & Err.Description)
    End If

End Sub

Private Sub Btn_Auftragsende_Click()
    Dim sql As String, t_case As String
    
    sql = "SELECT NameOfCompanyGroup, Client, BusinessEstablishmentOAreAvailable, Tool, Confirmation, TypeConfi, EngCode, EngContact, DatumAuftragsende, Anzahl_erhaltene_Salden, Final_Senior_Review_durch, Datum_final_Review, Versanddatum_Originale, Empfänger_Originale, Empfänger_Nachläufer, Datum_Feedbackanfrage, Datum_Reminder, Feedback_erhalten, CAD_Bearbeiter, CAD_Bearbeiter_vorherig, SB_Status From tCon_Orderbook" & strTestTable & ";"

    t_case = "Auftragsende"
    
    'Dim cnxn As ADODB.Connection
    'Set cnxn = CurrentProject.AccessConnection
    Dim myRecordSet As New ADODB.RecordSet
    
    ' Execute SQL Query
    'myRecordSet.Open SQL, cnxn, adOpenStatic, adLockOptimistic
    Set myRecordSet = getRS(sql)

    Call ExportRecordset(myRecordSet, t_case)
    
    'cnxn.Close
    
    'Set cnxn = Nothing
End Sub

Private Sub Btn_ExportAll_Click()

    Dim str_SQL As String, t_case As String

    str_SQL = "SELECT "
    'str_SQL = str_SQL & " ID," & vbCrLf
    str_SQL = str_SQL & " POrderNo," & vbCrLf
    str_SQL = str_SQL & " OrderNo," & vbCrLf
    str_SQL = str_SQL & " Storno," & vbCrLf
    str_SQL = str_SQL & " GISID," & vbCrLf
    str_SQL = str_SQL & " BelongingToCompanyGroup," & vbCrLf
    str_SQL = str_SQL & " [NameOfCompanyGroup]," & vbCrLf
    str_SQL = str_SQL & " Client," & vbCrLf
    str_SQL = str_SQL & " BusinessEstablishmentOAreAvailable," & vbCrLf
    str_SQL = str_SQL & " BusinessEstablishmentOAre," & vbCrLf
    str_SQL = str_SQL & " TypeFinancials," & vbCrLf
    str_SQL = str_SQL & " Tool," & vbCrLf
    str_SQL = str_SQL & " EYStand," & vbCrLf
    str_SQL = str_SQL & " Confirmation," & vbCrLf
    str_SQL = str_SQL & " TypeConfi," & vbCrLf
    str_SQL = str_SQL & " DateConfi," & vbCrLf
    str_SQL = str_SQL & " YearEnd," & vbCrLf
    str_SQL = str_SQL & " EngCode," & vbCrLf
    str_SQL = str_SQL & " EngName," & vbCrLf
    str_SQL = str_SQL & " Lang," & vbCrLf
    str_SQL = str_SQL & " EngPartner," & vbCrLf
    str_SQL = str_SQL & " EngManager," & vbCrLf
    str_SQL = str_SQL & " OtherContact," & vbCrLf
    str_SQL = str_SQL & " EngContact," & vbCrLf
    str_SQL = str_SQL & " DeliveryType," & vbCrLf
    str_SQL = str_SQL & " OtherTypeConfi," & vbCrLf
    str_SQL = str_SQL & " ConfiFlag," & vbCrLf
    str_SQL = str_SQL & " ConfiMailSent," & vbCrLf
    str_SQL = str_SQL & " TemplateSent," & vbCrLf
    str_SQL = str_SQL & " AuditiSent," & vbCrLf
    str_SQL = str_SQL & " Datum_Versand_EMail_Erstkontakt," & vbCrLf
    str_SQL = str_SQL & " Informationen_zum_Auftrag," & vbCrLf
    str_SQL = str_SQL & " Versand_durch," & vbCrLf
    str_SQL = str_SQL & " PlannedDelivery," & vbCrLf
    str_SQL = str_SQL & " Eingangsdatum_CAD," & vbCrLf
    str_SQL = str_SQL & " Versanddatum_CAD," & vbCrLf
    str_SQL = str_SQL & " Anzahl_Bestätigungen," & vbCrLf
    str_SQL = str_SQL & " Kommentar_Versand," & vbCrLf
    str_SQL = str_SQL & " First_Upload_Gewünscht," & vbCrLf
    str_SQL = str_SQL & " Antwort_Deadline," & vbCrLf
    str_SQL = str_SQL & " [2nd_Request_gewünscht]," & vbCrLf
    str_SQL = str_SQL & " Datum_2nd_Request," & vbCrLf
    str_SQL = str_SQL & " Date2ndReq," & vbCrLf
    str_SQL = str_SQL & " DatumAuftragsende," & vbCrLf
    str_SQL = str_SQL & " Anzahl_erhaltene_Salden," & vbCrLf
    str_SQL = str_SQL & " Final_Senior_Review_durch," & vbCrLf
    str_SQL = str_SQL & " Datum_final_Review," & vbCrLf
    str_SQL = str_SQL & " Versanddatum_Originale," & vbCrLf
    str_SQL = str_SQL & " Empfänger_Originale," & vbCrLf
    str_SQL = str_SQL & " Empfänger_Nachläufer," & vbCrLf
    str_SQL = str_SQL & " Datum_Feedbackanfrage," & vbCrLf
    str_SQL = str_SQL & " Datum_Reminder," & vbCrLf
    str_SQL = str_SQL & " Feedback_erhalten," & vbCrLf
    str_SQL = str_SQL & " CAD_Bearbeiter," & vbCrLf
    str_SQL = str_SQL & " CAD_Bearbeiter_vorherig," & vbCrLf
    str_SQL = str_SQL & " SB_Status," & vbCrLf
    str_SQL = str_SQL & " Abrechnungsbetrag," & vbCrLf
    str_SQL = str_SQL & " Abrechnungsdatum," & vbCrLf
    str_SQL = str_SQL & " ADDITIONAL_CONTACT_2," & vbCrLf
    str_SQL = str_SQL & " ADDITIONAL_CONTACT_3," & vbCrLf
    str_SQL = str_SQL & " ADDITIONAL_CONTACT_4," & vbCrLf
    str_SQL = str_SQL & " ADDITIONAL_CONTACT_5, " & vbCrLf
    '#JK-1/2020-15-12 START
    str_SQL = str_SQL & " Adressen_Status, " & vbCrLf
    '#JK-1/2020-15-12 END
    str_SQL = str_SQL & " Address_Validation, " & vbCrLf
    str_SQL = str_SQL & " Country " & vbCrLf
    
    
    str_SQL = str_SQL & " From tCon_Orderbook" & strTestTable & " Order By OrderNo DESC "
    
    t_case = "ExportAll"
    
    'Dim cnxn As ADODB.Connection
    'Set cnxn = CurrentProject.AccessConnection
    Dim myRecordSet As New ADODB.RecordSet
    
    ' Execute SQL Query
    'myRecordSet.Open str_SQL, cnxn, adOpenStatic, adLockOptimistic
    Set myRecordSet = getRS(str_SQL)

    Call ExportRecordset(myRecordSet, t_case)
    
    'cnxn.Close
    
    'Set cnxn = Nothing

End Sub

Private Sub BtnVersand_Click()

    Dim sql As String, t_case As String
    
    sql = "SELECT NameOfCompanyGroup, Client, BusinessEstablishmentOAreAvailable, Tool, Confirmation, TypeConfi, EngCode, EngContact, DeliveryType, DateConfi, Versand_durch, PlannedDelivery, Eingangsdatum_CAD, Versanddatum_CAD, Anzahl_Bestätigungen, Kommentar_Versand, First_Upload_Gewünscht, Antwort_Deadline, [2nd_Request_gewünscht], Datum_2nd_Request, Date2ndReq, CAD_Bearbeiter, CAD_Bearbeiter_vorherig, SB_Status From tCon_Orderbook" & strTestTable & ";"

    t_case = "Versand"
    
    'Dim cnxn As ADODB.Connection
    'Set cnxn = CurrentProject.AccessConnection
    Dim myRecordSet As New ADODB.RecordSet
    
    ' Execute SQL Query
    'myRecordSet.Open SQL, cnxn, adOpenStatic, adLockOptimistic
    Set myRecordSet = getRS(sql)

    Call ExportRecordset(myRecordSet, t_case)
    
    'cnxn.Close
    
    'Set cnxn = Nothing

End Sub
Private Sub ShowStorno_Click()

    If Me!ShowStorno.Caption = "Storno/Closed einblenden" Then
        Me!ShowStorno.Caption = "Storno/Closed ausblenden"
    Else
        Me!ShowStorno.Caption = "Storno/Closed einblenden"
    End If
    
    Call Refresh_Click

End Sub

Private Sub btneConfi_Click()

    If Me!btneConfi.Caption = "eConfirmations einblenden" Then
        Me!btneConfi.Caption = "eConfirmations ausblenden"
    Else
        Me!btneConfi.Caption = "eConfirmations einblenden"
    End If
    
    Call Refresh_Click

End Sub


Private Sub btnReset_Click()

1         Me.txtSearch.value = ""
          Me.TxtEntry.value = ""
          Me.LblHeader.Caption = ""
2         Call Refresh_Click

End Sub

Private Sub btnSearch_Click()

1         On Error GoTo Errormsg

          Dim filterCol As String, strSort As String, RSFiltered As ADODB.RecordSet
          Dim searchQuery As String, searchText As String

2         Me.cbFilterCol.SetFocus
3         filterCol = Me.cbFilterCol.Column(0)
          If filterCol = "Job_Status" Then filterCol = "SB_Status"
          If filterCol = "Group" Then filterCol = "NameOfCompanyGroup"
          If filterCol = "1_CAD_Preparer" Then filterCol = "CAD_Bearbeiter"
          If filterCol = "2_CAD_Preparer" Then filterCol = "CAD_Bearbeiter_vorherig"
5         If IsNull(Me.txtSearch.value) = False Then searchText = Me.txtSearch.value
          If Len(searchText) > 0 Then
              If filterCol = "Database" Then

                  If Me!btneConfi.Caption = "eConfirmations einblenden" Then
                      ' Start
                      If Me.Label76.Caption = "Start" Then
                          searchQuery = buildSQLStart(Me.Label87.Caption, "einblenden", 2, searchText, filterCol)
                      End If
                      '#TODO
                      If Me.Label76.Caption = "Versand" Then
                          searchQuery = buildSQLVersand(Me.Label87.Caption, "einblenden", 2, searchText, filterCol)
                      End If
                      If Me.Label76.Caption = "Auftragsende" Then
                          searchQuery = buildSQLAuftragsende(Me.Label87.Caption, "einblenden", 2, searchText, filterCol)
                      End If
                      If Me.Label76.Caption = "Rest" Then
                          searchQuery = buildSQLRest(Me.Label87.Caption, "einblenden", 2, searchText, filterCol)
                      End If
                  Else
                      If Me.Label76.Caption = "Start" Then
                          searchQuery = buildSQLStart(Me.Label87.Caption, "ausblenden", 2, searchText, filterCol)
                      End If
                      '#TODO
                      If Me.Label76.Caption = "Versand" Then
                          searchQuery = buildSQLVersand(Me.Label87.Caption, "ausblenden", 2, searchText, filterCol)
                      End If
                      If Me.Label76.Caption = "Auftragsende" Then
                          searchQuery = buildSQLAuftragsende(Me.Label87.Caption, "ausblenden", 2, searchText, filterCol)
                      End If
                      If Me.Label76.Caption = "Rest" Then
                          searchQuery = buildSQLRest(Me.Label87.Caption, "ausblenden", 2, searchText, filterCol)
                      End If
        
                  End If

              Else
              
                  If Me!btneConfi.Caption = "eConfirmations einblenden" Then
                      ' Start
                      If Me.Label76.Caption = "Start" Then
                          searchQuery = buildSQLStart(Me.Label87.Caption, "einblenden", 1, searchText, filterCol)
            
                      End If
                      '#TODO
                      If Me.Label76.Caption = "Versand" Then
                          searchQuery = buildSQLVersand(Me.Label87.Caption, "einblenden", 1, searchText, filterCol)
                      End If
                      If Me.Label76.Caption = "Auftragsende" Then
                          searchQuery = buildSQLAuftragsende(Me.Label87.Caption, "einblenden", 1, searchText, filterCol)
                      End If
                      If Me.Label76.Caption = "Rest" Then
                          searchQuery = buildSQLRest(Me.Label87.Caption, "einblenden", 1, searchText, filterCol)
                      End If
                  Else
                      If Me.Label76.Caption = "Start" Then
                          searchQuery = buildSQLStart(Me.Label87.Caption, "ausblenden", 1, searchText, filterCol)
                      End If
                      '#TODO
                      If Me.Label76.Caption = "Versand" Then
                          searchQuery = buildSQLVersand(Me.Label87.Caption, "ausblenden", 1, searchText, filterCol)
                      End If
                      If Me.Label76.Caption = "Auftragsende" Then
                          searchQuery = buildSQLAuftragsende(Me.Label87.Caption, "ausblenden", 1, searchText, filterCol)
                      End If
                      If Me.Label76.Caption = "Rest" Then
                          searchQuery = buildSQLRest(Me.Label87.Caption, "ausblenden", 1, searchText, filterCol)
                      End If
        
                  End If
              End If
7             Set RSFiltered = getRS(searchQuery)
8             Set Me.Alltasks.RecordSet = RSFiltered

              Select Case Me.Label76.Caption
    
                  Case "Start"
                      Call adjustStart(CStr(Me.Label87.Caption))
                      
                    '#TODO
            
                  Case "Versand"
                      Call adjustVersand(CStr(Me.Label87.Caption))
            
                  Case "Auftragsende"
                      Call adjustAuftragsende(CStr(Me.Label87.Caption))
            
                  Case "Rest"
                      Call adjustRest(CStr(Me.Label87.Caption))
    
            End Select
          End If
          On Error GoTo 0
9         Exit Sub
          
Errormsg:
10        If CStr(Err.Number) = "3044" Then
11            MsgBox ("No connection to the server. Please check the connection and try again")
12        ElseIf CStr(Err.Number) = "94" Then
13            MsgBox ("Please select an order")
14        Else
15            MsgBox ("An error is occured. Please restart the client. Error Info: " & Err.Description)
16        End If

End Sub

Private Sub btnSortRS_Click()

    On Error GoTo Errormsg
    
    
    If InStr(Me.ShowStorno.Caption, "einblenden") Then
    
        Dim RSFiltered As ADODB.RecordSet
        Set RSFiltered = Me.Alltasks.RecordSet
    
        RSFiltered.Filter = ("([Job_Status] <> 'Closed' AND [Job_Status] <> 'Storno') OR [Job_Status] = Null")
        
        Dim oStm As ADODB.Stream
        Dim RSFilteredout As ADODB.RecordSet

        ' Save records into an ADO Stream
        Set oStm = New ADODB.Stream
        oStm.Open
        RSFiltered.Save oStm, adPersistXML

        ' Create a new Recordset based on the initial recordset (via Stream)
        Set RSFilteredout = New ADODB.RecordSet
        RSFilteredout.Open oStm

        Set Me.Alltasks.RecordSet = Nothing
        Set Me.Alltasks.RecordSet = RSFilteredout

        Set RSFiltered = Nothing
        Set RSFilteredout = Nothing
        oStm.Close
    
    End If
    
    Dim sortcol As String, strSort As String, RSSorted As ADODB.RecordSet, SortDirection As String

    Me.cbSortCol.SetFocus
    sortcol = Me.cbSortCol.Column(0)

    If sortcol = "Group" Then sortcol = "GroupName"
      
    If Me.obtnAsc.value = True Then SortDirection = "ASC"
    If Me.obtnDesc.value = True Then SortDirection = "DESC"
   
    Set RSSorted = Me.Alltasks.RecordSet
    RSSorted.Sort = sortcol & " " & SortDirection

    Set Me.Alltasks.RecordSet = RSSorted


    On Error GoTo 0
    Exit Sub
    
Errormsg:
    If CStr(Err.Number) = "3044" Then
        MsgBox ("No connection to the server. Please check the connection and try again")
    Else
        MsgBox ("An error is occured. Please restart the client. Error info: " & Err.Description)
    End If

End Sub

Private Sub obtnDesc_Click()

    If Me.obtnAsc.value = True Then Me.obtnAsc.value = False
    
End Sub

Private Sub obtnAsc_Click()

    If Me.obtnDesc.value = True Then Me.obtnDesc.value = False
    
End Sub

Private Sub opentOrderbook_Click()
    On Error GoTo Errormsg

    DoCmd.OpenTable "Orderbook"
    DoCmd.OpenTable "BankConfirmation"
    DoCmd.OpenTable "ConfiClientContact"
    DoCmd.OpenTable "tUser"

    On Error GoTo 0
    Exit Sub

Errormsg:
    If CStr(Err.Number) = "3044" Then
        MsgBox ("No connection to the server. Please check the connection and try again")
    Else
        MsgBox ("An error is occured. Please restart the client. Error info: " & Err.Description)
    End If
    
End Sub

Private Sub closetOrderbook_Click()
    DoCmd.Close acTable, "Orderbook"
    DoCmd.Close acTable, "BankConfirmation"
    DoCmd.Close acTable, "ConfiClientContact"
    DoCmd.Close acTable, "tUser"
End Sub

' ######### BUTTON END ###########


' ######### FUNCTION START ###########

Private Sub Alltasks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer

    Dim RSpreparer As ADODB.RecordSet

    Dim Spl As String
    Dim ColumnID As Integer, anzCol As Integer
    Dim width As Double, stepWidth As Double

    Dim columnWidths() As Double
    width = Alltasks.width - 20
    Select Case Me.Label76.Caption
    
        Case "Start"
        
            If Me.Label87.Caption = "Start1" Then
                '#JK-6/2020-16-12 START
                ReDim columnWidths(1 To 18)
                anzCol = 18
                '#JK-6/2020-16-12 END
                stepWidth = Round(width / 45, 3)
            Else
                ReDim columnWidths(1 To 14)
                anzCol = 14
                stepWidth = Round(width / 45, 3)
            End If
    
        '#JK-7/2020-12-16 START
        Case "Adressen"
            ReDim columnWidths(1 To 16)
            anzCol = 16
            stepWidth = Round(width / 45, 3)
        '#JK-7/2020-12-16 START
    
        Case "Versand"
            ReDim columnWidths(1 To 15)
            anzCol = 15
            stepWidth = Round(width / 48, 3)
    
        Case "Auftragsende"
            ReDim columnWidths(1 To 13)
            anzCol = 13
            stepWidth = Round(width / 48, 3)
    
        Case "Rest"
            ReDim columnWidths(1 To 17)
            anzCol = 17
            stepWidth = Round(width / 48, 3)
    
    End Select


    ' Gleichtverteilung der Spaltenbreiten wäre 3*stepWidth pro Spalte
    ' Wenn alles addiert 45 * stepWidth ergibt, stimmt die Breite
    Select Case Me.Label87.Caption
    
        Case "Start1"
            columnWidths(1) = 3 * stepWidth
            columnWidths(2) = 4 * stepWidth
            columnWidths(3) = 4 * stepWidth
            columnWidths(4) = 2 * stepWidth
            columnWidths(5) = 3.5 * stepWidth
            columnWidths(6) = 1.5 * stepWidth
            columnWidths(7) = 3 * stepWidth
            columnWidths(8) = 3 * stepWidth
            columnWidths(9) = 1.5 * stepWidth
            columnWidths(10) = 2 * stepWidth
            columnWidths(11) = 2 * stepWidth
            columnWidths(12) = 2 * stepWidth
            columnWidths(13) = 2 * stepWidth
            '#JK-1/2020-12-17 START
            columnWidths(14) = 5 * stepWidth
            columnWidths(15) = 1.5 * stepWidth
            columnWidths(16) = 1.5 * stepWidth
            columnWidths(17) = 1.5 * stepWidth
            columnWidths(18) = width - (42.5 * stepWidth)
            '#JK-1/2020-12-17 START
    
        Case "Start2"
            columnWidths(1) = 3 * stepWidth
            columnWidths(2) = 4 * stepWidth
            columnWidths(3) = 4 * stepWidth
            columnWidths(4) = 2 * stepWidth
            columnWidths(5) = 4 * stepWidth
            columnWidths(6) = 1.5 * stepWidth
            columnWidths(7) = 3 * stepWidth
            columnWidths(8) = 3.5 * stepWidth
            columnWidths(9) = 2 * stepWidth
            columnWidths(10) = 3 * stepWidth
            columnWidths(11) = 2 * stepWidth
            columnWidths(12) = 2 * stepWidth
            columnWidths(13) = 2 * stepWidth
            columnWidths(14) = width - (42.5 * stepWidth)
            
        '#JK-8/2020-12-16 START
        Case "Adressen1"
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
        '#JK-8/2020-12-16 END
            
        Case "Versand1"
            columnWidths(1) = 3 * stepWidth
            columnWidths(2) = 4.5 * stepWidth
            columnWidths(3) = 4.5 * stepWidth
            columnWidths(4) = 4.5 * stepWidth
            columnWidths(5) = 2 * stepWidth
            columnWidths(6) = 3.5 * stepWidth
            columnWidths(7) = 2 * stepWidth
            
            columnWidths(8) = 2 * stepWidth
            columnWidths(9) = 2 * stepWidth
            columnWidths(10) = 2 * stepWidth
            columnWidths(11) = 2 * stepWidth
            columnWidths(12) = 2.5 * stepWidth
            columnWidths(13) = 2.5 * stepWidth
            columnWidths(14) = 2.5 * stepWidth
            columnWidths(15) = width - (45.5 * stepWidth)
            
        Case "Auftragsende1"
            columnWidths(1) = 3 * stepWidth
            columnWidths(2) = 4.5 * stepWidth
            columnWidths(3) = 4.5 * stepWidth
            columnWidths(4) = 3 * stepWidth
            columnWidths(5) = 2.5 * stepWidth
            columnWidths(6) = 3.5 * stepWidth
            columnWidths(7) = 2.5 * stepWidth
            
            columnWidths(8) = 3.5 * stepWidth
            columnWidths(9) = 2.5 * stepWidth
            columnWidths(10) = 2 * stepWidth
            columnWidths(11) = 2 * stepWidth
            columnWidths(12) = 4.5 * stepWidth
            columnWidths(13) = width - (45.5 * stepWidth)
    
        Case "Rest1"
            columnWidths(1) = 3 * stepWidth
            columnWidths(2) = 4 * stepWidth
            columnWidths(3) = 4 * stepWidth
            columnWidths(4) = 2 * stepWidth
            columnWidths(5) = 4 * stepWidth
            columnWidths(6) = 1.5 * stepWidth
            columnWidths(7) = 2.5 * stepWidth
            
            columnWidths(8) = 2 * stepWidth
            columnWidths(9) = 2.5 * stepWidth
            columnWidths(10) = 4 * stepWidth
            columnWidths(11) = 4 * stepWidth
            columnWidths(12) = 4 * stepWidth
            columnWidths(13) = 4 * stepWidth
            columnWidths(14) = 1.5 * stepWidth
            columnWidths(15) = 1.5 * stepWidth
            columnWidths(16) = 1.5 * stepWidth
            columnWidths(17) = width - (45.5 * stepWidth)
    
        Case "Rest2"

            columnWidths(1) = 3 * stepWidth
            columnWidths(2) = 4 * stepWidth
            columnWidths(3) = 4 * stepWidth
            columnWidths(4) = 2 * stepWidth
            columnWidths(5) = 4 * stepWidth
            columnWidths(6) = 1.5 * stepWidth
            columnWidths(7) = 3 * stepWidth
            
            columnWidths(8) = 1.5 * stepWidth
            columnWidths(9) = 1.5 * stepWidth
            columnWidths(10) = 2.5 * stepWidth
            columnWidths(11) = 5.5 * stepWidth
            columnWidths(12) = 4.5 * stepWidth
            columnWidths(13) = 2 * stepWidth
            columnWidths(14) = 2 * stepWidth
            columnWidths(15) = 2 * stepWidth
            columnWidths(16) = 2 * stepWidth
            columnWidths(17) = width - (45.5 * stepWidth)
            
            
    End Select

    Dim leftBound As Double, rightBound As Double
    leftBound = 0
    rightBound = 0

    For i = 0 To anzCol - 1
        rightBound = rightBound + columnWidths(i + 1)
        If i > 0 Then leftBound = leftBound + columnWidths(i)
        If leftBound < X And X < rightBound Then
            Select Case Me.Label87.Caption:
                Case "Start1":
                    Select Case i:
                        
                        Case 0: Spl = "OrderNo"
                        Case 1: Spl = "NameOfCompanyGroup"
                        Case 2: Spl = "Client"
                        Case 3: Spl = "Tool"
                        Case 4: Spl = "Confirmation"
                        Case 5: Spl = "Job_Status"
                        Case 6: Spl = "CAD_Bearbeiter"
                        Case 7: Spl = "CAD_Bearbeiter_vorherig"

                        Case 8: Spl = "PlannedDelivery"
                        Case 9: Spl = "TypeConfi"
                        Case 10: Spl = "DateConfi"
                        Case 11: Spl = "YearEnd"
                        Case 12: Spl = "EngCode"
                        Case 13: Spl = "EngContact"
                        Case 14: Spl = "Country"
                        '#JK-2/2020-12-09 START
                        Case 15: Spl = "AC_Status"
                        Case 16: Spl = "AC_Preparer"
                        Case 17: Spl = "DeliveryType"
                        '#JK-2/2020-12-09 END
                            
                    End Select
                    
                Case "Start2":
                    Select Case i:
                        Case 0: Spl = "OrderNo"
                        Case 1: Spl = "NameOfCompanyGroup"
                        Case 2: Spl = "Client"
                        Case 3: Spl = "Tool"
                        Case 4: Spl = "Confirmation"
                        Case 5: Spl = "Job_Status"
                        Case 6: Spl = "CAD_Bearbeiter"
                        Case 7: Spl = "CAD_Bearbeiter_vorherig"

                        Case 8: Spl = "Datum_Versand_EMail_Erstkontakt"
                        Case 9: Spl = "Versand_durch"
                        Case 10: Spl = "PlannedDelivery"
                        Case 11: Spl = "First_Upload_Gewünscht"
                        Case 12: Spl = "DatumAuftragsende"
                        Case 13: Spl = "Informationen_zum_Auftrag"
                        
                    End Select
                    
                '#JK-6/2020-12-09 START
                Case "Adressen1":
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
                '#JK-6/2020-12-09 START
                    
                Case "Versand1"
                 
                    Select Case i
                        Case 0: Spl = "OrderNo"
                        Case 1: Spl = "NameOfCompanyGroup"
                        Case 2: Spl = "Client"
                        Case 3: Spl = "Confirmation"
                        Case 4: Spl = "Job_Status"
                        Case 5: Spl = "CAD_Bearbeiter"
                        Case 6: Spl = "PlannedDelivery"
                        Case 7: Spl = "DeliveryType"
                        Case 8: Spl = "Versand_durch"
                        Case 9: Spl = "Eingangsdatum_CAD"
                        Case 10: Spl = "Versanddatum_CAD"
                        Case 11: Spl = "Anzahl_Anschr"
                        Case 12: Spl = "Antwort_Deadline"
                        Case 13: Spl = "Date2ndReq"
                        Case 14: Spl = "Kommentar_Versand"

                    End Select
                    
                Case "Auftragsende1"

                    Select Case i
                        Case 0: Spl = "OrderNo"
                        Case 1: Spl = "NameOfCompanyGroup"
                        Case 2: Spl = "Client"
                        Case 3: Spl = "Confirmation"
                        Case 4: Spl = "Job_Status"
                        Case 5: Spl = "CAD_Bearbeiter"
                        Case 6: Spl = "DatumAuftragsende"
                        Case 7: Spl = "Final_Senior_Review_durch"
                        Case 8: Spl = "Datum_final_Review"
                        Case 9: Spl = "Anzahl_erhaltene_Salden"
                        Case 10: Spl = "Versanddatum_Originale"
                        Case 11: Spl = "Empfänger_Originale"
                        Case 12: Spl = "Empfänger_Nachläufer"
                    End Select
                    
                Case "Rest1"

                    Select Case i
                        Case 0: Spl = "OrderNo"
                        Case 1: Spl = "NameOfCompanyGroup"
                        Case 2: Spl = "Client"
                        Case 3: Spl = "Tool"
                        Case 4: Spl = "Confirmation"
                        Case 5: Spl = "Job_Status"
                        Case 6: Spl = "TypeConfi"

                        Case 7: Spl = "YearEnd"
                        Case 8: Spl = "EngCode"
                        Case 9: Spl = "EngPartner"
                        Case 10: Spl = "EngManager"
                        Case 11: Spl = "EngContact"
                        Case 12: Spl = "OtherContact"
                        Case 13: Spl = "TypeFinancials"
                        Case 14: Spl = "Lang"
                        Case 15: Spl = "EYStand"
                        Case 16: Spl = "BelongingToCompanyGroup"
                        
                    End Select
                    
                Case "Rest2"
                 
                    Select Case i
                        Case 0: Spl = "OrderNo"
                        Case 1: Spl = "NameOfCompanyGroup"
                        Case 2: Spl = "Client"
                        Case 3: Spl = "Tool"
                        Case 4: Spl = "Confirmation"
                        Case 5: Spl = "Job_Status"
                        Case 6: Spl = "TypeConfi"

                        Case 7: Spl = "BusinessEstablishmentOAreAvailable"
                        Case 8: Spl = "BusinessEstablishmentOAre"
                        Case 9: Spl = "GISID"
                        Case 10: Spl = "EngName"
                        Case 11: Spl = "OtherTypeConfi"
                        Case 12: Spl = "ConfiMailSent"
                        Case 13: Spl = "TemplateSent"
                        Case 14: Spl = "AuditiSent"
                        Case 15: Spl = "Storno"
                        Case 16: Spl = "ConfiFlag"

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
            ElseIf Spl = "[1_CAD_Preparer]" Or Spl = "[2_CAD_Preparer]" Then
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
            '#JK-3/2020-12-09 START
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
            '#JK-3/2020-12-09 END
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

Private Sub adjustStart(str_input As String)

    Dim t_width As Integer
    Dim width As Double, stepWidth As Double
    
    Dim columnWidths() As Double
    width = Alltasks.width
    
    If Me.Label87.Caption = "Start1" Then
        '#JK-10/2020-12-16 START
        ReDim columnWidths(1 To 18)
        Alltasks.ColumnCount = 18
        t_width = 18
        '#JK-10/2020-12-16 START
    Else
        ReDim columnWidths(1 To 14)
        Alltasks.ColumnCount = 14
        t_width = 14
    End If


    width = Alltasks.width
    
    stepWidth = Round(width / 45, 3)
    ' Gleichtverteilung der Spaltenbreiten wäre 3*stepWidth pro Spalte
    ' Wenn alles addiert 45 * stepWidth ergibt, stimmt die Breite
    Select Case Me.Label87.Caption
    
        Case "Start1"
            columnWidths(1) = 3 * stepWidth
            columnWidths(2) = 4 * stepWidth
            columnWidths(3) = 4 * stepWidth
            columnWidths(4) = 2 * stepWidth
            columnWidths(5) = 3.5 * stepWidth
            columnWidths(6) = 1.5 * stepWidth
            columnWidths(7) = 3 * stepWidth
            columnWidths(8) = 3 * stepWidth
            columnWidths(9) = 1.5 * stepWidth
            columnWidths(10) = 2 * stepWidth
            columnWidths(11) = 2 * stepWidth
            columnWidths(12) = 2 * stepWidth
            columnWidths(13) = 2 * stepWidth
            '#JK-9/2020-12-16 START
            columnWidths(14) = 5 * stepWidth
            columnWidths(15) = 1.5 * stepWidth
            columnWidths(16) = 1.5 * stepWidth
            columnWidths(17) = 1.5 * stepWidth
            columnWidths(18) = width - (42.5 * stepWidth)
        '#JK-9/2020-12-16 END
    
        Case "Start2"
            columnWidths(1) = 3 * stepWidth
            columnWidths(2) = 4 * stepWidth
            columnWidths(3) = 4 * stepWidth
            columnWidths(4) = 2 * stepWidth
            columnWidths(5) = 4 * stepWidth
            columnWidths(6) = 1.5 * stepWidth
            columnWidths(7) = 3 * stepWidth
            columnWidths(8) = 3.5 * stepWidth
            columnWidths(9) = 2 * stepWidth
            columnWidths(10) = 3 * stepWidth
            columnWidths(11) = 2 * stepWidth
            columnWidths(12) = 2 * stepWidth
            columnWidths(13) = 2 * stepWidth
            columnWidths(14) = width - (42.5 * stepWidth)
    End Select

    
    Dim sum As Double
    Dim str_colWidths As String
    For i = 1 To t_width
        sum = sum + columnWidths(i)
        str_colWidths = str_colWidths & CStr(columnWidths(i)) & ";"
    Next i

    str_colWidths = Left(str_colWidths, Len(str_colWidths) - 1)
    
    Alltasks.columnWidths = str_colWidths

End Sub

'#JK-6/2020-12-09 START
Private Sub adjustAdressen(str_input As String)

    Dim t_width As Integer
    Dim width As Double, stepWidth As Double
    
    Dim columnWidths() As Double
    width = Alltasks.width
    
    If Me.Label87.Caption = "Adressen1" Then
        ReDim columnWidths(1 To 16)
        Alltasks.ColumnCount = 16
        t_width = 16
'    Else
'        ReDim columnWidths(1 To 14)
'        Alltasks.ColumnCount = 14
'        t_width = 14
    End If


    width = Alltasks.width
    
    stepWidth = Round(width / 45, 3)
    ' Gleichtverteilung der Spaltenbreiten wäre 3*stepWidth pro Spalte
    ' Wenn alles addiert 45 * stepWidth ergibt, stimmt die Breite
    Select Case Me.Label87.Caption
    
        Case "Adressen1"
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
    
    End Select

    
    Dim sum As Double
    Dim str_colWidths As String
    For i = 1 To t_width
        sum = sum + columnWidths(i)
        str_colWidths = str_colWidths & CStr(columnWidths(i)) & ";"
    Next i

    str_colWidths = Left(str_colWidths, Len(str_colWidths) - 1)
    
    Alltasks.columnWidths = str_colWidths

End Sub
'#JK-6/2020-12-09 END

Private Sub adjustVersand(str_input As String)

    Dim t_width As Integer
    Dim width As Double, stepWidth As Double
    
    Dim columnWidths() As Double
    width = Alltasks.width
    
    ReDim columnWidths(1 To 15)
    Alltasks.ColumnCount = 15
    t_width = 15

    width = Alltasks.width
    
    stepWidth = Round(width / 48, 3)
    ' Gleichtverteilung der Spaltenbreiten wäre 3*stepWidth pro Spalte
    ' Wenn alles addiert 45 * stepWidth ergibt, stimmt die Breite
    
    Select Case Me.Label87.Caption
    
        Case "Versand1"

            columnWidths(1) = 3 * stepWidth
            columnWidths(2) = 4.5 * stepWidth
            columnWidths(3) = 4.5 * stepWidth
            columnWidths(4) = 4.5 * stepWidth
            columnWidths(5) = 2 * stepWidth
            columnWidths(6) = 3.5 * stepWidth
            columnWidths(7) = 2 * stepWidth
            
            columnWidths(8) = 2 * stepWidth
            columnWidths(9) = 2 * stepWidth
            columnWidths(10) = 2 * stepWidth
            columnWidths(11) = 2 * stepWidth
            columnWidths(12) = 2.5 * stepWidth
            columnWidths(13) = 2.5 * stepWidth
            columnWidths(14) = 2.5 * stepWidth
            columnWidths(15) = width - (45.5 * stepWidth)
            
    End Select

    Dim sum As Double
    Dim str_colWidths As String
    For i = 1 To t_width
        sum = sum + columnWidths(i)
        str_colWidths = str_colWidths & CStr(columnWidths(i)) & ";"
    Next i

    str_colWidths = Left(str_colWidths, Len(str_colWidths) - 1)
    
    Alltasks.columnWidths = str_colWidths

End Sub

Private Sub adjustAuftragsende(str_input As String)

    Dim t_width As Integer
    Dim width As Double, stepWidth As Double
    
    Dim columnWidths() As Double
    width = Alltasks.width
    
    ReDim columnWidths(1 To 13)
    Alltasks.ColumnCount = 13
    t_width = 13

    width = Alltasks.width
    
    stepWidth = Round(width / 48, 3)
    ' Gleichtverteilung der Spaltenbreiten wäre 3*stepWidth pro Spalte
    ' Wenn alles addiert 45 * stepWidth ergibt, stimmt die Breite
    Select Case Me.Label87.Caption
    
        Case "Auftragsende1"
            columnWidths(1) = 3 * stepWidth
            columnWidths(2) = 4.5 * stepWidth
            columnWidths(3) = 4.5 * stepWidth
            columnWidths(4) = 3 * stepWidth
            columnWidths(5) = 2.5 * stepWidth
            columnWidths(6) = 3.5 * stepWidth
            columnWidths(7) = 2.5 * stepWidth
            
            columnWidths(8) = 3.5 * stepWidth
            columnWidths(9) = 2.5 * stepWidth
            columnWidths(10) = 2 * stepWidth
            columnWidths(11) = 2 * stepWidth
            columnWidths(12) = 4.5 * stepWidth
            columnWidths(13) = width - (45.5 * stepWidth)
            
    End Select

    Dim sum As Double
    Dim str_colWidths As String
    For i = 1 To t_width
        sum = sum + columnWidths(i)
        str_colWidths = str_colWidths & CStr(columnWidths(i)) & ";"
    Next i

    str_colWidths = Left(str_colWidths, Len(str_colWidths) - 1)
    
    Alltasks.columnWidths = str_colWidths

End Sub

Private Sub adjustRest(str_input As String)

    Dim t_width As Integer
    Dim width As Double, stepWidth As Double
    
    Dim columnWidths() As Double
    width = Alltasks.width
    
    ReDim columnWidths(1 To 17)
    Alltasks.ColumnCount = 17
    t_width = 17

    width = Alltasks.width
    
    stepWidth = Round(width / 48, 3)
    ' Gleichtverteilung der Spaltenbreiten wäre 3*stepWidth pro Spalte
    ' Wenn alles addiert 45 * stepWidth ergibt, stimmt die Breite
    Select Case Me.Label87.Caption
    
        Case "Rest1"
            columnWidths(1) = 3 * stepWidth
            columnWidths(2) = 4 * stepWidth
            columnWidths(3) = 4 * stepWidth
            columnWidths(4) = 2 * stepWidth
            columnWidths(5) = 4 * stepWidth
            columnWidths(6) = 1.5 * stepWidth
            columnWidths(7) = 2.5 * stepWidth
            
            columnWidths(8) = 2 * stepWidth
            columnWidths(9) = 2.5 * stepWidth
            columnWidths(10) = 4 * stepWidth
            columnWidths(11) = 4 * stepWidth
            columnWidths(12) = 4 * stepWidth
            columnWidths(13) = 4 * stepWidth
            columnWidths(14) = 1.5 * stepWidth
            columnWidths(15) = 1.5 * stepWidth
            columnWidths(16) = 1.5 * stepWidth
            columnWidths(17) = width - (45.5 * stepWidth)
    
        Case "Rest2"

            columnWidths(1) = 3 * stepWidth
            columnWidths(2) = 4 * stepWidth
            columnWidths(3) = 4 * stepWidth
            columnWidths(4) = 2 * stepWidth
            columnWidths(5) = 4 * stepWidth
            columnWidths(6) = 1.5 * stepWidth
            columnWidths(7) = 3 * stepWidth
            
            columnWidths(8) = 1.5 * stepWidth
            columnWidths(9) = 1.5 * stepWidth
            columnWidths(10) = 2.5 * stepWidth
            columnWidths(11) = 5.5 * stepWidth
            columnWidths(12) = 4.5 * stepWidth
            columnWidths(13) = 2 * stepWidth
            columnWidths(14) = 2 * stepWidth
            columnWidths(15) = 2 * stepWidth
            columnWidths(16) = 2 * stepWidth
            columnWidths(17) = width - (45.5 * stepWidth)
            
    End Select

    Dim sum As Double
    Dim str_colWidths As String
    For i = 1 To t_width
        sum = sum + columnWidths(i)
        str_colWidths = str_colWidths & CStr(columnWidths(i)) & ";"
    Next i

    str_colWidths = Left(str_colWidths, Len(str_colWidths) - 1)
    
    Alltasks.columnWidths = str_colWidths

End Sub

Private Function buildSQLStart(t_case As String, t_econfiBtn As String, Optional modus As Integer, Optional searchText As String, Optional filterCol As String) As String

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
    str_SQL = str_SQL & "[SB_Status] As Job_Status, " & vbCrLf
    str_SQL = str_SQL & "[CAD_Bearbeiter] As [1_CAD_Preparer], " & vbCrLf
    str_SQL = str_SQL & "[CAD_Bearbeiter_vorherig] As [2_CAD_Preparer], " & vbCrLf
    

    Select Case t_case
    
        Case "Start1"
            str_SQL = str_SQL & "PlannedDelivery ," & vbCrLf
            str_SQL = str_SQL & "TypeConfi ," & vbCrLf
            str_SQL = str_SQL & "DateConfi ," & vbCrLf
            str_SQL = str_SQL & "YearEnd ," & vbCrLf
            str_SQL = str_SQL & "EngCode ," & vbCrLf
            str_SQL = str_SQL & "EngContact ," & vbCrLf
            str_SQL = str_SQL & "Country ," & vbCrLf
            '#JK-1/2020-12-09 START
            'str_SQL = str_SQL & "Address_Validation as AC, " & vbCrLf
            str_SQL = str_SQL & "[AC_Status], " & vbCrLf
            str_SQL = str_SQL & "[AC_Preparer], " & vbCrLf
            '#JK-1/2020-12-09 END
            str_SQL = str_SQL & "DeliveryType " & vbCrLf

        Case "Start2"
            str_SQL = str_SQL & "[Datum_Versand_EMail_Erstkontakt] ," & vbCrLf
            str_SQL = str_SQL & "[Versand_durch] ," & vbCrLf
            str_SQL = str_SQL & "PlannedDelivery ," & vbCrLf
            str_SQL = str_SQL & "First_Upload_Gewünscht ," & vbCrLf
            str_SQL = str_SQL & "DatumAuftragsende, " & vbCrLf
            str_SQL = str_SQL & "[Informationen_zum_Auftrag] " & vbCrLf
            
    End Select

    str_SQL = str_SQL & "From tCon_Orderbook" & strTestTable & vbCrLf
    
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
    
    buildSQLStart = str_SQL


End Function

'#JK-5/2020-12-09 START
Private Function buildSQLAdressen(t_case As String, t_econfiBtn As String, Optional modus As Integer, Optional searchText As String, Optional filterCol As String) As String

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

    

    Select Case t_case
    
        Case "Adressen1"
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
            str_SQL = str_SQL & "Forensics_Status as [Forensics-Status] " & vbCrLf
            

            
            
            
    End Select

    str_SQL = str_SQL & "From tCon_Orderbook" & strTestTable & vbCrLf
    
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
    
    buildSQLAdressen = str_SQL


End Function
'#JK-5/2020-12-09 END

Private Function buildSQLVersand(t_case As String, t_econfiBtn As String, Optional modus As Integer, Optional searchText As String, Optional filterCol As String) As String

    Dim strSQL As String
    Dim b_buildQuery As Boolean
    b_buildQuery = False

    ' Improve SQL Query by selecting only necessary information
    If modus = 0 Then b_buildQuery = True
 
    ' Start Query
    str_SQL = "SELECT " & vbCrLf

    ' Select basic_info
    str_SQL = str_SQL & "OrderNo ," & vbCrLf
    str_SQL = str_SQL & "NameOfCompanyGroup as GroupName ," & vbCrLf
    str_SQL = str_SQL & "Client ," & vbCrLf
    str_SQL = str_SQL & "Confirmation ," & vbCrLf
    str_SQL = str_SQL & "[SB_Status] As Job_Status, " & vbCrLf
    str_SQL = str_SQL & "[CAD_Bearbeiter] As [1_CAD_Preparer] ," & vbCrLf

    Select Case t_case
    
        Case "Versand1"
            str_SQL = str_SQL & "[PlannedDelivery] ," & vbCrLf
            str_SQL = str_SQL & "DeliveryType ," & vbCrLf
            str_SQL = str_SQL & "Versand_durch ," & vbCrLf
            str_SQL = str_SQL & "[Eingangsdatum_CAD] ," & vbCrLf
            str_SQL = str_SQL & "[Versanddatum_CAD] ," & vbCrLf
            str_SQL = str_SQL & "[Anzahl_Bestätigungen] as Anzahl_Anschr ," & vbCrLf
            str_SQL = str_SQL & "[Antwort_Deadline], " & vbCrLf
            str_SQL = str_SQL & "[Date2ndReq], " & vbCrLf
            str_SQL = str_SQL & "[Kommentar_Versand] " & vbCrLf
     
    End Select
    

 
    ' Join table Orderbook x ConfiClientContact
    str_SQL = str_SQL & "From tCon_Orderbook" & strTestTable & vbCrLf
    
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
    
    buildSQLVersand = str_SQL

End Function
Private Function buildSQLAuftragsende(t_case As String, t_econfiBtn As String, Optional modus As Integer, Optional searchText As String, Optional filterCol As String) As String

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
    str_SQL = str_SQL & "Confirmation ," & vbCrLf
    str_SQL = str_SQL & "[SB_Status] As Job_Status, " & vbCrLf
    str_SQL = str_SQL & "[CAD_Bearbeiter] As [1_CAD_Preparer], " & vbCrLf
    str_SQL = str_SQL & "DatumAuftragsende ," & vbCrLf
    str_SQL = str_SQL & "Final_Senior_Review_durch ," & vbCrLf
    str_SQL = str_SQL & "Datum_final_Review ," & vbCrLf
    str_SQL = str_SQL & "Anzahl_erhaltene_Salden ," & vbCrLf
    str_SQL = str_SQL & "Versanddatum_Originale ," & vbCrLf
    str_SQL = str_SQL & "[Empfänger_Originale] ," & vbCrLf
    str_SQL = str_SQL & "[Empfänger_Nachläufer] " & vbCrLf

    str_SQL = str_SQL & "From tCon_Orderbook" & strTestTable & vbCrLf
    
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
    
    buildSQLAuftragsende = str_SQL


End Function

Private Function buildSQLRest(t_case As String, t_econfiBtn As String, Optional modus As Integer, Optional searchText As String, Optional filterCol As String) As String

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
    str_SQL = str_SQL & "[SB_Status] As Job_Status, " & vbCrLf
    str_SQL = str_SQL & "TypeConfi ," & vbCrLf
    
    Select Case t_case
    
        Case "Rest1"
            
            str_SQL = str_SQL & "YearEnd ," & vbCrLf
            str_SQL = str_SQL & "EngCode ," & vbCrLf
            str_SQL = str_SQL & "EngPartner ," & vbCrLf
            str_SQL = str_SQL & "EngManager ," & vbCrLf
            str_SQL = str_SQL & "EngContact ," & vbCrLf
            str_SQL = str_SQL & "OtherContact ," & vbCrLf
            str_SQL = str_SQL & "TypeFinancials ," & vbCrLf
            str_SQL = str_SQL & "Lang ," & vbCrLf
            str_SQL = str_SQL & "EYStand ," & vbCrLf
            str_SQL = str_SQL & "BelongingToCompanyGroup " & vbCrLf

            
        Case "Rest2"
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

       
    End Select


    str_SQL = str_SQL & "From tCon_Orderbook" & strTestTable & vbCrLf
    
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
    
    buildSQLRest = str_SQL


End Function

Private Sub ExportRecordset(ByVal rs As ADODB.RecordSet, t_case As String)

    Dim oExcel            As Object
    Dim oExcelWrkBk       As Object
    Dim oExcelWrkSht      As Object
    Dim oExcelWrkShtFinal As Object

    'Start Excel
    Set oExcel = CreateObject("Excel.Application")
    
    ' oExcel.Calculation = xlCalculationManual
    oExcel.DisplayStatusBar = False
    oExcel.EnableEvents = False
    oExcel.ScreenUpdating = False
    oExcel.Visible = False
    oExcel.displayalerts = False
    
    ' Start a new workbook
    Set oExcelWrkBk = oExcel.Workbooks.Add()
    Set oExcelWrkSht = oExcelWrkBk.Sheets(1)
    
    Dim i As Integer, pos As Integer
    Dim lastrow As Long

    ' Rename worksheet
    Select Case t_case
    
        Case "ExportAll"
            oExcelWrkSht.name = "Export all data"
            pos = 2
            
        Case "Versand"
            oExcelWrkSht.name = "Export Versand"
            pos = 1
            
        Case "Auftragsende"
            oExcelWrkSht.name = "Export Auftragsende"
            pos = 1
            
    End Select

    With rs
        If .RecordCount <> 0 Then
            .MoveFirst
            
            ' Output Header
            For i = 0 To rs.Fields.Count - 1
                '#TODO MAYBE
                If rs.Fields(i).name = "CAD_Bearbeiter" Then
                    oExcelWrkSht.Cells(pos, 1 + i).value = "1. CAD Bearbeiter"
                    
                ElseIf rs.Fields(i).name = "CAD_Bearbeiter_vorherig" Then
                    oExcelWrkSht.Cells(pos, 1 + i).value = "2. CAD Bearbeiter"

                ElseIf rs.Fields(i).name = "Orderbook.ID" Then
                    oExcelWrkSht.Cells(pos, 1 + i).value = "ID"
                    
                ElseIf rs.Fields(i).name = "Orderbook.OrderNo" Then
                    oExcelWrkSht.Cells(pos, 1 + i).value = "OrderNo"
                    
                ElseIf rs.Fields(i).name = "Anzahl_Bestätigungen" Then
                    oExcelWrkSht.Cells(pos, 1 + i).value = "Anzahl_Anschreiben"
                    
                Else
                    oExcelWrkSht.Cells(pos, 1 + i).value = rs.Fields(i).name
                
                End If

            Next i

            ' Rearrange "ExportAll" Sheet
            If t_case = "ExportAll" Then
            
                ' Copy the data from our query into Excel
                oExcelWrkSht.Cells(3, 1).CopyFromRecordset rs
            
                Dim b_deleted As Boolean
            
                i = 1

                Do While oExcelWrkSht.Cells(2, i) <> ""
                    
                    b_deleted = False

                    If (oExcelWrkSht.Cells(2, i) = "Einladung_Canvas_erhalten" Or oExcelWrkSht.Cells(2, i) = "Datum_erster_Upload" Or oExcelWrkSht.Cells(2, i) = "CAD_Reviewer" Or oExcelWrkSht.Cells(2, i) = "Auswahl_erhalten" Or oExcelWrkSht.Cells(2, i) = "Auswahl_erhalten" Or oExcelWrkSht.Cells(2, i) = "Datum_Auswahl_erhalten" Or oExcelWrkSht.Cells(2, i) = "Stunden_Bearbeiter") Then
                        ' Delete unused columns
                        oExcelWrkSht.Columns(i).EntireColumn.Delete
                        b_deleted = True
                    End If
                
                    If InStr(oExcelWrkSht.Cells(2, i).value, "ConfiClient") <> 0 Then
                        ' Delete unused columns
                        oExcelWrkSht.Columns(i).EntireColumn.Delete
                        b_deleted = True
                    End If

                    If b_deleted = True Then i = i - 1
                    If b_deleted = False Then i = i + 1
            
                Loop
            
                'For i = 1 To rs.Fields.Count + 1
                    'If oExcelWrkSht.Cells(2, i) <> "" Then oExcelWrkSht.Cells(1, i) = i
                'Next i
            
            
                ' Rearrange Header sequence
                'Call RearrangeColumns(1, oExcel)
                
                'Set oExcelWrkShtFinal = oExcelWrkBk.Sheets("Export all data")
                oExcelWrkSht.Rows(1).EntireRow.Delete
                lastrow = oExcelWrkSht.UsedRange.Rows(oExcelWrkSht.UsedRange.Rows.Count).Row
                'oExcelWrkSht.Delete

            End If
            
            'Fit the columns to the content
            Dim str_colWidth As String
            str_colWidth = "33|20,86|20|11,86|9,14|10,86|25,86|25,29|11,14|9,86|7,71|14,71|11,43|18,43|12,71|12,71|10,71|14,71|30|6,71|13,14|26,29|23,29|26,29|8,43|12,14|12,14|12,57|10,71|12|14|11,86|13|9,71|12|9,14|6,14|13,86|12|13|18|7,14|7,29|12,29|8,14|16|15,14|10,14|16,57|16,14|10,14|10,14|8,57|20,29|16,86|13,29|12,86|8,43"
            'str_colWidth = "33|17|7,29|8,57|7,43|10,86|11,86|19,29|19,29|10,71|14,43|7,86|11,43|11,57|13,86|12,71|25,57|7,43|11,71|11,43|7|23,43|12,43|12,43|8,43|12,14|12,14|12,57|7,57|12|13,14|13,14|15|14|11,86|13|9,71|7,57|9,14|6,14|8,29|8,71|8,71|18|7,14|7,29|14,57|5,43|16|15,14|10,14|10,14|10,14|10,14|10,14|8,57|16,57|16,86|10,29|12,86|13,86"
            
            Dim t_arr() As String
            t_arr = Split(str_colWidth, "|")
            

            ' AutoFilter
            If t_case = "ExportAll" Then
                oExcelWrkSht.Range("A1:BM" & 1).AutoFilter
                oExcelWrkSht.Range("A1:BM" & lastrow).WrapText = False
                ' Start at the top
                oExcelWrkSht.Activate
                oExcelWrkSht.Cells(1, 1).Select
                
                For i = 0 To UBound(t_arr) - 1
                    oExcelWrkSht.Columns(i + 1).ColumnWidth = CDbl(t_arr(i))
                Next i

                oExcelWrkSht.Rows("1:1").WrapText = True
                oExcelWrkSht.Rows("1:1").RowHeight = 63.75
                
                oExcelWrkSht.Rows("1:1").Select
                
                With oExcel.Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .TintAndShade = -0.249977111117893
                    .PatternTintAndShade = 0
                End With
                
                'oExcelWrkShtFinal.Cells(2, 9).Select
                'oExcel.ActiveWindow.FreezePanes = True
                
                With oExcel.ActiveWindow
                    .SplitColumn = 9
                    .SplitRow = 1
                    .FreezePanes = True
                End With
                
                'Set oExcelWrkSht = Nothing

            Else
                ' Copy the data from our query into Excel
                oExcelWrkSht.Cells(2, 1).CopyFromRecordset rs

                lastrow = oExcelWrkSht.UsedRange.Rows(oExcelWrkSht.UsedRange.Rows.Count).Row
                
                If t_case = "Versand" Then oExcelWrkSht.Range("A1:X" & 1).AutoFilter
                If t_case = "Auftragsende" Then oExcelWrkSht.Range("A1:U" & 1).AutoFilter
                
                oExcelWrkSht.Range("A1:X" & lastrow).WrapText = False
                ' Start at the top
                oExcelWrkSht.Activate
                oExcelWrkSht.Cells(1, 1).Select
                
                
                For i = 0 To UBound(t_arr) - 1
                    oExcel.Columns(i + 1).ColumnWidth = CDbl(t_arr(i))
                Next i
                
                oExcelWrkSht.Rows("1:1").WrapText = True
                oExcelWrkSht.Rows("1:1").RowHeight = 63.75
                
                oExcelWrkSht.Rows("1:1").Select
                
                With oExcel.Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .TintAndShade = -0.249977111117893
                    .PatternTintAndShade = 0
                End With

                oExcelWrkSht.Range("H1").Select
                oExcel.ActiveWindow.FreezePanes = True
                
            End If

            oExcel.Calculation = xlCalculationAutomatic
            oExcel.DisplayStatusBar = True
            oExcel.EnableEvents = True
            oExcel.ScreenUpdating = True
            oExcel.Visible = True
            oExcel.displayalerts = True
            
            Set oExcel = Nothing
            Set oExcelWrkBk = Nothing
            Set oExcelWrkSht = Nothing

        End If
    End With


End Sub

Private Sub btnStammdaten_Click()
1         'On Error GoTo ErrorMsg
          Dim OrderNo As String, confirmation As String, bEYstand As String
          Dim RSaux As ADODB.RecordSet, c_rows As Integer

2         OrderNo = Me.Alltasks.Column(0)
3         confirmation = getTableValue(OrderNo, "Confirmation")
          bEYstand = getTableValue(OrderNo, "EYStand")
          
'4         Set Me.ListStammdaten.Recordset = getRS("Select OrderNo, LGISID, BuchKrGISID, BuchKr, BuchKrName from OrderbookSub Where OrderNo='" & OrderNo & "'")
'        If Confirmation = "BankenPortal" Then
'             Set RSaux = getRS("SELECT * FROM tCON_BankConfirmation WHERE OrderNo = '" & OrderNo & "'")
'             c_rows = 0
'             Do While RSaux.EOF = False
'                c_rows = c_rows + 1
'                RSaux.MoveNext
'             Loop
'             RSaux.Close
'             If c_rows = 0 Then
'                Me.ListStammdaten.ColumnCount = 11
'                Set Me.ListStammdaten.RecordSet = getRS("SELECT OrderNo, GISID, EngName, EngPartner, EngManager, OtherContact, BusinessEstablishmentOAre, TypeFinancials, CAD_Reviewer, TypeConfi, ConfiMailSent from tCON_Orderbook WHERE OrderNo = '" & OrderNo & "'")
'             Else
'                Me.ListStammdaten.ColumnCount = 13
'                Set Me.ListStammdaten.RecordSet = getRS("SELECT Orderbook.OrderNo, Orderbook.GISID, Orderbook.EngName, Orderbook.EngPartner, Orderbook.EngManager, Orderbook.OtherContact, Orderbook.BusinessEstablishmentOAre, Orderbook.TypeFinancials, Orderbook.CAD_Reviewer, Orderbook.TypeConfi, Orderbook.ConfiMailSent, BankConfirmation.Name_Bank, BankConfirmation.Portal_Bank from tCON_Orderbook as Orderbook INNER JOIN tCON_BankConfirmation as BankConfirmation ON Orderbook.OrderNo = BankConfirmation.OrderNo WHERE Orderbook.OrderNo = '" & OrderNo & "'")
'            End If
'        ElseIf (InStr(Confirmation, "Kreditor") > 0 Or InStr(Confirmation, "Debitor")) And bEYstand = "True" Then
'            Set RSaux = getRS("SELECT * FROM tCON_ClientContact WHERE OrderNo = '" & OrderNo & "'")
'             c_rows = 0
'             Do While RSaux.EOF = False
'                c_rows = c_rows + 1
'                RSaux.MoveNext
'             Loop
'             RSaux.Close
'            If c_rows = 0 Then
'                Me.ListStammdaten.ColumnCount = 14
'                Set Me.ListStammdaten.RecordSet = getRS("SELECT OrderNo, GISID, EngName, EngPartner, EngManager, OtherContact, BusinessEstablishmentOAre, TypeFinancials, CAD_Reviewer, TypeConfi, EYStand, Lang As Sprache, ConfiMailSent, AuditiSent As eConfirmationsSent from tCON_Orderbook WHERE OrderNo = '" & OrderNo & "'")
'             Else
'                Me.ListStammdaten.ColumnCount = 17
'                Set Me.ListStammdaten.RecordSet = getRS("SELECT Orderbook.OrderNo, Orderbook.GISID, Orderbook.EngName, Orderbook.EngPartner, Orderbook.EngManager, Orderbook.OtherContact, Orderbook.BusinessEstablishmentOAre, Orderbook.TypeFinancials, Orderbook.CAD_Reviewer, Orderbook.TypeConfi, Orderbook.EYStand, Orderbook.Lang As Sprache, Orderbook.ConfiMailSent, Orderbook.AuditiSent As eConfirmationsSent, ConfiClientContact.Mandant_Anr, ConfiClientContact.Mandant_Name, ConfiClientContact.Mandant_Email from tCON_Orderbook  as Orderbook INNER JOIN tCON_ClientContact as ConfiClientContact ON Orderbook.OrderNo = ConfiClientContact.OrderNo WHERE Orderbook.OrderNo = '" & OrderNo & "'")
'            End If
'        ElseIf InStr(Confirmation, "Vorräte") > 0 Then
'            Me.ListStammdaten.ColumnCount = 12
'            Set Me.ListStammdaten.RecordSet = getRS("SELECT OrderNo, GISID, EngName, EngPartner, EngManager, OtherContact, BusinessEstablishmentOAre, TypeFinancials, Date2ndReq, CAD_Reviewer, TypeConfi, ConfiMailSent from tCON_Orderbook WHERE OrderNo = '" & OrderNo & "'")
'        ElseIf Confirmation = "Sonstige" Then
'            Me.ListStammdaten.ColumnCount = 13
'            Set Me.ListStammdaten.RecordSet = getRS("SELECT OrderNo, GISID, EngName, EngPartner, EngManager, OtherContact, BusinessEstablishmentOAre, TypeFinancials, Date2ndReq, OtherTypeConfi, CAD_Reviewer, TypeConfi, ConfiMailSent from tCON_Orderbook WHERE OrderNo = '" & OrderNo & "'")
'        Else
'            Me.ListStammdaten.ColumnCount = 13
'            Set Me.ListStammdaten.RecordSet = getRS("SELECT OrderNo, GISID, EngName, EngPartner, EngManager, OtherContact, BusinessEstablishmentOAre, TypeFinancials, CAD_Reviewer, TypeConfi, EYStand, Lang As Sprache, ConfiMailSent from tCON_Orderbook WHERE OrderNo = '" & OrderNo & "'")
'        End If

        
        Me.ListStammdaten.ColumnCount = 14
        Set Me.ListStammdaten.RecordSet = getRS("SELECT Orderbook.OrderNo, Orderbook.EngName, Orderbook.EngManager, Orderbook.OtherContact, Orderbook.TypeFinancials, Orderbook.CAD_Reviewer, Orderbook.Lang As Sprache, Orderbook.ConfiMailSent, Orderbook.AuditiSent As eConfirmationsSent, ConfiClientContact.Mandant_Anr, ConfiClientContact.Mandant_Name, ConfiClientContact.Mandant_Email,BankConfirmation.Name_Bank, BankConfirmation.Portal_Bank From tCon_Orderbook" & strTestTable & "  as Orderbook LEFT JOIN tCON_ClientContact" & strTestTable & " as ConfiClientContact ON Orderbook.OrderNo = ConfiClientContact.OrderNo left JOIN tCON_BankConfirmation" & strTestTable & " as BankConfirmation ON Orderbook.OrderNo = BankConfirmation.OrderNo WHERE Orderbook.OrderNo = '" & OrderNo & "'")
        Me.ListStammdaten.columnWidths = "3cm;3.5cm;4.5cm;4.5cm;1.5cm;3cm;1.5cm;3.5cm;3.5cm;2cm;3cm;3.5cm;4cm;4cm"

        ListStammdaten.Visible = True
        
6         On Error GoTo 0
7         Exit Sub

Errormsg:
8         If CStr(Err.Number) = "3044" Then
9             MsgBox ("No connection to the server. Please check the connection and try again")
10        ElseIf CStr(Err.Number) = "94" Then
11            MsgBox ("Please select an order")
12        Else
13            MsgBox ("An error is occured. Please restart the client. Error info: " & Err.Description)
14        End If
End Sub

Private Sub BtnEdit_Click()
DoCmd.OpenForm "EditForm"
End Sub
