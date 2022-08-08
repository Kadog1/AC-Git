VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_AC_Confi Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AV_Click()

    'Beim Klicken: erstellt die Folder & öffnet das mailtemaplte

    Dim path_2 As String
    Dim path_3 As String
    Dim pathConfirmationsWorkbook As String
    Dim tempname As String
    Dim OrderNo As String, docLang As String
    Dim period As String

    'On Error GoTo Errormsg

    OrderNo = Me.Alltasks.Column(0)
    docLang = Me.Alltasks.Column(8)
    
    If docLang = "DE" Then
        docLang = ""
    Else
        docLang = "_" & docLang
    End If

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
    
    pathConfirmationsWorkbook = path_2 & "*CAD Confirmations Workbook*.xlsm"
    If Len(Dir(pathConfirmationsWorkbook, vbDirectory)) > 0 Then
        answer = MsgBox("Es liegt bereits ein CAD Confirmations Workbook in ...\C Workplace\" & OrderNo & "\2. CAD_Abgleich\ ab. Möchten Sie dieses überschreiben?", vbQuestion + vbYesNo + vbDefaultButton2, "AC_Workbook öffnen")
        If Not answer = vbYes Then
            Exit Sub
        End If
    End If
    
    'copy template
    tempname = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\2_CAD-Adressabgleich Workbook Template" & docLang & ".xlsm"
    'tempname = "C:\TestLocal\AC adressabgliech\FIS maill project\2_CAD-Adressabgleich Workbook Template" & docLang & ".xlsm"
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile tempname, path_2
    '# New Feature - Run Excel Input import directly
    
    Dim str_fileName As String
    str_fileName = Dir$(path_2 & "*1_CAD-*" & ".xls*")
    
    If Len(str_fileName) > 0 Then

        Set XL = CreateObject("Excel.Application")
        Set wb = XL.Workbooks.Open(path_2 & "2_CAD-Adressabgleich Workbook Template" & docLang & ".xlsm")
        XL.Visible = True
        
        XL.Run "LoadInput_Click", path_2, str_fileName

    Else
    
        MsgBox "No Inputdata Workbook found! Please check folder manually!"
    
    End If
    
    Exit Sub


Errormsg:
    If CStr(Err.Number) = "3044" Then
        MsgBox ("No connection to the server. Please check the connection and try again")
    ElseIf CStr(Err.Number) = "94" Then
        MsgBox ("Please select an order")
        
    Else
        MsgBox ("Error processing documentation file.")
    
    
    End If


End Sub

Private Sub Btn_openFolder_Click()

    Dim FolderName As String, FolderOrderNo As String
    
    FolderOrderNo = Me.Alltasks.Column(0)

    FolderName = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & FolderOrderNo
    
    If Len(Dir("\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & FolderOrderNo, vbDirectory)) = 0 Then
        MkDir "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & FolderOrderNo
    End If

    Shell "C:\WINDOWS\explorer.exe """ & FolderName & "", vbNormalFocus

End Sub

Private Sub btnGAS_Click()

DoCmd.OpenForm "Update_Confi Client"
DoCmd.Close acForm, "AC_Confi Client", acSaveYes

End Sub

Private Sub Command102_Click()
Dim ctlList As Control, varItem As Variant, selOrderNo, selClient, selConfirmation, selEngCntct, selEngMngr, selOtherCntct, selYearEnd As String
    
    Set ctlList = Forms![AC_Confi Client].Alltasks
    
    For Each varItem In ctlList.ItemsSelected
        selOrderNo = selOrderNo & "," & "'" & ctlList.Column(0, varItem) & "'"
    Next varItem

    If selOrderNo = "" Then
        MsgBox "Please select an order to proceed."
        Exit Sub
    End If
    selOrderNo = Right(selOrderNo, Len(selOrderNo) - 1)
    Dim rsSQL As ADODB.RecordSet
    Set rsSQL = getRS("SELECT OrderNo, Client, Confirmation, EngContact, EngManager, YearEnd, GISID, DocLang, OtherContact FROM [CAD].[dbo].[tCON_Orderbook] WHERE OrderNo IN (" & selOrderNo & ")")

    Dim arr_rsSQL As Variant
    arr_rsSQL = rsSQL.GetRows
    
    For i = 0 To UBound(arr_rsSQL, 2) - LBound(arr_rsSQL, 2)
        Dim subject As String
        subject = "CAD Adressabgleich Bestellung: [Confirmation] [OrderNo] für [Client]"
        subject = Replace(subject, "[OrderNo]", arr_rsSQL(0, i))
        subject = Replace(subject, "[Client]", arr_rsSQL(1, i))
        subject = Replace(subject, "[Confirmation]", arr_rsSQL(2, i))
        
        Dim language As String
        language = "DE"
        If arr_rsSQL(7, i) = "EN" Then language = "EN"
        
        Dim templatePath As String
        'templatePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\DE\4_Team_Approval_Template_DE.htm"
        templatePath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\B Mail Templates\" & language & "\eCon_AC_ConfiMail.htm"
        
        Dim body As String
        Dim attPath As String
        attPath = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & arr_rsSQL(0, i) & "\1. Adressenabgleich\" & Format(arr_rsSQL(6, 0), "0000000000") & " 1_CAD-Adressabgleich Adressenabfrage Mandant " & Format(arr_rsSQL(5, i), "yyyyMMdd") & ".xlsm"
        If Dir(attPath, vbDirectory) = "" Then
            MsgBox "Bitte zuerst AC_Input erstellen."
            Exit Sub
        End If
        
        'Vorgefertigter Text
        body = parseBody(templatePath)
        body = Replace(body, "[OrderNo]", arr_rsSQL(0, i))
        body = Replace(body, "[Client]", arr_rsSQL(1, i))
        body = Replace(body, "[Confirmation]", arr_rsSQL(2, i))
        body = Replace(body, "[GISID]", arr_rsSQL(6, i))
        
        Dim engCntct As String, engCC As String
        
        engCntct = arr_rsSQL(3, i)
        engCC = arr_rsSQL(4, i) & ";" & Nz(arr_rsSQL(8, i), "")
        
        'create Email Draft
        Call createEmailDraft(attPath, body, subject, engCntct, engCC)
    Next i

End Sub

Private Sub Command97_Click()

    Dim ID As String
    Dim confirmation As String, confirmation2 As String, lang As String, OrderNo As String, client As String, BUKRS As String, yearend As String, typeFin As String, dateconfi As String, engpartner As String, engmanager As String, engcontact As String, engother As String
    Dim planneddelivery As String, deliverytype As String, date2ndrequest As String, otherconfitype As String, gisid As String, typeconfi As String, name As String
    

    Dim ctlList As Control, varItem As Variant, selOrderNo As String
    
    Set ctlList = Forms![AC_Confi Client].Alltasks
    
    For Each varItem In ctlList.ItemsSelected
        selOrderNo = selOrderNo & "," & "'" & ctlList.Column(0, varItem) & "'"
    Next varItem

    If selOrderNo = "" Then
        MsgBox "Please select an order to proceed."
        Exit Sub
    End If
    selOrderNo = Right(selOrderNo, Len(selOrderNo) - 1)
    Dim rsSQL As ADODB.RecordSet
    Set rsSQL = getRS("SELECT OrderNo, Client, Confirmation, EngCode, EngContact, EngManager, EngPartner, YearEnd, GISID, DocLang, OtherContact, DateConfi, BusinessEstablishmentOAre, ADDITIONAL_CONTACT_2, ADDITIONAL_CONTACT_3, ADDITIONAL_CONTACT_4, ADDITIONAL_CONTACT_5 FROM [CAD].[dbo].[tCON_Orderbook] WHERE OrderNo IN (" & selOrderNo & ")")
    
    Dim arr_rsSQL As Variant
    arr_rsSQL = rsSQL.GetRows

    Dim pathTemplates As String
    
    Dim folder As String
    Dim OutMail As Object
    Dim pathMailTemplates As String, templateName As String, wsTemplateName As String
    Dim body As String, subject As String
    Dim attachement1 As String, attachement2 As String, attachement3 As String
    
    Dim docLang As String
    pathMailTemplates = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\"
        
    For i = 0 To UBound(arr_rsSQL, 2) - LBound(arr_rsSQL, 2)
        OrderNo = arr_rsSQL(0, i)
        client = arr_rsSQL(1, i)
        yearend = arr_rsSQL(7, i)
        yearend = Format(CDate(yearend), "short date")
        docLang = arr_rsSQL(9, i)
        dateconfi = arr_rsSQL(11, i)
        dateconfi = Format(CDate(dateconfi), "short date")
        engpartner = arr_rsSQL(6, i)
        engmanager = arr_rsSQL(5, i)
        engcontact = arr_rsSQL(4, i)
        gisid = arr_rsSQL(8, i)

        engother = ""
        BUKRS = ""
        confirmation = ""
        If IsNull(arr_rsSQL(10, i)) = False Then engother = arr_rsSQL(10, i)
        If IsNull(arr_rsSQL(12, i)) = False Then BUKRS = arr_rsSQL(12, i)
        If IsNull(arr_rsSQL(2, i)) = False Then confirmation = arr_rsSQL(2, i)
        
        Dim additionalCntct As String
        additionalCntct = ""
        If Not IsNull(arr_rsSQL(13, i)) Then additionalCntct = additionalCntct & ";" & arr_rsSQL(13, i)
        If Not IsNull(arr_rsSQL(14, i)) Then additionalCntct = additionalCntct & ";" & arr_rsSQL(14, i)
        If Not IsNull(arr_rsSQL(15, i)) Then additionalCntct = additionalCntct & ";" & arr_rsSQL(15, i)
        If Not IsNull(arr_rsSQL(16, i)) Then additionalCntct = additionalCntct & ";" & arr_rsSQL(16, i)

        Dim wbTemplate As Object, wsTemplate As Object
        
        If docLang = "DE" Then
            docLang = ""
        Else
            docLang = "_" & docLang
        End If
        
        If InStr(OrderNo, "CON") > 0 Then
            pathTemplates = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\1_CAD-Adressabgleich Adressenabfrage Mandant" & docLang & ".xlsm"
        Else
        End If

        pathTemplates = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\D Dokumentation Templates\1_CAD-Adressabgleich Adressenabfrage Mandant" & docLang & ".xlsm"
        Set excelObj = CreateObject("Excel.Application")
        Set wbTemplate = excelObj.Workbooks.Open(pathTemplates)
        Set wsTemplate = wbTemplate.Worksheets("basic_info")

        wsTemplate.Cells(1, 2) = OrderNo
        wsTemplate.Cells(2, 2) = yearend
        wsTemplate.Cells(3, 2) = client
        wsTemplate.Cells(4, 2) = engcontact
        wsTemplate.Cells(5, 2) = engpartner
        wsTemplate.Cells(6, 2) = engmanager
        wsTemplate.Cells(7, 2) = engother
        wsTemplate.Cells(8, 2) = gisid
        wsTemplate.Cells(9, 2) = confirmation
        wsTemplate.Cells(10, 2) = arr_rsSQL(3, i)
          
        folder = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & OrderNo

        If Dir(folder, vbDirectory) = "" Then MkDir (folder)
        If Dir(folder & "\1. Adressenabgleich\", vbDirectory) = "" Then MkDir (folder & "\1. Adressenabgleich\")
        
        wbTemplate.SaveAs FileName:=folder & "\1. Adressenabgleich\" & Format(gisid, "0000000000") & " 1_CAD-Adressabgleich Adressenabfrage Mandant " & Format(yearend, "yyyyMMdd") & ".xlsm"

        wbTemplate.Close
        
        Set excelObj = Nothing
    Next i

End Sub

Private Sub Command98_Click()


    'Check if AC_Status = 'TeamApprovalReceived'
    
    Dim ctlList As Control, varItem As Variant, selOrderNo As String
    
    Set ctlList = Forms![AC_Confi Client].Alltasks
    
    For Each varItem In ctlList.ItemsSelected
    
        If ctlList.Column(11, varItem) = "TeamApprovalReceived" Then
            selOrderNo = selOrderNo & "|" & ctlList.Column(0, varItem)
        End If

    Next varItem
    
    If selOrderNo = "" Then
        MsgBox "Please select an order with AC_Status = TeamApprovalReceived to proceed"
        Exit Sub
    End If
    
    selOrderNo = Right(selOrderNo, Len(selOrderNo) - 1)
    
    Dim arr_Orders() As String
    arr_Orders = Split(selOrderNo, "|")

    'Check if .msg file exists
    Dim Logdir As String, wb_Documentation As String

    Dim pathMSG As String, pathWB As String
    
    Dim i As Integer
    
    For i = 0 To UBound(arr_Orders)
    
        Logdir = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & arr_Orders(i) & "\3. Team Approval\"
        wb_Documentation = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & arr_Orders(i) & "\2. CAD_Abgleich\"
    
        'Check if .msg File Exists
        pathMSG = ""
        pathMSG = fileExists(Logdir, ".msg")
        
        pathWB = ""
        pathWB = fileExists(wb_Documentation, "CAD Confirmations Workbook")
    
        If pathMSG <> "" And pathWB <> "" Then

            Dim obj_excel As Object, wb As Object
            Set obj_excel = CreateObject("Excel.Application")
            With obj_excel
                .Visible = True
                .displayalerts = False
                .Workbooks.Open (pathWB)
                .Run "process_TeamApproval", pathMSG, CStr(arr_Orders(i))
                    
                .ActiveWorkbook.Close (True)
                '.Quite
            End With
                
            Set obj_excel = Nothing

            'Call Excel Macro to embed .msg file

            'Close procedure
            
            'Update OrderNo in DB
        Else
        
        
        End If
    

    Next i


End Sub
Private Function fileExists(Logdir As String, searchkey As String) As String

    'Check if 'searchkey' file exists
    Set fso = CreateObject("Scripting.FileSystemObject")
                    
    Set source = fso.GetFolder(Logdir)
    Dim FileName As String

    For Each File In source.Files
        
        FileName = File.name
        Dim DestinationPathFile As String
        DestinationPathFile = ""
        DestinationPathFile = Logdir & "\" & FileName
                        
        If InStr(UCase(DestinationPathFile), searchkey) > 0 Then
        
            fileExists = DestinationPathFile
            
            Exit Function

        End If
    Next


End Function
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
    
    'Me.Next_Page.Visible = True
    'Me.Previous_Page.Visible = True
    
    Me.Label93.Caption = "Adressen1"

    If Me!btneConfi.Caption = "eConfirmations einblenden" Then

        AlltasksQuery = buildSQLAdressen(Me.Label93.Caption, "einblenden")

    Else
        AlltasksQuery = buildSQLAdressen(Me.Label93.Caption, "ausblenden")
    End If

    Set Me.Alltasks.RecordSet = getRS(AlltasksQuery)
    
    Call adjustAdressen(CStr(Me.Label93.Caption))
    
    Me.ShowStorno.Caption = "Storno/Closed ausblenden"
    
    Call setListbox
    'Call setUserListbox

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

Private Sub setListbox()

        Dim arrResult() As Variant, arrAlltasks() As Variant
        Me.ListStatus.RowSource = ""

        'Set rs = getRS(AlltasksQuery)

        arrAlltasks = Me.Alltasks.RecordSet.GetRows

        ReDim arrResult(0 To 2, 0 To 23)
    
        For i = 0 To UBound(arrAlltasks, 2)
            If IsNull(arrAlltasks(13, i)) Then arrResult(0, 1) = arrResult(0, 1) + 1
        Next i
    
        arrResult(0, 0) = "[Blank]"
        arrResult(0, 1) = "InputDataSent"
        arrResult(0, 2) = "WaitingForInputData"
        arrResult(0, 3) = "InputDataReceived"
        arrResult(0, 4) = "InProgress"
        arrResult(0, 5) = "ReadyForReview"
        arrResult(0, 6) = "TeamApprovalSent"
        arrResult(0, 7) = "TeamApprovalReceived"
        arrResult(0, 8) = "CanvasDone"
        arrResult(0, 9) = "Finalized"
        arrResult(0, 10) = "Storno"
        arrResult(0, 11) = "CADStorno"
        arrResult(0, 12) = "OnHold"
        
        Set RSFiltered = Me.Alltasks.RecordSet
        Dim RSFilteredCount As Integer
        
        RSFiltered.Filter = ("[AC_Status] = NULL")
        strListStatus = strListStatus & arrResult(0, 0) & ";" & RSFiltered.RecordCount & ";"
        For i = 1 To 11
            RSFiltered.Filter = ("[AC_Status] = " & "'" & arrResult(0, i) & "'")
            strListStatus = strListStatus & arrResult(0, i) & ";" & RSFiltered.RecordCount & ";"
        Next i
        Me.ListStatus.RowSource = strListStatus
        Me.ListStatus.columnWidths = "3cm;0.7cm"
        
        
        Call setUserListbox(arrAlltasks)

End Sub
Private Sub ListStatus_Click()
   ' On Error GoTo Errormsg
    
    Call Refresh_Click
    Dim FilterValue As String, RSFiltered As ADODB.RecordSet

    FilterValue = Me.ListStatus.Column(0)

    Set RSFiltered = Me.Alltasks.RecordSet

    'getRS (AlltasksQuery)

    If UCase(FilterValue) = "[BLANK]" Then
        RSFiltered.Filter = ("[AC_Status] = NULL")
    Else
        RSFiltered.Filter = ("[AC_Status] = " & "'" & FilterValue & "'")
    End If

    Set Me.Alltasks.RecordSet = RSFiltered

    'Call unselectLists
    
    Exit Sub
    
Errormsg:
    If CStr(Err.Number) = "3044" Then
        MsgBox ("No connection to the server. Please check the connection and try again")
    ElseIf CStr(Err.Number) = "94" Then
        MsgBox ("Please select an order")
    Else
        MsgBox ("An error is occured. Please restart the client. Error Info: " & Err.Description)
    End If

End Sub

'#JK-2/2020-12-16 START
Private Sub btn_OrderDetail_Click()

    On Error GoTo Errormsg:
    DoCmd.OpenForm "OrderDetail", OpenArgs:=Me.Alltasks.Column(0)
    Forms![OrderDetail].SetFocus
3         Exit Sub
Errormsg:
4         MsgBox ("Please select an OrderNo in the Mainlist and try again.")


End Sub
'#JK-2/2020-12-16 END


' ######### BUTTON START ##########

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

                Call RowsSelectedAC("", header)
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

                Call RowsSelectedAC(Me!TxtEntry.value, header)
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
                Call RowsSelectedAC("", header)
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
                Call RowsSelectedAC(Me!cbEntry.value, header)
            End If
        End If
    End If
    
    Call Refresh_Click

End Sub


Sub unselectLists()

            Me.ListStatus.Selected(Me.ListStatus.ListIndex) = False
End Sub
Private Sub setUserListbox(arrAlltasksUser As Variant)

        Dim arrResultUser() As Variant ', arrAlltasksUser() As Variant
        Me.lUserCAD.RowSource = ""
        
        'arrAlltasksUser = Me.Alltasks.RecordSet.GetRows

        ReDim arrResultUser(0 To 2, 0 To 23)
    
        For i = 0 To UBound(arrAlltasksUser, 2)
            If IsNull(arrAlltasksUser(12, i)) Then arrResultUser(0, 1) = arrResultUser(0, 1) + 1
            If arrAlltasksUser(13, i) = "Alexandra Staicu" Then arrResultUser(0, 3) = arrResultUser(0, 3) + 1
            If arrAlltasksUser(13, i) = "Tuan Anh Dang" Then arrResultUser(0, 5) = arrResultUser(0, 5) + 1
            If arrAlltasksUser(13, i) = "Chantal Berg" Then arrResultUser(0, 7) = arrResultUser(0, 7) + 1
            If arrAlltasksUser(13, i) = "Hassiba Ghobrini" Then arrResultUser(0, 9) = arrResultUser(0, 9) + 1
            If arrAlltasksUser(13, i) = "Marina Böhmer" Then arrResultUser(0, 11) = arrResultUser(0, 11) + 1
            If arrAlltasksUser(13, i) = "Saskia Dammann" Then arrResultUser(0, 13) = arrResultUser(0, 13) + 1
            If arrAlltasksUser(13, i) = "Wonyoung Lim" Then arrResultUser(0, 15) = arrResultUser(0, 15) + 1
            If arrAlltasksUser(13, i) = "Sonja Weinmann" Then arrResultUser(0, 17) = arrResultUser(0, 17) + 1
            If arrAlltasksUser(13, i) = "Rebecca Berg" Then arrResultUser(0, 19) = arrResultUser(0, 19) + 1
            If arrAlltasksUser(13, i) = "Mohamed Delfo" Then arrResultUser(0, 21) = arrResultUser(0, 21) + 1
            If arrAlltasksUser(13, i) = "Joanna Eleonora Ramuschkat" Then arrResultUser(0, 23) = arrResultUser(0, 23) + 1
        Next i
    
        arrResultUser(0, 0) = "[Blank]"
        arrResultUser(0, 1) = "Alexandra Staicu"
        arrResultUser(0, 2) = "Tuan Anh Dang"
        arrResultUser(0, 3) = "Chantal Berg"
        arrResultUser(0, 4) = "Hassiba Ghobrini"
        arrResultUser(0, 5) = "Marina Böhmer"
        arrResultUser(0, 6) = "Saskia Dammann"
        arrResultUser(0, 7) = "Wonyoung Lim"
        arrResultUser(0, 8) = "Sonja Weinmann"
        arrResultUser(0, 9) = "Rebecca Berg"
        arrResultUser(0, 10) = "Mohamed Delfo"
        arrResultUser(0, 11) = "Joanna Eleonora Ramuschkat"
        
        Set RSFiltered = Me.Alltasks.RecordSet
        Dim RSFilteredCount As Integer
        
        RSFiltered.Filter = ("[AC_Preparer] = NULL")
        strListStatusUser = strListStatusUser & arrResultUser(0, 0) & ";" & RSFiltered.RecordCount & ";"
        For i = 1 To 11
            RSFiltered.Filter = ("[AC_Preparer] = " & "'" & arrResultUser(0, i) & "'")
            strListStatusUser = strListStatusUser & arrResultUser(0, i) & ";" & RSFiltered.RecordCount & ";"
        Next i
        Me.lUserCAD.RowSource = strListStatusUser
        Me.lUserCAD.columnWidths = "3cm;0.7cm"

End Sub

Private Sub lUserCAD_Click()

    ' On Error GoTo Errormsg
    
    Call Refresh_Click
    Dim FilterValue As String, RSFiltered As ADODB.RecordSet

    FilterValue = Me.lUserCAD.Column(0)

    Set RSFiltered = Me.Alltasks.RecordSet

    'getRS (AlltasksQuery)

    If UCase(FilterValue) = "[BLANK]" Then
        RSFiltered.Filter = ("[AC_Preparer] = NULL")
    Else
        RSFiltered.Filter = ("[AC_Preparer] = " & "'" & FilterValue & "'")
    End If

    Set Me.Alltasks.RecordSet = RSFiltered

    'Call unselectLists
    
    Exit Sub
    
Errormsg:
    If CStr(Err.Number) = "3044" Then
        MsgBox ("No connection to the server. Please check the connection and try again")
    ElseIf CStr(Err.Number) = "94" Then
        MsgBox ("Please select an order")
    Else
        MsgBox ("An error is occured. Please restart the client. Error Info: " & Err.Description)
    End If

    
End Sub

Private Sub Refresh_Click()

    'On Error GoTo Errormsg

    If Me!btneConfi.Caption = "eConfirmations einblenden" Then

        AlltasksQuery = buildSQLAdressen(Me.Label93.Caption, "einblenden")

    Else

        AlltasksQuery = buildSQLAdressen(Me.Label93.Caption, "ausblenden")

    End If
    
    ' Debug.Print AlltasksQuery
    
    Set Me.Alltasks.RecordSet = getRS(AlltasksQuery)


        Call adjustAdressen(CStr(Me.Label93.Caption))
        '#JK-6/2020-12-16 START

    If Me.txtSearch <> "" Then Call btnSearch_Click
    
    Dim RSFiltered As ADODB.RecordSet

    Set RSFiltered = Me.Alltasks.RecordSet
    RSFiltered.Filter = Me.Alltasks.RecordSet.Filter

    If InStr(Me.ShowStorno.Caption, "einblenden") Then

        RSFiltered.Filter = ("([AC_Status] <> 'CanvasDone' AND [AC_Status] <> 'Storno') OR [AC_Status] = Null")
        Set Me.Alltasks.RecordSet = RSFiltered

    End If
    
    Call setListbox
    'Call setUserListbox
    'Call unselectLists

    On Error GoTo 0
    Exit Sub

Errormsg:
    If CStr(Err.Number) = "3044" Then
        MsgBox ("No connection to the server. Please check the connection and try again")
    Else
        MsgBox ("An error is occured. Please restart the client. Error info: " & Err.Description)
    End If

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

                      searchQuery = buildSQLAdressen(Me.Label93.Caption, "einblenden", 2, searchText, filterCol)

                  Else
                      searchQuery = buildSQLAdressen(Me.Label93.Caption, "ausblenden", 2, searchText, filterCol)
                  End If

              Else
              
                  If Me!btneConfi.Caption = "eConfirmations einblenden" Then
                      ' Start
                      searchQuery = buildSQLAdressen(Me.Label93.Caption, "einblenden", 1, searchText, filterCol)
                  Else

                      searchQuery = buildSQLAdressen(Me.Label93.Caption, "ausblenden", 1, searchText, filterCol)

                  End If
              End If
7             Set RSFiltered = getRS(searchQuery)
8             Set Me.Alltasks.RecordSet = RSFiltered


              Call adjustAdressen(CStr(Me.Label93.Caption))


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
    
        RSFiltered.Filter = ("([AC_Status] <> 'CanvasDone' AND [AC_Status] <> 'Storno') OR [AC_Status] = Null")
        
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

' ######### FUNCTION START ###########

Private Sub Alltasks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer

    Dim RSpreparer As ADODB.RecordSet

    Dim Spl As String
    Dim ColumnID As Integer, anzCol As Integer
    Dim width As Double, stepWidth As Double
    
    Call showFurtherDetails

    Dim columnWidths() As Double
    width = Alltasks.width - 20

    ReDim columnWidths(1 To 18)
    anzCol = 18
    stepWidth = Round(width / 45, 3)

    ' Gleichtverteilung der Spaltenbreiten wäre 3*stepWidth pro Spalte
    ' Wenn alles addiert 45 * stepWidth ergibt, stimmt die Breite

    columnWidths(1) = 3 * stepWidth 'OrderNo
    columnWidths(2) = 4 * stepWidth 'GroupName
    columnWidths(3) = 6 * stepWidth 'Client
    columnWidths(4) = 2 * stepWidth 'Tool
    columnWidths(5) = 2.5 * stepWidth 'Confirmation
    columnWidths(6) = 3 * stepWidth 'DeliveryType
    columnWidths(7) = 2 * stepWidth 'EngCode
    columnWidths(8) = 3.5 * stepWidth 'EngContact
    columnWidths(9) = 1.5 * stepWidth 'DokuLang
    columnWidths(10) = 2 * stepWidth 'Job_Status
    columnWidths(11) = 2.5 * stepWidth 'GAS_Preparer
    columnWidths(12) = 4.5 * stepWidth 'AC_STATUS
    columnWidths(13) = 4 * stepWidth 'AC_Preparer
    columnWidths(14) = 2.5 * stepWidth 'tsInputDataReceived


    Dim leftBound As Double, rightBound As Double
    leftBound = 0
    rightBound = 0

    For i = 0 To anzCol - 1
        rightBound = rightBound + columnWidths(i + 1)
        If i > 0 Then leftBound = leftBound + columnWidths(i)
        If leftBound < X And X < rightBound Then

            Select Case i:
                Case 0: Spl = "OrderNo"
                Case 1: Spl = "NameOfCompanyGroup"
                Case 2: Spl = "Client"
                Case 3: Spl = "Tool"
                Case 4: Spl = "Confirmation"
                Case 5: Spl = "DeliveryType"
                Case 6: Spl = "EngCode"
                Case 7: Spl = "EngContact"
                Case 8: Spl = "DocLang"
                Case 9: Spl = "Job_Status"
                Case 10: Spl = "1_CAD_Preparer"
                Case 11: Spl = "AC_Status"
                Case 12: Spl = "AC_Preparer"
                Case 13: Spl = "tsInputDataReceived"
                
            End Select
            '#JK-6/2020-12-09 START
            
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
             ElseIf Spl = "AC_Status" Then
                Me!TxtEntry.Visible = False
                Me!cbEntry.Visible = True
                Do While Me!cbEntry.ListCount > 0
                    Me!cbEntry.RemoveItem (0)
                Loop
                Me!cbEntry.AddItem "InputDataSent"
                Me!cbEntry.AddItem "WaitingForInputData"
                Me!cbEntry.AddItem "InProgress"
                Me!cbEntry.AddItem "ReadyForReview"
                Me!cbEntry.AddItem "TeamApprovalSent"
                Me!cbEntry.AddItem "TeamApprovalReceived"
                Me!cbEntry.AddItem "CanvasDone"
                Me!cbEntry.AddItem "Finalized"
                Me!cbEntry.AddItem "CADStorno"
                Me!cbEntry.AddItem "Storno"
                Me!cbEntry.AddItem "OnHold"
                
                Me!cbEntry.value = Me!Alltasks.Column(ColumnID, Me.Alltasks.ListIndex + 1)
             ElseIf Spl = "AC_Preparer" Then
                Me!TxtEntry.Visible = False
                Me!cbEntry.Visible = True
                Do While Me!cbEntry.ListCount > 0
                    Me!cbEntry.RemoveItem (0)
                Loop
                Me!cbEntry.AddItem "Alexandra Staicu"
                Me!cbEntry.AddItem "Tuan Anh Dang"
                Me!cbEntry.AddItem "Chantal Berg"
                Me!cbEntry.AddItem "Hassiba Ghobrini"
                Me!cbEntry.AddItem "Marina Böhmer"
                Me!cbEntry.AddItem "Saskia Dammann"
                Me!cbEntry.AddItem "Wonyoung Lim"
                Me!cbEntry.AddItem "Sonja Weinmann"
                Me!cbEntry.AddItem "Rebecca Berg"
                Me!cbEntry.AddItem "Mohamed Delfo"
                Me!cbEntry.AddItem "Joanna Eleonora Ramuschkat"
                
                Me!cbEntry.value = Me!Alltasks.Column(ColumnID, Me.Alltasks.ListIndex + 1)
                
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


'#JK-6/2020-12-09 START
Private Sub adjustAdressen(str_input As String)

    Dim t_width As Integer
    Dim width As Double, stepWidth As Double
    
    Dim columnWidths() As Double
    width = Alltasks.width
    
    If Me.Label93.Caption = "Adressen1" Then
        ReDim columnWidths(1 To 18)
        Alltasks.ColumnCount = 18
        t_width = 18
    End If

    width = Alltasks.width
    
    stepWidth = Round(width / 45, 3)
    ' Gleichtverteilung der Spaltenbreiten wäre 3*stepWidth pro Spalte
    ' Wenn alles addiert 45 * stepWidth ergibt, stimmt die Breite

    columnWidths(1) = 3 * stepWidth 'OrderNo
    columnWidths(2) = 4 * stepWidth 'GroupName
    columnWidths(3) = 6 * stepWidth 'Client
    columnWidths(4) = 2 * stepWidth 'Tool
    columnWidths(5) = 2.5 * stepWidth 'Confirmation
    columnWidths(6) = 3 * stepWidth 'DeliveryType
    columnWidths(7) = 2 * stepWidth 'EngCode
    columnWidths(8) = 3.5 * stepWidth 'EngContact
    columnWidths(9) = 1.5 * stepWidth 'DokuLang
    columnWidths(10) = 2 * stepWidth 'Job_Status
    columnWidths(11) = 2.5 * stepWidth 'GAS_Preparer
    columnWidths(12) = 4.5 * stepWidth 'AC_STATUS
    columnWidths(13) = 4 * stepWidth 'AC_Preparer
    columnWidths(14) = 2.5 * stepWidth 'tsInputDataReceived
    
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

Private Function buildSQLAdressen(t_case As String, t_econfiBtn As String, Optional modus As Integer, Optional searchText As String, Optional filterCol As String, Optional listOrder As String) As String

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
            'str_SQL = str_SQL & "PlannedDelivery, " & vbCrLf
            str_SQL = str_SQL & "DeliveryType, " & vbCrLf
            'str_SQL = str_SQL & "YearEnd," & vbCrLf
            str_SQL = str_SQL & "EngCode," & vbCrLf
            If modus = 3 Then str_SQL = str_SQL & "dateConfi," & vbCrLf
            If modus = 3 Then str_SQL = str_SQL & "EngPartner," & vbCrLf
            str_SQL = str_SQL & "EngContact," & vbCrLf
            'str_SQL = str_SQL & "[OtherContact], " & vbCrLf
            'str_SQL = str_SQL & "EngManager," & vbCrLf
            str_SQL = str_SQL & "DocLang," & vbCrLf
            If modus = 3 Then str_SQL = str_SQL & "GISID ," & vbCrLf
            If modus = 3 Then str_SQL = str_SQL & "OtherContact ," & vbCrLf
            If modus = 3 Then str_SQL = str_SQL & "BusinessEstablishmentOAre ," & vbCrLf
            If modus = 3 Then str_SQL = str_SQL & "ADDITIONAL_CONTACT_2 ," & vbCrLf
            If modus = 3 Then str_SQL = str_SQL & "ADDITIONAL_CONTACT_3 ," & vbCrLf
            If modus = 3 Then str_SQL = str_SQL & "ADDITIONAL_CONTACT_4 ," & vbCrLf
            If modus = 3 Then str_SQL = str_SQL & "ADDITIONAL_CONTACT_5 ," & vbCrLf
            
            'str_SQL = str_SQL & "ConfiMailSent ," & vbCrLf
            str_SQL = str_SQL & "[SB_Status] As Job_Status, " & vbCrLf
            str_SQL = str_SQL & "[CAD_Bearbeiter] As [1_CAD_Preparer], " & vbCrLf
            str_SQL = str_SQL & "AC_Status, " & vbCrLf
            str_SQL = str_SQL & "AC_Preparer, " & vbCrLf
            str_SQL = str_SQL & "convert(varchar(10), tsInputDataReceived, 120) AS tsInputDataReceived " & vbCrLf
           
    End Select

    '#JK-7/2021-04-14 START
    str_SQL = str_SQL & "From "

    str_SQL = str_SQL & conAndAc()
    '#JK-7/2021-04-14 START
    
    If b_buildQuery = True Then
    
        If t_econfiBtn = "einblenden" Then
            str_SQL = str_SQL & "WHERE Tool = 'CAD' AND (Storno is Null) AND Address_Validation = 1 Order By OrderNo ASC;"
        Else
            str_SQL = str_SQL & "WHERE (Storno is Null) AND Address_Validation = 1 Order By OrderNo ASC;"
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
            
        Case 3
        
            str_SQL = str_SQL & "WHERE OrderNo in (" & listOrder & ")"
        
        
    
    End Select
    
    modus = Empty
    
    buildSQLAdressen = str_SQL


End Function

Private Sub BtnEdit_Click()

DoCmd.OpenForm "EditForm"

End Sub
Private Function validateDateFormat(strDate)

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "\d\d\d\d-\d\d-\d\d"
    validateDateFormat = regex.test(strDate)

End Function
Private Function isManualOrderNo()

    Dim ctlList As Control
    Set ctlList = Forms![AC_Confi Client].Alltasks
    
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
Public Sub RowsSelectedAC(NewUser As String, orderbHeader As String)
   ' On Error GoTo Errormsg
    
    Call isTesting
    
    DoCmd.SetWarnings False
    
    Dim ctlList As Control, varItem As Variant, sql As String, selOrderNo As String, DBtable As String


    Set ctlList = Forms![AC_Confi Client].Alltasks
    

    For Each varItem In ctlList.ItemsSelected
    
        selOrderNo = ctlList.Column(0, varItem)
        
        If InStr(selOrderNo, "CON") > 0 Then
        DBtable = getTable(NewUser, orderbHeader, 1)
        
        Else
        'AC Orderbook
        DBtable = getTable(NewUser, orderbHeader, 0)
        End If

        sql = "UPDATE " & DBtable & " SET " & orderbHeader & "=" & handleSQLValues(NewUser) & " WHERE OrderNo='" & selOrderNo & "'"
        updateSQL (sql)
        If orderbHeader = "AC_Status" And handleSQLValues(NewUser) = "'CanvasDone'" Then
            sql = "UPDATE " & DBtable & " SET tsCanvasDone =" & Format(Now, "'yyyy-mm-dd hh:mm'") & " WHERE OrderNo='" & selOrderNo & "'"
            updateSQL (sql)
        End If
        
        If NewUser = "Storno" And orderbHeader <> "AC_Status" And orderbHeader <> "AC_Preparer" Then
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

Function getTable(NewUser As String, orderbHeader As String, modus As Integer) As String

' This function returns database table concerning solution path -> 0: AC, 1: regular CON

    Call isTesting

    Dim DBtable As String
    
    Select Case modus
    
        Case 0
        
            DBtable = "tAC_Orderbook" & strTestTable
                
        Case 1
    
            If orderbHeader = "Mandant_Anr" Or orderbHeader = "Mandant_Name" Or orderbHeader = "Mandant_Email" Then
                DBtable = "tCON_ClientContact" & strTestTable
            ElseIf orderbHeader = "Name_Bank" Or orderbHeader = "Portal_Bank" Then
                DBtable = "tCON_BankConfirmation" & strTestTable
            Else
                DBtable = "tCON_Orderbook" & strTestTable
            End If
    
    End Select
    
    getTable = DBtable


End Function

'#JK-8/2021-04-14 START
Function conAndAc() As String

    Call isTesting

    inclAC = Me.toggle_AC.value
    
    If inclAC = False Or IsNull(inclAC) Then
    
        conAndAc = "tCon_Orderbook" & vbCrLf
        
    ElseIf inclAC = True Then
    
        conAndAc = "(SELECT *" & vbCrLf
        conAndAc = conAndAc & "FROM tCON_Orderbook" & vbCrLf
        ' AC
        conAndAc = conAndAc & "UNION" & vbCrLf
        conAndAc = conAndAc & "SELECT NULL as [POrderNo]" & vbCrLf
        conAndAc = conAndAc & ",[OrderNo]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Storno]"
        conAndAc = conAndAc & ",[GISID]" & vbCrLf
        conAndAc = conAndAc & ",[Client]" & vbCrLf
        conAndAc = conAndAc & ",[EngCode]" & vbCrLf
        conAndAc = conAndAc & ",[EngName]" & vbCrLf
        conAndAc = conAndAc & ",engprtnr as [EngPartner]" & vbCrLf
        conAndAc = conAndAc & ",engmngr as [EngManager]" & vbCrLf
        conAndAc = conAndAc & ",[EngContact]" & vbCrLf
        conAndAc = conAndAc & ",[EngOtherContact] as [OtherContact]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [BusinessEstablishmentOAreAvailable]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [BusinessEstablishmentOAre]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [BelongingToCompanyGroup]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [NameOfCompanyGroup]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [TypeFinancials]" & vbCrLf
        conAndAc = conAndAc & ",[Periodend] as [YearEnd]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Confirmation]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [EYStand]" & vbCrLf
        conAndAc = conAndAc & ",[OrderDate] as  [DateConfi]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [TypeConfi]" & vbCrLf
        conAndAc = conAndAc & ",NuLL as [LangProduct]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [PlannedDelivery]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [DeliveryType]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Date2ndReq]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [OtherTypeConfi]" & vbCrLf
        conAndAc = conAndAc & ",[ADDITIONAL_CONTACT_2]" & vbCrLf
        conAndAc = conAndAc & ",[ADDITIONAL_CONTACT_3]" & vbCrLf
        conAndAc = conAndAc & ",[ADDITIONAL_CONTACT_4]" & vbCrLf
        conAndAc = conAndAc & ",[ADDITIONAL_CONTACT_5]" & vbCrLf
        conAndAc = conAndAc & ",[LangProduct] as  [DocLang]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Country]" & vbCrLf
        conAndAc = conAndAc & ",[tsConfiMailSent] as [ConfiMailSent]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [TemplateSent]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [AuditiSent]" & vbCrLf
        conAndAc = conAndAc & ",'CAD' as  [Tool]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [CAD_Bearbeiter]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [CAD_Bearbeiter_vorherig]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [CAD_Reviewer]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [SB_Status]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Anzahl_erhaltene_Salden]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Einladung_Canvas_erhalten]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Datum_erster_Upload]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Datum_Versand_EMail_Erstkontakt]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Informationen_zum_Auftrag]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Anzahl_Bestätigungen]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Auswahl_erhalten]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Datum_Auswahl_erhalten]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Versand_durch]" & vbCrLf
        conAndAc = conAndAc & ",[tsOrderImported] as [Eingangsdatum_CAD]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Versanddatum_CAD]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [2nd_Request_gewünscht]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Datum_2nd_Request]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Kommentar_Versand]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Final_Senior_Review_durch]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Datum_final_Review]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Versanddatum_Originale]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Empfänger_Originale]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Empfänger_Nachläufer]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Stunden_Bearbeiter]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Abrechnungsbetrag]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Abrechnungsdatum]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Datum_Feedbackanfrage]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Datum_Reminder]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Feedback_erhalten]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [ConfiFlag]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [First_Upload_Gewünscht]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Antwort_Deadline]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [DatumAuftragsende]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [StornoBit]" & vbCrLf
        conAndAc = conAndAc & ",1 as  [Address_Validation]" & vbCrLf
        conAndAc = conAndAc & ",[AC_Status] as  [AC_Status]" & vbCrLf
        conAndAc = conAndAc & ",[AC_Preparer] as  [AC_Preparer]" & vbCrLf
        conAndAc = conAndAc & ",[tsInputDataSent] as  [tsInputDataSent]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [ConfirmationsPlus]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [FiscalYearBegin]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [PlannedSendDate]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [ResponseTime]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [MandantKommunikation]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Confi_com_Status]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [AbfrageVollstaendigkeit]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [RL_Befuellung_Type]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Kommentar_eConfiPLUS]" & vbCrLf
        conAndAc = conAndAc & ",[tsInputDataReceived]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [tsTeamApprovalSent]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [tsTeamApprovalReceived]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [tsReminderTeamApprovalSent]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [tsReminderTeamApprovalReceived]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [tsInputDataAvailable]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [tsCanvasDone]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [Anzahl Banken]" & vbCrLf
        conAndAc = conAndAc & "FROM [CAD].[dbo].[tAC_Orderbook" & strTestTable & "]" & vbCrLf
        ' AP
        conAndAc = conAndAc & "UNION" & vbCrLf
        conAndAc = conAndAc & "SELECT NULL as [POrderNo]" & vbCrLf
        conAndAc = conAndAc & ", ap.OrderNo" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Storno], oi.GISID" & vbCrLf
        conAndAc = conAndAc & ", oi.GISName AS Client" & vbCrLf
        conAndAc = conAndAc & ", oi.EngCode" & vbCrLf
        conAndAc = conAndAc & ", oi.EngName" & vbCrLf
        conAndAc = conAndAc & ", oi.EngPartner_oi as [EngPartner]" & vbCrLf
        conAndAc = conAndAc & ", oi.EngManager_oi as [EngManager]" & vbCrLf
        conAndAc = conAndAc & ", oi.EngContact_oi AS EngContact" & vbCrLf
        conAndAc = conAndAc & ", oi.ADD_Contact1_oi as [OtherContact]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [BusinessEstablishmentOAreAvailable]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [BusinessEstablishmentOAre]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [BelongingToCompanyGroup]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [NameOfCompanyGroup]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [TypeFinancials]" & vbCrLf
        conAndAc = conAndAc & ", oh.PeriodEnd as [YearEnd]" & vbCrLf
        conAndAc = conAndAc & ", 'Kreditoren' as  [Confirmation]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [EYStand]" & vbCrLf
        conAndAc = conAndAc & ", ts.tsOrderDate as  [DateConfi]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [TypeConfi]" & vbCrLf
        conAndAc = conAndAc & ", NuLL as [LangProduct]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [PlannedDelivery]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [DeliveryType]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Date2ndReq]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [OtherTypeConfi]" & vbCrLf
        conAndAc = conAndAc & ", oi.ADD_Contact2_oi as [ADDITIONAL_CONTACT_2]" & vbCrLf
        conAndAc = conAndAc & ", oi.ADD_Contact3_oi as [ADDITIONAL_CONTACT_3]" & vbCrLf
        conAndAc = conAndAc & ", oi.ADD_Contact4_oi as [ADDITIONAL_CONTACT_4]" & vbCrLf
        conAndAc = conAndAc & ", oi.ADD_Contact5_oi as [ADDITIONAL_CONTACT_5]" & vbCrLf
        conAndAc = conAndAc & ", oi.Doculang_oi as [DocLang]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Country]" & vbCrLf
        conAndAc = conAndAc & ", ts.tsConfiMailSent as [ConfiMailSent]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [TemplateSent]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [AuditiSent]" & vbCrLf
        conAndAc = conAndAc & ", 'CAD' as  [Tool]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [CAD_Bearbeiter]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [CAD_Bearbeiter_vorherig]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [CAD_Reviewer]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [SB_Status]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Anzahl_erhaltene_Salden]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Einladung_Canvas_erhalten]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_erster_Upload]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_Versand_EMail_Erstkontakt]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Informationen_zum_Auftrag]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Anzahl_Bestätigungen]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Auswahl_erhalten]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_Auswahl_erhalten]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Versand_durch]" & vbCrLf
        conAndAc = conAndAc & ", NULL as [Eingangsdatum_CAD]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Versanddatum_CAD]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [2nd_Request_gewünscht]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_2nd_Request]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Kommentar_Versand]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Final_Senior_Review_durch]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_final_Review]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Versanddatum_Originale]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Empfänger_Originale]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Empfänger_Nachläufer]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Stunden_Bearbeiter]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Abrechnungsbetrag]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Abrechnungsdatum]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_Feedbackanfrage]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_Reminder]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Feedback_erhalten]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [ConfiFlag]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [First_Upload_Gewünscht]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Antwort_Deadline]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [DatumAuftragsende]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [StornoBit]" & vbCrLf
        conAndAc = conAndAc & ", 1 as  [Address_Validation]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [AC_Status]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [AC_Preparer]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [tsInputDataSent]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [ConfirmationsPlus]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [FiscalYearBegin]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [PlannedSendDate]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [ResponseTime]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [MandantKommunikation]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Confi_com_Status]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [AbfrageVollstaendigkeit]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [RL_Befuellung_Type]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Kommentar_eConfiPLUS]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [tsInputDataReceived]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [tsTeamApprovalSent]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [tsTeamApprovalReceived]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [tsReminderTeamApprovalSent]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [tsReminderTeamApprovalReceived]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [tsInputDataAvailable]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [tsCanvasDone]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Anzahl Banken]" & vbCrLf
        conAndAc = conAndAc & "FROM [CAD].[dbo].[tT_CADDB_AP] ap " & vbCrLf
        conAndAc = conAndAc & "LEFT JOIN [CAD].[dbo].[tT_CADDB_CAD_OrderedItems] oi ON ap.OrderNo = oi.OrderNo" & vbCrLf
        conAndAc = conAndAc & "LEFT JOIN [CAD].[dbo].[tT_CADDB_CAD_OrderHeader] oh ON oi.CADOrderNo = oh.CADOrderNo" & vbCrLf
        conAndAc = conAndAc & "LEFT JOIN [CAD].[dbo].[tT_CADDB_CAD_OrderedItemsTimeStamps] ts ON ap.OrderNo = ts.OrderNo WHERE ap.EConfi = 1" & vbCrLf
        ' AR
        conAndAc = conAndAc & "UNION" & vbCrLf
        conAndAc = conAndAc & "SELECT NULL as [POrderNo]" & vbCrLf
        conAndAc = conAndAc & ", ar.OrderNo" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Storno], oi.GISID" & vbCrLf
        conAndAc = conAndAc & ", oi.GISName AS Client" & vbCrLf
        conAndAc = conAndAc & ", oi.EngCode" & vbCrLf
        conAndAc = conAndAc & ", oi.EngName" & vbCrLf
        conAndAc = conAndAc & ", oi.EngPartner_oi as [EngPartner]" & vbCrLf
        conAndAc = conAndAc & ", oi.EngManager_oi as [EngManager]" & vbCrLf
        conAndAc = conAndAc & ", oi.EngContact_oi AS EngContact" & vbCrLf
        conAndAc = conAndAc & ", oi.ADD_Contact1_oi as [OtherContact]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [BusinessEstablishmentOAreAvailable]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [BusinessEstablishmentOAre]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [BelongingToCompanyGroup]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [NameOfCompanyGroup]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [TypeFinancials]" & vbCrLf
        conAndAc = conAndAc & ", oh.PeriodEnd as [YearEnd]" & vbCrLf
        conAndAc = conAndAc & ", 'Debitoren' as  [Confirmation]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [EYStand]" & vbCrLf
        conAndAc = conAndAc & ", ts.tsOrderDate as  [DateConfi]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [TypeConfi]" & vbCrLf
        conAndAc = conAndAc & ", NuLL as [LangProduct]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [PlannedDelivery]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [DeliveryType]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Date2ndReq]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [OtherTypeConfi]" & vbCrLf
        conAndAc = conAndAc & ", oi.ADD_Contact2_oi as [ADDITIONAL_CONTACT_2]" & vbCrLf
        conAndAc = conAndAc & ", oi.ADD_Contact3_oi as [ADDITIONAL_CONTACT_3]" & vbCrLf
        conAndAc = conAndAc & ", oi.ADD_Contact4_oi as [ADDITIONAL_CONTACT_4]" & vbCrLf
        conAndAc = conAndAc & ", oi.ADD_Contact5_oi as [ADDITIONAL_CONTACT_5]" & vbCrLf
        conAndAc = conAndAc & ", oi.Doculang_oi as [DocLang]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Country]" & vbCrLf
        conAndAc = conAndAc & ", ts.tsConfiMailSent as [ConfiMailSent]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [TemplateSent]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [AuditiSent]" & vbCrLf
        conAndAc = conAndAc & ", 'CAD' as  [Tool]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [CAD_Bearbeiter]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [CAD_Bearbeiter_vorherig]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [CAD_Reviewer]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [SB_Status]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Anzahl_erhaltene_Salden]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Einladung_Canvas_erhalten]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_erster_Upload]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_Versand_EMail_Erstkontakt]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Informationen_zum_Auftrag]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Anzahl_Bestätigungen]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Auswahl_erhalten]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_Auswahl_erhalten]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Versand_durch]" & vbCrLf
        conAndAc = conAndAc & ", NULL as [Eingangsdatum_CAD]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Versanddatum_CAD]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [2nd_Request_gewünscht]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_2nd_Request]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Kommentar_Versand]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Final_Senior_Review_durch]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_final_Review]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Versanddatum_Originale]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Empfänger_Originale]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Empfänger_Nachläufer]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Stunden_Bearbeiter]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Abrechnungsbetrag]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Abrechnungsdatum]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_Feedbackanfrage]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Datum_Reminder]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Feedback_erhalten]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [ConfiFlag]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [First_Upload_Gewünscht]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Antwort_Deadline]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [DatumAuftragsende]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [StornoBit]" & vbCrLf
        conAndAc = conAndAc & ", 1 as  [Address_Validation]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [AC_Status]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [AC_Preparer]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [tsInputDataSent]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [ConfirmationsPlus]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [FiscalYearBegin]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [PlannedSendDate]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [ResponseTime]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [MandantKommunikation]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Confi_com_Status]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [AbfrageVollstaendigkeit]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [RL_Befuellung_Type]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Kommentar_eConfiPLUS]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [tsInputDataReceived]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [tsTeamApprovalSent]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [tsTeamApprovalReceived]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [tsReminderTeamApprovalSent]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [tsReminderTeamApprovalReceived]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [tsInputDataAvailable]" & vbCrLf
        conAndAc = conAndAc & ",NULL as  [tsCanvasDone]" & vbCrLf
        conAndAc = conAndAc & ", NULL as  [Anzahl Banken]" & vbCrLf
        conAndAc = conAndAc & "FROM [CAD].[dbo].[tT_CADDB_AR] ar " & vbCrLf
        conAndAc = conAndAc & "LEFT JOIN [CAD].[dbo].[tT_CADDB_CAD_OrderedItems] oi ON ar.OrderNo = oi.OrderNo" & vbCrLf
        conAndAc = conAndAc & "LEFT JOIN [CAD].[dbo].[tT_CADDB_CAD_OrderHeader] oh ON oi.CADOrderNo = oh.CADOrderNo" & vbCrLf
        conAndAc = conAndAc & "LEFT JOIN [CAD].[dbo].[tT_CADDB_CAD_OrderedItemsTimeStamps] ts ON ar.OrderNo = ts.OrderNo WHERE ar.EConfi = 1" & vbCrLf
        conAndAc = conAndAc & ") as a" & vbCrLf
    End If

End Function
'#JK-8/2021-04-14 END

Private Sub showFurtherDetails()

    On Error GoTo Errorhandler
    
    Dim strOrderNo As String
    Dim rsData As ADODB.RecordSet, rsData2 As ADODB.RecordSet
    
    strOrderNo = Me.Alltasks.Column(0)
    
    Dim queryPOrderNo As String
    queryPOrderNo = getQueryPOrderNo(strOrderNo)
    
    Set rsData2 = getRS(queryPOrderNo)
        
    
    Me.FurtherDetails.ColumnCount = 8
    Me.FurtherDetails.columnWidths = "3 cm;3 cm;5 cm;3 cm;2 cm;5 cm;5 cm;5 cm"
    Set Me.FurtherDetails.RecordSet = rsData2
    
    Exit Sub
    
Errorhandler:
    '    rsData.Close
    Set rsData2 = Nothing

End Sub

