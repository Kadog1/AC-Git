Attribute VB_Name = "CON_Helper"
Public Sub updateSQL(strSQL As String)

    Dim RecordSet As ADODB.RecordSet
    Dim Connection As ADODB.Connection

    Set Connection = CreateObject("ADODB.Connection")

    Connection.ConnectionString = "Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;"

    Connection.Open
    Debug.Print strSQL

    Set RecordSet = Connection.Execute(strSQL)

    Connection.Close
End Sub
Sub readDatenSammler_CON()

          Dim I As Integer, j As Integer
          
          Dim indexAccount As Integer
1         For I = 1 To Application.Session.Accounts.Count
2             If LCase(Application.Session.Accounts.item(I)) = "adressabgleich@de.ey.com" Then
3                 indexAccount = I
4                 Exit For
5             End If
6         Next

          Dim accIND As Object, myInbox As Object, olFolder As Object
7         Set accIND = Application.Session.Accounts.item(indexAccount)
8         Set myInbox = accIND.DeliveryStore.GetDefaultFolder(6)
9         Set olFolder = myInbox.Folders("Team Reply")

          ' read e-mails in Confi-Inbox
          Dim c_mails As Long
          Dim mails() As Object
10        If olFolder.Items.Count >= 1 Then ReDim mails(1 To olFolder.Items.Count)
11        c_mails = 0
12        For Each item In olFolder.Items
14            c_mails = c_mails + 1
15            Set mails(c_mails) = item
16        Next

          Dim rsOrderbook As Object, rsSubTable As Object
          
          Dim RecordSet As ADODB.RecordSet
          Dim Connection As ADODB.Connection

          Set Connection = CreateObject("ADODB.Connection")
          Connection.ConnectionString = "Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;"
          Connection.Open

          Dim OutMail As Object
          Dim objAtt As Object, objApp As Object
          Dim subject As String, subj() As String, OrderNo As String
          Dim saveFolder As String
          Dim mappingcmd As String
          'Set objApp = CreateObject("Excel.Application")
          Dim wbTemplate As Object, wsTemplate As Object, wbOutput As Object
          Dim arrInput() As Variant, lastrow As Integer
          Dim outName As String
          Dim bSendToAuditi As Boolean
        
          For I = 1 To c_mails
              Call processMail(mails(I), Connection, "Team Reply Processed", myInbox, "InputDataAvailable")
          Next I
         
          Connection.Close
          Set rsOrderbook = Nothing

End Sub
Sub readDatenSammlerAuditi_CON()

          Dim I As Integer, j As Integer
          
          Dim indexAccount As Integer
1         For I = 1 To Application.Session.Accounts.Count
2             If LCase(Application.Session.Accounts.item(I)) = "adressabgleich@de.ey.com" Then
3                 indexAccount = I
4                 Exit For
5             End If
6         Next

          Dim accIND As Object, myInbox As Object, olFolder As Object
7         Set accIND = Application.Session.Accounts.item(indexAccount)
8         Set myInbox = accIND.DeliveryStore.GetDefaultFolder(6)
9         Set olFolder = myInbox.Folders("Team Reply")

          ' read e-mails in Confi-Inbox
          Dim c_mails As Long
          Dim mails() As Object
10        If olFolder.Items.Count >= 1 Then ReDim mails(1 To olFolder.Items.Count)
11        c_mails = 0
12        For Each item In olFolder.Items
14            c_mails = c_mails + 1
15            Set mails(c_mails) = item
16        Next
    
          Dim rsOrderbook As Object, rsSubTable As Object
          
          Dim RecordSet As ADODB.RecordSet
          Dim Connection As ADODB.Connection

          Set Connection = CreateObject("ADODB.Connection")
          Connection.ConnectionString = "Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;"
          Connection.Open

          Dim OutMail As Object
          
          Dim subject As String, subj() As String, OrderNo As String
          Dim saveFolder As String
          Dim mappingcmd As String

          Dim arrInput() As Variant, lastrow As Integer
          Dim outName As String
          Dim bSendToAuditi As Boolean
        
          For I = 1 To c_mails
              Call processMailAuditi(mails(I), Connection, "Team Reply Processed", myInbox, "InputDataAvailable")
          Next I
          
          Connection.Close
          Set rsOrderbook = Nothing

End Sub
Sub process_TeamApproval_CON()

          Dim I As Integer, j As Integer
          Dim strMailAccount As String, lngCountWB As Long
    
          Dim indexAccount As Integer
          Dim saveFolder As String
    
          'Liveumgebung
          strMailAccount = "adressabgleich@de.ey.com"

          For I = 1 To Application.Session.Accounts.Count
              If LCase(Application.Session.Accounts.item(I)) = strMailAccount Then
                  indexAccount = I
                  Exit For
              End If
          Next
    
          Dim accIND As Object, myInbox As Object, olFolder As Object
          Set accIND = Application.Session.Accounts.item(indexAccount)
          Set myInbox = accIND.DeliveryStore.GetDefaultFolder(6)

          Set olFolder = myInbox.Folders("Team Approval")
    
          ' read e-mails in Confi-Inbox
          Dim allmails As Outlook.Items
          Dim c_mails As Long
          Dim mails() As Object
          If olFolder.Items.Count >= 1 Then ReDim mails(1 To olFolder.Items.Count)
          c_mails = 0

          Dim item As Object

          Set allmails = olFolder.Items

          ' Collect all e-mails in New Order
          For Each item In allmails
              c_mails = c_mails + 1
              Set mails(c_mails) = item
          Next item
    
          Dim subj() As String
          Dim OrderNo As String
          
          Dim objAtt As Object, objApp As Object
          
          For I = 1 To c_mails
    
              OrderNo = ""
    
              subj = Split(mails(I).subject, ": ")
              If UBound(subj) >= 1 Then
                  OrderNo = subj(2)
              End If
        
              If OrderNo = "" Then Exit For

              ' create folder for ordernumber in customer data
        
              saveFolder = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & OrderNo & "\"
                      
              If Dir(saveFolder, vbDirectory) = "" Then MkDir (saveFolder)
              If Dir(saveFolder & "3. Team Approval\", vbDirectory) = "" Then MkDir (saveFolder & "3. Team Approval\")

              mails(I).SaveAs saveFolder & "3. Team Approval\" & OrderNo & " " & Format(Now(), "YYYYMMDD-HHMMSS") & ".msg", olMSG
        
				' Save attachment from mailItem'
              Debug.Print mails(I).Attachments.Count
              If mails(I).Attachments.Count > 0 Then

                  Set objAtt = mails(I).Attachments.item(1)
22                If InStr(objAtt.FileName, " 3_CAD-Adressenabgleich Team Approval") > 0 Or InStr(objAtt.FileName, " 3_CAD-Adressabgleich Team Approval") > 0 Then

                      objAtt.SaveAsFile saveFolder & "3. Team Approval\" & objAtt.FileName
                      Set objAtt = Nothing
                  End If
              End If
        
              mails(I).Move myInbox.Folders("Team Approval Processed")
        
              Dim timeNow As String
              timeNow = Format(Now(), "yyyy-MM-dd hh:mm:ss")
              Call updateSQL("UPDATE tCON_Orderbook Set AC_Status = 'TeamApprovalReceived' WHERE OrderNo = '" & OrderNo & "'")
              Call updateSQL("UPDATE tCON_Orderbook Set tsTeamApprovalReceived = '" & timeNow & "' WHERE OrderNo = '" & OrderNo & "'")
          Next I

End Sub

Sub SendNotification_CON()

    Dim I As Integer, j As Integer
    Dim myInbox As Object, item As Object, objAtt As Object, accIND As Object, newOrderFolder As Object
    Dim RecordSet As ADODB.RecordSet
    Dim Connection As ADODB.Connection

    Set Connection = CreateObject("ADODB.Connection")
    Connection.ConnectionString = "Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;"
    Connection.Open


    Dim indexAccount As Integer
    For I = 1 To Outlook.Application.Session.Accounts.Count
        If LCase(Outlook.Application.Session.Accounts.item(I)) = "adressabgleich@de.ey.com" Then ' PROD / TEST
            indexAccount = I
            Exit For
        End If
    Next

    Set accIND = Outlook.Application.Session.Accounts.item(indexAccount)
    Set myInbox = accIND.DeliveryStore.GetDefaultFolder(6)
    Set newOrderFolder = myInbox.Folders("Team Reply")

    Dim c_mails As Long
    Dim mails() As Object

    If newOrderFolder.Items.Count >= 1 Then ReDim mails(1 To newOrderFolder.Items.Count)

    Dim allmails As Object
    Set allmails = newOrderFolder.Items
    allmails.Sort "[ReceivedTime]"

    Dim DupItem As Object
    Set DupItem = CreateObject("Scripting.Dictionary")
    For I = allmails.Count To 1 Step -1
        If TypeOf allmails(I) Is mailItem Then
            Set item = allmails(I)
            If item.ReceivedTime >= allmails(I).ReceivedTime Then
                If DupItem.Exists(item.subject) Then
                    item.Move myInbox.Folders("Team Reply Processed")
                Else
                    DupItem.Add item.subject, 0
                End If
            End If
        End If
    Next I
    Set allmails = newOrderFolder.Items
    c_mails = 0
    For Each item In allmails
        c_mails = c_mails + 1
        Set mails(c_mails) = item
    Next
    Dim pathMailTempaltes, templateName As String
    pathMailTemplates = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich\B Mail Templates\German\"
    templateName = "AC_StatusMail.htm"

    Dim replaced As String

    For I = 1 To c_mails
        OrderNo = Right(mails(I).subject, 13)
        If Not Left(OrderNo, 6) = "CON000" Then
            subject = mails(I).subject
            subj = Split(subject, " ")
            If InStr(subject, "Beauftragung Adressabgleich") > 0 Then
                OrderNo = subj(3)
            Else
                OrderNo = subj(4)
            End If
        End If
        Debug.Print mails(I).subject
        Set rsOrderbook = Connection.Execute("SELECT [AC_Status], [AC_Preparer], [Client], [GISID], [YearEnd], [Tool] FROM tCON_Orderbook WHERE [OrderNo] = '" & OrderNo & "'") ' PROD / TEST
        Do While rsOrderbook.EOF = False
            If rsOrderbook("AC_Status") = "" Or rsOrderbook("AC_Status") = "InputDataSent" Or rsOrderbook("AC_Status") = "WaitingForInputData" Then
                Debug.Print ("Continue")
            ElseIf (rsOrderbook("AC_Status") = "TeamApprovalSent" Or rsOrderbook("AC_Status") = "TeamApprovalReceived" Or rsOrderbook("AC_Status") = "CanvasDone") And rsOrderbook("Tool") = "eConfirmations" Then
                mails(I).Move myInbox.Folders("Team Reply Processed")
            ElseIf rsOrderbook("AC_Status") = "InputDataReceived" Or rsOrderbook("AC_Status") = "InProgress" Or rsOrderbook("AC_Status") = "ReadyForReview" Or rsOrderbook("AC_Status") = "TeamApprovalSent" Or rsOrderbook("AC_Status") = "TeamApprovalReceived" Or rsOrderbook("AC_Status") = "CanvasDone" Then
                ' AC Nachlieferungsmanagement: bei Tool = eCon keine Statusaenderung
                Dim newStatus As String
                newStatus = "InputDataAvailable"
                If rsOrderbook("Tool") = "eConfirmations" Then
                    newStatus = rsOrderbook("AC_Status")
                End If
                saveFolder = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & OrderNo & "\2. CAD_Abgleich\"
                If Dir(saveFolder, vbDirectory) = "" Then MkDir (saveFolder)
                If Dir(saveFolder & "Previous InputDatenSheets\", vbDirectory) = "" Then MkDir (saveFolder & "Previous InputDatenSheets\")
                fileInputSheet = Dir(saveFolder & Format(rsOrderbook("GISID"), "0000000000") & " 1_CAD-Adressabgleich Adressenabfrage Mandant " & Format(rsOrderbook("YearEnd"), "yyyyMMdd") & ".xlsx")
                ' Verschiebe InputSheet falls vorhanden
                If Len(fileInputSheet) > 0 Then
                    Set fso = CreateObject("Scripting.Filesystemobject")
                    SourceFileName = saveFolder & fileInputSheet
                    DestinFileName = saveFolder & "Previous InputDatenSheets\" & Format(Now(), "YYYYMMDD-HHMM_") & fileInputSheet
                    fso.MoveFile source:=SourceFileName, Destination:=DestinFileName
                End If
                ' Speicher neue InputSheet
                Call processMail(mails(I), Connection, "Team Reply Nachlieferung", myInbox, newStatus)
                Call processMailAuditi(mails(I), Connection, "Team Reply Nachlieferung", myInbox, newStatus)
                If IsNull(rsOrderbook("AC_Preparer")) = False And Not (rsOrderbook("AC_Preparer") = "") Then
                    preparer = rsOrderbook("AC_Preparer")
                    preparer = Replace(preparer, "?", "ae")
                    preparer = Replace(preparer, "?", "oe")
                    preparer = Replace(preparer, "?", "ue")
                    replaced = Replace(preparer, " ", "")
                    Count = Len(preparer) - Len(replaced)
                    If Count > 1 Then
                        givenName = Left(preparer, InStr(preparer, " ") - 1)
                        restName = Mid(preparer, InStr(preparer, " ") + 1)
                        MiddleName = Left(restName, 1)
                        surname = Mid(restName, InStr(restName, " "))
                        preparerEmail = givenName + "." + MiddleName + "." + surname + "@de.ey.com"
                    Else
                        givenName = Left(preparer, InStr(preparer, " ") - 1)
                        surname = Mid(preparer, InStr(preparer, " ") + 1)
                        preparerEmail = givenName + "." + surname + "@de.ey.com"
                    End If
                    preparerEmail = Replace(preparerEmail, " ", "")
                    gisid = rsOrderbook("GISID")
                    client = rsOrderbook("Client")
                    body = parseBody(pathMailTemplates & templateName)
                    body = Replace(body, "[orderNo]", OrderNo)
                    body = Replace(body, "[GISID]", gisid)
                    body = Replace(body, "[client]", client)
                    Set OutMail = Outlook.Application.CreateItem(olMailItem)
                    OutMail.To = preparerEmail
                    OutMail.HTMLbody = body
                    OutMail.SentOnBehalfOfName = "adressabgleich@de.ey.com"
                    OutMail.subject = "Neue Adressdaten: " + OrderNo
                    OutMail.Attachments.Add saveFolder & Format(gisid, "0000000000") & " 1_CAD-Adressabgleich Adressenabfrage Mandant " & Format(rsOrderbook("YearEnd"), "yyyyMMdd") & ".xlsx"
                    OutMail.Send
                End If
                'mails(I).Move myInbox.Folders("Team Reply Nachlieferung")
            Else
                mails(I).Move myInbox.Folders("Team Reply Processed")
            End If
            rsOrderbook.MoveNext
        Loop
        rsOrderbook.Close
    Next
    Connection.Close
End Sub

Public Function parseBody(t_template As String) As String

          '##############################################
          '##### Import .HTM Content into Mail Body #####
          '##############################################
              
          Dim body As String
1         body = ""
          
2         If t_template <> "" Then
          
              ' Read htm content into body
              Dim objStream As Object
3             Set objStream = CreateObject("ADODB.Stream")
          
4             objStream.Charset = "utf-8"
5             objStream.Open
6             objStream.LoadFromFile (t_template)
          
7             body = objStream.ReadText()
              
              ' Clean up
8             objStream.Close
9             Set objStream = Nothing
              
10        End If
11        parseBody = body

End Function
Sub processMail(mailItem As Object, Connection As ADODB.Connection, moveToFolder As String, myInbox As Object, newStatus As String)
    bSendToAuditi = True
    ' Welche Art von Confirmation ist es
    subject = mailItem.subject
    If InStr(subject, "Beauftragung Adressabgleich") > 0 Or InStr(subject, "Order of debtor") > 0 Then
        subj = Split(subject, " ")
        If InStr(subject, "Beauftragung Adressabgleich") > 0 Then
            OrderNo = subj(3)
        Else
            OrderNo = subj(4)
        End If
        ' Check, ob Ordernummer schon storniert
        Set rsOrderbook = Connection.Execute("SELECT * FROM tCON_Orderbook WHERE [OrderNo] = '" & OrderNo & "'")
        Do While rsOrderbook.EOF = False
            If rsOrderbook("AC_Status") = "Storno" Or IsNull(rsOrderbook("Storno")) = False Then
                ' Keine Dateien an Auditi senden
                bSendToAuditi = False
                Exit Do
            End If
                rsOrderbook.MoveNext
        Loop
        rsOrderbook.Close
        If bSendToAuditi Then
            saveFolder = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & OrderNo & "\"
            If Dir(saveFolder, vbDirectory) = "" Then MkDir (saveFolder)
            If Dir(saveFolder & "2. CAD_Abgleich\", vbDirectory) = "" Then MkDir (saveFolder & "2. CAD_Abgleich\")
            'Debug.Print saveFolder
            Debug.Print mailItem.Attachments.Count
            If mailItem.Attachments.Count > 0 Then
                 Set objAtt = mailItem.Attachments.item(1)
                 If InStr(objAtt.FileName, "1_CAD-Adressabgleich Adressenabfrage Mandant") > 0 Or InStr(objAtt.FileName, "Debitoren Einzelposten") > 0 Then
                    'objApp.DisplayAlerts = False
                    Debug.Print saveFolder & "2. CAD_Abgleich\" & objAtt.FileName
                    objAtt.SaveAsFile saveFolder & "2. CAD_Abgleich\" & objAtt.FileName
                    'objApp.DisplayAlerts = True
                    mailItem.Move myInbox.Folders(moveToFolder)
                    Dim timeNow As String
                    timeNow = Format(Now(), "yyyy-MM-dd hh:mm:ss")
                    Call updateSQL("UPDATE tCON_Orderbook Set AC_Status = '" & newStatus & "' WHERE OrderNo = '" & OrderNo & "'")
                                        If newStatus = "InputDataAvailable" Then
                                                Call updateSQL("UPDATE tCON_Orderbook Set tsInputDataAvailable = '" & timeNow & "' WHERE OrderNo = '" & OrderNo & "'")
                                        End If
                End If
            Else
                 Debug.Print "Anhang umbenannt vom Team?!"
            End If
            Set objAtt = Nothing
        End If
    End If
End Sub

Sub processMailAuditi(mailItem As Object, Connection As ADODB.Connection, moveToFolder As String, myInbox As Object, newStatus As String)
    bSendToAuditi = True
    ' Welche Art von Confirmation ist es
    subject = mailItem.subject
    If InStr(subject, "CAD Adressabgleich") > 0 And InStr(subject, ",CON00") = 0 Then
        subj = Split(subject, " ")
        If InStr(subject, "CAD Adressabgleich") > 0 Then
            OrderNo = subj(2)
        Else
          GoTo JM_strangeReply
        End If
        
        Dim yearend As String, client As String, engcontact As String, engpartner As String, engmanager As String, additionalCntct As String, gisid As String, Confirmation As String, engcode As String, DocLang As String
        ' Check, ob Ordernummer schon storniert
        Dim querySQL As String, arapConfirmation As String
        arapConfirmation = "Kreditoren"
        querySQL = "SELECT * FROM tCON_Orderbook WHERE [OrderNo] = '" & OrderNo & "'"
        If InStr(OrderNo, "AP000") > 0 Or InStr(OrderNo, "AR000") > 0 Then
            If InStr(OrderNo, "AR000") > 0 Then arapConfirmation = "Debitoren"
            querySQL = "SELECT oi.GISID, oi.GISName AS client, oi.EngCode AS engCode, oi.Doculang_oi As DocLang, '" & arapConfirmation & "' AS Confirmation, oh.PeriodEnd AS YearEnd, CASE WHEN oi.Storno = 0 THEN NULL ELSE 1 END AS Storno, NULL AS AC_Status, "
            querySQL = querySQL & "oi.EngContact_oi AS engcontact, oi.EngManager_oi AS engmanager, oi.EngPartner_oi AS engpartner, oi.Add_Contact1_oi AS OtherContact, " & vbCrLf
            querySQL = querySQL & "oi.Add_Contact2_oi AS ADDITIONAL_CONTACT_2, oi.Add_Contact3_oi AS ADDITIONAL_CONTACT_3, oi.Add_Contact4_oi AS ADDITIONAL_CONTACT_4, oi.Add_Contact5_oi AS ADDITIONAL_CONTACT_5 " & vbCrLf
            querySQL = querySQL & " FROM [CAD].[dbo].[tT_CADDB_CAD_OrderedItems] oi LEFT JOIN [CAD].[dbo].[tT_CADDB_CAD_OrderHeader] oh ON oi.CADOrderNo = oh.CADOrderNo " & vbCrLf
            querySQL = querySQL & " WHERE OrderNo = '" & OrderNo & "'" & vbCrLf
        End If
        'Set rsOrderbook = Connection.Execute(querySQL)
        Set rsOrderbook = Connection.Execute("SELECT * FROM AC_SELECT('" & Left(OrderNo, 2) & "', '" & OrderNo & "')")
        Do While rsOrderbook.EOF = False
            yearend = rsOrderbook("YearEnd").Value
            client = rsOrderbook("client").Value
            DocLang = rsOrderbook("DocLang").Value
            engcontact = rsOrderbook("engcontact").Value
            engmanager = rsOrderbook("engmanager").Value
            engpartner = rsOrderbook("engpartner").Value
            additionalCntct = ""
            If Not IsNull(rsOrderbook("OtherContact").Value) Then additionalCntct = additionalCntct & ";" & rsOrderbook("OtherContact").Value
            If Not IsNull(rsOrderbook("ADDITIONAL_CONTACT_2").Value) Then additionalCntct = additionalCntct & ";" & rsOrderbook("ADDITIONAL_CONTACT_2").Value
            If Not IsNull(rsOrderbook("ADDITIONAL_CONTACT_3").Value) Then additionalCntct = additionalCntct & ";" & rsOrderbook("ADDITIONAL_CONTACT_3").Value
            If Not IsNull(rsOrderbook("ADDITIONAL_CONTACT_4").Value) Then additionalCntct = additionalCntct & ";" & rsOrderbook("ADDITIONAL_CONTACT_4").Value
            If Not IsNull(rsOrderbook("ADDITIONAL_CONTACT_5").Value) Then additionalCntct = additionalCntct & ";" & rsOrderbook("ADDITIONAL_CONTACT_5").Value
            gisid = rsOrderbook("GISID").Value
            Confirmation = rsOrderbook("Confirmation").Value
            engcode = rsOrderbook("engCode").Value
            If rsOrderbook("AC_Status") = "Storno" Or IsNull(rsOrderbook("Storno")) = False Then ' Keine Dateien an Auditi senden
                bSendToAuditi = False
                Exit Do
            End If
            rsOrderbook.MoveNext
        Loop
        rsOrderbook.Close
        If bSendToAuditi Then
          saveFolder = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & OrderNo & "\"
          If Dir(saveFolder, vbDirectory) = "" Then MkDir (saveFolder)
          If Dir(saveFolder & "2. CAD_Abgleich\", vbDirectory) = "" Then MkDir (saveFolder & "2. CAD_Abgleich\")
          'Debug.Print saveFolder
          Debug.Print mailItem.Attachments.Count
          If mailItem.Attachments.Count > 0 Then
            Set objAtt = mailItem.Attachments.item(1)
            If InStr(objAtt.FileName, "CAD Adressabgleich") > 0 Then
                Dim objApp As Object
                Set objApp = CreateObject("Excel.Application")
                Dim wbTemplate As Object, wsTemplate As Object, wbOutput As Object
                Debug.Print saveFolder & "2. CAD_Abgleich\" & objAtt.FileName
                  
                objAtt.SaveAsFile saveFolder & "2. CAD_Abgleich\" & objAtt.FileName
                  
                ' Load binfo!
                Set wbTemplate = objApp.Workbooks.Open(saveFolder & "2. CAD_Abgleich\" & objAtt.FileName)
                Set wsTemplate = wbTemplate.Worksheets("basic_info")
                Set wbOutput = wbTemplate.Worksheets("Adresscheck")
                  
                wsTemplate.Cells(1, 2) = OrderNo
                wsTemplate.Cells(2, 2) = yearend
                wsTemplate.Cells(3, 2) = client
                wsTemplate.Cells(4, 2) = engcontact
                wsTemplate.Cells(5, 2) = engpartner
                wsTemplate.Cells(6, 2) = engmanager
                wsTemplate.Cells(7, 2) = additionalCntct
                wsTemplate.Cells(8, 2) = gisid
                wsTemplate.Cells(9, 2) = Confirmation
                wsTemplate.Cells(10, 2) = engcode
                
                wbOutput.Activate
                wbOutput.Columns("E:E").Select
                wbOutput.Range("E:E").Insert

                wbOutput.Range("E15") = "Adresszusatz"
                If DocLang = "EN" Then
                    Call replaceAll(wbOutput, "Debitor", "Debtor")
                    Call replaceAll(wbOutput, "Kreditor", "Creditor")
                    Call replaceAll(wbOutput, "Rechtsberater", "Legal Advisor")
                    Call replaceAll(wbOutput, "Steuerberater", "Tax advisor")
                    Call replaceAll(wbOutput, "Wirtschaftsprüfer", "Auditor")
                    Call replaceAll(wbOutput, "Sonstiger Berater", "Other advisor")
                    Call replaceAll(wbOutput, "Sonstige", "Other")
                    wbOutput.Range("E15") = "Additional address information"
                    wbTemplate.Worksheets("Adresscheck").Name = "Address check"
                End If
                Folder = "\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\eConfirmations\Datenbank\C Workplace\" & OrderNo & "\"
                If Dir(Folder, vbDirectory) = "" Then MkDir (Folder)
                If Dir(Folder & "2. CAD_Abgleich\", vbDirectory) = "" Then MkDir (Folder & "2. CAD_Abgleich\")
                    wbTemplate.SaveAs FileName:=Folder & "2. CAD_Abgleich\" & Format(gisid, "0000000000") & " 1_CAD-Adressabgleich Adressenabfrage Mandant " & Format(yearend, "yyyyMMdd") & ".xlsx"
                    wbTemplate.Close
                    objApp.Quit
                    mailItem.Move myInbox.Folders(moveToFolder)
                    Dim timeNow As String
                    timeNow = Format(Now(), "yyyy-MM-dd hh:mm:ss")
                    Call updateSQL("UPDATE tCON_Orderbook Set AC_Status = '" & newStatus & "' WHERE OrderNo = '" & OrderNo & "'")
                                        If newStatus = "InputDataAvailable" Then
                                                Call updateSQL("UPDATE tCON_Orderbook Set tsInputDataAvailable = '" & timeNow & "' WHERE OrderNo = '" & OrderNo & "' AND tsInputDataAvailable IS NULL")
                                        End If
                End If
            Else
JM_strangeReply:
                Debug.Print "Anhang umbenannt vom Team?!"
26          End If
            Set objAtt = Nothing
        End If
    End If
End Sub
Sub replaceAll(wsAdresscheck As Object, what As String, replacement As String)

   wsAdresscheck.Cells.Replace what:=what, replacement:=replacement, _
   LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, _
   SearchFormat:=False, ReplaceFormat:=False
End Sub



