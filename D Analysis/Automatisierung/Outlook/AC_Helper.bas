Attribute VB_Name = "AC_Helper"
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
Sub readDatenSammler()

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
9         Set olFolder = myInbox.Folders("Team Reply AC")

          ' read e-mails in Confi-Inbox
          Dim c_mails As Long
          Dim mails() As Object
10        If olFolder.Items.Count >= 1 Then ReDim mails(1 To olFolder.Items.Count)
11        c_mails = 0
12        For Each item In olFolder.Items
14            c_mails = c_mails + 1
15            Set mails(c_mails) = item
16        Next
    
          Dim appAccess As Object
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
              bSendToAuditi = True
              ' Welche Art von Confirmation ist es
              subject = mails(I).subject
              If InStr(subject, "Beauftragung Adressabgleich") > 0 Or InStr(subject, "Order of debtor") > 0 Then
                  subj = Split(subject, " ")
                
                  If InStr(subject, "Beauftragung Adressabgleich") > 0 Then
                      OrderNo = subj(3)
                  Else
                      OrderNo = subj(4)
                  End If
                
                  ' Check, ob Ordernummer schon storniert
                  Set rsOrderbook = Connection.Execute("SELECT * FROM tAC_Orderbook WHERE [OrderNo] = '" & OrderNo & "'")
                  Do While rsOrderbook.EOF = False
                      If IsNull(rsOrderbook("tsStornoSent")) = False Then
                          ' Keine Dateien an Auditi senden
                          bSendToAuditi = False
                          Exit Do
                      End If
                      rsOrderbook.MoveNext
                  Loop
                  rsOrderbook.Close
                  
                  Set rsOrderbook = Connection.Execute("SELECT * FROM tAC_Orderbook WHERE [OrderNo] = '" & OrderNo & "'")
                  Do While rsOrderbook.EOF = False
                      If rsOrderbook("AC_Status") = "Storno" Or rsOrderbook("AC_Status") = "StornoSent" Then
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
                      Debug.Print mails(I).Attachments.Count
                      If mails(I).Attachments.Count > 0 Then
    
                          Set objAtt = mails(I).Attachments.item(1)
22                        If InStr(objAtt.FileName, "1_CAD-Adressabgleich Adressenabfrage Mandant") > 0 Or InStr(objAtt.FileName, "Debitoren Einzelposten") > 0 Then
                                

                              'objApp.DisplayAlerts = False
                              Debug.Print saveFolder & "2. CAD_Abgleich\" & objAtt.FileName
                              
                              objAtt.SaveAsFile saveFolder & "2. CAD_Abgleich\" & objAtt.FileName
                              'objApp.DisplayAlerts = True
                              
                              mails(I).Move myInbox.Folders("Team Reply Processed")
                              Dim timeNow As String
                              timeNow = Format(Now(), "yyyy-MM-dd hh:mm:ss")
                              Call updateSQL("UPDATE tAC_Orderbook Set AC_Status = 'InputDataReceived' WHERE OrderNo = '" & OrderNo & "'")
                              Call updateSQL("UPDATE tAC_Orderbook Set tsInputDataReceived = '" & timeNow & "' WHERE OrderNo = '" & OrderNo & "'")
                                        
                          End If
                                    
                      Else
                          Debug.Print "Anhang umbenannt vom Team?!"
26                    End If
                      Set objAtt = Nothing
                  End If
              End If
         
         
          Next I
         
          Connection.Close
          Set rsOrderbook = Nothing
End Sub
Sub process_TeamApproval()

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

          Set olFolder = myInbox.Folders("Team Approval AC")
    
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
              Call updateSQL("UPDATE tAC_Orderbook Set AC_Status = 'TeamApprovalReceived' WHERE OrderNo = '" & OrderNo & "'")
              Call updateSQL("UPDATE tAC_Orderbook Set tsTeamApprovalReceived = '" & timeNow & "' WHERE OrderNo = '" & OrderNo & "'")
    
          Next I

End Sub

Sub SendNotification()

Dim I As Integer, j As Integer
Dim myInbox As Object, item As Object, objAtt As Object, accIND As Object, newOrderFolder As Object

Dim RecordSet As ADODB.RecordSet
Dim Connection As ADODB.Connection

Set Connection = CreateObject("ADODB.Connection")
Connection.ConnectionString = "Provider=SQLNCLI11;Data Source=DERUSCMPDWASQ01.ey.net\INST02; Initial Catalog=CAD;Integrated Security=SSPI;DataTypeCompatibility=80;"
Connection.Open


Dim indexAccount As Integer
For I = 1 To Outlook.Application.Session.Accounts.Count
If LCase(Outlook.Application.Session.Accounts.item(I)) = "adressabgleich@de.ey.com" Then
    indexAccount = I
Exit For
End If
Next

Set accIND = Outlook.Application.Session.Accounts.item(indexAccount)
Set myInbox = accIND.DeliveryStore.GetDefaultFolder(6)
Set newOrderFolder = myInbox.Folders("Team Reply AC")

Dim c_mails As Long
Dim mails() As Object

If newOrderFolder.Items.Count >= 1 Then ReDim mails(1 To newOrderFolder.Items.Count)

Dim allmails As Outlook.Items
Set allmails = newOrderFolder.Items
allmails.Sort "[ReceivedTime]"

Dim DupItem As Object
Set DupItem = CreateObject("Scripting.Dictionary")

For I = allmails.Count To 1 Step -1
        If TypeOf allmails(I) Is MailItem Then
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
    OrderNo = Right(mails(I).subject, 12)
    Debug.Print mails(I).subject
    Set rsOrderbook = Connection.Execute("SELECT [AC_Status], [AC_Preparer], [client], [GISID] FROM tAC_Orderbook WHERE [OrderNo] = '" & OrderNo & "'")
    
    Do While rsOrderbook.EOF = False
        If rsOrderbook("AC_Status") = "" Or rsOrderbook("AC_Status") = "InputDataSent" Or rsOrderbook("AC_Status") = "WaitingforInputData" Then
            Debug.Print ("Continue")
        ElseIf rsOrderbook("AC_Status") = "InputDataReceived" Or rsOrderbook("AC_Status") = "InProgress" Or rsOrderbook("AC_Status") = "ReadyForReview" Then
            mails(I).Move myInbox.Folders("Team Reply Nachlieferung AC")
            If IsNull(rsOrderbook("AC_Preparer")) = False And Not (rsOrderbook("AC_Preparer") = "") Then
                preparer = rsOrderbook("AC_Preparer")
                preparer = Replace(preparer, "ä", "ae")
                preparer = Replace(preparer, "ö", "oe")
                preparer = Replace(preparer, "ü", "ue")
                replaced = Replace(preparer, " ", "")
                Count = Len(preparer) - Len(replaced)
                If Count > 1 Then
                    givenName = Left(preparer, InStr(preparer, " ") - 1)
                    restName = Mid(preparer, InStr(preparer, " ") + 1)
                    MiddleName = Left(restName, 1)
                    surname = Mid(restName, InStr(preparer, " "))
                    preparerEmail = givenName + "." + MiddleName + "." + surname + "@de.ey.com"
                Else
                    givenName = Left(preparer, InStr(preparer, " ") - 1)
                    surname = Mid(preparer, InStr(preparer, " ") + 1)
                    preparerEmail = givenName + "." + surname + "@de.ey.com"
                End If
                gisid = rsOrderbook("GISID")
                client = rsOrderbook("client")
                
                body = parseBody(pathMailTemplates & templateName)
                body = Replace(body, "[orderNo]", OrderNo)
                body = Replace(body, "[GISID]", gisid)
                body = Replace(body, "[client]", client)
                
                Set OutMail = Outlook.Application.CreateItem(olMailItem)
                OutMail.To = preparerEmail
                OutMail.HTMLbody = body
                OutMail.SentOnBehalfOfName = "adressabgleich@de.ey.com"
                OutMail.subject = "Neue Adressdaten: " + OrderNo
                OutMail.Send
            End If
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



