VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private WithEvents inboxItems As Outlook.Items
Attribute inboxItems.VB_VarHelpID = -1
Private WithEvents inboxItemsSwitch As Outlook.Items ' added for testing
Attribute inboxItemsSwitch.VB_VarHelpID = -1

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
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

Private Sub Application_Startup()

          Dim I As Integer

          Dim indexAccount As Integer
1         For I = 1 To Application.Session.Accounts.Count
2             If LCase(Application.Session.Accounts.item(I)) = "adressabgleich@de.ey.com" Then
3                 indexAccount = I
4                 Exit For
5             End If
6         Next

          Dim accIND As Object, myInbox As Object, olFolder As Object, olFolderSwitch As Object ' last one added for testing
7         Set accIND = Application.Session.Accounts.item(indexAccount)
8         Set myInbox = accIND.DeliveryStore.GetDefaultFolder(6)
9         Set olFolder = myInbox.Folders("Switch")
10        Set olFolderSwitch = olFolder.Folders("On")
11        Set inboxItems = olFolderSwitch.Items
                    
End Sub

Private Sub inboxItems_ItemAdd(ByVal item As Object)

1         If TypeName(item) = "MailItem" Then
2             Call RunMain
3         End If

End Sub

Sub RunMain()
        Call updateSQL("Insert into [CAD].[dbo].[tAC_Monitoring] ([Anstoﬂen]) values ('" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "');")
        
        Dim I As Integer
        Dim item As Object
        
         Dim indexAccount As Integer
         For I = 1 To Application.Session.Accounts.Count
             If LCase(Application.Session.Accounts.item(I)) = "adressabgleich@de.ey.com" Then
                 indexAccount = I
                 Exit For
             End If
         Next

          Dim accIND As Object, myInbox As Object, olFolder As Object, mail As Object
         Set accIND = Application.Session.Accounts.item(indexAccount)
         Set myInbox = accIND.DeliveryStore.GetDefaultFolder(6)
         Set olFolder = myInbox.Folders("Switch")

        ' Sollte nur EINE email enthalten
        For Each item In olFolder.Folders("On").Items
            Set mail = item
        Next

        ' AC
        Call SendNotification
        Call readDatenSammler
        Call process_TeamApproval

        
        ' CON
        Call SendNotification_CON
        Call readDatenSammler_CON
        Call readDatenSammlerAuditi_CON
        Call process_TeamApproval_CON
        
        Dim fso As Object
        
        mail.Move olFolder.Folders("Off")
        Call updateSQL("update [CAD].[dbo].[tAC_Monitoring] set [SwitchOFF] = '" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "' where [SwitchOFF] is null;")
        
End Sub




