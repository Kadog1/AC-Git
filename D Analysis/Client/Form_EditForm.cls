VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub BtnClose_Click()

DoCmd.Close acForm, Me.name

End Sub


Private Sub BtnSaveClose_Click()

Dim OrderNo As String, NameBank As String, sql As String, Tool As String
Dim OrderbookRS As ADODB.RecordSet, BankenRS As ADODB.RecordSet
Dim anweisungsmail As Boolean, bestaetigungsmail As Boolean, rechnung As Boolean, betrag As Boolean
Dim einladung As Integer, auswahl As Integer, feedback As Integer, req2ndW As Integer
    
' alle Datumseingaben konvertieren für Access: DD.MM.YYYY -> MM/DD/YYYY
Dim DoraWeitergeleitet As String, DatumAbrechnung As String
Dim DatumErstkontakt As String, DatumAuswahlErhalten As String, EingangsdatumCAD As String, ScndRequest As String, FinalReview As String, VersanddatumOriginale As String, Abrechnungsdatum As String, DatumFeedbackabfrage As String, DatumReminder As String, DatumAuftragsende As String
Dim UploadGew As String, AntwortDeadline As String, dateconfi As String

If MsgBox("Wollen Sie die eingegebenen Daten in der Datenbank speichern?", vbYesNo) = vbYes Then
    OrderNo = TxtOrderNo.value
    Tool = getTableValue(OrderNo, "Tool")
    Set OrderbookRS = getRS("SELECT * From tCon_Orderbook" & strTestTable & " WHERE OrderNo = '" & OrderNo & "'")
    Do While OrderbookRS.EOF = False
        If IsNull(TxtDateConfi.value) = False Then
            'DateConfi = Mid(TxtDateConfi.Value, 4, 2) & "/" & Mid(TxtDateConfi.Value, 1, 2) & "/" & Right(TxtDateConfi.Value, 4)
            sql = "UPDATE tCON_Orderbook" & strTestTable & " SET DateConfi = " & handleSQLValues(TxtDateConfi.value) & " WHERE OrderNo='" & OrderNo & "';"
            Call updateSQL(sql)
        Else
            sql = "UPDATE tCON_Orderbook" & strTestTable & " SET DateConfi = null WHERE OrderNo='" & OrderNo & "';"
            Call updateSQL(sql)
        End If
        sql = "UPDATE tCON_Orderbook" & strTestTable & " SET GISID = " & handleSQLValues(TxtGISID.value) & ", Client = " & handleSQLValues(TxtClient.value) & ", BusinessEstablishmentOAre = " & handleSQLValues(TxtBusinessEstablishment.value) & "," _
                & "TypeFinancials = " & handleSQLValues(TxtTypeFinancials.value) & ", YearEnd = " & handleSQLValues(TxtYearEnd.value) & ", NameOfCompanyGroup = " & handleSQLValues(TxtGroup.value) & ", EngCode = " & handleSQLValues(TxtEngCode.value) & ", EngName = " & handleSQLValues(TxtEngName.value) & ", EngPartner = " & handleSQLValues(TxtEngPartner.value) & ", EngManager = " & handleSQLValues(TxtManager.value) & "," _
                & "OtherContact = " & handleSQLValues(TxtOtherContact.value) & ", EngContact = " & handleSQLValues(TxtEngContact.value) & ", Lang = " & handleSQLValues(CblSprache.value) & ", Tool = " & handleSQLValues(CblTool.value) & ", Confirmation = " & handleSQLValues(CblConfirmation.value) & ", TypeConfi = " & handleSQLValues(CblTypeConfi.value) & ", PlannedDelivery = " & handleSQLValues(TxtPlannedDelivery.value) & "," _
                & "CAD_Bearbeiter = " & handleSQLValues(Cbl1CADPreparer.value) & ", CAD_Bearbeiter_vorherig = " & handleSQLValues(Cbl2CADPreparer.value) & ", DeliveryType = " & handleSQLValues(CblDeliveryType.value) & ", SB_Status = " & handleSQLValues(CblJobStatus.value) & " WHERE OrderNo='" & OrderNo & "';"
        Call updateSQL(sql)
    
        If OrderbookRS("Confirmation") = "BankenPortal" Then
            NameBank = TxtNameBank.value
            Set BankenRS = getRS("SELECT * From tCON_BankConfirmation" & strTestTable & " WHERE OrderNo = '" & OrderNo & "' AND Name_Bank = '" & NameBank & "'")
            Do While BankenRS.EOF = False
                If CblAnweisungsmail.value = "Ja" Then
                    anweisungsmail = 1
                Else
                    anweisungsmail = 0
                End If
                If CblBestätigungsmail.value = "Ja" Then
                    bestaetigungsmail = 1
                Else
                    bestaetigungsmail = 0
                End If
                If CblGeldZurückerstattet.value = "Ja" Then
                    rechnung = 1
                Else
                    rechnung = 0
                End If
                
                If IsNull(TxtAnDora.value) = False Then
                    'DoraWeitergeleitet = Mid(TxtAnDora.Value, 4, 2) & "/" & Mid(TxtAnDora.Value, 1, 2) & "/" & Right(TxtAnDora.Value, 4)
                    sql = "UPDATE tCON_BankConfirmation" & strTestTable & " SET DoraAbrechnung = " & handleSQLValues(TxtAnDora.value) & " WHERE OrderNo='" & OrderNo & "' AND Name_Bank = '" & NameBank & "';"
                    Call updateSQL(sql)
                Else
                    sql = "UPDATE tCON_BankConfirmation" & strTestTable & " SET DoraAbrechnung = null WHERE OrderNo='" & OrderNo & "' AND Name_Bank = '" & NameBank & "';"
                    Call updateSQL(sql)
                End If
                If IsNull(TxtDatumAbrechnung.value) = False Then
                    'DatumAbrechnung = Mid(TxtDatumAbrechnung.Value, 4, 2) & "/" & Mid(TxtDatumAbrechnung.Value, 1, 2) & "/" & Right(TxtDatumAbrechnung.Value, 4)
                    sql = "UPDATE tCON_BankConfirmation" & strTestTable & " SET Datum_Abrechung = " & handleSQLValues(TxtDatumAbrechnung.value) & " WHERE OrderNo='" & OrderNo & "' AND Name_Bank = '" & NameBank & "';"
                    Call updateSQL(sql)
                Else
                    sql = "UPDATE tCON_BankConfirmation" & strTestTable & " SET Datum_Abrechung = null WHERE OrderNo='" & OrderNo & "' AND Name_Bank = '" & NameBank & "';"
                    Call updateSQL(sql)
                End If
                sql = "UPDATE tCON_BankConfirmation" & strTestTable & " SET Portal_Bank_Status = " & handleSQLValues(CblPortalBankStatus.value) & ", Anweisungsmail_Authorizer = " & anweisungsmail & ", Bestätigungsmail_Auth_vor_Init = " & bestaetigungsmail & ", Anzahl_Forms = " & handleSQLValues(TxtAnzahlForms.value) & ", Kosten_USD = " & handleSQLValues(TxtKosten.value) & ", Geld_erstattet = " & rechnung & ", Portal_Bank_geb_Stunden = " & handleSQLValues(TxtGebuchteStunden.value) & ", Portal_Bank_Kommentar = " & handleSQLValues(TxtComments.value) & " WHERE OrderNo='" & OrderNo & "' AND Name_Bank = '" & NameBank & "';"
                Call updateSQL(sql)
                
                If IsNull(TxtDatumFeedbackabfrage.value) = False Then
                    'DatumFeedbackabfrage = Mid(TxtDatumFeedbackabfrage.Value, 4, 2) & "/" & Mid(TxtDatumFeedbackabfrage.Value, 1, 2) & "/" & Right(TxtDatumFeedbackabfrage.Value, 4)
                    sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_Feedbackanfrage = " & handleSQLValues(TxtDatumFeedbackabfrage.value) & " WHERE OrderNo='" & OrderNo & "';"
                    Call updateSQL(sql)
                Else
                    sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_Feedbackanfrage = null WHERE OrderNo='" & OrderNo & "';"
                    Call updateSQL(sql)
                End If
                If IsNull(TxtDatumReminder.value) = False Then
                    'DatumReminder = Mid(TxtDatumReminder.Value, 4, 2) & "/" & Mid(TxtDatumReminder.Value, 1, 2) & "/" & Right(TxtDatumReminder.Value, 4)
                    sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_Reminder = " & handleSQLValues(TxtDatumReminder.value) & " WHERE OrderNo='" & OrderNo & "';"
                    Call updateSQL(sql)
                Else
                    sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_Reminder = null WHERE OrderNo='" & OrderNo & "';"
                    Call updateSQL(sql)
                End If
                If CblFeedbackErhalten.value = "Ja" Then
                   feedback = True
               Else
                   feedback = False
               End If
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Feedback_erhalten = " & feedback & " WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
        
                ' ########################################################################################################################################################
                ' Es wird nur eine Bestellung abgearbeitet -> Falls auch der gleiche Namen der Bank eingegeben wird: PROBLEM
                ' ########################################################################################################################################################
                Exit Do
            Loop
            BankenRS.Close
        Else
            If Tool = "eConfirmations" Then
                If CblAuswahlErhalten.value = "Ja" Then
                    auswahl = 1
                Else
                    auswahl = 0
                End If
            End If
            
            If CblFeedbackErhalten.value = "Ja" Then
                feedback = 1
            Else
                feedback = 0
            End If
            
            If Cbl2ndRequest.value = "Ja" Then
                req2ndW = 1
            Else
                req2ndW = 0
            End If
            
            DoCmd.SetWarnings False
            
            If IsNull(TxtAuftragsende.value) = False Then
                'DatumAuftragsende = Mid(TxtAuftragsende.Value, 4, 2) & "/" & Mid(TxtAuftragsende.Value, 1, 2) & "/" & Right(TxtAuftragsende.Value, 4)
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET DatumAuftragsende = " & handleSQLValues(TxtAuftragsende.value) & " WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            Else
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET DatumAuftragsende = null WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            End If
            
            If IsNull(TxtErstkontakt.value) = False Then
                'DatumErstkontakt = Mid(TxtErstkontakt.Value, 4, 2) & "/" & Mid(TxtErstkontakt.Value, 1, 2) & "/" & Right(TxtErstkontakt.Value, 4)
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_Versand_EMail_Erstkontakt = " & handleSQLValues(TxtErstkontakt.value) & " WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            Else
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_Versand_EMail_Erstkontakt = null WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            End If
            
            If IsNull(TxtEingangsdatum.value) = False Then
                'EingangsdatumCAD = Mid(TxtEingangsdatum.Value, 4, 2) & "/" & Mid(TxtEingangsdatum.Value, 1, 2) & "/" & Right(TxtEingangsdatum.Value, 4)
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Eingangsdatum_CAD = " & handleSQLValues(TxtEingangsdatum.value) & " WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            Else
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Eingangsdatum_CAD = null WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            End If
            
            If IsNull(TxtVersanddatum.value) = False Then
                'VersanddatumCAD = Mid(TxtVersanddatum.Value, 4, 2) & "/" & Mid(TxtVersanddatum.Value, 1, 2) & "/" & Right(TxtVersanddatum.Value, 4)
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Versanddatum_CAD = " & handleSQLValues(TxtVersanddatum.value) & " WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            Else
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Versanddatum_CAD = null WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            End If
            
            If IsNull(TxtDatum2ndRequest.value) = False Then
                'ScndRequest = Mid(TxtDatum2ndRequest.Value, 4, 2) & "/" & Mid(TxtDatum2ndRequest.Value, 1, 2) & "/" & Right(TxtDatum2ndRequest.Value, 4)
                '#JK-11/2020-12-16 START
                sql = "UPDATE Orderbook SET Date2ndReq = '" & handleSQLValues(TxtDatum2ndRequest.value) & "' WHERE OrderNo='" & OrderNo & "';"
                '#JK-11/2020-12-16 END
                Call updateSQL(sql)
            Else
                '#JK-12/2020-12-16 START
                sql = "UPDATE Orderbook SET Date2ndReq = null WHERE OrderNo='" & OrderNo & "';"
                '#JK-12/2020-12-16 END
                Call updateSQL(sql)
            End If
            
            If IsNull(TxtDatumFinalReview.value) = False Then
                'FinalReview = Mid(TxtDatumFinalReview.Value, 4, 2) & "/" & Mid(TxtDatumFinalReview.Value, 1, 2) & "/" & Right(TxtDatumFinalReview.Value, 4)
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_final_Review = " & handleSQLValues(TxtDatumFinalReview.value) & " WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            Else
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_final_Review = null WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            End If
            
            If IsNull(TxtVersanddatumOriginale.value) = False Then
                'VersanddatumOriginale = Mid(TxtVersanddatumOriginale.Value, 4, 2) & "/" & Mid(TxtVersanddatumOriginale.Value, 1, 2) & "/" & Right(TxtVersanddatumOriginale.Value, 4)
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Versanddatum_Originale = " & handleSQLValues(TxtVersanddatumOriginale.value) & " WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            Else
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Versanddatum_Originale = null WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            End If
            
            If IsNull(TxtDatumFeedbackabfrage.value) = False Then
                'DatumFeedbackabfrage = Mid(TxtDatumFeedbackabfrage.Value, 4, 2) & "/" & Mid(TxtDatumFeedbackabfrage.Value, 1, 2) & "/" & Right(TxtDatumFeedbackabfrage.Value, 4)
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_Feedbackanfrage = " & handleSQLValues(TxtDatumFeedbackabfrage.value) & " WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            Else
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_Feedbackanfrage = null WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            End If
            
            If IsNull(TxtDatumReminder.value) = False Then
                'DatumReminder = Mid(TxtDatumReminder.Value, 4, 2) & "/" & Mid(TxtDatumReminder.Value, 1, 2) & "/" & Right(TxtDatumReminder.Value, 4)
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_Reminder = " & handleSQLValues(TxtDatumReminder.value) & " WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            Else
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_Reminder = null WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            End If
            
            If IsNull(TxtUploadGewünscht.value) = False Then
                'UploadGew = Mid(TxtUploadGewünscht.Value, 4, 2) & "/" & Mid(TxtUploadGewünscht.Value, 1, 2) & "/" & Right(TxtUploadGewünscht.Value, 4)
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET First_Upload_Gewünscht = " & handleSQLValues(TxtUploadGewünscht.value) & " WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            Else
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET First_Upload_Gewünscht = null WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            End If
            
            If IsNull(TxtAntwortDeadline.value) = False Then
                'AntwortDeadline = Mid(TxtAntwortDeadline.Value, 4, 2) & "/" & Mid(TxtAntwortDeadline.Value, 1, 2) & "/" & Right(TxtAntwortDeadline.Value, 4)
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Antwort_Deadline = " & handleSQLValues(TxtAntwortDeadline.value) & " WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            Else
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Antwort_Deadline = null WHERE OrderNo='" & OrderNo & "';"
                Call updateSQL(sql)
            End If
            
            If Tool = "eConfirmations" Then
                If IsNull(TxtDatumAuswahlErhalten.value) = False Then
                    'DatumAuswahlErhalten = Mid(TxtDatumAuswahlErhalten.Value, 4, 2) & "/" & Mid(TxtDatumAuswahlErhalten.Value, 1, 2) & "/" & Right(TxtDatumAuswahlErhalten.Value, 4)
                    sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_Auswahl_erhalten = " & handleSQLValues(TxtDatumAuswahlErhalten.value) & " WHERE OrderNo='" & OrderNo & "';"
                    Call updateSQL(sql)
                Else
                    sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Datum_Auswahl_erhalten = null WHERE OrderNo='" & OrderNo & "';"
                    Call updateSQL(sql)
                End If
                If IsNull(TxtAbrechnungsdatum.value) = False Then
                    'Abrechnungsdatum = Mid(TxtAbrechnungsdatum.Value, 4, 2) & "/" & Mid(TxtAbrechnungsdatum.Value, 1, 2) & "/" & Right(TxtAbrechnungsdatum.Value, 4)
                    sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Abrechnungsdatum = " & handleSQLValues(TxtAbrechnungsdatum.value) & " WHERE OrderNo='" & OrderNo & "';"
                    Call updateSQL(sql)
                Else
                    sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Abrechnungsdatum = null WHERE OrderNo='" & OrderNo & "';"
                    Call updateSQL(sql)
                End If
                sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Anzahl_erhaltene_Salden = " & handleSQLValues(TxtAnzahlErhalteneSalden.value) & ", Informationen_zum_Auftrag = " & handleSQLValues(TxtInformationenZumAuftrag.value) & ", Anzahl_Bestätigungen = " & handleSQLValues(TxtAnzahlBestätigungen.value) & ", Auswahl_erhalten = " & auswahl & "," _
                    & "Versand_durch = " & handleSQLValues(cblVersanddurch.value) & ", Kommentar_Versand = " & handleSQLValues(TxtKommentarZumVersand.value) & ", Final_Senior_Review_durch = " & handleSQLValues(CblFinalReview.value) & ", Empfänger_Originale = " & handleSQLValues(TxtEmpfängerOriginale.value) & ", [2nd_Request_gewünscht] = " & req2ndW & ", Empfänger_Nachläufer = " & handleSQLValues(TxtEmpfängerNachläufer.value) & ", Abrechnungsbetrag = " & handleSQLValues(TxtAbrechnungsbetrag.value) & ", Feedback_erhalten = " & feedback & " WHERE OrderNo='" & OrderNo & "';"
            Else
                If OrderbookRS("Confirmation") = "Sonstige" Then
                    sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Anzahl_erhaltene_Salden = " & handleSQLValues(TxtAnzahlErhalteneSalden.value) & ", Informationen_zum_Auftrag = " & handleSQLValues(TxtInformationenZumAuftrag.value) & ", Anzahl_Bestätigungen = " & handleSQLValues(TxtAnzahlBestätigungen.value) & "," _
                        & "Versand_durch = " & handleSQLValues(cblVersanddurch.value) & ", Kommentar_Versand = " & handleSQLValues(TxtKommentarZumVersand.value) & ", Final_Senior_Review_durch = " & handleSQLValues(CblFinalReview.value) & ", OtherTypeConfi = " & handleSQLValues(TxtOtherTypeConfi.value) & ", Empfänger_Originale = " & handleSQLValues(TxtEmpfängerOriginale.value) & ", [2nd_Request_gewünscht] = " & req2ndW & ", Empfänger_Nachläufer = " & handleSQLValues(TxtEmpfängerNachläufer.value) & ", Feedback_erhalten = " & feedback & " WHERE OrderNo='" & OrderNo & "';"
                Else
                    sql = "UPDATE tCON_Orderbook" & strTestTable & " SET Anzahl_erhaltene_Salden = " & handleSQLValues(TxtAnzahlErhalteneSalden.value) & ", Informationen_zum_Auftrag = " & handleSQLValues(TxtInformationenZumAuftrag.value) & ", Anzahl_Bestätigungen = " & handleSQLValues(TxtAnzahlBestätigungen.value) & "," _
                        & "Versand_durch = " & handleSQLValues(cblVersanddurch.value) & ", Kommentar_Versand = " & handleSQLValues(TxtKommentarZumVersand.value) & ", Final_Senior_Review_durch = " & handleSQLValues(CblFinalReview.value) & ", Empfänger_Originale = " & handleSQLValues(TxtEmpfängerOriginale.value) & ", [2nd_Request_gewünscht] = " & req2ndW & ", Empfänger_Nachläufer = " & handleSQLValues(TxtEmpfängerNachläufer.value) & ", Feedback_erhalten = " & feedback & " WHERE OrderNo='" & OrderNo & "';"
                End If
            End If
            Call updateSQL(sql)
            'DoCmd.SetWarnings True
        
        End If
        ' Es sollte nur eine Ordernummer geben
        Exit Do
    Loop
    OrderbookRS.Close
    DoCmd.Close acForm, Me.name
    
End If

End Sub

Private Sub Form_DataChange(ByVal Reason As Long)

End Sub

Private Sub Form_Open(Cancel As Integer)

Dim OrderNo As String, confirmation As String, NameBank As String, Tool As String, NamePortal As String, bEYstand As String
Dim BankenRS As ADODB.RecordSet, OrderbookRS As ADODB.RecordSet
Dim TaskEditQuery As String

If IsNull(Forms![Update_Confi Client].Alltasks.Column(0)) = False Then
    
    ' Lade alle Daten aus dem Orderbook
    OrderNo = Forms![Update_Confi Client].Alltasks.Column(0)
    confirmation = getTableValue(OrderNo, "Confirmation")
    Tool = getTableValue(OrderNo, "Tool")
    bEYstand = getTableValue(OrderNo, "EYStand")
    
    ' Ein/Ausblenden der Banken-Daten
    If confirmation = "BankenPortal" Then
        ' Lese zu bearbeitende Bestellung aus ListStammdaten (zu einer Ordernummer kann es mehrere Banken geben)
        If IsNull(Forms![Update_Confi Client].ListStammdaten.Column(0)) = False Then
            LblDeliveryType.Visible = False
            CblDeliveryType.Visible = False
            LblDateConfi.Visible = False
            TxtDateConfi.Visible = False
            LblErstkontakt.Visible = False
            TxtErstkontakt.Visible = False
            LblInformationenZumAuftrag.Visible = False
            TxtInformationenZumAuftrag.Visible = False
            LblVersandDurch.Visible = False
            cblVersanddurch.Visible = False
            LblPlannedDelivery.Visible = False
            TxtPlannedDelivery.Visible = False
            LblEingangsdatum.Visible = False
            TxtEingangsdatum.Visible = False
            LblVersanddatum.Visible = False
            TxtVersanddatum.Visible = False
            LblAnzahlBestätigungen.Visible = False
            TxtAnzahlBestätigungen.Visible = False
            LblKommentarVersand.Visible = False
            TxtKommentarZumVersand.Visible = False
            LblUploadGewünscht.Visible = False
            TxtUploadGewünscht.Visible = False
            LblAntwortDeadline.Visible = False
            TxtAntwortDeadline.Visible = False
            Lbl2ndRequest.Visible = False
            Cbl2ndRequest.Visible = False
            LblDatum2ndRequest.Visible = False
            TxtDatum2ndRequest.Visible = False
            LblAuftragsende.Visible = False
            TxtAuftragsende.Visible = False
            LblAnzahlErhalteneSalden.Visible = False
            TxtAnzahlErhalteneSalden.Visible = False
            LblFinalReview.Visible = False
            CblFinalReview.Visible = False
            LblDatumFinalReview.Visible = False
            TxtDatumFinalReview.Visible = False
            LblVersanddatumOriginale.Visible = False
            TxtVersanddatumOriginale.Visible = False
            LblEmpfängerOriginale.Visible = False
            TxtEmpfängerOriginale.Visible = False
            LblEmpfängerNachläufer.Visible = False
            TxtEmpfängerNachläufer.Visible = False
            Lbl1CADPreparer.Visible = False
            Cbl1CADPreparer.Visible = False
            Lbl2CADPreparer.Visible = False
            Cbl2CADPreparer.Visible = False
            LblJobStatus.Visible = False
            CblJobStatus.Visible = False

            ' Alle BankenPortal-Felder invisible machen
            LblNameBank.Visible = True
            TxtNameBank.Visible = True
            LblNamePortal.Visible = True
            TxtNamePortal.Visible = True
            LblPortalBankStatus.Visible = True
            CblPortalBankStatus.Visible = True
            LblAnweisungsmail.Visible = True
            CblAnweisungsmail.Visible = True
            LblBestätigungsmail.Visible = True
            CblBestätigungsmail.Visible = True
            LblAnzahlForms.Visible = True
            TxtAnzahlForms.Visible = True
            LblKosten.Visible = True
            TxtKosten.Visible = True
            LblAnDora.Visible = True
            TxtAnDora.Visible = True
            LblDatumAbrechnung.Visible = True
            TxtDatumAbrechnung.Visible = True
            LblGeldZurückerstattet.Visible = True
            CblGeldZurückerstattet.Visible = True
            LblGebuchteStunden.Visible = True
            TxtGebuchteStunden.Visible = True
            LblComments.Visible = True
            TxtComments.Visible = True
        
            OrderNo = Forms![Update_Confi Client].ListStammdaten.Column(0)
            NameBank = Forms![Update_Confi Client].ListStammdaten.Column(11)
            NamePortal = Forms![Update_Confi Client].ListStammdaten.Column(12)
            Set BankenRS = getRS("SELECT * From tCON_BankConfirmation" & strTestTable & " WHERE OrderNo = '" & OrderNo & "' AND Name_Bank = '" & NameBank & "'")
            Do While BankenRS.EOF = False
                TxtNameBank.value = BankenRS("Name_Bank")
                TxtNamePortal.value = BankenRS("Portal_Bank")
                If IsNull(BankenRS("DoraAbrechnung")) = False Then TxtAnDora.value = BankenRS("DoraAbrechnung")
                If IsNull(BankenRS("Portal_Bank_Status")) = False Then CblPortalBankStatus.value = BankenRS("Portal_Bank_Status")
                If IsNull(BankenRS("Anweisungsmail_Authorizer")) = False Then
                    If BankenRS("Anweisungsmail_Authorizer") = True Then
                        CblAnweisungsmail.value = "Ja"
                    Else
                        CblAnweisungsmail.value = "Nein"
                    End If
                End If
                If IsNull(BankenRS("Bestätigungsmail_Auth_vor_Init")) = False Then
                    If BankenRS("Bestätigungsmail_Auth_vor_Init") = True Then
                        CblBestätigungsmail.value = "Ja"
                    Else
                        CblBestätigungsmail.value = "Nein"
                    End If
                End If
                If IsNull(BankenRS("Anzahl_Forms")) = False Then TxtAnzahlForms.value = BankenRS("Anzahl_Forms")
                If IsNull(BankenRS("Kosten_USD")) = False Then TxtKosten.value = BankenRS("Kosten_USD")
                If IsNull(BankenRS("Geld_erstattet")) = False Then
                    If BankenRS("Geld_erstattet") = True Then
                        CblGeldZurückerstattet.value = "Ja"
                    Else
                        CblGeldZurückerstattet.value = "Nein"
                    End If
                End If
                If IsNull(BankenRS("Datum_Abrechung")) = False Then TxtDatumAbrechnung.value = BankenRS("Datum_Abrechung")
                If IsNull(BankenRS("Portal_Bank_geb_Stunden")) = False Then TxtGebuchteStunden.value = BankenRS("Portal_Bank_geb_Stunden")
                If IsNull(BankenRS("Portal_Bank_Kommentar")) = False Then TxtComments.value = BankenRS("Portal_Bank_Kommentar")
                ' ########################################################################################################################################################
                ' Es wird nur eine Bestellung abgearbeitet -> Falls auch der gleiche Namen der Bank eingegeben wird: PROBLEM
                ' ########################################################################################################################################################
                Exit Do
            Loop
            BankenRS.Close
        Else
            MsgBox "Bitte wählen Sie eine Bestellung in der Stammdaten-Übersicht aus!"
            DoCmd.Close acForm, Me.name
            Exit Sub
        End If
    End If
    
    If Tool = "eConfirmations" Then
        CblAuswahlErhalten.Visible = True
        TxtDatumAuswahlErhalten.Visible = True
        LblAuswahlErhalten.Visible = True
        LblDatumAuswahlErhalten.Visible = True
        TxtAbrechnungsbetrag.Visible = True
        TxtAbrechnungsdatum.Visible = True
        LblAbrechnungsbetrag.Visible = True
        LblAbrechnungsdatum.Visible = True
        LbleConfirmationsSent.Visible = True
        TxteConfirmationsSent.Visible = True
        
        CblDeliveryType.Enabled = False
        TxtDateConfi.Enabled = False
        TxtErstkontakt.Enabled = False
        TxtInformationenZumAuftrag.Enabled = False
        cblVersanddurch.Enabled = False
        TxtPlannedDelivery.Enabled = False
        TxtEingangsdatum.Enabled = False
        TxtVersanddatum.Enabled = False
        TxtAnzahlBestätigungen.Enabled = False
        TxtKommentarZumVersand.Enabled = False
        TxtUploadGewünscht.Enabled = False
        TxtAntwortDeadline.Enabled = False
        Cbl2ndRequest.Enabled = False
        TxtDatum2ndRequest.Enabled = False
        TxtAuftragsende.Enabled = False
        TxtAnzahlErhalteneSalden.Enabled = False
        CblFinalReview.Enabled = False
        TxtDatumFinalReview.Enabled = False
        Cbl1CADPreparer.Enabled = False
        Cbl2CADPreparer.Enabled = False
        CblJobStatus.Enabled = False
        TxtOrderNo.Enabled = False
        TxtGISID.Enabled = False
        TxtClient.Enabled = False
        TxtBusinessEstablishment.Enabled = False
        TxtTypeFinancials.Enabled = False
        TxtYearEnd.Enabled = False
        TxtEngCode.Enabled = False
        TxtEngName.Enabled = False
        TxtEngPartner.Enabled = False
        TxtManager.Enabled = False
        TxtOtherContact.Enabled = False
        TxtEngContact.Enabled = False
        CblSprache.Enabled = False
        CblTool.Enabled = False
        CblEYStand.Enabled = False
        CblConfirmation.Enabled = False
        CblTypeConfi.Enabled = False
        TxtGroup.Enabled = False
    End If
    
    Dim RSaux As ADODB.RecordSet
    
    If (InStr(confirmation, "Kreditor") > 0 Or InStr(confirmation, "Debitor")) And bEYstand = "True" Then
        Set RSaux = getRS("SELECT * FROM tCON_ClientContact" & strTestTable & " WHERE OrderNo = '" & OrderNo & "'")
         Do While RSaux.EOF = False
            LblAnredeMandant.Visible = True
            TxtAnredeMandant.Visible = True
            LblNameMandant.Visible = True
            TxtNameMandant.Visible = True
            LblEMailMandant.Visible = True
            TxtEMailMandant.Visible = True
            If IsNull(RSaux("Mandant_Anr")) = False Then TxtAnredeMandant.value = RSaux("Mandant_Anr")
            If IsNull(RSaux("Mandant_Name")) = False Then TxtNameMandant.value = RSaux("Mandant_Name")
            If IsNull(RSaux("Mandant_Email")) = False Then TxtEMailMandant.value = RSaux("Mandant_Email")
            Exit Do
         Loop
         RSaux.Close
    End If

    ' Alle anderen Daten füllen
    Set OrderbookRS = getRS("SELECT * From tCon_Orderbook" & strTestTable & " WHERE OrderNo = '" & OrderNo & "'")
    
    TxtOrderNo.value = OrderNo
    If IsNull(OrderbookRS("GISID")) = False Then TxtGISID.value = OrderbookRS("GISID")
    If IsNull(OrderbookRS("Client")) = False Then TxtClient.value = OrderbookRS("Client")
    If IsNull(OrderbookRS("BusinessEstablishmentOAre")) = False Then TxtBusinessEstablishment.value = OrderbookRS("BusinessEstablishmentOAre")
    If IsNull(OrderbookRS("TypeFinancials")) = False Then TxtTypeFinancials.value = OrderbookRS("TypeFinancials")
    If IsNull(OrderbookRS("YearEnd")) = False Then TxtYearEnd.value = OrderbookRS("YearEnd")
    If IsNull(OrderbookRS("NameOfCompanyGroup")) = False Then TxtGroup.value = OrderbookRS("NameOfCompanyGroup")
    If IsNull(OrderbookRS("EngCode")) = False Then TxtEngCode.value = OrderbookRS("EngCode")
    If IsNull(OrderbookRS("EngName")) = False Then TxtEngName.value = OrderbookRS("EngName")
    If IsNull(OrderbookRS("EngPartner")) = False Then TxtEngPartner.value = OrderbookRS("EngPartner")
    If IsNull(OrderbookRS("EngManager")) = False Then TxtManager.value = OrderbookRS("EngManager")
    If IsNull(OrderbookRS("OtherContact")) = False Then TxtOtherContact.value = OrderbookRS("OtherContact")
    If IsNull(OrderbookRS("EngContact")) = False Then TxtEngContact.value = OrderbookRS("EngContact")
    If IsNull(OrderbookRS("ConfiMailSent")) = False Then TxtConfiMailSent.value = OrderbookRS("ConfiMailSent")
       
    If IsNull(OrderbookRS("Lang")) = False Then CblSprache.value = OrderbookRS("Lang")
    If IsNull(OrderbookRS("Tool")) = False Then CblTool.value = OrderbookRS("Tool")
    If IsNull(OrderbookRS("Confirmation")) = False Then CblConfirmation.value = OrderbookRS("Confirmation")
    If IsNull(OrderbookRS("TypeConfi")) = False Then CblTypeConfi.value = OrderbookRS("TypeConfi")
    
    If IsNull(OrderbookRS("PlannedDelivery")) = False Then TxtPlannedDelivery.value = OrderbookRS("PlannedDelivery")
    If IsNull(OrderbookRS("CAD_Bearbeiter")) = False Then Cbl1CADPreparer.value = OrderbookRS("CAD_Bearbeiter")
    If IsNull(OrderbookRS("CAD_Bearbeiter_vorherig")) = False Then Cbl2CADPreparer.value = OrderbookRS("CAD_Bearbeiter_vorherig")
    If IsNull(OrderbookRS("DeliveryType")) = False Then CblDeliveryType.value = OrderbookRS("DeliveryType")
    If IsNull(OrderbookRS("DateConfi")) = False Then TxtDateConfi.value = OrderbookRS("DateConfi")
    If IsNull(OrderbookRS("SB_Status")) = False Then CblJobStatus.value = OrderbookRS("SB_Status")
    
    If confirmation = "Sonstige" Then
        LblOtherTypeConfi.Visible = True
        TxtOtherTypeConfi.Visible = True
        If IsNull(OrderbookRS("OtherTypeConfi")) = False Then TxtOtherTypeConfi.value = OrderbookRS("OtherTypeConfi")
    End If
    
    If IsNull(OrderbookRS("EYStand")) = False Then
        If OrderbookRS("EYStand") = True Then
            CblEYStand.value = "Ja"
        Else
            CblEYStand.value = "Nein"
        End If
    End If
    
    Do While OrderbookRS.EOF = False
        If confirmation <> "BankenPortal" Then
            If IsNull(OrderbookRS("CAD_Bearbeiter")) = False Then Cbl1CADPreparer.value = OrderbookRS("CAD_Bearbeiter")
            If IsNull(OrderbookRS("CAD_Bearbeiter_vorherig")) = False Then Cbl2CADPreparer.value = OrderbookRS("CAD_Bearbeiter_vorherig")
            If IsNull(OrderbookRS("CAD_Reviewer")) = False Then CblFinalReview.value = OrderbookRS("CAD_Reviewer")
            If IsNull(OrderbookRS("Anzahl_erhaltene_Salden")) = False Then TxtAnzahlErhalteneSalden.value = OrderbookRS("Anzahl_erhaltene_Salden")
            If IsNull(OrderbookRS("DatumAuftragsende")) = False Then TxtAuftragsende.value = OrderbookRS("DatumAuftragsende")
            If IsNull(OrderbookRS("Datum_Versand_EMail_Erstkontakt")) = False Then TxtErstkontakt.value = OrderbookRS("Datum_Versand_EMail_Erstkontakt")
            If IsNull(OrderbookRS("Informationen_zum_Auftrag")) = False Then TxtInformationenZumAuftrag.value = OrderbookRS("Informationen_zum_Auftrag")
            If IsNull(OrderbookRS("Anzahl_Bestätigungen")) = False Then TxtAnzahlBestätigungen.value = OrderbookRS("Anzahl_Bestätigungen")
            
            If Tool = "eConfirmations" Then
                If IsNull(OrderbookRS("Auswahl_erhalten")) = False Then
                    If OrderbookRS("Auswahl_erhalten") = True Then
                        CblAuswahlErhalten.value = "Ja"
                    Else
                        CblAuswahlErhalten.value = "Nein"
                    End If
                End If
                If IsNull(OrderbookRS("Datum_Auswahl_erhalten")) = False Then TxtDatumAuswahlErhalten.value = OrderbookRS("Datum_Auswahl_erhalten")
            End If
            
            If IsNull(OrderbookRS("Eingangsdatum_CAD")) = False Then TxtEingangsdatum.value = OrderbookRS("Eingangsdatum_CAD")
            If IsNull(OrderbookRS("Versanddatum_CAD")) = False Then TxtVersanddatum.value = OrderbookRS("Versanddatum_CAD")
            If IsNull(OrderbookRS("Datum_2nd_Request")) = False Then TxtDatum2ndRequest.value = OrderbookRS("Datum_2nd_Request")
            If IsNull(OrderbookRS("Kommentar_Versand")) = False Then TxtKommentarZumVersand.value = OrderbookRS("Kommentar_Versand")
            If IsNull(OrderbookRS("Final_Senior_Review_durch")) = False Then CblFinalReview.value = OrderbookRS("Final_Senior_Review_durch")
            If IsNull(OrderbookRS("Datum_final_Review")) = False Then TxtDatumFinalReview.value = OrderbookRS("Datum_final_Review")
            If IsNull(OrderbookRS("Versand_durch")) = False Then cblVersanddurch.value = OrderbookRS("Versand_durch")
            If IsNull(OrderbookRS("Versanddatum_Originale")) = False Then TxtVersanddatumOriginale.value = OrderbookRS("Versanddatum_Originale")
            If IsNull(OrderbookRS("Empfänger_Originale")) = False Then TxtEmpfängerOriginale.value = OrderbookRS("Empfänger_Originale")
            If IsNull(OrderbookRS("First_Upload_Gewünscht")) = False Then TxtUploadGewünscht.value = OrderbookRS("First_Upload_Gewünscht")
            If IsNull(OrderbookRS("Antwort_Deadline")) = False Then TxtAntwortDeadline.value = OrderbookRS("Antwort_Deadline")
            If IsNull(OrderbookRS("Empfänger_Nachläufer")) = False Then TxtEmpfängerNachläufer.value = OrderbookRS("Empfänger_Nachläufer")
            If IsNull(OrderbookRS("2nd_Request_gewünscht")) = False Then
                If OrderbookRS("2nd_Request_gewünscht") = True Then
                    Cbl2ndRequest.value = "Ja"
                Else
                    Cbl2ndRequest.value = "Nein"
                End If
            End If
            
            If Tool = "eConfirmations" Then
                If IsNull(OrderbookRS("AuditiSent")) = False Then TxteConfirmationsSent.value = OrderbookRS("AuditiSent")
                If IsNull(OrderbookRS("Abrechnungsbetrag")) = False Then TxtAbrechnungsbetrag.value = OrderbookRS("Abrechnungsbetrag")
                If IsNull(OrderbookRS("Abrechnungsdatum")) = False Then TxtAbrechnungsdatum.value = OrderbookRS("Abrechnungsdatum")
            End If
        End If
        
        If IsNull(OrderbookRS("Datum_Feedbackanfrage")) = False Then TxtDatumFeedbackabfrage.value = OrderbookRS("Datum_Feedbackanfrage")
        If IsNull(OrderbookRS("Datum_Reminder")) = False Then TxtDatumReminder.value = OrderbookRS("Datum_Reminder")
        If IsNull(OrderbookRS("Feedback_erhalten")) = False Then
            If OrderbookRS("Feedback_erhalten") = True Then
                CblFeedbackErhalten.value = "Ja"
            Else
                CblFeedbackErhalten.value = "Nein"
            End If
        End If
        ' Es sollte nur eine Ordernummer geben
        Exit Do
    Loop
    OrderbookRS.Close
Else
    MsgBox "Bitte wählen Sie eine Bestellung aus!"
    DoCmd.Close acForm, Me.name
End If

End Sub
