Attribute VB_Name = "validation"
Option Explicit

   ' http://en.wikipedia.org/wiki/International_Bank_Account_Number
Private Const IbanCountryLengths As String = "AL28AD24AT20AZ28BH22BE16BA20BR29BG22CR21HR21CY28CZ24DK18DO28EE20FO18" & _
                                             "FI18FR27GE22DE22GI23GR27GL18GT28HU28IS26IE22IL23IT27KZ20KW30LV21LB28" & _
                                             "LI21LT20LU20MK19MT31MR27MU30MC27MD24ME22NL18NO15PK24PS29PL28PT25RO24" & _
                                             "SM27SA24RS22SK24SI19ES24SE24CH21TN24TR26AE23GB22VG24QA29"

'validates ZIP Code iternationally
Function validateZipCode(zip As String, country As String) As Boolean

    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    validateZipCode = False
        
    'makes VBA Dictionary with regex to different country ISO codes
    dict.Add "GB", "GIR[ ]?0AA|((AB|AL|B|BA|BB|BD|BH|BL|BN|BR|BS|BT|CA|CB|CF|CH|CM|CO|CR|CT|CV|CW|DA|DD|DE|DG|DH|DL|DN|DT|DY|E|EC|EH|EN|EX|FK|FY|G|GL|GY|GU|HA|HD|HG|HP|HR|HS|HU|HX|IG|IM|IP|IV|JE|KA|KT|KW|KY|L|LA|LD|LE|LL|LN|LS|LU|M|ME|MK|ML|N|NE|NG|NN|NP|NR|NW|OL|OX|PA|PE|PH|PL|PO|PR|RG|RH|RM|S|SA|SE|SG|SK|SL|SM|SN|SO|SP|SR|SS|ST|SW|SY|TA|TD|TF|TN|TQ|TR|TS|TW|UB|W|WA|WC|WD|WF|WN|WR|WS|WV|YO|ZE)(\d[\dA-Z]?[ ]?\d[ABD-HJLN-UW-Z]{2}))|BFPO[ ]?\d{1,4}"
    dict.Add "JE", "JE\d[\dA-Z]?[ ]?\d[ABD-HJLN-UW-Z]{2}"
    dict.Add "GG", "GY\d[\dA-Z]?[ ]?\d[ABD-HJLN-UW-Z]{2}"
    dict.Add "IM", "IM\d[\dA-Z]?[ ]?\d[ABD-HJLN-UW-Z]{2}"
    dict.Add "US", "\d{5}([ \-]\d{4})?"
    dict.Add "CA", "[ABCEGHJKLMNPRSTVXY]\d[ABCEGHJ-NPRSTV-Z][ ]?\d[ABCEGHJ-NPRSTV-Z]\d"
    dict.Add "DE", "\d{5}"
    dict.Add "JP", "\d{3}-\d{4}"
    dict.Add "FR", "\d{2}[ ]?\d{3}"
    dict.Add "AU", "\d{4}"
    dict.Add "IT", "\d{5}"
    dict.Add "CH", "\d{4}"
    dict.Add "AT", "\d{4}"
    dict.Add "ES", "\d{5}"
    dict.Add "NL", "\d{4}[ ]?[A-Z]{2}"
    dict.Add "BE", "\d{4}"
    dict.Add "DK", "\d{4}"
    dict.Add "SE", "\d{3}[ ]?\d{2}"
    dict.Add "NO", "\d{4}"
    dict.Add "BR", "\d{5}[\-]?\d{3}"
    dict.Add "PT", "\d{4}([\-]\d{3})?"
    dict.Add "FI", "\d{5}"
    dict.Add "AX", "22\d{3}"
    dict.Add "KR", "(\d{3}[\-]\d{3}|\d{5})"
    dict.Add "CN", "\d{6}"
    dict.Add "TW", "\d{3}(\d{2})?"
    dict.Add "SG", "\d{6}"
    dict.Add "DZ", "\d{5}"
    dict.Add "AD", "AD\d{3}"
    dict.Add "AR", "([A-HJ-NP-Z])?\d{4}([A-Z]{3})?"
    dict.Add "AM", "(37)?\d{4}"
    dict.Add "AZ", "\d{4}"
    dict.Add "BH", "((1[0-2]|[2-9])\d{2})?"
    dict.Add "BD", "\d{4}"
    dict.Add "BB", "(BB\d{5})?"
    dict.Add "BY", "\d{6}"
    dict.Add "BM", "[A-Z]{2}[ ]?[A-Z0-9]{2}"
    dict.Add "BA", "\d{5}"
    dict.Add "IO", "BBND 1ZZ"
    dict.Add "BN", "[A-Z]{2}[ ]?\d{4}"
    dict.Add "BG", "\d{4}"
    dict.Add "KH", "\d{5}"
    dict.Add "CV", "\d{4}"
    dict.Add "CL", "\d{7}"
    dict.Add "CR", "\d{4,5}|\d{3}-\d{4}"
    dict.Add "HR", "\d{5}"
    dict.Add "CY", "\d{4}"
    dict.Add "CZ", "\d{3}[ ]?\d{2}"
    dict.Add "DO", "\d{5}"
    dict.Add "EC", "([A-Z]\d{4}[A-Z]|(?:[A-Z]{2})?\d{6})?"
    dict.Add "EG", "\d{5}"
    dict.Add "EE", "\d{5}"
    dict.Add "FO", "\d{3}"
    dict.Add "GE", "\d{4}"
    dict.Add "GR", "\d{3}[ ]?\d{2}"
    dict.Add "GL", "39\d{2}"
    dict.Add "GT", "\d{5}"
    dict.Add "HT", "\d{4}"
    dict.Add "HN", "(?:\d{5})?"
    dict.Add "HU", "\d{4}"
    dict.Add "IS", "\d{3}"
    dict.Add "IN", "\d{6}"
    dict.Add "ID", "\d{5}"
    dict.Add "IL", "\d{5}"
    dict.Add "JO", "\d{5}"
    dict.Add "KZ", "\d{6}"
    dict.Add "KE", "\d{5}"
    dict.Add "KW", "\d{5}"
    dict.Add "LA", "\d{5}"
    dict.Add "LV", "\d{4}"
    dict.Add "LB", "(\d{4}([ ]?\d{4})?)?"
    dict.Add "LI", "(948[5-9])|(949[0-7])"
    dict.Add "LT", "\d{5}"
    dict.Add "LU", "\d{4}"
    dict.Add "MK", "\d{4}"
    dict.Add "MY", "\d{5}"
    dict.Add "MV", "\d{5}"
    dict.Add "MT", "[A-Z]{3}[ ]?\d{2,4}"
    dict.Add "MU", "(\d{3}[A-Z]{2}\d{3})?"
    dict.Add "MX", "\d{5}"
    dict.Add "MD", "\d{4}"
    dict.Add "MC", "980\d{2}"
    dict.Add "MA", "\d{5}"
    dict.Add "NP", "\d{5}"
    dict.Add "NZ", "\d{4}"
    dict.Add "NI", "((\d{4}-)?\d{3}-\d{3}(-\d{1})?)?"
    dict.Add "NG", "(\d{6})?"
    dict.Add "OM", "(PC )?\d{3}"
    dict.Add "PK", "\d{5}"
    dict.Add "PY", "\d{4}"
    dict.Add "PH", "\d{4}"
    dict.Add "PL", "\d{2}-\d{3}"
    dict.Add "PR", "00[679]\d{2}([ \-]\d{4})?"
    dict.Add "RO", "\d{6}"
    dict.Add "RU", "\d{6}"
    dict.Add "SM", "4789\d"
    dict.Add "SA", "\d{5}"
    dict.Add "SN", "\d{5}"
    dict.Add "SK", "\d{3}[ ]?\d{2}"
    dict.Add "SI", "\d{4}"
    dict.Add "ZA", "\d{4}"
    dict.Add "LK", "\d{5}"
    dict.Add "TJ", "\d{6}"
    dict.Add "TH", "\d{5}"
    dict.Add "TN", "\d{4}"
    dict.Add "TR", "\d{5}"
    dict.Add "TM", "\d{6}"
    dict.Add "UA", "\d{5}"
    dict.Add "UY", "\d{5}"
    dict.Add "UZ", "\d{6}"
    dict.Add "VA", "00120"
    dict.Add "VE", "\d{4}"
    dict.Add "ZM", "\d{5}"
    dict.Add "AS", "96799"
    dict.Add "CC", "6799"
    dict.Add "CK", "\d{4}"
    dict.Add "RS", "\d{6}"
    dict.Add "ME", "8\d{4}"
    dict.Add "CS", "\d{5}"
    dict.Add "YU", "\d{5}"
    dict.Add "CX", "6798"
    dict.Add "ET", "\d{4}"
    dict.Add "FK", "FIQQ 1ZZ"
    dict.Add "NF", "2899"
    dict.Add "FM", "(9694[1-4])([ \-]\d{4})?"
    dict.Add "GF", "9[78]3\d{2}"
    dict.Add "GN", "\d{3}"
    dict.Add "GP", "9[78][01]\d{2}"
    dict.Add "GS", "SIQQ 1ZZ"
    dict.Add "GU", "969[123]\d([ \-]\d{4})?"
    dict.Add "GW", "\d{4}"
    dict.Add "HM", "\d{4}"
    dict.Add "IQ", "\d{5}"
    dict.Add "KG", "\d{6}"
    dict.Add "LR", "\d{4}"
    dict.Add "LS", "\d{3}"
    dict.Add "MG", "\d{3}"
    dict.Add "MH", "969[67]\d([ \-]\d{4})?"
    dict.Add "MN", "\d{6}"
    dict.Add "MP", "9695[012]([ \-]\d{4})?"
    dict.Add "MQ", "9[78]2\d{2}"
    dict.Add "NC", "988\d{2}"
    dict.Add "NE", "\d{4}"
    dict.Add "VI", "008(([0-4]\d)|(5[01]))([ \-]\d{4})?"
    dict.Add "PF", "987\d{2}"
    dict.Add "PG", "\d{3}"
    dict.Add "PM", "9[78]5\d{2}"
    dict.Add "PN", "PCRN 1ZZ"
    dict.Add "PW", "96940"
    dict.Add "RE", "9[78]4\d{2}"
    dict.Add "SH", "(ASCN|STHL) 1ZZ"
    dict.Add "SJ", "\d{4}"
    dict.Add "SO", "\d{5}"
    dict.Add "SZ", "[HLMS]\d{3}"
    dict.Add "TC", "TKCA 1ZZ"
    dict.Add "WF", "986\d{2}"
    dict.Add "XK", "\d{5}"
    dict.Add "YT", "976\d{2}"
    
    'if ISO Code is in Dictionary then check regex
    If dict.exists(country) Then
    
        With RegEx
            .Pattern = dict.Item(country)
        End With
    
        If RegEx.test(zip) Then
            validateZipCode = True
        End If
    
    Else
        validateZipCode = True
    End If
    
    Set dict = Nothing

End Function

'validate email address via regex
Function validateEmail(address As String) As Boolean
    
    validateEmail = False
    
    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    
    With RegEx
        .Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    End With
    
    If RegEx.test(address) Or address = "" Then
        validateEmail = True
    End If
    
End Function

'##############################
'Validation of IBAN
Private Function ValidateIbanCountryLength(CountryCode As String, IbanLength As Integer) As Boolean
    Dim i As Integer
    For i = 0 To Len(IbanCountryLengths) / 4 - 1
        If Mid(IbanCountryLengths, i * 4 + 1, 2) = CountryCode And _
            CInt(Mid(IbanCountryLengths, i * 4 + 3, 2)) = IbanLength Then
            ValidateIbanCountryLength = True
            Exit Function
        End If
    Next i
    ValidateIbanCountryLength = False
End Function

Private Function Mod97(Num As String) As Integer
    Dim lngTemp As Long
    Dim strTemp As String

    Do While Val(Num) >= 97
        If Len(Num) > 5 Then
            strTemp = Left(Num, 5)
            Num = Right(Num, Len(Num) - 5)
        Else
            strTemp = Num
            Num = ""
        End If
        lngTemp = CLng(strTemp)
        lngTemp = lngTemp Mod 97
        strTemp = CStr(lngTemp)
        Num = strTemp & Num
    Loop
    Mod97 = CInt(Num)
End Function

Public Function ValidateIban(IBAN As String) As Boolean
    Dim strIban As String
    Dim i As Integer

    strIban = UCase(IBAN)
    ' Remove spaces
    strIban = Replace(strIban, " ", "")

    ' Check if IBAN contains only uppercase characters and numbers
    For i = 1 To Len(strIban)
        If Not ((Asc(Mid(strIban, i, 1)) <= Asc("9") And Asc(Mid(strIban, i, 1)) >= Asc("0")) Or _
                (Asc(Mid(strIban, i, 1)) <= Asc("Z") And Asc(Mid(strIban, i, 1)) >= Asc("A"))) Then
            ValidateIban = False
            Exit Function
        End If
    Next i

    ' Check if length of IBAN equals expected length for country
    If Not ValidateIbanCountryLength(Left(strIban, 2), Len(strIban)) Then
        ValidateIban = False
        Exit Function
    End If

    ' Rearrange
    strIban = Right(strIban, Len(strIban) - 4) & Left(strIban, 4)

    ' Replace characters
    For i = 0 To 25
        strIban = Replace(strIban, Chr(i + Asc("A")), i + 10)
    Next i

    ' Check remainder
    ValidateIban = Mod97(strIban) = 1
End Function
'###################################

'finds range of header by cell contents "laufende Nummer (automatisch)" and "Falls notwendig Kommentar"
Function getHeader(sheet As Variant) As Range

    Dim objul As Object
    Dim objur As Object
    Dim lrow As String
    Dim lcol As String
    Dim rrow As String
    Dim rcol As String
    Dim tmp() As String
    
    Set objul = sheet.Range("B1:B100").Find(what:="Sequence number (automatic)", LookIn:=xlValues)
    
    If Not objul Is Nothing Then
        tmp = Split(objul.address, "$")
        lcol = tmp(1)
        lrow = tmp(2)
    End If
    
    Set objur = Range(lcol & lrow).Resize(2, 50).Find("If necessary comment", LookIn:=xlValues)
    
    If Not objur Is Nothing Then
        tmp = Split(objur.address, "$")
        rcol = tmp(1)
        rrow = tmp(2)
    End If
    
    Set getHeader = Range(lcol & lrow & ":" & rcol & rrow)
    
End Function
    
'finds all columns that contain a obligatory field in header
Function getMandatory(sheet As Variant, header As Range) As String()

    Dim tmp() As String
    Dim mandatory As String
    Dim c As Variant
    Dim i As Integer

    For Each c In header
        If InStr(1, c.Value, "required") <> 0 Or InStr(1, c.Value, "Debtor/ Creditor/ Other") <> 0 Or InStr(1, c.Value, "Service related to the address") <> 0 Then
            If c.MergeCells = True Then
                If c.MergeArea.Columns.Count = 1 Then
                    tmp = Split(c.address, "$")
                    mandatory = mandatory & "," & tmp(1)
                Else
                    For i = 0 To c.MergeArea.Columns.Count - 1
                        mandatory = mandatory & "," & Split(Cells(c.Row, c.Column + i).address, "$")(1)
                    Next i
                End If
            Else
                tmp = Split(c.address, "$")
                mandatory = mandatory & "," & tmp(1)
            End If
        End If
    Next
    
    If mandatory <> "" Then
        getMandatory = Split(Right(mandatory, Len(mandatory) - 1), ",")
    Else
        getMandatory = Null
    End If
    
End Function

'determines last row by determining length of obligatory fields
'RETUR VALUES: -1: if no entries to loop through, -2: indicates difference in entries in obligatory fields, else: last row
Function getLastRow(sheet As Variant, header As Range, mandatoryFields() As String) As Long

    Dim mand As Variant
    getLastRow = -1
    
    For Each mand In mandatoryFields
        If getLastRow = -1 Then
            getLastRow = sheet.Range(mand & sheet.Rows.Count).End(xlUp).Row
        ElseIf getLastRow <> sheet.Range(mand & sheet.Rows.Count).End(xlUp).Row And sheet.Range(mand & sheet.Rows.Count).End(xlUp).Row > 17 Then
            getLastRow = -2
            Exit Function
        End If
    Next

 End Function

'main sub
Sub exe()

    Dim wsTemplate As Variant
    Dim hdr As Range
    Dim sheetNumber As Integer
    Dim sheet As Variant
    Dim mand() As String
    Dim lastrow As Long
    Dim bValid1 As Boolean
    Dim bValid2 As Boolean
    Dim bValid3 As Boolean
    Dim bTemp As Boolean
    Dim strOrderNo As String
    Dim cell As Variant
    Dim header() As String
    Dim fields As String
    Dim i As Integer
    
    Set wsTemplate = ActiveWorkbook
    wsTemplate.Save
    
    bValid1 = True
    bValid2 = True
    bValid3 = False
    
    'loops through worksheets in workbook
    For Each sheet In wsTemplate.Worksheets
       
        'exclude hidden sheets
        If sheet.Name <> "basic_info" And sheet.Name <> "Inhalte" And sheet.Name <> "ISO" And sheet.Name <> "Summary" Then
            sheet.Activate
    
            'check if headers are unchanged
            Set hdr = getHeader(sheet)
            If sheet.Name = "Debtor_Creditor_Other" Then
                fields = "Sequence number (automatic)|Debtor/ Creditor/ Other|Account/ Invoice number|Name of company (required)|Additional address information|Contact person (optional)||Street + house number / P.O. Box (required)|Postcode (required)|City                                      (required)|Country                           (required, ISO code if possible)|E-mail" & Chr(10) & "(optional)|If necessary comment"
                
            ElseIf sheet.Name = "Bank" Then
                fields = "Sequence number (automatic)|Number of the general" & Chr(10) & "ledger account (optional)|IBAN" & Chr(10) & "(optional)|Name of the bank (required)|Additional address information|Contact person (optional)||Street + house number / P.O. Box (required)|Postcode (required)|City                                   (required)|Country                                     (required, ISO code if possible)|E-mail" & Chr(10) & "(optional)|If necessary comment"
                
            ElseIf sheet.Name = "Legal_Tax Advisors" Then
                fields = "Sequence number (automatic)|Type of service|Name of the law firm (required)|Additional address information|Contact person (optional)||Street + house number / P.O. Box (required)|Postcode" & Chr(10) & "(required)|City" & Chr(10) & "(required)|Country" & Chr(10) & "(required, ISO code if possible)|E-mail" & Chr(10) & "(optional)|If necessary comment"
                
            ElseIf sheet.Name = "Address check" Then
                fields = "Sequence number (automatic)|Service related to the address|Name of company (required)|Additional address information|Contact person (optional)||Street + house number/ P.O. Box (required)|Postcode" & Chr(10) & "(required)|City" & Chr(10) & "(required)|Country" & Chr(10) & "(required, ISO code if possible)|E-mail" & Chr(10) & "(optional)|If necessary comment"

                
            End If
            
            header = Split(fields, "|")
            
            i = 0
            
            For Each cell In hdr
            
                If cell.Value <> header(i) Then
                    Debug.Print cell & " vs " & header(i) & " " & sheet.Name
                    MsgBox ("Please put the columns according to the initial order of the document!")
                    Exit Sub
                End If
                
                i = i + 1
                
            Next cell
        
            'get columns with obligatory fields
            mand = getMandatory(sheet, hdr)
        
            'get last row of entries
            lastrow = getLastRow(sheet, hdr, mand)
        
            'checks if there was a different amount of numbers of entries in obligatory fields
            If lastrow = -2 Then
                bValid2 = False
            End If
            
            Debug.Print sheet.Name
            '            If sheet.Name = "Rechts-_Steuerberater" Then
            '
            '                'Setze lastrow um 1 zurück, da Ansprechpartner (benötigt) nicht gemerged ist
            '                If lastrow = 16 Then lastrow = 15
            '
            '
            '            End If
        
            'if there are entries
            If lastrow <> -1 And lastrow > 16 Then
            
                bTemp = validate(hdr, sheet, lastrow)
                
                If bValid1 Then
                    bValid1 = bTemp
                End If
                bValid3 = True
                
            ElseIf lastrow >= 15 Then
        
                Debug.Print sheet.Name
                'bValid3 = False
                'Exit For
            End If
        
        End If
    Next sheet
    
    'empty sheet found
    If bValid3 = False Then
        MsgBox ("No entry was made!")
    Else
    
        'problem with validity of values
        If bValid1 = False Then
            MsgBox ("Please check the orange marked entries in all sheets!")
        End If
        
        'problem with mising obligatory entries
        If bValid2 = False Then
            MsgBox ("Please make sure that in each row all fields, which are marked as 'reqired', are filled!")
        End If
    End If
    
    'if any problem occured
    If bValid1 = False Or bValid2 = False Or bValid3 = False Then
    
        Exit Sub

        'if no problems occured
    Else
        Call sendMail(wsTemplate, bValid1 And bValid2, False)
        
    End If

    On Error GoTo 0
    Exit Sub

errorSend:
    
    Call sendMail(wsTemplate, False, True)
    ActiveWorkbook.Close

End Sub

'loops through columns and rows and calls validation subs for single values
Function validate(hdr As Range, sheet As Variant, lastrow As Long) As Boolean

    Dim i As Variant
    Dim k As Integer
    Dim c As Variant
    Dim answer As Integer
    Dim bIBAN As Boolean
    Dim country As String
    
    bIBAN = True
    validate = True
    
    For Each i In hdr
        If InStr(1, i, "Postcode") <> 0 Then
      
            For Each c In sheet.Range(Cells(hdr.Row + 2, i.Column).address & ":" & Cells(lastrow, i.Column).address)
            
                c.Interior.Color = RGB(255, 245, 153)
                
                c.Value = Trim(c.Value)
                
                country = c.Offset(0, 2).Value
                
                If validateZipCode(c.Value, country) = False Then
              
                    validate = False
                    c.Interior.Color = RGB(255, 150, 0)
                End If
  
            Next c
          
        ElseIf InStr(1, i, "E-mail") <> 0 Then
      
            For Each c In sheet.Range(Cells(hdr.Row + 2, i.Column).address & ":" & Cells(lastrow, i.Column).address)
            
                c.Interior.Color = RGB(255, 245, 153)
                
                c.Value = Trim(c.Value)
          
                If validateEmail(c.Value) = False Then
              
                    validate = False
                    c.Interior.Color = RGB(255, 150, 0)
                End If
  
            Next c
          
          'uncomment to add validation of IBAN
'        ElseIf InStr(1, i, "IBAN") <> 0 Then
'
'            For Each c In sheet.Range(Cells(hdr.Row + 2, i.Column).address & ":" & Cells(lastrow, i.Column).address)
'
'                c.Interior.Color = RGB(255, 245, 153)
'
'                c.Value = Trim(c.Value)
'
'                If c.Value <> "" And Not ValidateIban(c.Value) Then
'
'                    bIBAN = False
'                    c.Interior.Color = RGB(255, 150, 0)
'
'                End If
'
'            Next c
'
        End If
          
        'checks obligatory fields
        If InStr(1, i, "required") <> 0 Then
        
            'when cells in header are merged
            If i.MergeArea.Columns.Count > 1 Then
            
                For k = 1 To i.MergeArea.Columns.Count
                
                    For Each c In sheet.Range(Cells(hdr.Row + 2, i.Column - 1 + k).address & ":" & Cells(lastrow, i.Column - 1 + k).address)
                    
                        c.Interior.Color = RGB(255, 245, 153)
                        
                        c.Value = Trim(c.Value)
          
                        If c.Value = "" Then
                            validate = False
                            c.Interior.Color = RGB(255, 150, 0)
                        ElseIf c.Interior.Color <> RGB(255, 150, 0) Then
                            c.Interior.Color = RGB(255, 245, 153)
                        End If
                    Next c
                    
                Next k
            
            Else
            
                For Each c In sheet.Range(Cells(hdr.Row + 2, i.Column).address & ":" & Cells(lastrow, i.Column).address)
                    
                    c.Value = Trim(c.Value)
          
                    If c.Value = "" Then
                        validate = False
                        c.Interior.Color = RGB(255, 150, 0)
                    ElseIf c.Interior.Color <> RGB(255, 150, 0) Then
                        c.Interior.Color = RGB(255, 245, 153)
                    End If
                Next c
                
            End If
          
        End If
          
    Next i
    
    'alert for IBAN that was determined invalid
    If validate And Not bIBAN Then
        answer = MsgBox("One of the entered IBANs seems to be not correct. Do you want to send anyway?", vbQuestion + vbYesNo + vbDefaultButton2, "IBAN error")
        
        If answer = vbNo Then
            validate = False
        End If
    End If

End Function

'sends Mail in case of valid document and document with error
Sub sendMail(wsTemplate As Variant, bValid As Boolean, bError As Boolean)

    Dim OutMail As Object
    Dim OutApp As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = Outlook.Application.CreateItem(olMailItem)
    Dim body As String
    Dim subject As String
    Dim savePath As String
    
    Dim orderNo As String
    Dim clientName As String

    If bValid Or bError Then
        ' Verschicke E-Mail
        orderNo = wsTemplate.Sheets(1).Range("C2")
        clientName = wsTemplate.Sheets(1).Range("C4")
        
        Application.DisplayAlerts = False
    
        savePath = "C:\Users\" & Environ("username") & "\Documents\" & wsTemplate.Name
    
        Workbooks(wsTemplate.Name).SaveAs Filename:=savePath, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        wsTemplate.Saved = True
    
        If bError Then
    
            MsgBox "A technical error occured. Still, this template has been saved on your desctop and will now be sent via E-mail to adressabgleich@de.ey.com. Soon, the CAD-Team will get in touch with you." & vbCrLf & "We apologize for the error and thank you for the entry!"
    
            subject = "RE: ERROR Beauftragung Adressabgleich " & orderNo & " für " & clientName
         
        ElseIf bValid Then
        
            MsgBox "This template has been successfully saved on your desctop and will now be sent cia E-mail to adressabgleich@de.ey.com." & vbCrLf & "Thank you for the entry!"
    
            subject = "RE_VAL: Beauftragung Adressabgleich " & orderNo & " für " & clientName
    
        End If
        
        'Vorgefertigter Text
        body = "Hello CAD Team, <br><br> Please find the completed templatre for Address comparison attached."
    
        'Email Erstellung
        With OutMail
            .To = "adressabgleich@de.ey.com"
            .subject = subject
            .Attachments.Add savePath
            .HTMLBody = body
    
            'Outlook Draft Fenster wird geöffnet
            .Display
    
            'Simulierte "STRG" + ENTER Eingabe, um die Nachricht zu versenden
            'SendKeys "^{ENTER}"
    
        End With
    
        'Schließe Outlook Object
        Set OutMail = Nothing
    End If
    
End Sub


