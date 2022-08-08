Attribute VB_Name = "CheckBox_Conclusion"
Sub cb_FisTrue_Click()
    Call uncheckOtherBoxes("Check Box 3", 0)

    If ActiveSheet.CheckBoxes("Check Box 3").Value = 1 Then
    
        ActiveSheet.Cells(23, 8) = "ûFIS"
        Call colorConclusion("H23", 1)
        ActiveSheet.Cells(23, 8).Interior.Color = RGB(255, 255, 255)
    Else
        
        ActiveSheet.Cells(23, 8) = ""

    End If

End Sub
Sub cb_FisFalse_Click()
    Call uncheckOtherBoxes("Check Box 4", 1)
    If ActiveSheet.CheckBoxes("Check Box 4").Value = 1 Then
    
        ActiveSheet.Cells(34, 8) = ""
        ActiveSheet.Cells(34, 8).Interior.Color = RGB(217, 217, 217)
        
        'Call colorConclusion("H34", 1)
    Else
        
        ActiveSheet.Cells(34, 8) = ""
        ActiveSheet.Cells(34, 8).Interior.Color = RGB(255, 255, 255)

    End If

End Sub

Sub cb_greyAdress_Click()
    Call uncheckOtherBoxes("Check Box 9", 0)
    If ActiveSheet.CheckBoxes("Check Box 9").Value = 1 Then
    
        ActiveSheet.Cells(23, 8) = ""
        ActiveSheet.Cells(23, 8).Interior.Color = RGB(217, 217, 217)
        
        'Call colorConclusion("H23", 1)
    Else
        
        ActiveSheet.Cells(23, 8) = ""
        ActiveSheet.Cells(23, 8).Interior.Color = RGB(255, 255, 255)

    End If

End Sub


Sub cb_xTrue_Click()
    Call uncheckOtherBoxes("Check Box 5", 0)
    If ActiveSheet.CheckBoxes("Check Box 5").Value = 1 Then
    
        ActiveSheet.Cells(23, 8) = "û"
        Call colorConclusion("H23", 2)
        ActiveSheet.Cells(23, 8).Interior.Color = RGB(255, 255, 255)
    Else
        
        ActiveSheet.Cells(23, 8) = ""

    End If
    
End Sub
Sub cb_xFalse_Click()
    Call uncheckOtherBoxes("Check Box 6", 1)

    If ActiveSheet.CheckBoxes("Check Box 6").Value = 1 Then
    
        ActiveSheet.Cells(34, 8) = "û"
        Call colorConclusion("H34", 2)
    Else
        
        ActiveSheet.Cells(34, 8) = ""

    End If

End Sub
Sub cb_okTrue_Click()
    Call uncheckOtherBoxes("Check Box 7", 0)

    If ActiveSheet.CheckBoxes("Check Box 7").Value = 1 Then
    
        ActiveSheet.Cells(23, 8) = "ü"
        Call colorConclusion("H23", 3)
        ActiveSheet.Cells(23, 8).Interior.Color = RGB(255, 255, 255)
    Else
        
        ActiveSheet.Cells(23, 8) = ""

    End If

End Sub
Sub cb_okFalse_Click()
    Call uncheckOtherBoxes("Check Box 8", 1)

    If ActiveSheet.CheckBoxes("Check Box 8").Value = 1 Then
    
        ActiveSheet.Cells(34, 8) = "ü"
        Call colorConclusion("H34", 3)
    Else
        
        ActiveSheet.Cells(34, 8) = ""

    End If

End Sub

Sub uncheckOtherBoxes(currentBox As String, modus As Integer)

    Dim cmd As Object
    
    With ActiveSheet
    
        For Each cmd In ActiveSheet.Shapes
            
            If (cmd.Name <> currentBox) And InStr(1, cmd.Name, "Button") = 0 And cmd.Type = 8 Then
                'Debug.Print cmd.name
                If modus = 0 Then
                
                    If cmd.Name = "Check Box 3" Or cmd.Name = "Check Box 5" Or cmd.Name = "Check Box 7" Or cmd.Name = "Check Box 9" Then
                        'Debug.Print cmd.name
                        cmd.ControlFormat.Value = False
                 
                    End If
                
                Else
                    If cmd.Name = "Check Box 4" Or cmd.Name = "Check Box 6" Or cmd.Name = "Check Box 8" Or cmd.Name = "Check Box 9" Then
                        Debug.Print cmd.Name
                        cmd.ControlFormat.Value = False
                 
                    End If
                
                End If
                

            End If
                
        Next

    End With

End Sub



