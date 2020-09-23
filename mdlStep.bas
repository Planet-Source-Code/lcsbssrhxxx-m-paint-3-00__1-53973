Attribute VB_Name = "mdlStep"
Public UpDownRed As Integer
Public UpDownGreen As Integer
Public UpDownBlue As Integer
Public UpDownColor As Integer
Public Sub StepRGB()
    Select Case UpDownRed
        Case 0
            If frmMain.red.Value >= 255 Then
                UpDownRed = UpDownRed + 1
            End If
            If Not frmMain.red.Value = 255 Then
                frmMain.red.Value = frmMain.red.Value + 1
            End If
        Case 1
            If frmMain.red.Value <= 0 Then
                UpDownRed = UpDownRed - 1
            End If
            If Not frmMain.red.Value = 0 Then
                frmMain.red.Value = frmMain.red.Value - 1
            End If
        End Select
    
    Select Case UpDownGreen
        Case 0
            If frmMain.green.Value >= 255 Then
                UpDownGreen = UpDownGreen + 1
            End If
            If Not frmMain.green.Value = 255 Then
                frmMain.green.Value = frmMain.green.Value + 1
            End If
        Case 1
            If frmMain.green.Value = 0 Then
                UpDownGreen = UpDownGreen - 1
            End If
            If Not frmMain.green.Value <= 0 Then
                frmMain.green.Value = frmMain.green.Value - 1
            End If
        End Select
        
    Select Case UpDownBlue
        Case 0
            If frmMain.blue.Value = 255 Then
                UpDownBlue = UpDownBlue + 1
            End If
            If Not frmMain.blue.Value >= 255 Then
                frmMain.blue.Value = frmMain.blue.Value + 1
            End If
        Case 1
            If Not frmMain.blue.Value = 0 Then
                frmMain.blue.Value = frmMain.blue.Value - 1
            End If
            If frmMain.blue.Value <= 0 Then
                UpDownBlue = UpDownBlue - 1
            End If
        End Select
End Sub
Public Sub StepGradientBrush()
    Select Case UpDownColor
        Case 0
            If frmMain.red.Value >= 255 Or frmMain.green.Value >= 255 Or frmMain.blue.Value >= 255 Then
                UpDownColor = UpDownColor + 1
            End If
            If Not frmMain.red.Value >= 255 Then
                frmMain.red.Value = frmMain.red.Value + 1
            End If
            If Not frmMain.green.Value >= 255 Then
                frmMain.green.Value = frmMain.green.Value + 1
            End If
            If Not frmMain.blue.Value >= 255 Then
                frmMain.blue.Value = frmMain.blue.Value + 1
            End If
        Case 1
            If frmMain.red.Value <= 0 Or frmMain.green.Value <= 0 Or frmMain.blue.Value <= 0 Then
                UpDownColor = UpDownColor - 1
            End If
            If Not frmMain.red.Value = 0 Then
                frmMain.red.Value = frmMain.red.Value - 1
            End If
            If Not frmMain.green.Value = 0 Then
                frmMain.green.Value = frmMain.green.Value - 1
            End If
            If Not frmMain.blue.Value = 0 Then
                frmMain.blue.Value = frmMain.blue.Value - 1
            End If
        End Select
End Sub
Public Sub StepRedBrush()
    Select Case UpDownRed
        Case 0
            If frmMain.red.Value >= 255 Then
                UpDownRed = UpDownRed + 1
            End If
            If Not frmMain.red.Value = 255 Then
                frmMain.red.Value = frmMain.red.Value + 1
            End If
        Case 1
            If frmMain.red.Value <= 0 Then
                UpDownRed = UpDownRed - 1
            End If
            If Not frmMain.red.Value = 0 Then
                frmMain.red.Value = frmMain.red.Value - 1
            End If
    End Select
End Sub
Public Sub StepGreenBrush()
    Select Case UpDownGreen
        Case 0
            If frmMain.green.Value >= 255 Then
                UpDownGreen = UpDownGreen + 1
            End If
            If Not frmMain.green.Value = 255 Then
                frmMain.green.Value = frmMain.green.Value + 1
            End If
        Case 1
            If frmMain.green.Value <= 0 Then
                UpDownGreen = UpDownGreen - 1
            End If
            If Not frmMain.green.Value = 0 Then
                frmMain.green.Value = frmMain.green.Value - 1
            End If
    End Select
End Sub
Public Sub StepBlueBrush()
    Select Case UpDownBlue
        Case 0
            If frmMain.blue.Value >= 255 Then
                UpDownBlue = UpDownBlue + 1
            End If
            If Not frmMain.blue.Value = 255 Then
                frmMain.blue.Value = frmMain.blue.Value + 1
            End If
        Case 1
            If frmMain.blue.Value <= 0 Then
                UpDownBlue = UpDownBlue - 1
            End If
            If Not frmMain.blue.Value = 0 Then
                frmMain.blue.Value = frmMain.blue.Value - 1
            End If
    End Select
End Sub
