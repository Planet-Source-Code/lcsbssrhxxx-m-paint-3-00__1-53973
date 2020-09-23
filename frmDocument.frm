VERSION 5.00
Begin VB.Form frmDocument 
   Caption         =   "Image 1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4455
   Begin VB.PictureBox PicBox 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000B&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   277
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmDocument' Is A MDI Child For The MDI Prarent / Form 'frmMain'
'A MDI Child Is A Form Inside Of A Form, Giving The Program A Nice Effect

'This Form Has Pretty Much All The Stuff The Program Uses To Draw

'I Found And Fixed A Major Problem (When Ever A Window Overlaped or
'The Picture Was Minimized, The Image Was Cleared In The Spot It Got
'Overlapped.)
'I Fixed The Problem By Turnin
'View mdlSplash.bas (Module) For Sub Main

'[VARIABLES]
Dim b As Long 'b Is The Point Where You Clicked, This Is Used For The Color Picker
Dim oldx 'oldx This Is Used So The Line Draws Contoniously Using The Equation (X1, Y1)-(X2, Y2)
Dim oldy 'oldy This Is Used So The Line Draws Contoniously Using The Equation (X1, Y1)-(X2, Y2)
Dim StripeCase As Integer

'################################## LOAD / RESIZE ############################################

Private Sub Form_Load()
    Form_Resize
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    PicBox.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
End Sub

'################################### MOUSE DOWN ##############################################

Private Sub picBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'[COLOR PICKER]

    'If Mouse Is Down And Color Picker Is Selected Then
    If frmMain.DrawType.Caption = "Color Picker" Then
        'Set Variable B To The Point (X,Y) You Click
        b = PicBox.Point(X, Y)
        'If Mouse1 Then Set Color1 To B
        If Button = 1 Then frmMain.Color1.BackColor = b
        'If Mouse2 Then Set Color2 To B
        If Button = 2 Then frmMain.Color2.BackColor = b
        
        '[COLOR TO RGB]
        texr = PicBox.Point(X, Y) And 255
        texg = (PicBox.Point(X, Y) And 65280) / 256
        texb = (PicBox.Point(X, Y) And 16711680) / 65535
        frmMain.red.Value = texr
        frmMain.green.Value = texg
        frmMain.blue.Value = texb
    End If
    
    '[CUSTOM BRUSH]
    If frmMain.BrushStyle = "Custom Brush" Then
        TransparentBlt PicBox.hdc, X, Y, frmBrush.picBrush.ScaleWidth, frmBrush.picBrush.ScaleHeight, frmBrush.picBrush.hdc, 0, 0, frmBrush.picBrush.ScaleWidth, frmBrush.picBrush.ScaleHeight, frmBrush.Transparent.BackColor
        PicBox.Refresh
    End If

'[FLOOD FILL]

    If frmMain.DrawType.Caption = "Flood Fill" Then
        
        '[Fill Styles]

        If frmMain.xFillStyle.Text = "Solid" Then
            PicBox.FillStyle = 0 'solid
        ElseIf frmMain.xFillStyle.Text = "Horizontal" Then
            PicBox.FillStyle = 2 'horizontal lines
        ElseIf frmMain.xFillStyle.Text = "Vertical" Then
            PicBox.FillStyle = 3 'vertical lines
        ElseIf frmMain.xFillStyle.Text = "Upward Diagonal" Then
            PicBox.FillStyle = 4 'upward diagonal lines
        ElseIf frmMain.xFillStyle.Text = "Downward Diagonal" Then
            PicBox.FillStyle = 5 'downward diagonal lines
        ElseIf frmMain.xFillStyle.Text = "Cross" Then
            PicBox.FillStyle = 6 'crossed lines
        ElseIf frmMain.xFillStyle.Text = "Diagonol Cross" Then
            PicBox.FillStyle = 7 'diagonal crossed lines
        End If
        
        'If Mouse1 Then Fill Color = Color1
        If Button = 1 Then PicBox.FillColor = frmMain.Color1.BackColor
        'If Mouse2 Then Fill Color = Color2
        If Button = 2 Then PicBox.FillColor = frmMain.Color2.BackColor
        PicBox.DrawMode = vbCopyPen
        'API Function ExtFloodFill Get PicBox's .hdc at (X,Y), In PicBox, Fill (X,Y), Fill Type = 1
        ExtFloodFill PicBox.hdc, X, Y, PicBox.Point(X, Y), 1
    End If
    
    
    
    
    
    '[ERASER]
    If frmMain.DrawType.Caption = "Eraser" Then
        frmMain.PenSize.Enabled = True
        PicBox.DrawWidth = frmMain.PenSize.Text
        If Button = 1 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : White
            PicBox.Line (oldx, oldy)-(X, Y), vbWhite
        End If
        If Button = 2 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : White
            PicBox.Line (oldx, oldy)-(X, Y), vbWhite
        End If
    End If
'[PENCIL]

    If frmMain.DrawType.Caption = "Pencil" Then
        frmMain.PenSize.Enabled = False
        PicBox.DrawMode = 13 'solid brush
        PicBox.DrawWidth = 1
        
        If Button > 0 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : Color1
            PicBox.Line (oldx, oldy)-(X, Y), frmMain.Color1.BackColor
        End If
        If Button > 1 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : Color2
            PicBox.Line (oldx, oldy)-(X, Y), frmMain.Color2.BackColor
        End If
    End If
    
'[PEN / MARKER ]

    If frmMain.DrawType.Caption = "Pen" Or frmMain.DrawType.Caption = "Marker" Then
        frmMain.PenSize.Enabled = True
    
    '[SOLID BRUSH]
    If frmMain.BrushStyle = "Solid Brush" Then
        PicBox.DrawMode = 13 'solid brush
        frmMain.PenSize.Enabled = True
        PicBox.DrawWidth = frmMain.PenSize.Text
        If Button > 0 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : Color1
            PicBox.Line (oldx, oldy)-(X, Y), frmMain.Color1.BackColor
        End If
        If Button > 1 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : Color2
            PicBox.Line (oldx, oldy)-(X, Y), frmMain.Color2.BackColor
        End If
    End If
        
    '[INVERT BRUSH]
    If frmMain.BrushStyle = "Invert Brush" Then
        PicBox.DrawMode = 6 'invert brush
        frmMain.PenSize.Enabled = True
        PicBox.DrawWidth = frmMain.PenSize.Text
        If Button = 1 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : black
            PicBox.Line (oldx, oldy)-(X, Y), vbBlack
        End If
    End If
    
    '[CIRCLE BRUSH]
    If frmMain.BrushStyle = "Circle Brush" Then
        PicBox.DrawMode = 13 'solid brush
        PicBox.DrawWidth = 1
        PicBox.FillStyle = 1
        frmMain.PenSize.Enabled = True
        If Button > 0 Then
            PicBox.Circle (X, Y), frmMain.PenSize.Text, frmMain.Color1.BackColor
        End If
        If Button > 1 Then
            PicBox.Circle (X, Y), frmMain.PenSize.Text, frmMain.Color2.BackColor
        End If
    End If
    
    '[AIR BRUSH]
    If frmMain.BrushStyle = "Air Brush" Then
        frmMain.PenSize.Enabled = True
        PicBox.DrawWidth = 1
        PicBox.DrawMode = 13 'Solid Brush
        If Button = 1 Then

            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        End If
        
        If Button = 2 Then

            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        End If
    End If
        
    '[INVERT AIR BRUSH]
    If frmMain.BrushStyle = "Invert Air Brush" Then
        frmMain.PenSize.Enabled = True
        PicBox.DrawWidth = 1
        PicBox.DrawMode = 6 'Solid Brush
        If Button = 1 Or Button = 2 Then

            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        End If
    End If
        
    '[SPRAY PAINT]
    If frmMain.BrushStyle = "Spray Paint" Then
        frmMain.PenSize.Enabled = True
        PicBox.DrawWidth = 1
        PicBox.DrawMode = 13 'Solid Brush
        If Button = 1 Then
           
            PicBox.PSet (X, Y), frmMain.Color1.BackColor
            
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor

            PicBox.PSet (X + 2, Y + 2), frmMain.Color1.BackColor
            PicBox.PSet (X - 2, Y + 2), frmMain.Color1.BackColor
            PicBox.PSet (X + 2, Y - 2), frmMain.Color1.BackColor
            PicBox.PSet (X - 2, Y - 2), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 3, Y + 2), frmMain.Color1.BackColor
            PicBox.PSet (X - 3, Y + 2), frmMain.Color1.BackColor
            PicBox.PSet (X + 3, Y - 2), frmMain.Color1.BackColor
            PicBox.PSet (X - 3, Y - 2), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 2, Y + 3), frmMain.Color1.BackColor
            PicBox.PSet (X - 2, Y + 3), frmMain.Color1.BackColor
            PicBox.PSet (X + 2, Y - 3), frmMain.Color1.BackColor
            PicBox.PSet (X - 2, Y - 3), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 3, Y + 3), frmMain.Color1.BackColor
            PicBox.PSet (X - 3, Y + 3), frmMain.Color1.BackColor
            PicBox.PSet (X + 3, Y - 3), frmMain.Color1.BackColor
            PicBox.PSet (X - 3, Y - 3), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 4, Y + 4), frmMain.Color1.BackColor
            PicBox.PSet (X - 4, Y + 4), frmMain.Color1.BackColor
            PicBox.PSet (X + 4, Y - 4), frmMain.Color1.BackColor
            PicBox.PSet (X - 4, Y - 4), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 5, Y + 6), frmMain.Color1.BackColor
            PicBox.PSet (X - 5, Y + 6), frmMain.Color1.BackColor
            PicBox.PSet (X + 5, Y - 6), frmMain.Color1.BackColor
            PicBox.PSet (X - 5, Y - 6), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 6, Y + 5), frmMain.Color1.BackColor
            PicBox.PSet (X - 6, Y + 5), frmMain.Color1.BackColor
            PicBox.PSet (X + 6, Y - 5), frmMain.Color1.BackColor
            PicBox.PSet (X - 6, Y - 5), frmMain.Color1.BackColor
            
            PicBox.PSet (X, Y + 2), frmMain.Color1.BackColor
            PicBox.PSet (X, Y + 2), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 2), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 2), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 2, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 2, Y), frmMain.Color1.BackColor
            PicBox.PSet (X + 2, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 2, Y), frmMain.Color1.BackColor
            
            PicBox.PSet (X, Y + 3), frmMain.Color1.BackColor
            PicBox.PSet (X, Y + 3), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 3), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 3), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 3, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 3, Y), frmMain.Color1.BackColor
            PicBox.PSet (X + 3, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 3, Y), frmMain.Color1.BackColor
            
            PicBox.PSet (X, Y + 4), frmMain.Color1.BackColor
            PicBox.PSet (X, Y + 4), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 4), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 4), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 4, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 4, Y), frmMain.Color1.BackColor
            PicBox.PSet (X + 4, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 4, Y), frmMain.Color1.BackColor
            
            PicBox.PSet (X, Y + 5), frmMain.Color1.BackColor
            PicBox.PSet (X, Y + 5), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 5, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 5, Y), frmMain.Color1.BackColor
            PicBox.PSet (X + 5, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 5, Y), frmMain.Color1.BackColor
            
            PicBox.PSet (X, Y + 5), frmMain.Color1.BackColor
            PicBox.PSet (X, Y + 5), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 6, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 6, Y), frmMain.Color1.BackColor
            PicBox.PSet (X + 6, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 6, Y), frmMain.Color1.BackColor
        End If
        
        If Button = 2 Then
           
            PicBox.PSet (X, Y), frmMain.Color2.BackColor
            
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
        
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 2, Y + 2), frmMain.Color2.BackColor
            PicBox.PSet (X - 2, Y + 2), frmMain.Color2.BackColor
            PicBox.PSet (X + 2, Y - 2), frmMain.Color2.BackColor
            PicBox.PSet (X - 2, Y - 2), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 3, Y + 2), frmMain.Color2.BackColor
            PicBox.PSet (X - 3, Y + 2), frmMain.Color2.BackColor
            PicBox.PSet (X + 3, Y - 2), frmMain.Color2.BackColor
            PicBox.PSet (X - 3, Y - 2), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 2, Y + 3), frmMain.Color2.BackColor
            PicBox.PSet (X - 2, Y + 3), frmMain.Color2.BackColor
            PicBox.PSet (X + 2, Y - 3), frmMain.Color2.BackColor
            PicBox.PSet (X - 2, Y - 3), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 3, Y + 3), frmMain.Color2.BackColor
            PicBox.PSet (X - 3, Y + 3), frmMain.Color2.BackColor
            PicBox.PSet (X + 3, Y - 3), frmMain.Color2.BackColor
            PicBox.PSet (X - 3, Y - 3), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 4, Y + 4), frmMain.Color2.BackColor
            PicBox.PSet (X - 4, Y + 4), frmMain.Color2.BackColor
            PicBox.PSet (X + 4, Y - 4), frmMain.Color2.BackColor
            PicBox.PSet (X - 4, Y - 4), frmMain.Color2.BackColor
        
            PicBox.PSet (X + 5, Y + 6), frmMain.Color2.BackColor
            PicBox.PSet (X - 5, Y + 6), frmMain.Color2.BackColor
            PicBox.PSet (X + 5, Y - 6), frmMain.Color2.BackColor
            PicBox.PSet (X - 5, Y - 6), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 6, Y + 5), frmMain.Color2.BackColor
            PicBox.PSet (X - 6, Y + 5), frmMain.Color2.BackColor
            PicBox.PSet (X + 6, Y - 5), frmMain.Color2.BackColor
            PicBox.PSet (X - 6, Y - 5), frmMain.Color2.BackColor
            
            PicBox.PSet (X, Y + 2), frmMain.Color2.BackColor
            PicBox.PSet (X, Y + 2), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 2), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 2), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 2, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 2, Y), frmMain.Color2.BackColor
            PicBox.PSet (X + 2, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 2, Y), frmMain.Color2.BackColor
        
            PicBox.PSet (X, Y + 3), frmMain.Color2.BackColor
            PicBox.PSet (X, Y + 3), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 3), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 3), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 3, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 3, Y), frmMain.Color2.BackColor
            PicBox.PSet (X + 3, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 3, Y), frmMain.Color2.BackColor
            
            PicBox.PSet (X, Y + 4), frmMain.Color2.BackColor
            PicBox.PSet (X, Y + 4), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 4), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 4), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 4, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 4, Y), frmMain.Color2.BackColor
            PicBox.PSet (X + 4, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 4, Y), frmMain.Color2.BackColor
            
            PicBox.PSet (X, Y + 5), frmMain.Color2.BackColor
            PicBox.PSet (X, Y + 5), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 5, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 5, Y), frmMain.Color2.BackColor
            PicBox.PSet (X + 5, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 5, Y), frmMain.Color2.BackColor
            
            PicBox.PSet (X, Y + 5), frmMain.Color2.BackColor
            PicBox.PSet (X, Y + 5), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 6, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 6, Y), frmMain.Color2.BackColor
            PicBox.PSet (X + 6, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 6, Y), frmMain.Color2.BackColor
        End If
    End If

 '[STEP GRADIENT BRUSH]
    If frmMain.BrushStyle = "Step Gradient Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepGradientBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If

    '[STEP RED GREEN BLUE BRUSH]
    If frmMain.BrushStyle = "Step Red Green Blue  Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepRGB
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
    
    '[STRIPE BRUSH]
    If frmMain.BrushStyle = "Stripe Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
        If Button = 1 Or Button = 2 Then
        Select Case StripeCase
            Case 0
                PicBox.Line (oldx, oldy)-(X, Y), frmMain.Color1.BackColor
                StripeCase = StripeCase + 1
            Case 1
                PicBox.Line (oldx, oldy)-(X, Y), frmMain.Color2.BackColor
                StripeCase = StripeCase - 1
            End Select
        End If
    End If
    
 '[STEP RED BRUSH]
    If frmMain.BrushStyle = "Step Red Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepRedBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
 '[STEP GREEN BRUSH]
    If frmMain.BrushStyle = "Step Green Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepGreenBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
 '[STEP BLUE BRUSH]
    If frmMain.BrushStyle = "Step Blue Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepBlueBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
 '[STEP RED GREEN BRUSH]
    If frmMain.BrushStyle = "Step Red Green Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepRedBrush
                Call StepGreenBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
 '[STEP RED BLUE BRUSH]
    If frmMain.BrushStyle = "Step Red Blue Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepRedBrush
                Call StepBlueBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
 '[STEP GREEN BLUE BRUSH]
    If frmMain.BrushStyle = "Step Green Blue Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepGreenBrush
                Call StepBlueBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
End If 'This Is The End If For "If Pen Or Marker"
oldx = X
oldy = Y
End Sub
'################################### MOUSE MOVE ##############################################
Private Sub picBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'[MOUSE COORDINATES]
frmMain.StatusBar.Panels.Item(1).Text = "X : " & X & " Y : " & Y
'[COLOR PICKER]
    'If Mouse Is Moving And Color Picker Is Selected Then
    If frmMain.DrawType.Caption = "Color Picker" Then
        On Error Resume Next
        'Set Variable B To The Point (X,Y) You Move Over (If Clicked)
        b = PicBox.Point(X, Y)
        
        If Button <> 1 And Button <> 2 Then
            Exit Sub
        End If
        
        If Button = 1 Then frmMain.Color1.BackColor = b 'If Mouse1 Then Set Color1 To B
        If Button = 2 Then frmMain.Color2.BackColor = b 'If Mouse2 Then Set Color2 To B
        
        '[COLOR TO RGB]
        If Button = 1 Or 2 Then
            texr = PicBox.Point(X, Y) And 255
            texg = (PicBox.Point(X, Y) And 65280) / 256
            texb = (PicBox.Point(X, Y) And 16711680) / 65535
            frmMain.red.Value = texr
            frmMain.green.Value = texg
            frmMain.blue.Value = texb
        End If
    End If
    '[ERASER]
    If frmMain.DrawType.Caption = "Eraser" Then
        frmMain.PenSize.Enabled = True
        PicBox.DrawWidth = frmMain.PenSize.Text
        If Button = 1 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : White
            PicBox.Line (oldx, oldy)-(X, Y), vbWhite
        End If
        If Button = 2 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : White
            PicBox.Line (oldx, oldy)-(X, Y), vbWhite
        End If
    End If
'[PENCIL]

    If frmMain.DrawType.Caption = "Pencil" Then
        frmMain.PenSize.Enabled = False
        PicBox.DrawMode = 13 'solid brush
        PicBox.DrawWidth = 1
        
        If Button > 0 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : Color1
            PicBox.Line (oldx, oldy)-(X, Y), frmMain.Color1.BackColor
        End If
        If Button > 1 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : Color2
            PicBox.Line (oldx, oldy)-(X, Y), frmMain.Color2.BackColor
        End If
    End If
    
'[PEN / MARKER ]

    If frmMain.DrawType.Caption = "Pen" Or frmMain.DrawType.Caption = "Marker" Then
        frmMain.PenSize.Enabled = True
    
    '[CUSTOM BRUSH]
    If frmMain.BrushStyle = "Custom Brush" Then
        If Button = 1 Then
            TransparentBlt PicBox.hdc, X, Y, frmBrush.picBrush.ScaleWidth, frmBrush.picBrush.ScaleHeight, frmBrush.picBrush.hdc, 0, 0, frmBrush.picBrush.ScaleWidth, frmBrush.picBrush.ScaleHeight, frmBrush.Transparent.BackColor
            PicBox.Refresh
        End If
    End If
    
    '[SOLID BRUSH]
    If frmMain.BrushStyle = "Solid Brush" Then
        PicBox.DrawMode = 13 'solid brush
        frmMain.PenSize.Enabled = True
        PicBox.DrawWidth = frmMain.PenSize.Text
        If Button > 0 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : Color1
            PicBox.Line (oldx, oldy)-(X, Y), frmMain.Color1.BackColor
        End If
        If Button > 1 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : Color2
            PicBox.Line (oldx, oldy)-(X, Y), frmMain.Color2.BackColor
        End If
    End If
        
    '[INVERT BRUSH]
    If frmMain.BrushStyle = "Invert Brush" Then
        PicBox.DrawMode = 6 'invert brush
        frmMain.PenSize.Enabled = True
        PicBox.DrawWidth = frmMain.PenSize.Text
        If Button = 1 Then
            'Draw Line Using The Equation (X1, Y1)-(X2, Y2), In Color : black
            PicBox.Line (oldx, oldy)-(X, Y), vbBlack
        End If
    End If
    
    '[CIRCLE BRUSH]
    If frmMain.BrushStyle = "Circle Brush" Then
        PicBox.DrawMode = 13 'solid brush
        PicBox.DrawWidth = 1
        PicBox.FillStyle = 1
        frmMain.PenSize.Enabled = True
        If Button > 0 Then
            PicBox.Circle (X, Y), frmMain.PenSize.Text, frmMain.Color1.BackColor
        End If
        If Button > 1 Then
            PicBox.Circle (X, Y), frmMain.PenSize.Text, frmMain.Color2.BackColor
        End If
    End If
    
    '[AIR BRUSH]
    If frmMain.BrushStyle = "Air Brush" Then
        frmMain.PenSize.Enabled = True
        PicBox.DrawWidth = 1
        PicBox.DrawMode = 13 'Solid Brush
        If Button = 1 Then

            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        End If
        
        If Button = 2 Then

            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        End If
    End If
        
    '[INVERT AIR BRUSH]
    If frmMain.BrushStyle = "Invert Air Brush" Then
        frmMain.PenSize.Enabled = True
        PicBox.DrawWidth = 1
        PicBox.DrawMode = 6 'Solid Brush
        If Button = 1 Or Button = 2 Then

            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        End If
    End If
        
    '[SPRAY PAINT]
    If frmMain.BrushStyle = "Spray Paint" Then
        frmMain.PenSize.Enabled = True
        PicBox.DrawWidth = 1
        PicBox.DrawMode = 13 'Solid Brush
        If Button = 1 Then
           
            PicBox.PSet (X, Y), frmMain.Color1.BackColor
            
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color1.BackColor
            
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color1.BackColor
            
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color1.BackColor
            
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color1.BackColor

            PicBox.PSet (X + 2, Y + 2), frmMain.Color1.BackColor
            PicBox.PSet (X - 2, Y + 2), frmMain.Color1.BackColor
            PicBox.PSet (X + 2, Y - 2), frmMain.Color1.BackColor
            PicBox.PSet (X - 2, Y - 2), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 3, Y + 2), frmMain.Color1.BackColor
            PicBox.PSet (X - 3, Y + 2), frmMain.Color1.BackColor
            PicBox.PSet (X + 3, Y - 2), frmMain.Color1.BackColor
            PicBox.PSet (X - 3, Y - 2), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 2, Y + 3), frmMain.Color1.BackColor
            PicBox.PSet (X - 2, Y + 3), frmMain.Color1.BackColor
            PicBox.PSet (X + 2, Y - 3), frmMain.Color1.BackColor
            PicBox.PSet (X - 2, Y - 3), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 3, Y + 3), frmMain.Color1.BackColor
            PicBox.PSet (X - 3, Y + 3), frmMain.Color1.BackColor
            PicBox.PSet (X + 3, Y - 3), frmMain.Color1.BackColor
            PicBox.PSet (X - 3, Y - 3), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 4, Y + 4), frmMain.Color1.BackColor
            PicBox.PSet (X - 4, Y + 4), frmMain.Color1.BackColor
            PicBox.PSet (X + 4, Y - 4), frmMain.Color1.BackColor
            PicBox.PSet (X - 4, Y - 4), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 5, Y + 6), frmMain.Color1.BackColor
            PicBox.PSet (X - 5, Y + 6), frmMain.Color1.BackColor
            PicBox.PSet (X + 5, Y - 6), frmMain.Color1.BackColor
            PicBox.PSet (X - 5, Y - 6), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 6, Y + 5), frmMain.Color1.BackColor
            PicBox.PSet (X - 6, Y + 5), frmMain.Color1.BackColor
            PicBox.PSet (X + 6, Y - 5), frmMain.Color1.BackColor
            PicBox.PSet (X - 6, Y - 5), frmMain.Color1.BackColor
            
            PicBox.PSet (X, Y + 2), frmMain.Color1.BackColor
            PicBox.PSet (X, Y + 2), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 2), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 2), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 2, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 2, Y), frmMain.Color1.BackColor
            PicBox.PSet (X + 2, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 2, Y), frmMain.Color1.BackColor
            
            PicBox.PSet (X, Y + 3), frmMain.Color1.BackColor
            PicBox.PSet (X, Y + 3), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 3), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 3), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 3, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 3, Y), frmMain.Color1.BackColor
            PicBox.PSet (X + 3, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 3, Y), frmMain.Color1.BackColor
            
            PicBox.PSet (X, Y + 4), frmMain.Color1.BackColor
            PicBox.PSet (X, Y + 4), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 4), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 4), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 4, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 4, Y), frmMain.Color1.BackColor
            PicBox.PSet (X + 4, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 4, Y), frmMain.Color1.BackColor
            
            PicBox.PSet (X, Y + 5), frmMain.Color1.BackColor
            PicBox.PSet (X, Y + 5), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 5, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 5, Y), frmMain.Color1.BackColor
            PicBox.PSet (X + 5, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 5, Y), frmMain.Color1.BackColor
            
            PicBox.PSet (X, Y + 5), frmMain.Color1.BackColor
            PicBox.PSet (X, Y + 5), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color1.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color1.BackColor
            
            PicBox.PSet (X + 6, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 6, Y), frmMain.Color1.BackColor
            PicBox.PSet (X + 6, Y), frmMain.Color1.BackColor
            PicBox.PSet (X - 6, Y), frmMain.Color1.BackColor
        End If
        
        If Button = 2 Then
           
            PicBox.PSet (X, Y), frmMain.Color2.BackColor
            
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * frmMain.PenSize.Text), Y - Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
        
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * frmMain.PenSize.Text), Y + Int(Rnd * frmMain.PenSize.Text)), frmMain.Color2.BackColor
            
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
        
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y + Int(Rnd * 20)), frmMain.Color2.BackColor
            
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 10), Y + Int(Rnd * 10)), frmMain.Color2.BackColor
            
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X - Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            PicBox.PSet (X + Int(Rnd * 20), Y - Int(Rnd * 20)), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 2, Y + 2), frmMain.Color2.BackColor
            PicBox.PSet (X - 2, Y + 2), frmMain.Color2.BackColor
            PicBox.PSet (X + 2, Y - 2), frmMain.Color2.BackColor
            PicBox.PSet (X - 2, Y - 2), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 3, Y + 2), frmMain.Color2.BackColor
            PicBox.PSet (X - 3, Y + 2), frmMain.Color2.BackColor
            PicBox.PSet (X + 3, Y - 2), frmMain.Color2.BackColor
            PicBox.PSet (X - 3, Y - 2), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 2, Y + 3), frmMain.Color2.BackColor
            PicBox.PSet (X - 2, Y + 3), frmMain.Color2.BackColor
            PicBox.PSet (X + 2, Y - 3), frmMain.Color2.BackColor
            PicBox.PSet (X - 2, Y - 3), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 3, Y + 3), frmMain.Color2.BackColor
            PicBox.PSet (X - 3, Y + 3), frmMain.Color2.BackColor
            PicBox.PSet (X + 3, Y - 3), frmMain.Color2.BackColor
            PicBox.PSet (X - 3, Y - 3), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 4, Y + 4), frmMain.Color2.BackColor
            PicBox.PSet (X - 4, Y + 4), frmMain.Color2.BackColor
            PicBox.PSet (X + 4, Y - 4), frmMain.Color2.BackColor
            PicBox.PSet (X - 4, Y - 4), frmMain.Color2.BackColor
        
            PicBox.PSet (X + 5, Y + 6), frmMain.Color2.BackColor
            PicBox.PSet (X - 5, Y + 6), frmMain.Color2.BackColor
            PicBox.PSet (X + 5, Y - 6), frmMain.Color2.BackColor
            PicBox.PSet (X - 5, Y - 6), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 6, Y + 5), frmMain.Color2.BackColor
            PicBox.PSet (X - 6, Y + 5), frmMain.Color2.BackColor
            PicBox.PSet (X + 6, Y - 5), frmMain.Color2.BackColor
            PicBox.PSet (X - 6, Y - 5), frmMain.Color2.BackColor
            
            PicBox.PSet (X, Y + 2), frmMain.Color2.BackColor
            PicBox.PSet (X, Y + 2), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 2), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 2), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 2, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 2, Y), frmMain.Color2.BackColor
            PicBox.PSet (X + 2, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 2, Y), frmMain.Color2.BackColor
        
            PicBox.PSet (X, Y + 3), frmMain.Color2.BackColor
            PicBox.PSet (X, Y + 3), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 3), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 3), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 3, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 3, Y), frmMain.Color2.BackColor
            PicBox.PSet (X + 3, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 3, Y), frmMain.Color2.BackColor
            
            PicBox.PSet (X, Y + 4), frmMain.Color2.BackColor
            PicBox.PSet (X, Y + 4), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 4), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 4), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 4, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 4, Y), frmMain.Color2.BackColor
            PicBox.PSet (X + 4, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 4, Y), frmMain.Color2.BackColor
            
            PicBox.PSet (X, Y + 5), frmMain.Color2.BackColor
            PicBox.PSet (X, Y + 5), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 5, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 5, Y), frmMain.Color2.BackColor
            PicBox.PSet (X + 5, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 5, Y), frmMain.Color2.BackColor
            
            PicBox.PSet (X, Y + 5), frmMain.Color2.BackColor
            PicBox.PSet (X, Y + 5), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color2.BackColor
            PicBox.PSet (X, Y - 5), frmMain.Color2.BackColor
            
            PicBox.PSet (X + 6, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 6, Y), frmMain.Color2.BackColor
            PicBox.PSet (X + 6, Y), frmMain.Color2.BackColor
            PicBox.PSet (X - 6, Y), frmMain.Color2.BackColor
        End If
    End If

 '[STEP GRADIENT BRUSH]
    If frmMain.BrushStyle = "Step Gradient Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepGradientBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If

    '[STEP RED GREEN BLUE BRUSH]
    If frmMain.BrushStyle = "Step Red Green Blue  Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepRGB
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
    
    '[STRIPE BRUSH]
    If frmMain.BrushStyle = "Stripe Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
        If Button = 1 Or Button = 2 Then
        Select Case StripeCase
            Case 0
                PicBox.Line (oldx, oldy)-(X, Y), frmMain.Color1.BackColor
                StripeCase = StripeCase + 1
            Case 1
                PicBox.Line (oldx, oldy)-(X, Y), frmMain.Color2.BackColor
                StripeCase = StripeCase - 1
            End Select
        End If
    End If
    
 '[STEP RED BRUSH]
    If frmMain.BrushStyle = "Step Red Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepRedBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
 '[STEP GREEN BRUSH]
    If frmMain.BrushStyle = "Step Green Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepGreenBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
 '[STEP BLUE BRUSH]
    If frmMain.BrushStyle = "Step Blue Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepBlueBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
 '[STEP RED GREEN BRUSH]
    If frmMain.BrushStyle = "Step Red Green Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepRedBrush
                Call StepGreenBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
 '[STEP RED BLUE BRUSH]
    If frmMain.BrushStyle = "Step Red Blue Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepRedBrush
                Call StepBlueBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
 '[STEP GREEN BLUE BRUSH]
    If frmMain.BrushStyle = "Step Green Blue Brush" Then
        PicBox.DrawWidth = frmMain.PenSize.Text
        PicBox.DrawMode = 13 'Solid Brush
            If Button = 1 Or Button = 2 Then
                Call StepGreenBrush
                Call StepBlueBrush
            If Button = 1 Then
                frmMain.m1 = True
                frmMain.m2 = False
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
            If Button = 2 Then
                frmMain.m1 = False
                frmMain.m2 = True
                PicBox.Line (oldx, oldy)-(X, Y), RGB(frmMain.red.Value, frmMain.green.Value, frmMain.blue.Value)
            End If
        End If
    End If
End If 'This Is The End If For "If Pen Or Marker"
oldx = X
oldy = Y
End Sub
