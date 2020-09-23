VERSION 5.00
Begin VB.Form frmBrush 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brush Editor"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2520
   Icon            =   "frmBrush.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   2520
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton ClearBrush 
         Caption         =   "Clear Current Brush"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Caption         =   "Transparent Color"
         Height          =   1095
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   2175
         Begin VB.PictureBox Transparent 
            BackColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   120
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   15
            ToolTipText     =   "RGB(255,255,255)"
            Top             =   240
            Width           =   735
         End
         Begin VB.HScrollBar tb 
            Height          =   255
            Left            =   840
            Max             =   255
            TabIndex        =   14
            Top             =   720
            Value           =   255
            Width           =   1215
         End
         Begin VB.HScrollBar tg 
            Height          =   255
            Left            =   840
            Max             =   255
            TabIndex        =   13
            Top             =   480
            Value           =   255
            Width           =   1215
         End
         Begin VB.HScrollBar tr 
            Height          =   255
            Left            =   840
            Max             =   255
            TabIndex        =   12
            Top             =   240
            Value           =   255
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Editing Color"
         Height          =   1095
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   2175
         Begin VB.PictureBox color 
            BackColor       =   &H00000000&
            Height          =   735
            Left            =   120
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   10
            ToolTipText     =   "RGB(0,0,0)"
            Top             =   240
            Width           =   735
         End
         Begin VB.HScrollBar b 
            Height          =   255
            Left            =   840
            Max             =   255
            TabIndex        =   9
            Top             =   720
            Width           =   1215
         End
         Begin VB.HScrollBar g 
            Height          =   255
            Left            =   840
            Max             =   255
            TabIndex        =   8
            Top             =   480
            Width           =   1215
         End
         Begin VB.HScrollBar r 
            Height          =   255
            Left            =   840
            Max             =   255
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Brush"
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   975
         Begin VB.PictureBox picBrush 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   120
            ScaleHeight     =   45
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   45
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox PenSize 
         Height          =   315
         ItemData        =   "frmBrush.frx":08CA
         Left            =   1200
         List            =   "frmBrush.frx":08EC
         TabIndex        =   3
         Text            =   "1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Pen 
         Caption         =   "Pen"
         Height          =   255
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton FloodFill 
         Caption         =   "Flood Fill"
         Height          =   255
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldx
Dim oldy
Private Sub exit_Click()
End
End Sub

'################################ BRUSH EDIT #################################

Private Sub PenSize_Change()
If PenSize.Text > 20 Then
PenSize.Text = 20
MsgBox "Size Must Be Between 1 And 20", vbCritical, "Brush Editor"
End If
If PenSize.Text <= 0 Then
PenSize.Text = 1
MsgBox "Size Must Be Between 1 And 20", vbCritical, "Brush Editor"
End If
picBrush.DrawWidth = PenSize.Text
End Sub

Private Sub picBrush_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If Pen.Value = True Then
            picBrush.Line (X, Y)-(X, Y), color.BackColor
        End If
        If FloodFill.Value = True Then
            picBrush.FillStyle = 0
            ExtFloodFill picBrush.hdc, X, Y, picBrush.Point(X, Y), 1
        End If
    End If
End Sub
Private Sub picBrush_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xR, xG, xB
    xR = picBrush.Point(X, Y) And 255
    xG = (picBrush.Point(X, Y) And 65280) / 256
    xB = (picBrush.Point(X, Y) And 16711680) / 65535
    picBrush.ToolTipText = "RGB(" & xR & "," & xG & "," & Int(xB) & ")"
    picBrush.FillColor = RGB(r.Value, g.Value, b.Value)
    picBrush.DrawWidth = PenSize.Text
    If Button = 1 Then
        If Pen.Value = True Then
            picBrush.Line (oldx, oldy)-(X, Y), color.BackColor
        End If
    End If
    oldx = X
    oldy = Y
End Sub
'################################# COLOR MIXERS ########################################
'[EDIT COLOR]
Private Sub r_Change()
color.ToolTipText = "RGB(" & r.Value & "," & g.Value & "," & b.Value & ")"
color.BackColor = RGB(r.Value, g.Value, b.Value)
End Sub
Private Sub b_Change()
color.ToolTipText = "RGB(" & r.Value & "," & g.Value & "," & b.Value & ")"
color.BackColor = RGB(r.Value, g.Value, b.Value)
End Sub
Private Sub g_Change()
color.ToolTipText = "RGB(" & r.Value & "," & g.Value & "," & b.Value & ")"
color.BackColor = RGB(r.Value, g.Value, b.Value)
End Sub
Private Sub r_Scroll()
color.ToolTipText = "RGB(" & r.Value & "," & g.Value & "," & b.Value & ")"
color.BackColor = RGB(r.Value, g.Value, b.Value)
End Sub
Private Sub g_Scroll()
color.ToolTipText = "RGB(" & r.Value & "," & g.Value & "," & b.Value & ")"
color.BackColor = RGB(r.Value, g.Value, b.Value)
End Sub
Private Sub b_Scroll()
color.ToolTipText = "RGB(" & r.Value & "," & g.Value & "," & b.Value & ")"
color.BackColor = RGB(r.Value, g.Value, b.Value)
End Sub
'[TRANSPARENT COLOR]
Private Sub tr_Change()
Transparent.ToolTipText = "RGB(" & tr.Value & "," & tg.Value & "," & tb.Value & ")"
Transparent.BackColor = RGB(tr.Value, tg.Value, tb.Value)
End Sub
Private Sub tb_Change()
Transparent.ToolTipText = "RGB(" & tr.Value & "," & tg.Value & "," & tb.Value & ")"
Transparent.BackColor = RGB(tr.Value, tg.Value, tb.Value)
End Sub
Private Sub tg_Change()
Transparent.ToolTipText = "RGB(" & tr.Value & "," & tg.Value & "," & tb.Value & ")"
Transparent.BackColor = RGB(tr.Value, tg.Value, tb.Value)
End Sub
Private Sub tr_Scroll()
Transparent.ToolTipText = "RGB(" & tr.Value & "," & tg.Value & "," & tb.Value & ")"
Transparent.BackColor = RGB(tr.Value, tg.Value, tb.Value)
End Sub
Private Sub tg_Scroll()
Transparent.ToolTipText = "RGB(" & tr.Value & "," & tg.Value & "," & tb.Value & ")"
Transparent.BackColor = RGB(tr.Value, tg.Value, tb.Value)
End Sub
Private Sub tb_Scroll()
Transparent.ToolTipText = "RGB(" & tr.Value & "," & tg.Value & "," & tb.Value & ")"
Transparent.BackColor = RGB(tr.Value, tg.Value, tb.Value)
End Sub


