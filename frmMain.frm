VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "M-Paint Version 3.00"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11160
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6960
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolBar 
      Left            =   6840
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "Pencil"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D1C
            Key             =   "Pen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":116E
            Key             =   "Marker"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22EC
            Key             =   "Fill"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":273E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3272
            Key             =   "Options"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6240
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B4C
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C5E
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D70
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E82
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F94
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40A6
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41B8
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42CA
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4424
            Key             =   "horizontally"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":457E
            Key             =   "vertically"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46D8
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4832
            Key             =   "rotate"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cascade"
            Object.ToolTipText     =   "Cascade"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tile Horizontal"
            Object.ToolTipText     =   "Tile Horizontal"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tile Vertical"
            Object.ToolTipText     =   "Tile Vertical"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   9165
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14023
            Text            =   "Status"
            TextSave        =   "Status"
            Key             =   "status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "5/28/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "5:28 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   3  'Align Left
      Height          =   8550
      Left            =   0
      TabIndex        =   2
      Top             =   615
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   15081
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imlToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pencil"
            Object.ToolTipText     =   "Pencil"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pen"
            Object.ToolTipText     =   "Pen"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Marker"
            Object.ToolTipText     =   "Marker"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Eraser"
            Object.ToolTipText     =   "Eraser"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Color Picker"
            Object.ToolTipText     =   "Color Picker"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Flood Fill"
            Object.ToolTipText     =   "Flood Fill"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar3 
      Align           =   4  'Align Right
      Height          =   8550
      Left            =   7590
      TabIndex        =   3
      Top             =   615
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   15081
      ButtonWidth     =   5292
      ButtonHeight    =   5292
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   0
         TabIndex        =   23
         Top             =   360
         Width           =   3495
         Begin VB.ComboBox PenSize 
            Height          =   315
            ItemData        =   "frmMain.frx":498C
            Left            =   600
            List            =   "frmMain.frx":4A26
            TabIndex        =   26
            Text            =   "1"
            Top             =   120
            Width           =   2775
         End
         Begin VB.ComboBox BrushStyle 
            Height          =   315
            ItemData        =   "frmMain.frx":4AEE
            Left            =   600
            List            =   "frmMain.frx":4B22
            TabIndex        =   25
            Text            =   "Solid Brush"
            Top             =   480
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox xFillStyle 
            Height          =   315
            ItemData        =   "frmMain.frx":4C3B
            Left            =   600
            List            =   "frmMain.frx":4C54
            TabIndex        =   24
            Text            =   "Solid"
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label5 
            Caption         =   "Style"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Size"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Height          =   375
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   3495
         Begin VB.Label DrawType 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pencil"
            Height          =   255
            Left            =   600
            TabIndex        =   22
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5295
         Left            =   0
         TabIndex        =   4
         Top             =   1440
         Width           =   3495
         Begin VB.OptionButton m2 
            Caption         =   "Apply To Mouse 2"
            Height          =   435
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   4680
            Width           =   3255
         End
         Begin VB.OptionButton m1 
            Caption         =   "Apply To Mouse 1"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   4200
            Value           =   -1  'True
            Width           =   3255
         End
         Begin VB.TextBox txtblue 
            Height          =   285
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   14
            Text            =   "0"
            Top             =   3840
            Width           =   495
         End
         Begin VB.TextBox txtgreen 
            Height          =   285
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   13
            Text            =   "0"
            Top             =   3480
            Width           =   495
         End
         Begin VB.TextBox txtred 
            Height          =   285
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   12
            Text            =   "0"
            Top             =   3120
            Width           =   495
         End
         Begin VB.PictureBox ShapeColor 
            BackColor       =   &H00000000&
            Height          =   975
            Left            =   2760
            ScaleHeight     =   915
            ScaleWidth      =   555
            TabIndex        =   11
            Top             =   3120
            Width           =   615
         End
         Begin VB.HScrollBar blue 
            Height          =   285
            Left            =   600
            Max             =   255
            TabIndex        =   10
            Top             =   3840
            Width           =   1575
         End
         Begin VB.HScrollBar green 
            Height          =   285
            Left            =   600
            Max             =   255
            TabIndex        =   9
            Top             =   3480
            Width           =   1575
         End
         Begin VB.HScrollBar red 
            Height          =   285
            Left            =   600
            Max             =   255
            TabIndex        =   8
            Top             =   3120
            Width           =   1575
         End
         Begin VB.PictureBox Picture3 
            Height          =   560
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   5
            Top             =   240
            Width           =   560
            Begin VB.PictureBox Color1 
               BackColor       =   &H00000000&
               Height          =   255
               Left            =   60
               ScaleHeight     =   195
               ScaleWidth      =   195
               TabIndex        =   6
               Top             =   60
               Width           =   255
            End
            Begin VB.PictureBox Color2 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   160
               ScaleHeight     =   195
               ScaleWidth      =   195
               TabIndex        =   7
               Top             =   160
               Width           =   255
            End
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   2175
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   3836
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Swatches"
            TabPicture(0)   =   "frmMain.frx":4CB0
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Swatches"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Gradient"
            TabPicture(1)   =   "frmMain.frx":4CCC
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Gradient"
            Tab(1).ControlCount=   1
            Begin VB.PictureBox Gradient 
               AutoSize        =   -1  'True
               Height          =   900
               Left            =   -74840
               MousePointer    =   2  'Cross
               Picture         =   "frmMain.frx":4CE8
               ScaleHeight     =   840
               ScaleWidth      =   2880
               TabIndex        =   17
               Top             =   480
               Width           =   2940
            End
            Begin VB.PictureBox Swatches 
               AutoSize        =   -1  'True
               Height          =   1520
               Left            =   120
               MousePointer    =   2  'Cross
               Picture         =   "frmMain.frx":CB2A
               ScaleHeight     =   1455
               ScaleWidth      =   2985
               TabIndex        =   16
               Top             =   480
               Width           =   3045
            End
         End
         Begin VB.Label Label3 
            Caption         =   "Blue"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   3840
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Green"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Red"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   3120
            Width           =   375
         End
      End
   End
   Begin MSComctlLib.ProgressBar PG 
      Align           =   1  'Align Top
      Height          =   195
      Left            =   0
      TabIndex        =   31
      Top             =   420
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "Filter"
      Begin VB.Menu mnuColor 
         Caption         =   "Colors"
         Begin VB.Menu mnuColorize 
            Caption         =   "Colorize"
            Begin VB.Menu RedFilter 
               Caption         =   "Red Filter"
            End
            Begin VB.Menu GreenFilter 
               Caption         =   "Green Filter"
            End
            Begin VB.Menu BlueFilter 
               Caption         =   "Blue Filter"
            End
         End
         Begin VB.Menu mnuGrey_Scale 
            Caption         =   "Grey Scale"
            Begin VB.Menu grey 
               Caption         =   "Grey Scale"
            End
            Begin VB.Menu greyTop 
               Caption         =   "Grey Scale (Top Half)"
            End
            Begin VB.Menu greyBottom 
               Caption         =   "Grey Scale (Bottom Half)"
            End
            Begin VB.Menu greyLeft 
               Caption         =   "Grey Scale (Left Half)"
            End
            Begin VB.Menu greyRight 
               Caption         =   "Grey Scale (Right Half)"
            End
         End
         Begin VB.Menu mnuInvert_Colors 
            Caption         =   "Invert Colors"
            Begin VB.Menu invert 
               Caption         =   "Invert Colors"
            End
            Begin VB.Menu invertTop 
               Caption         =   "Invert (Top Half)"
            End
            Begin VB.Menu invertBottom 
               Caption         =   "Invert (Bottom Half)"
            End
            Begin VB.Menu invertLeft 
               Caption         =   "Invert (Left Half)"
            End
            Begin VB.Menu invertRight 
               Caption         =   "Invert (Right Half)"
            End
         End
         Begin VB.Menu NoRed 
            Caption         =   "No Red"
         End
         Begin VB.Menu NoGreen 
            Caption         =   "No Green"
         End
         Begin VB.Menu NoBlue 
            Caption         =   "No Blue"
         End
      End
      Begin VB.Menu mnuBrighten 
         Caption         =   "Brighten"
         Begin VB.Menu bright1 
            Caption         =   "Brighten By 1"
         End
         Begin VB.Menu bright5 
            Caption         =   "Brighten By 5"
         End
         Begin VB.Menu bright10 
            Caption         =   "Brighten By 10"
         End
      End
      Begin VB.Menu mnuDarken 
         Caption         =   "Darken"
         Begin VB.Menu dark1 
            Caption         =   "Darken By 1"
         End
         Begin VB.Menu dark5 
            Caption         =   "Darken By 5"
         End
         Begin VB.Menu dark10 
            Caption         =   "Darken By 10"
         End
      End
      Begin VB.Menu mnuNoise 
         Caption         =   "Noise"
         Begin VB.Menu AddBrightNoise 
            Caption         =   "Add Bright Noise"
         End
         Begin VB.Menu fastBrightNoise 
            Caption         =   "Add Fast Bright Noise "
         End
         Begin VB.Menu AddDarkNoise 
            Caption         =   "Add Dark Noise"
         End
         Begin VB.Menu fastDarkNoise 
            Caption         =   "Add Fast Dark Noise"
         End
      End
      Begin VB.Menu Stripes 
         Caption         =   "Stripes"
         Begin VB.Menu addVertical 
            Caption         =   "Add Vertical"
         End
         Begin VB.Menu addHorizontal 
            Caption         =   "Add Horizontal"
         End
         Begin VB.Menu addDotGrid 
            Caption         =   "Add Dot Grid"
         End
      End
      Begin VB.Menu fad 
         Caption         =   "Fade"
         Begin VB.Menu fadeRightDown 
            Caption         =   "Fade (Top Right Corner)"
         End
         Begin VB.Menu fadeLeftUp 
            Caption         =   "Fade (Bottom Left Corner)"
         End
         Begin VB.Menu fadeBottom 
            Caption         =   "Fade (Bottom Half)"
         End
         Begin VB.Menu fadeRight 
            Caption         =   "Fade (Right Half)"
         End
      End
      Begin VB.Menu mnuDodge_Burn 
         Caption         =   "Dodge / Burn"
         Begin VB.Menu mnuDodge 
            Caption         =   "Dodge"
            Begin VB.Menu Dodge 
               Caption         =   "Dodge"
            End
            Begin VB.Menu fastDodge 
               Caption         =   "Fast Dodge (Lower Quality)"
            End
            Begin VB.Menu dodgeTop 
               Caption         =   "Dodge (Top Half)"
            End
            Begin VB.Menu dodgeBottom 
               Caption         =   "Dodge (Bottom Half)"
            End
            Begin VB.Menu dodgeLeft 
               Caption         =   "Dodge (Left Half)"
            End
            Begin VB.Menu dodgeRight 
               Caption         =   "Dodge (Right Half)"
            End
         End
         Begin VB.Menu mnuBurn 
            Caption         =   "Burn"
            Begin VB.Menu Burn 
               Caption         =   "Burn"
            End
            Begin VB.Menu fastBurn 
               Caption         =   "Fast Burn (Lower Quality)"
            End
            Begin VB.Menu burnTop 
               Caption         =   "Burn (Top Half)"
            End
            Begin VB.Menu burnBottom 
               Caption         =   "Burn (Bottom Half)"
            End
            Begin VB.Menu burnLeft 
               Caption         =   "Burn (Left Half)"
            End
            Begin VB.Menu burnRight 
               Caption         =   "Burn (Right Half) "
            End
         End
      End
   End
   Begin VB.Menu win 
      Caption         =   "Window"
      Begin VB.Menu mnuVertical 
         Caption         =   "Tile Vertical"
      End
      Begin VB.Menu mnuHorizontal 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'           ***************************************************
'           *              M-Paint Version 3.00               *
'           *  By Mike Plaehn (LCSBSSRHXXX) 4/10/04 - 5/8/04 *
'           ***************************************************

'Toatal         : 3473  Lines

'frmSplash      : 9     Lines
'mdlSplash      : 15    Lines
'mdlStep        : 140   Lines
'mdlAPI         : 21    Lines
'frmMain        : 1688  Lines
'frmDocument    : 1600  Lines

'As You May Notice This Is Heavily Commented, Feel Free To Learn From This
'But Please Don't Use Any Of My Code With Out My Permission Thank You!
'Enjoi!
' - LCSBSSRHXXX

'This Is A Paint Program That I Am Working On.

'frmMain' Is A MDI Parent / Form For The MDI Child 'frmDocument'
'A MDI Parent Has A Form Inside Of It, Giving The Program A Nice Effect

'View mdlSplash.bas (Module) For Sub Main

'       ************************
'       * M-Paint Version 3.00 *
'       *      Fetures:        *
'       ************************

'[NEW FETURES]
'###NEW!### Color To RGB           (4/16/04)
'###NEW!### Mouse Coordinates      (4/16/04)
'###NEW!### Air Brush              (4/21/04)
'###NEW!### Spray Paint            (4/21/04)

'###NEW!### Circle Brush            (5/7/04)
'###NEW!### Step Gradient Brush     (5/8/04)
'###NEW!### Step Brush RGB         (4/29/04)
'###NEW!### Step Brush RG          (4/29/04)
'###NEW!### Step Brush RB          (4/29/04)
'###NEW!### Step Brush GB          (4/29/04)
'###NEW!### Step Brush R           (4/29/04)
'###NEW!### Step Brush G           (4/29/04)
'###NEW!### Step Brush B           (4/29/04)
'###NEW!### Stripe Brush           (4/21/04)

'*******************************************
'############# NEW! Filters ################
'*******************************************
'# 1# Grey Scale                (4/16/04)###
'# 2# Grey Scale Top Half       (4/18/04)###
'# 3# Grey Scale Bottom Half    (4/18/04)###
'# 4# Grey Scale Left Half      (4/18/04)###
'# 5# Grey Scale Right Half     (4/18/04)###
'# 6# Invert Colors             (4/17/04)###
'# 7# Invert Colors Top Half    (4/18/04)###
'# 8# Invert Colors Bottom Half (4/18/04)###
'# 9# Invert Colors Left Half   (4/18/04)###
'#10# Invert Colors Right Half  (4/18/04)###
'#11# Brighten 1                (4/17/04)###
'#12# Brighten 5                (4/17/04)###
'#13# Brighten 10               (4/17/04)###
'#14# Darken 1                  (4/17/04)###
'#15# Darken 5                  (4/17/04)###
'#16# Darken 10                 (4/17/04)###
'#17# Add Bright Noise          (4/17/04)###
'#18# Add Dark Noise            (4/17/04)###
'#19# Fade Bottom               (4/17/04)###
'#20# Fade Bottom Left          (4/17/04)###
'#21# Fade Right                (4/17/04)###
'#22# Fade Top Right            (4/17/04)###
'#23# No Red                    (4/17/04)###
'#24# No Green                  (4/17/04)###
'#25# No Blue                   (4/17/04)###
'#26# Red Filter                (4/17/04)###
'#27# Green Filter              (4/17/04)###
'#28# Blue Filter               (4/17/04)###
'#29# Dodge                     (4/17/04)###
'#30# Dodge Top Half            (4/18/04)###
'#31# Dodge Bottom Half         (4/18/04)###
'#32# Dodge Left Half           (4/18/04)###
'#33# Dodge Right Half          (4/18/04)###
'#34# Burn                      (4/17/04)###
'#35# Burn Top Half             (4/18/04)###
'#36# Burn Bottom Half          (4/18/04)###
'#37# Burn Left Half            (4/18/04)###
'#38# Burn Right Half           (4/18/04)###
'#39# Fast Burn                  (5/1/04)###
'#40# Fast Dodge                 (5/1/04)###
'#38# Fast Bright Noise          (5/1/04)###
'#38# Fast Dark Noise            (5/1/04)###
'#39# Vertical Lines             (5/7/04)###
'#40# Horizontal Lines           (5/7/04)###
'#41# Dot Grid                   (5/7/04)###
'*******************************************
'###########################################
'*******************************************
'[WORKING TOOLS]
'Pencil                         (4/10/04)
'Pen                            (4/10/04)
'Eraser                         (4/11/04)
'Color Picker                   (4/13/04)
'Flood Fill                     (4/13/04)
'[WORKING UTILITIES]
'Color Swatches                 (4/14/04)
'Color Gradients                (4/14/04)
'Custom Color Mixer             (4/11/04)
'[OTHER]
'MDI Forms                      (4/10/04)
'Save                           (4/13/04)
'Load                           (4/13/04)
'Tile Windows Vertically        (4/10/04)
'Tile Windows Horizontally      (4/10/04)
'Cascade Windows                (4/10/04)
'Exit                           (4/10/04)

'[UNFINISHED FEATURES]
'Marker
'Print
'Rotate
'Magnify
'Cut
'Copy
'Paste
'Move
'Rectangle
'Elipse

'[VARIABLES]
Dim a 'a Is Used As The Number To Add To frmD's Caption When It Is Loaded
Dim frmD As frmDocument 'frmD Is The frmDocument That Is Currently Active
Dim greycolor As Integer 'greycolor Is A Variable I Used For The Grey Scale Filter, _
Grey Scale Is The Pixles RGB Values / 3 (The Avearge) Then It Is Plugged In _
As The Colors New RGB Values Since The Are All The Same, They Are In Grey Scale

'##################################### FILTERS ###############################################
'              The Next 849 Lines Of Code Are Filters ^^
'              I Spent A Few Days Writing These Using Math And Trial And Error.
'              So Please Don't Use Use My Code For The Filters With Out My Permission
'
'              How Alot Of These Work Is They Scroll Throgh The Pictures Pixles
'              Then Get The RGB Value Of The Pixle Then Modify The Pixle's RGB
'              Value Then Set A New Pixle Over The Old One, Applying The New Effect.
'
'              The Progress Bar That Moniters The Progress Of The Filter Is Equal To
'              The Y Value That The Filter Is Scrolling Through, Making It Accurate
'#############################################################################################

'##################################### BRIGHTEN / DARKEN #####################################

'[BRIGHTEN + 1]
Private Sub bright1_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Brighter Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr + 1, texg + 1, texb + 1)
    Next X
    Next Y
End Sub
'[BRIGHTEN + 5]
Private Sub bright5_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Brighter Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr + 5, texg + 5, texb + 5)
    Next X
    Next Y
End Sub
'[BRIGHTEN + 10]
Private Sub bright10_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Brighter Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr + 10, texg + 10, texb + 10)
    Next X
    Next Y
End Sub

Private Sub BrushStyle_Change()
    If BrushStyle.Text = "Custom Brush" Then
        frmBrush.Show
    Else
        Unload frmBrush
    End If
End Sub
Private Sub BrushStyle_Click()
    If BrushStyle.Text = "Custom Brush" Then
        frmBrush.Show
    Else
        Unload frmBrush
    End If
End Sub

'[DARKEN - 1]
Private Sub dark1_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Darker Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr - 1, texg - 1, texb - 1)
    Next X
    Next Y
End Sub
'[DARKEN - 5]
Private Sub dark5_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Darker Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr - 5, texg - 5, texb - 5)
    Next X
    Next Y
End Sub
'[DARKEN - 10]
Private Sub dark10_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Darker Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr - 10, texg - 10, texb - 10)
    Next X
    Next Y
End Sub

Private Sub DrawType_Change()
    If Not DrawType.Caption = "Pen" Or DrawType.Caption = "Marker" Then
        Unload frmBrush
    End If
End Sub

'######################################## GREY SCALE #########################################

'[GREY SCALE]
Private Sub grey_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'The Grey Color = Average Of RGB In Pixle
        greycolor = (texr + texg + texb) / 3
        'Set Pixle With Grey Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(greycolor, greycolor, greycolor)
    Next X
    Next Y
End Sub
'[GREY SCALE (TOP HALF)]
Private Sub greyTop_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY) / 2
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY) / 2
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'The Grey Color = Average Of RGB In Pixle
        greycolor = (texr + texg + texb) / 3
        'Set Pixle With Grey Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(greycolor, greycolor, greycolor)
    Next X
    Next Y
End Sub
'[GREY SCALE (BOTTOM HALF)]
Private Sub greyBottom_Click()
    For Y = (frmD.PicBox.Height / Screen.TwipsPerPixelY) / 2 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'The Grey Color = Average Of RGB In Pixle
        greycolor = (texr + texg + texb) / 3
        'Set Pixle With Grey Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(greycolor, greycolor, greycolor)
    Next X
    Next Y
End Sub
'[GREY SCALE (LEFT HALF)]
Private Sub greyLeft_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX) / 2
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'The Grey Color = Average Of RGB In Pixle
        greycolor = (texr + texg + texb) / 3
        'Set Pixle With Grey Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(greycolor, greycolor, greycolor)
    Next X
    Next Y
End Sub
'[GREY SCALE (RIGHT HALF)]
Private Sub greyRight_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = (frmD.PicBox.Height / Screen.TwipsPerPixelX) / 2 To (frmD.PicBox.Height / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'The Grey Color = Average Of RGB In Pixle
        greycolor = (texr + texg + texb) / 3
        'Set Pixle With Grey Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(greycolor, greycolor, greycolor)
    Next X
    Next Y
End Sub

'####################################### INVERT COLORS #######################################

'[INVERT COLORS]
Private Sub invert_Click()
    GetObject frmD.PicBox.Image, Len(PicInfo), PicInfo
    BytesPerLine = (PicInfo.bmWidth * 3 + 3) And &HFFFFFFFC
    ReDim PicBits(1 To BytesPerLine * PicInfo.bmHeight * 3) As Byte
    GetBitmapBits frmD.PicBox.Image, UBound(PicBits), PicBits(1)
    For Cnt = 1 To UBound(PicBits)
        PicBits(Cnt) = 255 - PicBits(Cnt)
    Next Cnt
    SetBitmapBits frmD.PicBox.Image, UBound(PicBits), PicBits(1)
    frmD.PicBox.Refresh
End Sub
'[INVERT (TOP HALF)]
Private Sub invertTop_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY) / 2
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY) / 2
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Inverted Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(255 - texr, 255 - texg, 255 - texb)
    Next X
    Next Y
End Sub
'[INVERT (BOTTOM HALF)]
Private Sub invertBottom_Click()
    For Y = (frmD.PicBox.Height / Screen.TwipsPerPixelY) / 2 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Inverted Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(255 - texr, 255 - texg, 255 - texb)
    Next X
    Next Y
End Sub
'[INVERT (LEFT HALF)]
Private Sub invertLeft_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX) / 2
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Grey Color'
        SetPixel frmD.PicBox.hdc, X, Y, RGB(255 - texr, 255 - texg, 255 - texb)
    Next X
    Next Y
End Sub
'[INVERT (RIGHT HALF)]
Private Sub invertRight_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = (frmD.PicBox.Height / Screen.TwipsPerPixelX) / 2 To (frmD.PicBox.Height / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Grey Color'
        SetPixel frmD.PicBox.hdc, X, Y, RGB(255 - texr, 255 - texg, 255 - texb)
    Next X
    Next Y
End Sub

'######################################### ADD NOISE #########################################

'[ADD BRIGHT NOISE]
Private Sub AddBrightNoise_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Bright Noise
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr + Int(Rnd * 100), texg + Int(Rnd * 100), texb + Int(Rnd * 100))
    Next X
    Next Y
End Sub
'[ADD DARK NOISE]
Private Sub AddDarkNoise_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Dark Noise Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr - Int(Rnd * 100), texg - Int(Rnd * 100), texb - Int(Rnd * 100))
    Next X
    Next Y
End Sub

'####################################### DODGE / BURN ########################################

'[DODGE]
Private Sub Dodge_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Dodged Colors
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr * 2, texg * 2, texb * 2)
    Next X
    Next Y
End Sub
'[DODGE (TOP HALF)]
Private Sub dodgeTop_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY) / 2
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY) / 2
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Grey Color'
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr * 2, texg * 2, texb * 2)
    Next X
    Next Y
End Sub
'[DODGE (BOTTOM HALF)]
Private Sub dodgeBottom_Click()
    For Y = (frmD.PicBox.Height / Screen.TwipsPerPixelY) / 2 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Dodged Colors
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr * 2, texg * 2, texb * 2)
    Next X
    Next Y
End Sub
'[DODGE (LEFT HALF)]
Private Sub dodgeLeft_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX) / 2
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Dodged Colors
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr * 2, texg * 2, texb * 2)
    Next X
    Next Y
End Sub
'[DODGE (RIGHT HALF)]
Private Sub dodgeRight_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = (frmD.PicBox.Height / Screen.TwipsPerPixelX) / 2 To (frmD.PicBox.Height / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'The Grey Color = Average Of RGB In Pixle
        greycolor = (texr + texg + texb) / 3
        'Set Pixle With Dodged Colors
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr * 2, texg * 2, texb * 2)
    Next X
    Next Y
End Sub
'[BURN]
Private Sub Burn_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Dodged Colors
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr / 2, texg / 2, texb / 2)
    Next X
    Next Y
End Sub
'[BURN (TOP HALF)]
Private Sub burnTop_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY) / 2
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY) / 2
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Burnt Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr / 2, texg / 2, texb / 2)
    Next X
    Next Y
End Sub
'[BURN (BOTTOM HALF)]
Private Sub burnBottom_Click()
    For Y = (frmD.PicBox.Height / Screen.TwipsPerPixelY) / 2 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Burnt Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr / 2, texg / 2, texb / 2)
    Next X
    Next Y
End Sub
'[BURN (LEFT HALF)]
Private Sub burnLeft_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX) / 2
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Burnt Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr / 2, texg / 2, texb / 2)
    Next X
    Next Y
End Sub
'[BURN (RIGHT HALF)]
Private Sub burnRight_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = (frmD.PicBox.Height / Screen.TwipsPerPixelX) / 2 To (frmD.PicBox.Height / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Burnt Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr / 2, texg / 2, texb / 2)
    Next X
    Next Y
End Sub

'########################################### FADE ############################################

'[FADE BOTTOM]
Private Sub FadeBottom_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Y Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr + Y, texg + Y, texb + Y)
    Next X
    Next Y
End Sub
'[FADE RIGHT]
Private Sub FadeRight_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With X Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr + X, texg + X, texb + X)
    Next X
    Next Y
End Sub
'[FADE BOTTOM LEFT]
Private Sub FadeLeftUp_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Y-X Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr + Y - X, texg + Y - X, texb + Y - X)
    Next X
    Next Y
End Sub
'[FADE TOP RIGHT]
Private Sub FadeRightDown_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With X-Y Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr + X - Y, texg + X - Y, texb + X - Y)
    Next X
    Next Y
End Sub

'################################### REMOVE COLOR (RGB) ######################################

'[NO RED]
Private Sub NoRed_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Green And Blue Colors
        SetPixel frmD.PicBox.hdc, X, Y, RGB(0, texg, texb)
    Next X
    Next Y
End Sub
'[NO GREEN]
Private Sub NoGreen_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Red And Blue Colors
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr, 0, texb)
    Next X
    Next Y
End Sub
'[NO BLUE]
Private Sub NoBlue_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Red And Green Colors
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr, texg, 0)
    Next X
    Next Y
End Sub

'###################################### COLORIZE (RGB) #######################################

'[RED FILTER]
Private Sub RedFilter_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Red Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(texr, 0, 0)
    Next X
    Next Y
End Sub
'[GREEN FILTER]
Private Sub GreenFilter_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Green Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(0, texg, 0)
    Next X
    Next Y
End Sub
'[BLUE FILTER]
Private Sub BlueFilter_Click()
On Error Resume Next
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Blue Color
        SetPixel frmD.PicBox.hdc, X, Y, RGB(0, 0, texb)
    Next X
    Next Y
End Sub

'###################################### STRIPES ########################################

Private Sub addDotGrid_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY) Step 10
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX) Step 10
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Burnt Color
        SetPixel frmD.PicBox.hdc, X, Y, Color1.BackColor
    Next X
    Next Y
End Sub

Private Sub addHorizontal_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY) Step 10
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX)
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Burnt Color
        SetPixel frmD.PicBox.hdc, X, Y, Color1.BackColor
    Next X
    Next Y
End Sub

Private Sub addVertical_Click()
    For Y = 0 To (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Max = (frmD.PicBox.Height / Screen.TwipsPerPixelY)
    PG.Value = Y
        If PG.Value = PG.Max Then
            PG.Value = 0
        End If
    For X = 0 To (frmD.PicBox.Width / Screen.TwipsPerPixelX) Step 10
        'Pixle's Red Value
        texr = frmD.PicBox.Point(X, Y) And 255
        'Pixle's Green Value
        texg = (frmD.PicBox.Point(X, Y) And 65280) / 256
        'Pixle's Blue Value
        texb = (frmD.PicBox.Point(X, Y) And 16711680) / 65535
        'Set Pixle With Burnt Color
        SetPixel frmD.PicBox.hdc, X, Y, Color1.BackColor
    Next X
    Next Y
End Sub

'#################################### FAST GRAPHICS ####################################

Private Sub fastBrightNoise_Click()
On Error Resume Next
    GetObject frmD.PicBox.Image, Len(PicInfo), PicInfo
    BytesPerLine = (PicInfo.bmWidth * 3 + 3) And &HFFFFFFFC
    ReDim PicBits(1 To BytesPerLine * PicInfo.bmHeight * 3) As Byte
    GetBitmapBits frmD.PicBox.Image, UBound(PicBits), PicBits(1)
    For Cnt = 1 To UBound(PicBits)
        PicBits(Cnt) = PicBits(Cnt) + Int(Rnd * 100)
    Next Cnt
    SetBitmapBits frmD.PicBox.Image, UBound(PicBits), PicBits(1)
    frmD.PicBox.Refresh
End Sub
Private Sub fastDarkNoise_Click()
On Error Resume Next
    GetObject frmD.PicBox.Image, Len(PicInfo), PicInfo
    BytesPerLine = (PicInfo.bmWidth * 3 + 3) And &HFFFFFFFC
    ReDim PicBits(1 To BytesPerLine * PicInfo.bmHeight * 3) As Byte
    GetBitmapBits frmD.PicBox.Image, UBound(PicBits), PicBits(1)
    For Cnt = 1 To UBound(PicBits)
        PicBits(Cnt) = PicBits(Cnt) - Int(Rnd * 100)
    Next Cnt
    SetBitmapBits frmD.PicBox.Image, UBound(PicBits), PicBits(1)
    frmD.PicBox.Refresh
End Sub
Private Sub fastBurn_Click()
On Error Resume Next
    GetObject frmD.PicBox.Image, Len(PicInfo), PicInfo
    BytesPerLine = (PicInfo.bmWidth * 3 + 3) And &HFFFFFFFC
    ReDim PicBits(1 To BytesPerLine * PicInfo.bmHeight * 3) As Byte
    GetBitmapBits frmD.PicBox.Image, UBound(PicBits), PicBits(1)
    For Cnt = 1 To UBound(PicBits)
        PicBits(Cnt) = PicBits(Cnt) / 2
    Next Cnt
    SetBitmapBits frmD.PicBox.Image, UBound(PicBits), PicBits(1)
    frmD.PicBox.Refresh
End Sub
Private Sub fastDodge_Click()
On Error Resume Next
    GetObject frmD.PicBox.Image, Len(PicInfo), PicInfo
    BytesPerLine = (PicInfo.bmWidth * 3 + 3) And &HFFFFFFFC
    ReDim PicBits(1 To BytesPerLine * PicInfo.bmHeight * 3) As Byte
    GetBitmapBits frmD.PicBox.Image, UBound(PicBits), PicBits(1)
    For Cnt = 1 To UBound(PicBits)
        PicBits(Cnt) = PicBits(Cnt) * 2
    Next Cnt
    SetBitmapBits frmD.PicBox.Image, UBound(PicBits), PicBits(1)
    frmD.PicBox.Refresh
End Sub








'################################## LOAD / UNLOAD ############################################
'[LOAD]
Private Sub MDIForm_Load() 'MDIForm(frmMain)
a = 0
    'Get The Dimensions Of frmMain
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    LoadNewDoc
End Sub
'[UNLOAD]
Private Sub MDIForm_Unload(Cancel As Integer) 'MDIForm(frmMain)
    If Me.WindowState <> vbMinimized Then
        'Save The Settings Of The Window So It Will Open That Way Next Time
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub
'##################################### FILE ##################################################
'[NEW]
Private Sub mnuNew_Click()
    LoadNewDoc
End Sub
Private Sub LoadNewDoc()
a = a + 1
Set frmD = New frmDocument
    frmD.Caption = "Image " & a
    frmD.Show
End Sub
'[OPEN]
Private Sub mnuOpen_Click()
a = a + 1
Set frmD = New frmDocument
frmD.Caption = "Image " & a
frmD.Show
cd1.ShowOpen
frmD.PicBox.Picture = LoadPicture(cd1.FileName)
End Sub
'[SAVE]
Private Sub mnuSave_Click()
    'With cd1 (Common Dialog)
    With cd1
        'Set The Title Of The Save Window To "Choose a filename to save"
        .DialogTitle = "Choose a filename to save"
        'Set The Save Filter Of The Save Window To "24-bit Bitmap (*.bmp)|*.bmp"
        .Filter = "24-bit Bitmap (*.bmp)|*.bmp"
        'Set The Filters Index To 1
        .FilterIndex = 1
        'Set The File Name To "" (Blank)
        .FileName = ""
        'Show The Save Window
        .ShowSave
        
    'If The File Name Is "" (Blank) Then Exit Sub
    If .FileName = "" Then
        Exit Sub
    End If
        'Save Picture PicBox's Image as .FileName
        SavePicture frmDocument.PicBox.Image, .FileName
    'End With cd1 (Common Dialog)
    End With
End Sub
'[PRINT]
Private Sub mnuPrint_Click()
    'Common Dialog Show Print
    cd1.ShowPrinter
End Sub
'[EXIT]
Private Sub mnuExit_Click()
    'End (Exit App)
    End
End Sub

'################################### ARRANGEMENTS ############################################

'[ARRANGE ICONS]
Private Sub mnuArrangeIcons_Click()
    'Arrange The Icons On frmMain
    Me.Arrange vbArrangeIcons
End Sub

'[WINDOW ARRANGEMENTS]

'[HORIZONTAL]
Private Sub mnuHorizontal_Click()
    'Arrange The Windows On frmMain Horizontaly
    Me.Arrange vbTileHorizontal
End Sub
'[VERTICAL]
Private Sub mnuVertical_Click()
    'Arrange The Windows On frmMain Verticaly
    Me.Arrange vbTileVertical
End Sub
'[CASCADE]
Private Sub mnuCascade_Click()
    'Cascade The Windows On frmMain
    Me.Arrange vbCascade
End Sub

'################################### TOOLBAR CASES ###########################################

'[TOOLBAR 1]

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next

    'Select Case Button.Key (Tool Bar Button Keys)
    Select Case Button.Key

        '[NEW]
        Case "New"
            LoadNewDoc

        '[SAVE]
        Case "Save"
            mnuSave_Click

        '[OPEN]
        Case "Open"
            mnuOpen_Click

        '[PRINT]
        Case "Print"
            mnuPrint_Click

        '[CASCADE]
        Case "Cascade"
            mnuCascade_Click

        '[HORIZONTAL]
        Case "Tile Horizontal"
            mnuHorizontal_Click

        '[VERTICAL]
        Case "Tile Vertical"
            mnuVertical_Click
            
    'End Selection For Button.Key
    End Select
    
End Sub

'[TOOLBAR 2]

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next

    'Select Case Button.Key (Tool Bar Button Keys)
    Select Case Button.Key

        '[PENCIL]

        Case "Pencil"
            'Mouse Pointer Is The Default (Arrow)
            frmD.PicBox.MousePointer = 0
            'Draw Mode = Copy Pen
            PicBox.DrawMode = 13
            'Draw Width = 1
            PicBox.DrawWidth = 1
            'Tool Title = "Pencil"
            DrawType.Caption = "Pencil"
            'Hide Fill Style Options
            xFillStyle.Visible = False
            'Hide Brush Style Options
            BrushStyle.Visible = False
            'Show The Pen Size
            PenSize.Visible = True
            'Disable Pen Size (Draw Width) Options
            PenSize.Enabled = False
            'Hide Labels
            Label4.Visible = True
            Label5.Visible = False
            'Show The Frame
            Frame2.Visible = True
            
        '[PEN]

        Case "Pen"
            'Mouse Pointer Is The Default (Arrow)
            frmD.PicBox.MousePointer = 0
            'Tool Title = "Pen"
            DrawType.Caption = "Pen"
            'Hide Fill Style Options
            xFillStyle.Visible = False
            'Show The Pen Size
            PenSize.Visible = True
            'Enable Pen Size (Draw Width) Options
            PenSize.Enabled = True
            'Show Labels
            Label4.Visible = True
            Label5.Visible = True
            'Show Brush Style Options
            BrushStyle.Visible = True
            'Show The Frame
            Frame2.Visible = True

        '[MARKER]

        Case "Marker"
            'Mouse Pointer Is The Default (Arrow)
            frmD.PicBox.MousePointer = 0
            'Tool Title = "Marker"
            DrawType.Caption = "Marker"
            'Hide Fill Style Options
            FillStyle.Visible = False
            'Show The Pen Size
            PenSize.Visible = True
            'Enable Pen Size (Draw Width) Options
            PenSize.Enabled = True
            'Show Labels
            Label4.Visible = True
            Label5.Visible = True
            'Hide Fill Style Options
            xFillStyle.Visible = False
            'Show Brush Style Options
            BrushStyle.Visible = True
            'Show The Frame
            Frame2.Visible = True
        '[ERASER]

        Case "Eraser"
            'Mouse Pointer Is The Default (Arrow)
            frmD.PicBox.MousePointer = 0
            'Tool Title = "Eraser"
            DrawType.Caption = "Eraser"
            'Hide Fill Style Options
            xFillStyle.Visible = False
            'Show The Pen Size
            PenSize.Visible = True
            'Enable Pen Size (Draw Width) Options
            PenSize.Enabled = True
            'Show / Hide Labels
            Label4.Visible = True
            Label5.Visible = False
            'Hide Brush Style Options
            BrushStyle.Visible = False
            'Show The Frame
            Frame2.Visible = True

        '[COLOR PICKER]

        Case "Color Picker"
            'Mouse Pointer Is A Cross
            frmD.PicBox.MousePointer = 2
            'Tool Title = "Color Picker"
            DrawType.Caption = "Color Picker"
            'Hide The Frame
            Frame2.Visible = False

        '[FLOOD FILL]

        Case "Flood Fill"
            'Mouse Pointer Is The Default (Arrow)
            frmD.PicBox.MousePointer = 0
            'Tool Title = "Flood Fill"
            DrawType.Caption = "Flood Fill"
            'Show Fill Style Options
            xFillStyle.Visible = True
            'Hide The Pen Size
            PenSize.Visible = False
            'Show / Hide Labels
            Label4.Visible = False
            Label5.Visible = True
            'Hide Brush Style Options
            BrushStyle.Visible = False
            'Show The Frame
            Frame2.Visible = True
            
    'End Selection For Button.Key
    End Select
    
End Sub

'#################################### COLOR MIXER ############################################

'[TEXT TO RED VALUE]
Private Sub txtred_Change()

    'Custom Color (ShapeColor.BackColor) = RGB(red.Value + green.Value + blue.Value)
    ShapeColor.BackColor = RGB(red.Value, green.Value, blue.Value)

    On Error Resume Next

    'If Greater Then 255
    If txtred.Text > 255 Then
        MsgBox "Red Value Must Be Between 0 And 255"
        txtred.Text = 255
    End If

    'If Less Then 0
    If txtred.Text < 0 Then
        MsgBox "Red Value Must Be Between 0 And 255"
        txtred.Text = 0
    End If

    'Value = Text
    red.Value = txtred.Text

End Sub
'[TEXT TO GREEN VALUE]
Private Sub txtgreen_Change()

    'Custom Color (ShapeColor.BackColor) = RGB(red.Value + green.Value + blue.Value)
    ShapeColor.BackColor = RGB(red.Value, green.Value, blue.Value)

    On Error Resume Next

    'If Greater Then 255
    If txtgreen.Text > 255 Then
        MsgBox "Green Value Must Be Between 0 And 255"
        txtgreen.Text = 255
    End If

    'If Less Then 0
    If txtgreen.Text < 0 Then
        MsgBox "Green Value Must Be Between 0 And 255"
        txtgreen.Text = 0
    End If

    'Value = Text
    green.Value = txtgreen.Text

End Sub
'[TEXT TO BLUE VALUE]
Private Sub txtblue_Change()

    'Custom Color (ShapeColor.BackColor) = RGB(red.Value + green.Value + blue.Value)
    ShapeColor.BackColor = RGB(red.Value, green.Value, blue.Value)

    On Error Resume Next

    'If Greater Then 255
    If txtblue.Text > 255 Then
        MsgBox "Blue Value Must Be Between 0 And 255"
        txtblue.Text = 255
    End If

    'If Less Then 0
    If txtblue.Text < 0 Then
        MsgBox "Blue Value Must Be Between 0 And 255"
        txtblue.Text = 0
    End If

    'Value = Text
    blue.Value = txtblue.Text

End Sub
'[RED VALUE TO TEXT]
Private Sub red_Change()
    'Custom Color (ShapeColor.BackColor) = RGB(red.Value + green.Value + blue.Value)
    ShapeColor.BackColor = RGB(red.Value, green.Value, blue.Value)
If m1.Value = True Then Color1.BackColor = ShapeColor.BackColor
If m2.Value = True Then Color2.BackColor = ShapeColor.BackColor
    'Text = Value
    txtred.Text = red.Value

End Sub
Private Sub red_Scroll()
    'Custom Color (ShapeColor.BackColor) = RGB(red.Value + green.Value + blue.Value)
    ShapeColor.BackColor = RGB(red.Value, green.Value, blue.Value)
If m1.Value = True Then Color1.BackColor = ShapeColor.BackColor
If m2.Value = True Then Color2.BackColor = ShapeColor.BackColor
    'Text = Value
    txtred.Text = red.Value
End Sub
'[GREEN VALUE TO TEXT]
Private Sub green_Change()
    'Custom Color (ShapeColor.BackColor) = RGB(red.Value + green.Value + blue.Value)
    ShapeColor.BackColor = RGB(red.Value, green.Value, blue.Value)
If m1.Value = True Then Color1.BackColor = ShapeColor.BackColor
If m2.Value = True Then Color2.BackColor = ShapeColor.BackColor
    'Text = Value
    txtgreen.Text = green.Value

End Sub
Private Sub green_Scroll()
    'Custom Color (ShapeColor.BackColor) = RGB(red.Value + green.Value + blue.Value)
    ShapeColor.BackColor = RGB(red.Value, green.Value, blue.Value)
If m1.Value = True Then Color1.BackColor = ShapeColor.BackColor
If m2.Value = True Then Color2.BackColor = ShapeColor.BackColor
    'Text = Value
    txtgreen.Text = green.Value
End Sub
'[BLUE VALUE TO TEXT]
Private Sub blue_Change()
    'Custom Color (ShapeColor.BackColor) = RGB(red.Value + green.Value + blue.Value)
    ShapeColor.BackColor = RGB(red.Value, green.Value, blue.Value)
If m1.Value = True Then Color1.BackColor = ShapeColor.BackColor
If m2.Value = True Then Color2.BackColor = ShapeColor.BackColor
    'Text = Value
    txtblue.Text = blue.Value

End Sub
Private Sub blue_Scroll()
    'Custom Color (ShapeColor.BackColor) = RGB(red.Value + green.Value + blue.Value)
    ShapeColor.BackColor = RGB(red.Value, green.Value, blue.Value)
If m1.Value = True Then Color1.BackColor = ShapeColor.BackColor
If m2.Value = True Then Color2.BackColor = ShapeColor.BackColor
    'Text = Value
    txtblue.Text = blue.Value
End Sub
'[GRADIENT MOUSE DOWN]
Private Sub gradient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo 1
    
    'B = Where You Clicked On The Gradient
    'If You Click The Gradient It Gets The Point Where You Clicked Then
    'It Sets The Color You Clicked On The Gradient, To Color1
    b = Gradient.Point(X, Y)
    If Button = 1 And m1.Value = True Then
    Color1.BackColor = b
    End If
    If Button = 2 And m1.Value = True Then
    m1.Value = False
    m2.Value = True
    End If
    
    If Button = 1 And m2.Value = True Then
    m1.Value = True
    m2.Value = False
    End If
    If Button = 2 And m2.Value = True Then
    Color2.BackColor = b
    End If
    '[COLOR TO RGB]
        'Get Red Value
        texr = Gradient.Point(X, Y) And 255
        'Get Green Value
        texg = (Gradient.Point(X, Y) And 65280) / 256
        'Get Blue Value
        texb = (Gradient.Point(X, Y) And 16711680) / 65535
    red.Value = texr
    green.Value = texg
    blue.Value = texb
    
1 End Sub
'[GRADIENT MOUSE MOVE]
Private Sub gradient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo 1

    'B = Where You Clicked On The Gradient
    b = Gradient.Point(X, Y)

    'If Mouse1 and Mouse2 Are Pressed Then Exit Sub
    If Button <> 1 And Button <> 2 Then
        Exit Sub
    End If

    'If Mouse1 Then Set The Color You Clicked On The Gradient, To Color1
    If Button = 1 And m1.Value = True Then
        frmMain.Color1.BackColor = b
        'Get Red Value
        texr = Gradient.Point(X, Y) And 255
        'Get Green Value
        texg = (Gradient.Point(X, Y) And 65280) / 256
        'Get Blue Value
        texb = (Gradient.Point(X, Y) And 16711680) / 65535
        red.Value = texr
        green.Value = texg
        blue.Value = texb
    End If
        'If Mouse2 Then Set The Color You Clicked On The Gradient, To Color2
    If Button = 2 And m1.Value = True Then
        frmMain.Color1.BackColor = b
        'Get Red Value
        texr = Gradient.Point(X, Y) And 255
        'Get Green Value
        texg = (Gradient.Point(X, Y) And 65280) / 256
        'Get Blue Value
        texb = (Gradient.Point(X, Y) And 16711680) / 65535
        red.Value = texr
        green.Value = texg
        blue.Value = texb
    End If
    
    'If Mouse1 Then Set The Color You Clicked On The Gradient, To Color2
    If Button = 1 And m2.Value = True Then
        frmMain.Color2.BackColor = b
        'Get Red Value
        texr = Gradient.Point(X, Y) And 255
        'Get Green Value
        texg = (Gradient.Point(X, Y) And 65280) / 256
        'Get Blue Value
        texb = (Gradient.Point(X, Y) And 16711680) / 65535
        red.Value = texr
        green.Value = texg
        blue.Value = texb
    End If
        'If Mouse2 Then Set The Color You Clicked On The Gradient, To Color2
    If Button = 2 And m2.Value = True Then
        frmMain.Color2.BackColor = b
        'Get Red Value
        texr = Gradient.Point(X, Y) And 255
        'Get Green Value
        texg = (Gradient.Point(X, Y) And 65280) / 256
        'Get Blue Value
        texb = (Gradient.Point(X, Y) And 16711680) / 65535
        red.Value = texr
        green.Value = texg
        blue.Value = texb
    End If
1 End Sub
Private Sub Swatches_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo 1
    
    'B = Where You Clicked On The Swatch
    'If You Click The Swatch It Gets The Point Where You Clicked Then
    'It Sets The Color You Clicked On The Swatch, To Color1
    
    b = Swatches.Point(X, Y)
    
    If Button = 1 And m1.Value = True Then
    Color1.BackColor = b
    End If
    If Button = 2 And m1.Value = True Then
    m1.Value = False
    m2.Value = True
    End If
    
    If Button = 1 And m2.Value = True Then
    m1.Value = True
    m2.Value = False
    End If
    If Button = 2 And m2.Value = True Then
    Color2.BackColor = b
    End If
    '[COLOR TO RGB]
        'Get Red Value
        texr = Swatches.Point(X, Y) And 255
        'Get Green Value
        texg = (Swatches.Point(X, Y) And 65280) / 256
        'Get Blue Value
        texb = (Swatches.Point(X, Y) And 16711680) / 65535
    red.Value = texr
    green.Value = texg
    blue.Value = texb
1 End Sub
Private Sub Swatches_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo 1

    'B = Where You Clicked On The Gradient
    b = Swatches.Point(X, Y)

    'If Mouse1 and Mouse2 Are Pressed Then Exit Sub
    If Button <> 1 And Button <> 2 Then
        Exit Sub
    End If

    'If Mouse1 Then Set The Color You Clicked On The Gradient, To Color1
    If Button = 1 And m1.Value = True Then
        frmMain.Color1.BackColor = b
        'Get Red Value
        texr = Swatches.Point(X, Y) And 255
        'Get Green Value
        texg = (Swatches.Point(X, Y) And 65280) / 256
        'Get Blue Value
        texb = (Swatches.Point(X, Y) And 16711680) / 65535
        red.Value = texr
        green.Value = texg
        blue.Value = texb
    End If
        'If Mouse2 Then Set The Color You Clicked On The Gradient, To Color2
    If Button = 2 And m1.Value = True Then
        frmMain.Color1.BackColor = b
        'Get Red Value
        texr = Swatches.Point(X, Y) And 255
        'Get Green Value
        texg = (Swatches.Point(X, Y) And 65280) / 256
        'Get Blue Value
        texb = (Swatches.Point(X, Y) And 16711680) / 65535
        red.Value = texr
        green.Value = texg
        blue.Value = texb
    End If
    
    'If Mouse1 Then Set The Color You Clicked On The Gradient, To Color2
    If Button = 1 And m2.Value = True Then
        frmMain.Color2.BackColor = b
        'Get Red Value
        texr = Swatches.Point(X, Y) And 255
        'Get Green Value
        texg = (Swatches.Point(X, Y) And 65280) / 256
        'Get Blue Value
        texb = (Swatches.Point(X, Y) And 16711680) / 65535
        red.Value = texr
        green.Value = texg
        blue.Value = texb
    End If
        'If Mouse2 Then Set The Color You Clicked On The Gradient, To Color2
    If Button = 2 And m2.Value = True Then
        frmMain.Color2.BackColor = b
        'Get Red Value
        texr = Swatches.Point(X, Y) And 255
        'Get Green Value
        texg = (Swatches.Point(X, Y) And 65280) / 256
        'Get Blue Value
        texb = (Swatches.Point(X, Y) And 16711680) / 65535
        red.Value = texr
        green.Value = texg
        blue.Value = texb
    End If
1 End Sub

