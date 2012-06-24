VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIMain 
   Caption         =   "Flow Chart"
   ClientHeight    =   7755
   ClientLeft      =   1350
   ClientTop       =   780
   ClientWidth     =   9765
   Icon            =   "Flow_main.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7755
   ScaleWidth      =   9765
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ilsColours 
      Left            =   2880
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":05A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":0702
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":0862
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":09C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":0B22
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":0C82
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbColour 
      Height          =   420
      Left            =   480
      TabIndex        =   14
      Top             =   6840
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilsColours"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Text Colour"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Line Colour"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Line Thickness"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Line Style"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Background Colour"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Background Style"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Format Object"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Timer timMsg 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2640
      Top             =   2400
   End
   Begin VB.PictureBox picSelection 
      Align           =   3  'Align Left
      Height          =   6915
      Left            =   0
      ScaleHeight     =   6855
      ScaleWidth      =   195
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "Layer Selection Tool"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox picGrid 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   120
      MouseIcon       =   "Flow_main.frx":10D6
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picExport 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      MouseIcon       =   "Flow_main.frx":1228
      ScaleHeight     =   375
      ScaleWidth      =   615
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2760
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FilterIndex     =   1
      FontName        =   "Fixedsys"
      FromPage        =   1
      Max             =   1
      Min             =   1
      MaxFileSize     =   512
      ToPage          =   1
   End
   Begin MSComctlLib.ImageList ilsTools 
      Left            =   4560
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":137A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":17CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":1C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":2076
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":24CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":291E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":2D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":31C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":361A
            Key             =   "small_icon"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":3A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":3EC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":4316
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":476A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.VScrollBar vsbBar 
      Height          =   2895
      LargeChange     =   1500
      Left            =   5280
      SmallChange     =   100
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   255
   End
   Begin VB.HScrollBar hsbBar 
      Height          =   255
      LargeChange     =   1500
      Left            =   720
      SmallChange     =   100
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3480
      Width           =   3615
   End
   Begin VB.PictureBox picView 
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   960
      ScaleHeight     =   2475
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   840
      Width           =   3375
      Begin VB.PictureBox picBay 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         FillColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   0
         MouseIcon       =   "Flow_main.frx":4BBE
         ScaleHeight     =   1455
         ScaleWidth      =   1335
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
         Begin VB.TextBox txtEdit 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin FlowProg.ctlBox ctlBox 
            Height          =   120
            Index           =   0
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   120
            _ExtentX        =   212
            _ExtentY        =   212
            BackColor       =   -2147483635
            ForeColor       =   -2147483626
         End
         Begin VB.Line linSize 
            BorderStyle     =   3  'Dot
            Visible         =   0   'False
            X1              =   720
            X2              =   1560
            Y1              =   120
            Y2              =   600
         End
         Begin VB.Shape shpMove 
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            FillColor       =   &H80000005&
            Height          =   615
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
      End
   End
   Begin MSComctlLib.ImageList ilsPic 
      Left            =   4560
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":4D10
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":4E24
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":4F38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":504C
            Key             =   "left"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":5160
            Key             =   "centre"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Flow_main.frx":5274
            Key             =   "right"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbDraw 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilsTools"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Pointer"
            Object.ToolTipText     =   "Pointer"
            ImageIndex      =   12
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   13
            Style           =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkAutoAdd 
         Caption         =   "Auto-add"
         Height          =   255
         Left            =   4200
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   60
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar tlbFont 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   420
      Visible         =   0   'False
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ilsPic"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   3500
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Description     =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Description     =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Description     =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Description     =   "Left Align"
            Object.ToolTipText     =   "Left Align"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Centre"
            Description     =   "Centre Align"
            Object.ToolTipText     =   "Centre Align"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Description     =   "Right Align"
            Object.ToolTipText     =   "Right Align"
            ImageIndex      =   6
            Style           =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ComboBox cboSize 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   0
         Width           =   735
      End
      Begin VB.ComboBox cboFont 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Flow_main.frx":5388
         Left            =   0
         List            =   "Flow_main.frx":538F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Font"
         Top             =   0
         Width           =   2535
      End
   End
   Begin MSComctlLib.StatusBar sbrBottom 
      Height          =   300
      Left            =   480
      TabIndex        =   8
      Top             =   7320
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Coordinates"
            TextSave        =   "Coordinates"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "Size"
            TextSave        =   "Size"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   917
            MinWidth        =   917
            Text            =   "Name"
            TextSave        =   "Name"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save &As..."
         Index           =   1
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "Print Set&up..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "Print Pr&eview"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditWhat 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditEdit 
         Caption         =   "Toggle &Edit"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditCancel 
         Caption         =   "Cancel Edit"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditToFront 
         Caption         =   "&Bring to front"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEditToBack 
         Caption         =   "&Send to back"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuText 
      Caption         =   "&Text"
      Begin VB.Menu mnuTextLeft 
         Caption         =   "&Left"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuTextCentre 
         Caption         =   "&Centre"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuTextRight 
         Caption         =   "&Right"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuTextSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextBold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuTextItalic 
         Caption         =   "&Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuTextUnderline 
         Caption         =   "&Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuTextSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextFont 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuTextSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextSpell 
         Caption         =   "&Spell Check"
         Begin VB.Menu mnuTextSpellThis 
            Caption         =   "&This Item"
         End
         Begin VB.Menu mnuTextSpellAll 
            Caption         =   "&All Items"
         End
      End
   End
   Begin VB.Menu mnuShape 
      Caption         =   "&Shape"
      Begin VB.Menu mnuShapeItem 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mnuShapeButton 
         Caption         =   "Button..."
      End
      Begin VB.Menu mnuShapeArea 
         Caption         =   "Area..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Too&ls"
      Begin VB.Menu mnuToolsAutoFitText 
         Caption         =   "&Auto-fit text"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuToolsDuplicate 
         Caption         =   "&Duplicate"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuToolsSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsSplit 
         Caption         =   "Split Line"
      End
      Begin VB.Menu mnuToolsRefreshSingle 
         Caption         =   "R&efresh"
      End
      Begin VB.Menu mnuToolsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsPic 
         Caption         =   "&Picture..."
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "For&mat"
      Begin VB.Menu mnuFormatObjProp 
         Caption         =   "&Object..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuFormatSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatAlign 
         Caption         =   "&Align"
         Begin VB.Menu mnuFormatAlignOne 
            Caption         =   "To &Grid"
         End
         Begin VB.Menu mnuFormatAlignAll 
            Caption         =   "&All to Grid"
         End
      End
      Begin VB.Menu mnuFormatGrid 
         Caption         =   "&Grid"
         Begin VB.Menu mnuFormatGridSnap 
            Caption         =   "&Snap to Grid"
            Shortcut        =   ^G
         End
         Begin VB.Menu mnuFormatGridShow 
            Caption         =   "Sho&w Persistent Grid"
            Shortcut        =   ^H
         End
      End
      Begin VB.Menu mnuFormatSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatExportWord 
         Caption         =   "Export word"
      End
      Begin VB.Menu mnuFormatExport 
         Caption         =   "&Export as bmp..."
      End
      Begin VB.Menu mnuFormatOpt 
         Caption         =   "O&ptions..."
      End
      Begin VB.Menu mnuFormatZoom 
         Caption         =   "&Zoom..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowFontBar 
         Caption         =   "F&ont Bar"
      End
      Begin VB.Menu mnuWindowSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowLayer 
         Caption         =   "&Layer Tool"
      End
      Begin VB.Menu mnuWindowMultis 
         Caption         =   "Grou&p Box"
      End
      Begin VB.Menu mnuWindowShift 
         Caption         =   "Shif&t Box"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpPrintInfo 
         Caption         =   "&Printer..."
      End
      Begin VB.Menu mnuHelpOInfo 
         Caption         =   "&Window..."
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Flow Chart..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPText 
         Caption         =   "Text"
         Visible         =   0   'False
         Begin VB.Menu mnuPTextAdd 
            Caption         =   "&Add Field..."
         End
         Begin VB.Menu mnuPTextRemove 
            Caption         =   "&Remove Field"
         End
      End
      Begin VB.Menu mnuPPic 
         Caption         =   "Picture"
         Visible         =   0   'False
         Begin VB.Menu mnuPPicRatio 
            Caption         =   "R&atio..."
         End
         Begin VB.Menu mnuPPicReload 
            Caption         =   "&Reload Picture"
         End
         Begin VB.Menu mnuPPicReset 
            Caption         =   "R&eset Picture"
         End
         Begin VB.Menu mnuPPicChange 
            Caption         =   "&Change Picture"
         End
         Begin VB.Menu mnuPPicSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPPicRelative 
            Caption         =   "Relativ&e"
         End
         Begin VB.Menu mnuPPicAbs 
            Caption         =   "A&bsolute"
         End
      End
      Begin VB.Menu mnuPLine 
         Caption         =   "Line"
         Visible         =   0   'False
         Begin VB.Menu mnuPLineNormal 
            Caption         =   "&Normal"
         End
         Begin VB.Menu mnuPLineEng 
            Caption         =   "&Engineering"
         End
         Begin VB.Menu mnuPLineSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPLineSolid 
            Caption         =   "&Solid"
         End
         Begin VB.Menu mnuPLineHollow 
            Caption         =   "&Hollow"
         End
         Begin VB.Menu mnuPLineSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPLineSize 
            Caption         =   "50%"
            Index           =   0
         End
         Begin VB.Menu mnuPLineSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPLineNoArrow 
            Caption         =   "N&o Arrow"
         End
         Begin VB.Menu mnuPLineFlowArrow 
            Caption         =   "&Flow Arrow"
         End
         Begin VB.Menu mnuPLinePointingArrow 
            Caption         =   "&Pointing Arrow"
         End
      End
      Begin VB.Menu mnuPButton 
         Caption         =   "Button"
         Visible         =   0   'False
         Begin VB.Menu mnuPButtonRunMacro 
            Caption         =   "&Run Macro"
         End
         Begin VB.Menu mnuPButtonEditMacro 
            Caption         =   "&Edit Macro"
         End
         Begin VB.Menu mnuPButtonDelMacro 
            Caption         =   "&Delete Macro"
         End
      End
      Begin VB.Menu mnuPGroup 
         Caption         =   "Group"
         Visible         =   0   'False
         Begin VB.Menu mnuPGroupGroup 
            Caption         =   "&Group"
         End
         Begin VB.Menu mnuPGroupUngroup 
            Caption         =   "&Ungroup"
         End
         Begin VB.Menu mnuPGroupSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPGroupSelGroup 
            Caption         =   "&Select Group"
         End
         Begin VB.Menu mnuPGroupSelOne 
            Caption         =   "Select &One"
         End
         Begin VB.Menu mnuPGroupUnselGroup 
            Caption         =   "U&nselect Group"
         End
      End
      Begin VB.Menu mnuPSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPopupCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuPopupCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPopupPasteHere 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupBringFront 
         Caption         =   "&Bring to Font"
      End
      Begin VB.Menu mnuPopupSendBack 
         Caption         =   "&Send to Back"
      End
      Begin VB.Menu mnuPopupSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuPCancel 
         Caption         =   "Cance&l"
      End
      Begin VB.Menu mnuPSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPProp 
         Caption         =   "Prop&erties..."
      End
   End
End
Attribute VB_Name = "frmIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Richard Fung.  August 12, 2000.
'Updated Aug. 2002
'F1  (reserved) Windows
'F2     edit on/off mnuEditEdit
'F3     edit cancel txtEdit
'F4  (reserved) Windows - like Ctrl+F4,Alt+F4
'F5     force redraw Form_KeyUp
'F6     auto-fit text
'F7
'F8     Format Object
'F9
'F10 (reserved) Windows
'F11    edit off txtEdit
'F12    edit off txtEdit
'above Tab        next item

'Objects
Public mFlowChart   As FlowChart
Public mClipboard   As PClipboard
Private mRegistry   As PRegistry
Private frmMultis   As frmIMulti

'for interface
Private mMousePos   As Rect
Private mSize       As Rect
Private mlngAddType As FAddType
Private mblnHitItem As Boolean

'Scale values
Private msngScale   As Single
Private msngGrid    As Single 'for alignment

'for SetView()
Private mlngDefaultTextFlags As Long

'for undo
Private WithEvents mUndo As PUndo
Attribute mUndo.VB_VarHelpID = -1

'for drawing
Private mblnToolRedraw As Boolean
Private mblnSelected   As Boolean
Private mblnMouseDown  As Boolean
Private mblnMovedSized As Boolean
Private mblnMouseMovedSensitive As Boolean 'half-grid point
Private mblnMouseMovedGrid      As Boolean 'grid point

'for point tracing (mouse clicking on multiple items)
Private mcolSelections As Collection

'for toolbar
Private mblnFontBarBusy As Boolean

'variables manipulated by other variables
Private m__objSelected As FlowItem
Private m__blnSizing As Boolean
Private m__blnMoving As Boolean
Private m__blnAdding As Boolean
Private m__blnSelecting As Boolean

Private Enum EnumHandle
    conNW = 1
    conN
    conNE
    conE
    conSE
    conS
    conSW
    conW
End Enum

Private Enum EnumTest
    conMiss = 0
    conHit = 1
    conHitDefault = 2
End Enum

Private Const conHMin = 1
Private Const conHMax = 8 'count of EnumHandle
Private Const conPixel = 15
Private Const conTxtEditDisableUndo = "d-u" 'disable undo
Private Const conDigits = 2 'round to 2 digits

'SetScrollBars()
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXHSCROLL = 21
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CYVSCROLL = 20

'Aligns a FlowItem to the grid.
Public Sub Align(Item As FlowItem, Optional NoChange As Boolean)
    If Not Item Is Nothing Then
        'reposition
        Item.P.Left = CLng(Item.P.Left / msngGrid) * msngGrid
        Item.P.Top = CLng(Item.P.Top / msngGrid) * msngGrid
        Item.P.Width = CLng(Item.P.Width / msngGrid) * msngGrid
        Item.P.Height = CLng(Item.P.Height / msngGrid) * msngGrid
        'straighten lines
        If IsObjLine(Item) Then
            If Abs(Item.P.Width) < Abs(Item.P.Height) Then   'vertical line
                If Abs(Item.P.Width) < msngGrid * msngScale Then
                    Item.P.Width = 0
                End If
            ElseIf Abs(Item.P.Height) < Abs(Item.P.Width) Then   'horizontal line
                If Abs(Item.P.Height) < msngGrid * msngScale Then
                    Item.P.Height = 0
                End If
            End If
        End If
        If NoChange = False Then
            mFlowChart.Changed = True 'change flag
        End If
    End If
End Sub



'Turns off auto-redraw.
Private Sub AutoRedrawOff()
    'turn off autoredraw & it will paint
    If picBay.AutoRedraw = False Then
        picBay_Paint
    Else
        'must clear background graphics or it will stay
        'back there
        picBay.Cls
        picBay.AutoRedraw = False 'paint will be called
    End If
End Sub

'Turns on auto-redraw, usually for sizing or moving
'objects on picBay.
Private Sub AutoRedrawOn()
    'This remarked items protect swappings,
    'which occurs when auto-redraw is turned on in
    'a large picture box.
    If mFlowChart.ZoomPercent < 150 Then 'memory protection code to prevent HD reswapping
        picBay.AutoRedraw = True
    End If
End Sub


'Calculates the visible view area.  Used by picBayPaint()
Private Function CalcView() As Rect
    CalcView.X1 = picBay.ScaleX(-picBay.Left, picView.ScaleMode, picBay.ScaleMode)
    CalcView.Y1 = picBay.ScaleY(-picBay.Top, picView.ScaleMode, picBay.ScaleMode)
    CalcView.X2 = CalcView.X1 + picBay.ScaleX(picView.ScaleWidth, picView.ScaleMode, picBay.ScaleMode)
    CalcView.Y2 = CalcView.Y1 + picBay.ScaleY(picView.ScaleHeight, picView.ScaleMode, picBay.ScaleMode)
End Function



Private Sub ConnectY(Item As FlowItem, OldP As Properties, NewP As Properties, TopChanged As Boolean, BottomChanged As Boolean)
'December 27, 2002
'Purpose:   Since text auto-fit moves the bottom off of connected lins,
'   this function reconnects those lines
'   so gaps in between are not visible.
    Dim objEach As FlowItem
    Dim lngX1   As Long
    Dim lngX2   As Long
    Dim lngCorrection As Long
    
    lngCorrection = picBay.ScaleY(2, vbPixels, picBay.ScaleMode)
    
    lngX1 = NewP.Left
    lngX2 = NewP.Left + NewP.Width
    
    If TopChanged Then
        Dim lngTopNew As Long
        Dim lngTopOld As Long
        Dim lngTopOld1 As Long
        Dim lngTopOld2 As Long
        
        lngTopOld = OldP.Top
        lngTopOld1 = lngTopOld - lngCorrection
        lngTopOld2 = lngTopOld + lngCorrection
        lngTopNew = NewP.Top
        
        For Each objEach In mFlowChart
            If Not objEach Is Item And IsBetween(lngTopOld1, objEach.P.Top + objEach.P.Height, lngTopOld2) _
            And IsBetween(lngX1, objEach.P.Left + objEach.P.Width, lngX2) Then
                ' objEach.Top is fixed property
                If (Not IsObjLine(Item) And lngTopNew - objEach.P.Top >= 0) Or IsObjLine(Item) Then
                    objEach.P.Height = lngTopNew - objEach.P.Top
                End If
            End If
        Next objEach
    End If
    
    If BottomChanged Then
        Dim lngBottomNew As Long
        Dim lngBottomOld As Long
        Dim lngBottomOld1 As Long
        Dim lngBottomOld2 As Long
        Dim lngBottomEach As Long
        
        lngBottomOld = OldP.Top + OldP.Height
        lngBottomOld1 = lngBottomOld - lngCorrection
        lngBottomOld2 = lngBottomOld + lngCorrection
        lngBottomNew = NewP.Top + NewP.Height
        
        For Each objEach In mFlowChart
            If Not objEach Is Item And IsBetween(lngBottomOld1, objEach.P.Top, lngBottomOld2) _
            And IsBetween(lngX1, objEach.P.Left, lngX2) Then
                'objEach.Bottom is the fixed property
                lngBottomEach = objEach.P.Top + objEach.P.Height
                If (Not IsObjLine(Item) And lngBottomEach - lngBottomNew >= 0) Or IsObjLine(Item) Then
                    objEach.P.Top = lngBottomNew
                    objEach.P.Height = lngBottomEach - lngBottomNew
                End If
            End If
        Next objEach
    End If
End Sub

Private Sub EditBoxOnPic()
'Called from EditBoxOn.
'    Dim strFilename As String
    Dim objFile     As FileDlg
    Dim objPicture  As FPicture
    Dim strPathPart As String
    Dim strPathFLC  As String 'of flow chart
    
    On Error Resume Next
    'Change Pictures
    Set objFile = New FileDlg
    objFile.Title = "Open Picture"
    objFile.Initialize hwnd, "", "Picture Files|*.bmp;*.jpg;*.gif;*.wmf;*.emf;.ico;*.cur", mSelected.P.Text, True, False, cdlOFNHideReadOnly
    
    Select Case objFile.ShowOpen
    'After new picture file selected
    Case CDERR_OK
        Set objPicture = mSelected
        strPathPart = objFile.RetPath
        strPathFLC = mFlowChart.GetPath()
        
        If LCase$(strPathPart) = LCase$(strPathFLC) Then
            objPicture.LoadPicture2 objFile.RetFile, objFile.RetPath
        Else
            objPicture.LoadPicture2 objFile.FileName, ""
        End If
        
        objPicture.SetDefaultSize picBay
        Set objPicture = Nothing
        mFlowChart.Changed = True
    Case CDERR_CANCELLED
    Case Else
        MsgBox "Error showing Open Picture dialog box.", vbExclamation, "Open Picture"
    End Select
    
    'remove if without picture on adding
    If mblnAdding And Len(mSelected.P.Text) = 0 Then 'no picture specified
        mFlowChart.RemoveObj mSelected
        Set mSelected = Nothing
        SelectionChanged
    End If
    
    'redraw screen
    picBayPaint
    RedrawHandles mSelected 'redraw handles
End Sub

Public Function Grid(ByVal Num As Long) As Long
'If the grid is turned on, any value is rounded to the nearest grid value.
    If frmMain.mnuFormatGridSnap.Checked Then
        Grid = CLng(Num / msngGrid) * msngGrid
    Else
        Grid = Num
    End If
End Function

Public Function IsDrawing() As Boolean
    IsDrawing = mblnMoving Or mblnSizing Or mblnAdding Or mblnSelecting
End Function

Private Function IsMultipleSelected() As Boolean
    If gblnWindowMultis Then IsMultipleSelected = frmMultis.GetCount()
End Function


Public Sub LoadFont()
    Dim lngFontNo As Long

    'add font sizes
    cboSize.AddItem ""
    For lngFontNo = 6 To 12
        cboSize.AddItem lngFontNo
    Next lngFontNo
    
    For lngFontNo = 14 To 20 Step 2
        cboSize.AddItem lngFontNo
    Next lngFontNo
    
    For lngFontNo = 24 To 40 Step 4
        cboSize.AddItem lngFontNo
    Next lngFontNo
    
    'add font face names
    For lngFontNo = 0 To Screen.FontCount - 1
        cboFont.AddItem Screen.Fonts(lngFontNo)
        DoEvents
    Next lngFontNo
End Sub

Public Sub Prompt(Optional ByVal Str As String)
    If Len(Str) = 0 Then Str = " "
    sbrBottom.Panels(1).Text = Str
    
End Sub

Public Sub RedrawGrid2()
    Dim sngWidth    As Single
    Dim sngHeight   As Single
    
    On Error GoTo Handler
    
    sngWidth = picView.ScaleWidth / mFlowChart.ZoomPercent * 100
    sngHeight = picView.ScaleHeight / mFlowChart.ZoomPercent * 100
    
    'make the picGrid a bit smaller
    If sngWidth > picBay.Width Then sngWidth = picBay.Width / mFlowChart.ZoomPercent * 100
    If sngHeight > picBay.Height Then sngHeight = picBay.Height / mFlowChart.ZoomPercent * 100
    
    'grid is drawn directly onto paper
    Dim sngX As Single, sngY As Single
    

    For sngY = -Grid(picBay.Top) To -Grid(picBay.Top) + sngHeight Step picBay.ScaleY(msngGrid, picBay.ScaleMode, vbTwips) / mFlowChart.ZoomPercent * 100
        For sngX = -Grid(picBay.Left) To -Grid(picBay.Left) + sngWidth Step picBay.ScaleY(msngGrid, picBay.ScaleMode, vbTwips) / mFlowChart.ZoomPercent * 100
            picBay.PSet (sngX, sngY), vb3DShadow
        Next sngX
    Next sngY

   
    Exit Sub
Handler:
    Exit Sub
End Sub

Public Sub RedrawSelection(ByVal IsSelected As Boolean)
    mFlowChart.DrawSelected picBay, IsSelected
End Sub

Private Sub SelectionChanged()
'remove multiple selection if item clicked on is not in multiple selection
    If gblnWindowMultis Then
        If frmMultis.GetIndex(mSelected) = 0 Then
            'not a selected item clicked
            Unload frmMultis
        End If
    End If
End Sub



Private Sub ctlBoxSizing(Index As Integer)
    'make the changes to mSize, with grid applied if applicable
    Select Case Index
        Case conNW
            mSize.X1 = Grid(mSelected.P.Left + mMousePos.X2 - mMousePos.X1)
            mSize.Y1 = Grid(mSelected.P.Top + mMousePos.Y2 - mMousePos.Y1)
        Case conN
            mSize.Y1 = Grid(mSelected.P.Top + mMousePos.Y2 - mMousePos.Y1)
        Case conNE
            mSize.X2 = Grid(mSelected.P.Left + mSelected.P.Width + mMousePos.X2 - mMousePos.X1)
            mSize.Y1 = Grid(mSelected.P.Top + mMousePos.Y2 - mMousePos.Y1)
        Case conE
            mSize.X2 = Grid(mSelected.P.Left + mSelected.P.Width + mMousePos.X2 - mMousePos.X1)
        Case conSE
            mSize.X2 = Grid(mSelected.P.Left + mSelected.P.Width + mMousePos.X2 - mMousePos.X1)
            mSize.Y2 = Grid(mSelected.P.Top + mSelected.P.Height + mMousePos.Y2 - mMousePos.Y1)
        Case conS
            mSize.Y2 = Grid(mSelected.P.Top + mSelected.P.Height + mMousePos.Y2 - mMousePos.Y1)
        Case conSW
            mSize.X1 = Grid(mSelected.P.Left + mMousePos.X2 - mMousePos.X1)
            mSize.Y2 = Grid(mSelected.P.Top + mSelected.P.Height + mMousePos.Y2 - mMousePos.Y1)
        Case conW
            mSize.X1 = Grid(mSelected.P.Left + mMousePos.X2 - mMousePos.X1)
    End Select
'   'keep on the page
'    If mSize.x1 < 0 Then mSize.x1 = 0
'    If mSize.x2 < 0 Then mSize.x2 = 0
'      If mSize.y1 < 0 Then mSize.y1 = 0
'      If mSize.y2 < 0 Then mSize.y2 = 0
'    If mSize.x1 > picBay.ScaleWidth Then mSize.x1 = picBay.ScaleWidth
'    If mSize.x2 > picBay.ScaleWidth Then mSize.x2 = picBay.ScaleWidth
'      If mSize.y1 > picBay.ScaleHeight Then mSize.y1 = picBay.ScaleHeight
'      If mSize.y2 > picBay.ScaleHeight Then mSize.y2 = picBay.ScaleHeight

    'snap the line (to horizontal or vertical) in any mode
    If IsSelLine Then
        If Index = conNW Then 'snap NW
            If Abs(mSize.Y2 - mSize.Y1) < Abs(mSize.X2 - mSize.X1) Then 'vertical line
                If Abs(mSize.Y1 - mSize.Y2) < msngGrid / 2 Then
                    mSize.Y1 = mSize.Y2
                End If
            ElseIf Abs(mSize.Y2 - mSize.Y1) > Abs(mSize.X2 - mSize.X1) Then 'horizontal line
                If Abs(mSize.X1 - mSize.X2) < msngGrid / 2 Then
                    mSize.X1 = mSize.X2
                End If
            End If
        ElseIf Index = conSE Then 'snap SE
            If Abs(mSize.Y1 - mSize.Y2) < Abs(mSize.X1 - mSize.X2) Then 'vertical line
                If Abs(mSize.Y2 - mSize.Y1) < msngGrid / 2 Then
                    mSize.Y2 = mSize.Y1
                End If
            ElseIf Abs(mSize.Y1 - mSize.Y2) > Abs(mSize.X1 - mSize.X2) Then 'horizontal line
                If Abs(mSize.X2 - mSize.X1) < msngGrid / 2 Then
                    mSize.X2 = mSize.X1
                End If
            End If
        End If
    End If
End Sub

Private Sub ctlBoxRectToShape(Shape As Shape, ByRef Rect As Rect)
    Dim sngLeft As Single, sngTop As Single
    Dim sngWidth As Single, sngHeight As Single

    #If DebugMode Then
        Debug.Print "ctlBoxRectToShape()"
    #End If
    
    'keep on the page
    If Rect.X1 < 0 Then Rect.X1 = 0
    If Rect.X2 < 0 Then Rect.X2 = 0
      If Rect.Y1 < 0 Then Rect.Y1 = 0
      If Rect.Y2 < 0 Then Rect.Y2 = 0
    If Rect.X1 > picBay.ScaleWidth Then Rect.X1 = picBay.ScaleWidth
    If Rect.X2 > picBay.ScaleWidth Then Rect.X2 = picBay.ScaleWidth
      If Rect.Y1 > picBay.ScaleHeight Then Rect.Y1 = picBay.ScaleHeight
      If Rect.Y2 > picBay.ScaleHeight Then Rect.Y2 = picBay.ScaleHeight

    'reverses the negative values
    'or transfers x,y coordinates
    'so that shapes do not have
    ' -width or -height
    'Also: add in one pixel to width & height to correct indented pixel
    If Rect.X2 - Rect.X1 < 0 Then
        sngWidth = Rect.X1 - Rect.X2 ' + conPixel
        sngLeft = Rect.X1 + (Rect.X2 - Rect.X1)
    Else
        sngWidth = Rect.X2 - Rect.X1 '+ conPixel
        sngLeft = Rect.X1
    End If
    If Rect.Y2 - Rect.Y1 < 0 Then
        sngHeight = Rect.Y1 - Rect.Y2 '+ conPixel
        sngTop = Rect.Y1 + (Rect.Y2 - Rect.Y1)
    Else
        sngHeight = Rect.Y2 - Rect.Y1 '+ conPixel
        sngTop = Rect.Y1
    End If
    
    Shape.Move sngLeft, sngTop, sngWidth, sngHeight
    
    UpdateBottomCoord sngLeft, sngTop, sngWidth, sngHeight
End Sub






Private Sub EditBoxOff()
    If IsEditing Then
        mFlowChart.Changed = True
        'do not undo on first add
        If txtEdit.Tag <> conTxtEditDisableUndo Then
            mUndo.Add mSelected, conUndoTextChange
        End If
        txtEdit.Tag = "" 'clear any special commands
        'update the text to the object
        RedrawHandles mSelected
        mSelected.P.Text = txtEdit
        txtEdit.Visible = False
        DrawText picBay, mSelected, mFlowChart
    End If
End Sub

Private Sub EditBoxOn()
    Dim sngAdjust   As Single 'adjustment for display
    Dim sngAdjusted As Single
    
    sngAdjust = 15 'left,top increment
    sngAdjusted = 15  'width,height decrement
    
    On Error GoTo Handler
    If IsSelected And Not IsBusy Then 'not editing
        If Not mSelected.P.CanEdit Then
            Exit Sub 'leave if cannot edit
        End If
        If IsSelLine Then
            'use textbox for lines
            mSelected.P.Text = InputBox("Enter in the text below for the line.", "Edit Text", mSelected.P.Text) 'GetText(mSelected))
            mFlowChart.Changed = True
            Redraw
        ElseIf TypeOf mSelected Is FPicture Then
            EditBoxOnPic
        Else
            RedrawHandles mSelected, False
            
            txtEdit.Move mSelected.P.Left + mSelected.TextLeftMrg + sngAdjust, _
                         mSelected.P.Top + mSelected.TextTopMrg + sngAdjust, _
                         mSelected.P.Width - mSelected.TextLeftMrg - mSelected.TextRightMrg - sngAdjusted, _
                         mSelected.P.Height - mSelected.TextBottomMrg - mSelected.TextTopMrg - sngAdjusted
            'prepares edit box
            SetEditFont mSelected
            txtEdit.Text = mSelected.P.Text 'GetText(mSelected)
            'if adding for the first time, disable undo
            txtEdit.Tag = IIf(mblnAdding, conTxtEditDisableUndo, "")
'            'changes alignment, but cannot be changed
'            'in a text box at run-time
'            Select Case GetTextFlags(mSelected.P.Text)
'                Case DT_LEFT: txtEdit.Alignment = vbLeftJustify
'                Case DT_CENTER: txtEdit.Alignment = vbCenter
'                Case DT_RIGHT: txtEdit.Alignment = vbRightJustify
'                Case Else: txtEdit.Alignment = vbLeftJustify
'            End Select
            txtEdit.Visible = True
            mFlowChart.Changed = True 'note that edit changes text
            txtEdit.SetFocus
        End If
    End If
    Exit Sub
Handler:
    'in case txtEdit does not size properly,
    'the text area isn't visible so the user
    'will use the InputBox.
    RedrawHandles mSelected
    mSelected.P.Text = InputBox("Enter in the text below.", "Edit Text", mSelected.P.Text) 'GetText(mSelected))
    mFlowChart.Changed = True
End Sub





Private Function IsBusy() As Boolean
'Busy: editing, moving, sizing, selecting
    IsBusy = IsEditing Or mblnMoving Or mblnSizing Or mblnSelecting
End Function

Private Function IsEditing() As Boolean
    IsEditing = txtEdit.Visible
End Function

Public Sub OpenFile(ByVal FileName As String)
    Dim lngErr As Long
    Const conDiffers = "differs from the last time you edited."
    
    On Error Resume Next 'at top
    
    MousePointer = vbHourglass
    
    Set mFlowChart = New FlowChart
    ResetForDraw
    
    picBay.Visible = False
    mClipboard.ClearLink
    ResetForNewOpen
    
    lngErr = mFlowChart.Load(FileName, picBay)
    Select Case lngErr
    Case conFail
        MsgBox "Problem opening the flow chart.  The flow chart may be corrupted or the file is too new for this version of the program.", vbExclamation
        mFlowChart.FileName = "" 'erase filename
    Case 0
        'no problem
    Case Else
        MsgBox "Failed opening file because of the following problem: " & Error(lngErr), vbExclamation
        mFlowChart.FileName = "" 'erase filename
    End Select
    
    'parts of ResetForDraw()
    Recaption
'    mFlowChart.Changed = False
'    frmLog.txtLog = "" 'log beginning
    
    'check version information
    If mFlowChart.Version > conCurrentVersion Then
        MsgBox "This file was made in a newer version of Flow Chart.  This file may be read incorrectly because of major differences.", vbExclamation
        mFlowChart.FileName = "" 'remove filename
    End If
    
    'prepare messages
    sbrBottom.SimpleText = ""
    
    'set other data
    If mFlowChart.Version >= 4 Then
        With mFlowChart
            'check last printer used
            If .Header1PDevName <> Printer.DeviceName And Len(.Header1PDevName) Then
                sbrBottom.SimpleText = "Printer has changed from " & .Header1PDevName & " to " & Printer.DeviceName & ".  "
            End If
            'correct orientation of the page
            If .Orientation <> Printer.Orientation Then
                If .Orientation <> 0 Then
                    Printer.Orientation = .Orientation
                End If
            End If
            If .PaperSize <> Printer.PaperSize Then
                sbrBottom.SimpleText = sbrBottom.SimpleText & "Paper size is different.  "
            End If
            If (.PScaleHeight > Printer.ScaleHeight Or _
              .PScaleWidth > Printer.ScaleWidth) And Not .Customized Then
                sbrBottom.SimpleText = sbrBottom.SimpleText & "The drawing area is smaller.  "
'            ElseIf (.PScaleHeight <> Printer.ScaleHeight Or _
'              .PScaleWidth <> Printer.ScaleWidth) And Not .Customized Then
            End If

            On Error Resume Next
            'check scroll X and Y values
            If .ScrollX <= conScrollParts And .ScrollY <= conScrollParts Then
                hsbBar.Value = .ScrollX / conScrollParts * hsbBar.Max
                vsbBar.Value = .ScrollY / conScrollParts * vsbBar.Max
            End If
        End With
    End If
    
    If Len(sbrBottom.SimpleText) Then
        sbrBottom.Style = sbrSimple
        timMsg.Enabled = True
    End If
    
    If mFlowChart.Version >= 2 Then 'V2
        SetView mFlowChart.ZoomPercent, mFlowChart.FontName, mFlowChart.FontSize
    Else 'V1 uses default font name and size
        SetView FontName:=conFontName, FontSize:=conFontSize
    End If
    
    Set mcolSelections = New Collection
    CheckFileBounds mFlowChart, picBay.ScaleWidth, picBay.ScaleHeight, mcolSelections
    If mcolSelections.Count Then picSelection_Paint
    
    'show layers box?
    Dim lngMin As Long, lngMax As Long
    mFlowChart.Layers.Requery
    mFlowChart.Layers.MinMaxLayer lngMin, lngMax
    If lngMax > 1 Then
        mnuWindowLayer_Click
    End If
    mFlowChart.Layers.CloseQuery
    
    'if nothing in the log, unload the window
'    If Len(frmLog.txtLog) = 0 Then
'        Unload frmLog
'    End If
    picBay.Visible = True
    MousePointer = vbDefault
End Sub

Public Function IsInMode() As Boolean
'In a mode, which could be: editing, moving, sizing, adding
'Similar to IsBusy() except it also includes adding
    IsInMode = IsEditing Or mblnMoving Or mblnSizing Or mblnAdding Or mblnSelecting
End Function

Public Function IsSelected() As Boolean
    IsSelected = Not mSelected Is Nothing
End Function


Private Function IsWithoutText() As Boolean
    'pictures do not have text
    IsWithoutText = (mSelected.Number = conAddPicture) '(TypeOf mSelected Is FCircle) Or (TypeOf mSelected Is FPicture)
End Function


Public Sub MainShow()
'Called on startup.

    Form_Resize
    
    Set frmLayer = New frmLayer
    Set frmMultis = New frmIMulti
    Set frmCPick = New frmCPick
    Set frmStyleBack = New frmStyleBack
    Set frmStyleLine = New frmStyleLine
    Set frmStyleThickness = New frmStyleThickness
    
    Show
    Refresh
End Sub

Private Property Let mblnAdding(ByVal pAdding As Boolean)
    Static blnBusy As Boolean
    
    If pAdding Then
        Prompt "Click and drag to add object."
    Else
        Prompt
    End If
    
    If blnBusy = False Then
        blnBusy = True
        m__blnAdding = pAdding
        picBay.MousePointer = IIf(pAdding, vbCrosshair, vbDefault)
        If pAdding = False Then
            tlbDrawButtonsOff
        End If
        blnBusy = False
    End If
End Property


Private Property Get mblnAdding() As Boolean
    mblnAdding = m__blnAdding
End Property

Private Property Let mblnMoving(ByVal pMoving As Boolean)
    If pMoving Then
        Prompt "Drag the object to another location."
    Else
        Prompt
    End If
    
    m__blnMoving = pMoving
    picBay.MousePointer = vbDefault 'ResetForDraw to default
    'other code for moving must change icon
End Property

Private Property Get mblnMoving() As Boolean
    mblnMoving = m__blnMoving
End Property

Private Property Let mblnSizing(ByVal pSizing As Boolean)
    If Not mblnAdding Then
        If pSizing Then
            Prompt "Drag the handle to another location."
        Else
            Prompt
        End If
    End If
    
    m__blnSizing = pSizing
    picBay.MousePointer = IIf(pSizing, vbCrosshair, vbDefault)
End Property

Private Property Get mblnSizing() As Boolean
    mblnSizing = m__blnSizing
End Property

Private Property Get mblnSelecting() As Boolean
    mblnSelecting = m__blnSelecting
End Property

Private Property Let mblnSelecting(ByVal pSelecting As Boolean)
    If pSelecting Then
        Prompt "Drag the rubber rectangle around objects."
    Else
        Prompt
    End If
    
    m__blnSelecting = pSelecting
    If Not pSelecting Then
        picBay.MousePointer = vbDefault
    End If
End Property


'Called from picBay::MouseUp
Private Function picBayAdd(ByVal CopyFrom As FlowItem) As Boolean
    Dim objSample As FlowItem
    Dim blnNoEdit As Boolean
    Dim sngW        As Single, sngH As Single
    Dim lngColour   As Long
    Dim i           As Long

    sngW = 3000 'in twips; default sizes for shapes
    sngH = 1000
    
    If CopyFrom Is Nothing Then
        MsgBox "DEBUG: How is this possible?  Adding from nothing.", vbQuestion
        Exit Function
    End If

    Select Case mlngAddType
        Case conAddRect: Set objSample = New FRect
        Case conAddInOut: Set objSample = New FInOut
        Case conAddDecision: Set objSample = New FDecision
        Case conAddTerminator: Set objSample = New FTerminator
        Case conAddCircle: Set objSample = New FCircle: sngW = sngH   'width same as height
        'BEGIN these will not be automatically edited
        Case conAddLine: Set objSample = New FLine: blnNoEdit = True
        Case conAddMidArrowLine: Set objSample = New FMidArrowLine: blnNoEdit = True
        Case conAddEndArrowLine: Set objSample = New FArrowLine: blnNoEdit = True
        'END
        Case conAddText: Set objSample = New FText
        Case conAddPicture: Set objSample = New FPicture
        Case conAddEllipse: Set objSample = New FEllipse
        Case Else: Set objSample = New FlowItem
    End Select

    'check to see if too small
    If (Not IsObjLine(objSample) And Abs(CopyFrom.P.Width) < msngGrid And Abs(CopyFrom.P.Height) < msngGrid) Or _
    (IsObjLine(objSample) And Abs(CopyFrom.P.Width) < msngGrid / 2 And Abs(CopyFrom.P.Height) < msngGrid / 2) Then
        If objSample.Number <> conAddPicture Then
            Prompt "Select another point to size object."
            
            picBayAdd = False 'return error
            Exit Function
        Else
            CopyFrom.P.Width = sngW
            CopyFrom.P.Height = sngH
            Prompt "Select a picture file."
        End If
    End If
    
    
    'set up its properties
    'must do this first because default location will be set up
    CopyProperties mFlowChart.Layers.DefaultShape(mFlowChart.DefaultLayer).P, objSample.P
    
    If chkAutoAdd = vbUnchecked Then
        tlbDraw.Buttons(mlngAddType).Value = tbrUnpressed 'de-select selected button
    
        'add the item
        Set mSelected = mFlowChart.AddParam(objSample, CopyFrom.P.Left, CopyFrom.P.Top, CopyFrom.P.Width, CopyFrom.P.Height, "")
        
        AutoRedrawOff 'turn this off
        
        If Not blnNoEdit And (mRegistry.AutoEditText Or mSelected.Number = conAddPicture) Then
            'Let edit code run
            EditBoxOn
        End If
        
        'Change handle colour
        Randomize Timer
        lngColour = QBColor(2 * (Int(Rnd * 3) + 1))
        For i = conHMin To conHMax
            ctlBox(i).BackColor = lngColour ' &H8000& 'Qbcolor(2)
        Next i
        
        mblnAdding = False 'turn off flag
        
        UpdateFont
    Else
        mFlowChart.AddParam objSample, CopyFrom.P.Left, CopyFrom.P.Top, CopyFrom.P.Width, CopyFrom.P.Height, ""
        objSample.Draw picBay, mFlowChart
        
        If Not blnNoEdit And objSample.Number = conAddPicture Then
            Set mSelected = objSample
            EditBoxOnPic
            Set mSelected = Nothing
        End If
        
        mblnAdding = True
    End If
   
    mFlowChart.Changed = True 'note changes
    
    picBayAdd = True
End Function

Private Sub picBayMoving()
    Dim sngLeft As Single, sngTop As Single
    
    'moves the shape that correlates to the selected item
    'called by picBay_MouseDown, Over, etc.
    'And keeps item in page bounds
    #If DebugMode Then
        Debug.Print "picBayMoving()"
    #End If
    
    sngLeft = Grid(mSelected.P.Left + mMousePos.X2 - mMousePos.X1)
    sngTop = Grid(mSelected.P.Top + mMousePos.Y2 - mMousePos.Y1)
    If sngLeft < 0 Then sngLeft = 0
    If sngLeft + mSelected.P.Width < 0 Then sngLeft = -mSelected.P.Width
     If sngTop < 0 Then sngTop = 0
     If sngTop + mSelected.P.Height < 0 Then sngTop = -mSelected.P.Height
     
    If sngLeft > picBay.ScaleWidth Then sngLeft = picBay.ScaleWidth
    If sngLeft + mSelected.P.Width > picBay.ScaleWidth Then sngLeft = picBay.ScaleWidth - mSelected.P.Width
     
     If sngTop > picBay.ScaleHeight Then sngTop = picBay.ScaleHeight
     If sngTop + mSelected.P.Height > picBay.ScaleHeight Then sngTop = picBay.ScaleHeight - mSelected.P.Height
    shpMove.Left = sngLeft
    shpMove.Top = sngTop
    
    UpdateBottomCoord sngLeft, sngTop, -1, -1
End Sub

Private Sub picBayPaint(Optional All As Boolean)
    Dim objItem     As FlowItem
    Dim sngMrg      As Single
    Dim tView       As Rect

    
    'properties
    SetDrawProps picBay, Nothing, mFlowChart
    
    
    'grid
    If mnuFormatGridShow.Checked And (mFlowChart.CustomBack = -1 Or Not mFlowChart.Customized) Then
        On Error GoTo Skip 'sometimes, memory problem raises an error
        Set objItem = New FlowItem
        With objItem.P
            .Left = -picBay.ScaleX(picBay.Left, picView.ScaleMode, picBay.ScaleMode)
            .Top = -picBay.ScaleY(picBay.Top, picView.ScaleMode, picBay.ScaleMode)
        End With
        Align objItem, True
        picBay.PaintPicture picGrid.Image, objItem.P.Left, objItem.P.Top

Skip:
        Err.Clear
        Set objItem = Nothing
    Else
        RedrawGrid2
    End If
    
    'margins
    sngMrg = picBay.ScaleX(mRegistry.Margin, vbInches, picBay.ScaleMode) * mFlowChart.Percentage
    picBay.Line (sngMrg, sngMrg)-Step(picBay.ScaleWidth - sngMrg * 2, picBay.ScaleHeight - sngMrg * 2), vb3DFace, B
  
    'calculates visible area
    tView = CalcView()
    
    mFlowChart.DrawFile picBay, False, False, tView.X1, tView.Y1, tView.X2, tView.Y2
End Sub

Private Sub picBaySelect(Box As Rect)
'Selects items in a given Box dimension.
    Dim colSelect As Collection
    Dim objItem   As FlowItem
    
    Set colSelect = New Collection
    For Each objItem In mFlowChart
        If objItem.P.Enabled Then
            With objItem.P
                If .Left >= Box.X1 And .Top >= Box.Y1 And _
                    .Left + .Width < Box.X2 And .Top + .Height < Box.Y2 Then
                        colSelect.Add objItem
                End If
            End With
        End If
    Next objItem
    
    If colSelect.Count = 1 Then
        Unload frmMultis
        Set mSelected = colSelect(1) 'select the first item
    ElseIf colSelect.Count > 0 Then
        mnuWindowMultis_Click
        frmMultis.Clear
        For Each objItem In colSelect
            frmMultis.AddItem2 objItem
        Next objItem
        Set mSelected = colSelect(1) 'select the first item
        frmMultis.Update
    End If
    
    Set colSelect = Nothing
End Sub

Private Sub Recaption()
'Updates the caption.  Should be called on New, Open, Save.
    Dim strCaption  As String
    Dim intTrim     As Long
    
    If Len(mFlowChart.FileName) Then
        strCaption = String(255, vbNull)
        intTrim = GetFileTitle(mFlowChart.FileName & vbNullChar, strCaption, 255)
        strCaption = Left$(strCaption, InStr(1, strCaption, vbNullChar, vbBinaryCompare) - 1)
        Caption = strCaption & " - Flow Chart"
    Else
        Caption = "Flow Chart"
    End If
End Sub

Private Sub ResetForNewOpen() 'for new, open
    mlngDefaultTextFlags = 0 'ResetForDraw to left alignment
    mlngAddType = 0 'nothing to be added
    hsbBar.Value = 0 'set back at the
    vsbBar.Value = 0 'beginning corner
    sbrBottom.SimpleText = ""
    sbrBottom.Style = sbrNormal
    Set mUndo = New PUndo 'clear undo
    Call mUndo_UndoItemChanged(Nothing) 'and corresponding menu item
     If gblnWindowLayer Then Unload frmLayer
End Sub
Private Sub RedrawGrid()
    Dim hDesktop    As Long
    Dim tRect       As apiRECT
    Dim sngWidth    As Single
    Dim sngHeight   As Single
    
    On Error GoTo Handler
    
    'desktop dimensions retrieved by win API code
    'Screen.Width & .Height are not adjusted when
    'screen mode changes in run-time
    hDesktop = GetDesktopWindow()
    GetWindowRect hDesktop, tRect
    
    sngWidth = ScaleX(tRect.Right - tRect.Left, vbPixels, vbTwips)
    sngHeight = ScaleY(tRect.Bottom - tRect.Top, vbPixels, vbTwips)
    
    'make the picGrid a bit smaller
    If sngWidth > picBay.Width Then sngWidth = picBay.Width
    If sngHeight > picBay.Height Then sngHeight = picBay.Height
    'grid is drawn into another picture box
    'that is the size of the screen
    If mnuFormatGridShow.Checked Then
        Dim sngX As Single, sngY As Single
        
        Screen.MousePointer = vbHourglass
        picGrid.AutoRedraw = True
        picGrid.Cls
        picGrid.Move 0, 0, sngWidth, sngHeight ' Screen.Width, Screen.Height 'max ever possible
        'picGrid.Move 0, 0, Screen.Width, Screen.Height 'max ever possible
        For sngY = 0 To picGrid.Height Step picBay.ScaleY(msngGrid, picBay.ScaleMode, vbTwips)
            For sngX = 0 To picGrid.Width Step picBay.ScaleX(msngGrid, picBay.ScaleMode, vbTwips)
                picGrid.PSet (sngX, sngY), vb3DShadow
            Next sngX
        Next sngY
        'Debug.Print sngX, sngY
        Screen.MousePointer = vbDefault
    Else
Out:
'        Set picBay.Picture = Nothing
        picGrid.Cls
        picGrid.AutoRedraw = False
        picGrid.Move 0, 0, 15, 15
    End If
    
'    Redraw
    Exit Sub
Handler:
    Screen.MousePointer = vbDefault
    Resume Out
End Sub


Public Sub RedrawSingle(ByVal SingleObj As FlowItem)
'Redraws a single FlowItem on the page.
'The background is covered by a white box of background colour.
    picBay.Line (SingleObj.P.Left, SingleObj.P.Top)-Step(SingleObj.P.Width, SingleObj.P.Height), picBay.BackColor, BF
    SingleObj.Draw picBay, mFlowChart
    
    If IsEditing Then
        SetEditFont mSelected
        txtEdit.Refresh
    End If
End Sub

Private Function IsSelLine() As Boolean
'Uses IsObjLine()
    If IsSelected Then IsSelLine = IsObjLine(mSelected)
End Function



Public Property Set mSelected(ByVal pSel As FlowItem)
    EditBoxOff 'save text if opened
    Set m__objSelected = pSel
    'redraw handles
    RedrawHandles pSel
    
    'redraw selection bar
    Dim objItem As FlowItem
    Dim blnSel  As Boolean
    
    If Not mcolSelections Is Nothing Then
        For Each objItem In mcolSelections
            If objItem Is pSel Then
                blnSel = True
            End If
        Next objItem
        
        If Not blnSel Then
            Set mcolSelections = Nothing
        End If
        picSelection_Paint
    End If
    
    'redraw toolbar stuff
    UpdateBottom
    If tlbFont.Visible Then UpdateFont
    
    ToolboxUpdate
    
    'pop up menus
    If IsSelected Then
        mnuPText.Visible = (mSelected.Number = conAddText)
        mnuPPic.Visible = (mSelected.Number = conAddPicture)
        mnuPLine.Visible = IsObjLine(pSel)
        mnuPButton.Visible = (mSelected.Number = conAddButton)
        mnuPGroup.Visible = (mSelected.P.GroupNo <> 0) Or IsMultipleSelected
        mnuPSep1.Visible = mnuPText.Visible Or mnuPPic.Visible Or mnuPLine.Visible Or mnuPButton.Visible Or mnuPGroup.Visible
    Else
        mnuPText.Visible = False
        mnuPPic.Visible = False
        mnuPLine.Visible = False
        mnuPButton.Visible = False
        mnuPGroup.Visible = False
        mnuPSep1.Visible = False
        
        If gblnWindowBack Then Unload frmStyleBack
        If gblnWindowLine Then Unload frmStyleLine
        If gblnWindowThick Then Unload frmStyleThickness
        If gblnWindowPick Then Unload frmCPick
    End If
End Property

Public Property Get mSelected() As FlowItem
    Set mSelected = m__objSelected
End Property

Private Sub TabSelectItem(Back As Boolean)
'The following code of Feb. 15, 2001, during the
'long weekend of grade 11, is to select another
'item using the keyboard, and at the same time,
'make it in the viewable screen area.
'Updated 29 Jan 2002 to go forwards and backwards.

'    Dim objItem         As FlowItem 'item enumerated
'    Dim objItemUse      As FlowItem 'item to use
'    Dim blnGetNextItem  As Boolean 'flag to get next item
    Dim lngItem         As Long
    Dim tView           As Rect
    Dim sngToValue      As Single
    Dim lngCurrentItem  As Long

    If IsBusy Then Exit Sub
    
    MousePointer = vbArrowHourglass
    If IsSelected Then
        lngCurrentItem = mFlowChart.GetIndex(mSelected)
    
        If Back Then
            For lngItem = lngCurrentItem - 1 To 1 Step -1
                If mFlowChart(lngItem).P.Enabled Then
                    SetSelected mFlowChart(lngItem)
                    GoTo Finished
                End If
            Next lngItem

            For lngItem = mFlowChart.Count To lngCurrentItem + 1 Step -1
                If mFlowChart(lngItem).P.Enabled Then
                    SetSelected mFlowChart(lngItem)
                    GoTo Finished
                End If
            Next lngItem
        Else 'forward
            For lngItem = lngCurrentItem + 1 To mFlowChart.Count
                If mFlowChart(lngItem).P.Enabled Then
                    SetSelected mFlowChart(lngItem)
                    GoTo Finished
                End If
            Next lngItem
            
            For lngItem = 1 To lngCurrentItem - 1
                If mFlowChart(lngItem).P.Enabled Then
                    SetSelected mFlowChart(lngItem)
                    GoTo Finished
                End If
            Next lngItem
        End If
    ElseIf mFlowChart.Count > 0 Then
        'select the first enabled item
        For lngItem = 1 To mFlowChart.Count
            If mFlowChart(lngItem).P.Enabled Then
                SetSelected mFlowChart(lngItem)
                RedrawHandles mSelected
                GoTo Finished
            End If
        Next lngItem
    End If
    
Finished:
    
    'show the selected item in the viewable area
    If IsSelected Then
        'put into visible view
        'I had trouble with this code
        'because I reversed the horizontal
        'and vertical assignments to X and Y.
        tView = CalcView()
        If mSelected.P.Left < tView.X1 Or mSelected.P.Left + mSelected.P.Width > tView.X2 Then
            If mSelected.P.Left < mSelected.P.Left + mSelected.P.Width Then
                sngToValue = picBay.ScaleX(mSelected.P.Left, picBay.ScaleMode, conScrollScale)
            Else 'a line
                sngToValue = picBay.ScaleX(mSelected.P.Left + mSelected.P.Width, picBay.ScaleMode, conScrollScale)
            End If
            sngToValue = sngToValue - 8 'to see the handles, in pixels (UNDONE)
            If sngToValue > hsbBar.Max Then sngToValue = hsbBar.Max
            If sngToValue < hsbBar.Min Then sngToValue = hsbBar.Min
            hsbBar.Value = sngToValue
        End If
        If mSelected.P.Top < tView.Y1 Or mSelected.P.Top + mSelected.P.Height > tView.Y2 Then
            If mSelected.P.Top < mSelected.P.Top + mSelected.P.Height Then
                sngToValue = picBay.ScaleY(mSelected.P.Top, picBay.ScaleMode, conScrollScale)
            Else 'a line
                sngToValue = picBay.ScaleY(mSelected.P.Top + mSelected.P.Height, picBay.ScaleMode, conScrollScale)
            End If
            sngToValue = sngToValue - 120 '8 * 15 to see the handles, in pixels
            If sngToValue > vsbBar.Max Then sngToValue = vsbBar.Max
            If sngToValue < vsbBar.Min Then sngToValue = vsbBar.Min
            vsbBar.Value = sngToValue
        End If
    End If
    
    SelectionChanged
    
'        Set objItemUse = Nothing
'        Set objItem = Nothing
    MousePointer = vbDefault
End Sub

Private Sub SetEditFont(FlowItem As FlowItem)
    If Not FlowItem Is Nothing Then
        SetFontNmSz txtEdit, FlowItem.P.FontFace, FlowItem.P.TextSize, mFlowChart
'        txtEdit.Font.Name = GetFontName(FlowItem, mFlowChart)
'        txtEdit.Font.Size = FlowItem.P.TextSize * mFlowChart.Percentage 'assumes all objects have font size
        txtEdit.Font.Bold = FlowItem.P.TextBold 'GetTextFormatBold(FlowItem)
        txtEdit.Font.Italic = FlowItem.P.TextItalic 'GetTextFormatItalic(FlowItem)
        txtEdit.Font.Underline = FlowItem.P.TextUnderline
    Else
        SetFontNmSz txtEdit, "", 0, mFlowChart
'        txtEdit.Font.Name = mFlowChart.FontName  'picBay.Font.Name
'        txtEdit.Font.Size = mFlowChart.FontSize * mFlowChart.Percentage
        txtEdit.Font.Bold = False 'picBay.Font.Bold
        txtEdit.Font.Italic = False 'picBay.Font.Italic
        txtEdit.Font.Underline = False
    End If
End Sub

Public Sub SetSelected(ByVal Selected As FlowItem)
'with group selection enabled
    Set mSelected = Selected
    
    If IsSelected Then
        If mSelected.P.GroupNo <> 0 Then 'in a group
            mnuWindowMultis_Click 'show multiple selection window
            frmMultis.AddGroup mSelected.P.GroupNo 'add this group to multiple selection
        Else
            SelectionChanged
        End If
    End If
End Sub

Public Sub SetView(Optional ByVal Percentage As Long, Optional FontName As String, Optional ByVal FontSize As Currency)
'Necessary for both Viewer and main program.
'Sets up the necessary stuff for drawing to picBay
'and to the Printer.
    Dim hVal As Single, vVal As Single
        
    Form_Paint 'redraw normal icon
    picBay.Visible = False
    'adjust values
    If Percentage = 0 Then Percentage = mFlowChart.ZoomPercent
    If Len(FontName) = 0 Then FontName = mFlowChart.FontName
    If FontSize = 0 Then FontSize = mFlowChart.FontSize
    
    'note changed values (if it was changed)
    mFlowChart.ZoomPercent = Percentage
    mFlowChart.FontName = FontName
    mFlowChart.FontSize = FontSize
    
    'Adjust scale
    mFlowChart.ZoomPercent = Percentage
    msngScale = mFlowChart.Percentage
    
    'Grid spacing units
    Select Case mFlowChart.UnitScale
    Case vbCentimeters, vbMillimeters
        msngGrid = ScaleX(0.25, vbCentimeters, vbTwips)
    Case vbInches
        msngGrid = ScaleX(0.125, vbInches, vbTwips)
    Case vbPixels
        msngGrid = ScaleX(10, vbPixels, vbTwips)
    Case vbCharacters
        msngGrid = ScaleX(1, vbCharacters, vbTwips)
    End Select
    
    On Error Resume Next
    'Match font
'    picBay.Font.Name = FontName
'    picBay.Font.Size = FontSize
'    picBay.Font.Bold = False
'    picBay.Font.Italic = False
    
    If Err > 0 And mFlowChart.PrinterError = False Then
        MsgBox "Screen font and printer font cannot be matched.", vbExclamation
        Line (hsbBar.Left + hsbBar.Width, vsbBar.Top + vsbBar.Height)-(ScaleWidth, ScaleHeight), vbRed, BF
    End If
    Err.Clear
   
    'Size and scale to the flow chart file
    picBay.Move 0, 0, mFlowChart.PScaleWidth * msngScale, mFlowChart.PScaleHeight * msngScale
    picBay.Scale (0, 0)-(mFlowChart.PScaleWidth, mFlowChart.PScaleHeight)
    
    'scale font size
    picBay.Font.Size = FontSize * (Percentage / 100)
    
    'scale font size for printer
    If Not mFlowChart.PrinterError Then
        'test printer's font name and size to window's
        Printer.Font.Name = FontName
        Printer.Font.Size = FontSize
        
        picBay.Font.Size = Printer.Font.Size
        
        'tell user about any discrepancies
        If FontSize <> picBay.Font.Size Then
            'red colour to tell of font discrepancies
            Line (hsbBar.Left + hsbBar.Width, vsbBar.Top + vsbBar.Height)-(ScaleWidth, ScaleHeight), vbRed, BF
        End If
    End If
    
    'remember the percentage position of scroll bar
    On Error Resume Next
    If vsbBar.Max > 0 Then vVal = vsbBar.Value / vsbBar.Max
    If hsbBar.Max > 0 Then hVal = hsbBar.Value / hsbBar.Max
    'change view position to prevent errors
    'when the scroll bar max values are changed
    vsbBar.Value = 0
    hsbBar.Value = 0
    
    'update scroll bars by calling ::Resize
    UpdateScrollBars
    
    'set the percentage position of scroll bars
    vsbBar.Value = vVal * vsbBar.Max
    hsbBar.Value = hVal * hsbBar.Max
    
    'finish drawing
    If Not mFlowChart.CustomBack And picBay.BackColor <> vbWindowBackground Then
        picBay.BackColor = vbWindowBackground
    End If
    
    RedrawGrid
    picBay.Visible = True
End Sub
Private Sub SetScrollBars()
    Dim sngWidth As Single 'of scroll bar
    Dim sngHeight As Single
    
    'takes the image of the arrow as a measurement
    'for the width of the scroll bars
    sngWidth = GetSystemMetrics(SM_CXVSCROLL) '+ conScrollDif
    sngHeight = GetSystemMetrics(SM_CYHSCROLL) '+ conScrollDif
    sngWidth = ScaleX(sngWidth, vbPixels, ScaleMode)
    sngHeight = ScaleY(sngHeight, vbPixels, ScaleMode)
    vsbBar.Width = sngWidth
    hsbBar.Height = sngHeight
End Sub

Public Sub Shift(Optional Left As Boolean, Optional Right As Boolean, Optional Up As Boolean, Optional Down As Boolean)
    Dim objItem As FlowItem
    Dim sngShiftValue As Single
    
    sngShiftValue = msngGrid
    If Left Then
        For Each objItem In mFlowChart
            objItem.P.Left = objItem.P.Left - sngShiftValue
        Next objItem
    End If
    If Down Then
        'down
        For Each objItem In mFlowChart
            objItem.P.Top = objItem.P.Top + sngShiftValue
        Next objItem
    End If
    If Right Then
        For Each objItem In mFlowChart
            objItem.P.Left = objItem.P.Left + sngShiftValue
        Next objItem
    End If
    If Up Then
        For Each objItem In mFlowChart
            objItem.P.Top = objItem.P.Top - sngShiftValue
        Next objItem
    End If
    Redraw
End Sub

Private Sub SpellCheck(ByVal AllItems As Boolean)
    Dim objDoc      As Word.Document 'Object
    Dim blnVisible  As Boolean
    Dim strSpell    As String
    Dim strOrg      As String
    Dim fItem       As FlowItem
    Dim blnSingle   As Boolean

    On Error GoTo Handler
    
    'ask to save
    If mFlowChart.Changed And mFlowChart.AskToSave Then
        Select Case MsgBox("Do you want to save the file before spell check (just in case there is a problem)?  If you click No, this will be the last time you see this message.", vbQuestion Or vbYesNoCancel, "Save and Spell Check")
        Case vbYes
            mnuFileSave_Click 0
            If mFlowChart.Changed = True Then Exit Sub
        Case vbNo
            mFlowChart.AskToSave = False
        Case vbCancel
            Exit Sub
        End Select
    End If
    
    blnSingle = Not AllItems
    
    MousePointer = vbHourglass
    
    Set objDoc = CreateObject("Word.Document")
    blnVisible = objDoc.Application.Visible
    
    objDoc.Application.Visible = True
    
    If blnSingle Then
        Set fItem = mSelected
        GoSub SpellCheck
        Set fItem = Nothing
    Else
        For Each fItem In mFlowChart
            GoSub SpellCheck
        Next fItem
    End If
    
    AppActivate Caption
    
    If blnVisible Then
        objDoc.Close SaveChanges:=False
    Else
        objDoc.Application.Quit SaveChanges:=False
    End If
    
    Set objDoc = Nothing
    
    MousePointer = vbDefault
    
    Exit Sub
Handler:
    MsgBox "Failed to spell check using Microsoft Word 97 or greater.", vbExclamation, "Spell Check"
        
    Err.Clear
    On Error Resume Next
    If Not objDoc Is Nothing Then
        objDoc.Application.Visible = True 'in case Word loaded, leave it showing
    End If
    
    Set objDoc = Nothing
    MousePointer = vbDefault
    Exit Sub
SpellCheck:
    If Not TypeOf fItem Is FPicture Then
        strOrg = fItem.P.Text
    Else
        strOrg = "" 'picture filenames cannot be spell checked
    End If
    strSpell = strOrg
    If Len(strSpell) Then 'if there is text to spell check, then do so
        objDoc.Range.Text = strSpell
        AppActivate objDoc.Application.Caption 'keep reactivating MS Word
        objDoc.Range.CheckSpelling
        strSpell = objDoc.Range.Text
        strSpell = ConvCRtoCRLF(Left$(strSpell, Len(strSpell) - 1))
        If strSpell <> strOrg Then
            
            fItem.P.Text = strSpell
        End If
    End If
    Return
End Sub

Private Function t__func(TestItem As FlowItem, ByVal DefaultItem As FlowItem, ByVal X As Single, ByVal Y As Single) As EnumTest
'Called by the Test() function only!
    Dim tLoc As Rect 'changing rect
    Dim tToc As Rect 'non-changing rect
    
    Const conVr = 90 '* conScale 'variance 'around 6 pixels
    
    If TestItem.P.Enabled = False Then Exit Function
    
    tLoc = GetRect(TestItem)
    If TestItem.P.Width < 0 Then
        tToc = GetRect(TestItem)
        tLoc.X1 = tToc.X2
        tLoc.X2 = tToc.X1
    End If
    If TestItem.P.Height < 0 Then
        tToc = GetRect(TestItem)
        tLoc.Y1 = tToc.Y2
        tLoc.Y2 = tToc.Y1
    End If
    
    If X >= tLoc.X1 - conVr And Y >= tLoc.Y1 - conVr And _
    X <= (tLoc.X2 + conVr) And Y <= (tLoc.Y2 + conVr) Then
        If IsObjLine(TestItem) Then
            Dim sngSlope As Single
            Dim sngCalc  As Single
            
            If Abs(TestItem.P.Width) < 0.1 And Abs(TestItem.P.Height) > 0.1 Then
                sngSlope = TestItem.P.Width / TestItem.P.Height
                sngCalc = (X - TestItem.P.Left) / (Y - TestItem.P.Top)
            Else
                sngSlope = TestItem.P.Height / TestItem.P.Width
                sngCalc = (Y - TestItem.P.Top) / (X - TestItem.P.Left)
            End If
            
            If sngSlope = 0 Then
                If Abs(sngCalc - sngSlope) > 0.1 Then
                    Exit Function
                End If
            Else
                If Abs(sngCalc - sngSlope) / sngSlope > 0.3 Then
                    Exit Function
                End If
            End If
        End If
        'if the TestItem is within range, and it is not mSelected,
        'then choose this one, or when this one is the same as mSelected
        'and KeepSelected = True
        If TestItem Is DefaultItem And Not DefaultItem Is Nothing Then
            t__func = conHitDefault
        Else
            t__func = conHit
        End If
        If Not mcolSelections Is Nothing Then
            mcolSelections.Add TestItem
        End If
    End If
End Function



Private Function QuerySave() As Boolean
    If (Not mFlowChart.Changed) Or mFlowChart.Count = 0 Then
        'saved, pass good
        QuerySave = True 'True = saved
    Else
        Select Case MsgBox("The flow chart " & IIf(Len(mFlowChart.FileName), mFlowChart.FileName, "Untitled") & " has changed.  Do you want to save these changes?", vbExclamation Or vbYesNoCancel)
            Case vbNo
                QuerySave = True
            Case vbYes
                mnuFileSave_Click 0
                QuerySave = Not mFlowChart.Changed 'True = not saved
            Case vbCancel
                QuerySave = False
        End Select
    End If
End Function

'Redraws the selected item & handles.
Public Sub Redraw()
    Screen.MousePointer = vbArrowHourglass
    picBay.Cls
    picBayPaint
    RedrawHandles mSelected
    Screen.MousePointer = vbDefault
End Sub

Private Sub RedrawHandles(ByVal Obj As FlowItem, Optional ByVal Focus As Boolean = True)
    Dim i As Long
    Dim sngBox As Single
    
    sngBox = ctlBox(0).Width
    
    For i = conHMin To conHMax   'hide all of 'em
        ctlBox(i).Visible = False
        ctlBox(i).BackColor = IIf(Focus, vbHighlight, vb3DShadow)
    Next i
    
    If Obj Is Nothing Then
        'leave 'em hidden
    ElseIf IsObjLine(Obj) Then
        ctlBox(conNW).Move Obj.P.Left - sngBox / 2, Obj.P.Top - sngBox / 2
        ctlBox(conSE).Move Obj.P.Left + Obj.P.Width - sngBox / 2, Obj.P.Top + Obj.P.Height - sngBox / 2
        ctlBox(conNW).MousePointer = vbCrosshair
        ctlBox(conSE).MousePointer = vbCrosshair
        ctlBox(conNW).Visible = True
        ctlBox(conSE).Visible = True
    Else
        ctlBox(conNW).Move Obj.P.Left - sngBox, Obj.P.Top - sngBox
        ctlBox(conN).Move Obj.P.CenterX - sngBox / 2, Obj.P.Top - sngBox
        'changed
        ctlBox(conNE).Move Obj.P.Left + Obj.P.Width + conPixel, Obj.P.Top - sngBox
        ctlBox(conE).Move Obj.P.Left + Obj.P.Width + conPixel, Obj.P.CenterY - sngBox / 2
        ctlBox(conSE).Move Obj.P.Left + Obj.P.Width + conPixel, Obj.P.Top + Obj.P.Height + conPixel
        ctlBox(conS).Move Obj.P.CenterX - sngBox / 2, Obj.P.Top + Obj.P.Height + conPixel
        ctlBox(conSW).Move Obj.P.Left - sngBox, Obj.P.Top + Obj.P.Height + conPixel
        ctlBox(conW).Move Obj.P.Left - sngBox, Obj.P.CenterY - sngBox / 2
        For i = conHMin To conHMax 'show them
            Select Case i 'change mouse pointer
                Case conNW, conSE: ctlBox(i).MousePointer = vbSizeNWSE
                Case conN, conS: ctlBox(i).MousePointer = vbSizeNS
                Case conNE, conSW: ctlBox(i).MousePointer = vbSizeNESW
                Case conW, conE: ctlBox(i).MousePointer = vbSizeWE
            End Select
            ctlBox(i).Visible = True
        Next i
    End If
End Sub

Private Sub ResetForDraw() 'for drawing interface
    Set mSelected = Nothing 'release selected object
    SelectionChanged
    Redraw 'redraw the screen
    mblnAdding = False 'turn these off
    mFlowChart.Changed = False
    tlbDrawButtonsOff
End Sub

Private Function Test(ByVal X As Single, ByVal Y As Single, ByVal CanRotate As Boolean) As FlowItem
'Called by Hit() or from
' 1. picBay_MouseDown() KeepSelected:=True
' 2. Hit() KeepSelected:=False
    Dim objItem     As FlowItem
    Dim lngOrder    As Long
    Dim blnHit      As Boolean
    Dim blnFound    As Boolean
    Dim lngItem     As Long
   
    'The following code is to:
    ' 1. Maintain layering of objects
    '    when clicking.
    ' 2. Stop the selecting of lower-layer
    '    non-visible objects.
    ' 3. Realisitic of other draw programs.
    If Not mSelected Is Nothing Then
        'only hold selected item if it is clicked on
        If t__func(mSelected, Nothing, X, Y) Then
            Set Test = mSelected
        End If
    End If
    
    Set mcolSelections = New Collection
    
    For lngOrder = conTop To conBottom Step -1
        For lngItem = mFlowChart.Count To 1 Step -1
            Set objItem = mFlowChart(lngItem)
            If objItem.P.DrawOrder = lngOrder Then
                If Not blnFound Then
                    If CanRotate Then
                        Select Case t__func(objItem, Test, X, Y)
                        Case conHit
                            Set Test = objItem
                            blnFound = True
                            blnHit = True
                        Case conHitDefault
                            Set Test = objItem
                            blnHit = True
                        End Select
                    Else
                        Select Case t__func(objItem, Test, X, Y)
                        Case conHitDefault
                            Set Test = objItem
                            blnFound = True
                            blnHit = True
                        Case conHit
                            If Test Is Nothing Then
                                Set Test = objItem
                                blnFound = True
                                blnHit = True
                            End If
                            'if object was found & KeepSelected = True, then need not look no more
                            'If (Test Is mSelected) Then Exit Function
                        End Select
                    End If
                Else 'already found items, just go through and find other items in the same area
                    t__func objItem, Test, X, Y
                End If
            End If
        Next lngItem
    Next lngOrder
'    'middle
'    For Each objItem In mFlowChart
'        If objItem.P.DrawOrder = conMiddle Then
'            If t__func(objItem, Test, X, Y) <> conMiss Then
'                'if object was found  & KeepSelected = True, then need not look no more
'                If (Test Is mSelected) Then Exit Function
'            End If
'        End If
'    Next objItem
'
'    'top
'    For Each objItem In mFlowChart
'        If objItem.P.DrawOrder = conTop Then
'            If t__func(objItem, Test, X, Y) <> conMiss Then
'                'if object was found  & KeepSelected = True, then need not look no more
'                If (Test Is mSelected) Then Exit Function
'            End If
'        End If
'    Next objItem

    picSelection_Paint

    If Not blnHit Then
        'nothing selected
        Set Test = Nothing
    End If
End Function

Private Sub TextAutoSize(ByVal Object As FlowItem, ByVal WordWrap As Boolean, ByVal UserCalled As Boolean)
'called from mnuToolsAutoFitText
'    Dim curSize As Currency, blnBold As Boolean, blnItalic As Boolean
    Dim tRect As apiRECT ', strName As String
    Dim sngWidth As Single
    'Dim lngLoop As Long
    
    If UserCalled Then 'save undo data
        mUndo.Add Object, conUndoAutoSize
    End If
    
'    strName = picBay.Font.Name
'    curSize = picBay.Font.Size
'    blnBold = picBay.Font.Bold
'    blnItalic = picBay.Font.Italic
    'the size factor has been removed because
    'win API codes are in pixels at 100% zoom
    'supposedly
    SetFontNmSz picBay, Object.P.FontFace, Object.P.TextSize, mFlowChart
    'picBay.Font.Name = GetFontName(Object, mFlowChart)
    'picBay.Font.Size = Object.P.TextSize
    'always force bold to make the boxes larger
    picBay.Font.Bold = True 'Object.P.TextBold
    picBay.Font.Italic = Object.P.TextItalic
    picBay.Font.Underline = Object.P.TextUnderline
    
    If WordWrap Then 'word wrap, two or more lines
        'Problems could occur in this code for diamond shapes
        'because the size of the diamond affects the text
        'area inside, which the code does not account for.
        'Do this once.
        tRect.Right = ScaleX(Object.P.Width - Object.TextLeftMrg - Object.TextRightMrg - conIndent * 2, vbTwips, vbPixels) 'for apiDrawText()
        tRect.Bottom = ScaleY(Object.P.Height - Object.TextTopMrg - Object.TextBottomMrg - conIndent * 2, vbTwips, vbPixels)
        Call apiDrawText(picBay.hdc, Object.P.Text & vbNullChar, -1, tRect, DT_CALCRECT Or DT_WORDBREAK)
        Object.P.Height = Object.TextTopMrg + Object.TextBottomMrg + ScaleY(tRect.Bottom - tRect.Top, vbPixels, vbTwips) + conIndent * 2 'picBay.TextHeight(GetText(Object)) 'in twips
        sngWidth = Object.TextLeftMrg + Object.TextRightMrg + ScaleX(tRect.Right - tRect.Left, vbPixels, vbTwips) + conIndent * 2
        'Only change width if the text doesn't fit
        If Object.P.Width <= sngWidth Then
            Object.P.Width = sngWidth
        End If
    Else
        'word wrap to one line
        Object.P.Height = Object.TextTopMrg + Object.TextBottomMrg + picBay.TextHeight(Object.P.Text) + conIndent * 2   'in twips
        sngWidth = Object.TextLeftMrg + Object.TextRightMrg + picBay.TextWidth(Object.P.Text) + conIndent * 2 'in twips
        'Only change width if the text doesn't fit
        If Object.P.Width <= sngWidth Then
            Object.P.Width = sngWidth
        End If
    End If
    
    'ResetForDraw changes to font back to orginal values
'    picBay.Font.Name = strName
'    picBay.Font.Size = curSize
'    picBay.Font.Bold = blnBold
'    picBay.Font.Italic = blnItalic

    'keep within page dimensions
    If mFlowChart.PrinterError Then
        If Object.P.Width > mFlowChart.PScaleWidth - msngGrid Then
            Object.P.Width = mFlowChart.PScaleWidth - msngGrid
        End If
        If Object.P.Height > mFlowChart.PScaleHeight - msngGrid Then
            Object.P.Height = mFlowChart.PScaleHeight - msngGrid
        End If
    Else
        If Object.P.Width > Printer.ScaleWidth - msngGrid Then
            Object.P.Width = Printer.ScaleWidth - msngGrid
        End If
        If Object.P.Height > Printer.ScaleHeight - msngGrid Then
            Object.P.Height = Printer.ScaleHeight - msngGrid
        End If
    End If
End Sub

'Private Sub tlbChangeUpdate()
'    Dim lngItem As FAddType
'    Dim objButton As Button
'
'    'deselect all buttons
'    For Each objButton In tlbChange.Buttons
'        If objButton.Value = tbrPressed Or objButton.MixedState = True Then
'            objButton.MixedState = False
'            objButton.Value = tbrUnpressed 'deselect all items
'        End If
'    Next objButton
'
'    If IsSelected And Not IsInMode Then
'        'select the right item
'        lngItem = GetFItemNo(mSelected)
'        If lngItem > 0 Then
'            On Error Resume Next
'            tlbChange.Buttons(lngItem).Value = tbrPressed
'        End If
'        tlbChange.Visible = True
'    Else
'        tlbChange.Visible = False
'    End If
'    Form_Resize
'End Sub

Private Sub tlbDrawButtonsOff()
    'Go through each button and de-select it.
    Dim objButton As Button
    
    For Each objButton In tlbDraw.Buttons
        objButton.Value = tbrUnpressed
    Next objButton
    mlngAddType = 0 'clear this value
    mblnAdding = False 'turns off crosshair
    tlbDraw.Buttons(1).Value = tbrPressed 'mouse pointer selected
End Sub

Private Sub tlbDrawInit()
    Dim lngCounter As Long
    Dim lngItem    As Long
    Dim objItem    As FlowItem
    
    'Toolbar
    For lngItem = 2 To 11
        lngCounter = lngCounter + 1
        Set objItem = Duplicate(, lngCounter)
        With tlbDraw.Buttons(lngItem)
            If Len(objItem.DescriptionF) Then
                .Description = objItem.DescriptionF
            Else
                .Description = objItem.Description
            End If
            .ToolTipText = "Add " & .Description
        End With
        'Load menu items
        If lngCounter <= conAddPicture Then
            If lngCounter > 1 Then Load mnuShapeItem(lngCounter)
            mnuShapeItem(lngCounter).Caption = tlbDraw.Buttons(lngItem).Description
            If lngCounter = conAddPicture Then mnuShapeItem(lngCounter).Caption = mnuShapeItem(lngCounter).Caption & "..."
        End If
    Next lngItem
End Sub


Public Sub ToolBoxDone(ByVal FillStyle As Boolean, ByVal LineStyle As Boolean, ByVal LineWidth As Boolean, _
ByVal ArrowEngg As Boolean, ByVal ArrowSize As Boolean, ByVal BackColour As Boolean, ByVal ForeColour As Boolean, ByVal TextColour As Boolean)
    If Not IsInMode Then
        If gblnWindowMultis Then
            frmMultis.DoCopyAppearance FillStyle, LineStyle, LineWidth, ArrowEngg, ArrowSize, BackColour, ForeColour, TextColour
        Else
            RedrawSingle mSelected
        End If
        mblnToolRedraw = True
    End If
End Sub

Public Sub ToolBoxDoneNoCopy()
    If Not IsInMode Then
         mblnToolRedraw = True
    End If
End Sub


Private Sub ToolboxUpdate()
    'redraw other windows that are persistent
    'If gblnWindowColour Then frmColour.Update
    'If gblnWindowFont Then frmFont.Update
    If gblnWindowMultis Then frmMultis.Update
    If gblnWindowLayer Then frmLayer.Update
    
    '29-Dec-02, pop up menus
    If gblnWindowBack Then frmStyleBack.UpdateToForm mSelected, mFlowChart
    If gblnWindowLine Then frmStyleLine.UpdateToForm mSelected, mFlowChart
    If gblnWindowThick Then frmStyleThickness.UpdateToForm mSelected, mFlowChart
    If gblnWindowPick Then frmCPick.UpdateToForm mSelected, mFlowChart
End Sub

Private Sub UpdateBottom()
    Dim strEnd As String
    
    If IsSelected Then
        Select Case mFlowChart.UnitScale
        Case vbCentimeters: strEnd = " cm)"
        Case vbMillimeters: strEnd = " mm)"
        Case vbInches: strEnd = " in)"
        Case vbPixels: strEnd = " px)"
        Case vbCharacters: strEnd = " ch)"
        Case Else
            'convert to cm by default in case units not specified or invalid
            mFlowChart.UnitScale = vbCentimeters
            strEnd = " cm)"
        End Select
        
        sbrBottom.Panels(3).Text = "(" & Round(ScaleX(mSelected.P.Left, vbTwips, mFlowChart.UnitScale), 2) & ", " & Round(ScaleY(mSelected.P.Top, vbTwips, mFlowChart.UnitScale), conDigits) & strEnd
        sbrBottom.Panels(5).Text = "(" & Round(ScaleX(mSelected.P.Width, vbTwips, mFlowChart.UnitScale), 2) & " x " & Round(ScaleY(mSelected.P.Height, vbTwips, mFlowChart.UnitScale), conDigits) & strEnd
        
        If mSelected.Number = conAddButton Then
            sbrBottom.Panels(6).Text = "Macro"
            sbrBottom.Panels(7).Text = mSelected.P.Tag3
        ElseIf mSelected.Number = conAddText And Len(mSelected.P.Tag3) > 0 Then
            sbrBottom.Panels(6).Text = "Field code"
            If LCase$(Left$(mSelected.P.Tag3, 5)) = "type " Then
                sbrBottom.Panels(7).Text = Mid$(mSelected.P.Tag3, 6)
            End If
        ElseIf mSelected.Number = conAddPicture Then
            sbrBottom.Panels(6).Text = "Picture"
            sbrBottom.Panels(7).Text = mSelected.P.Text
        ElseIf mSelected.P.GroupNo <> 0 Then
            sbrBottom.Panels(6).Text = "Group"
            sbrBottom.Panels(7).Text = mSelected.P.GroupNo
        Else
            sbrBottom.Panels(6).Text = "Name"
            sbrBottom.Panels(7).Text = mSelected.P.Name
        End If
    Else
        sbrBottom.Panels(3).Text = ""
        sbrBottom.Panels(5).Text = ""
        sbrBottom.Panels(7).Text = ""
    End If
    sbrBottom.Panels(7).ToolTipText = sbrBottom.Panels(7).Text
End Sub

Private Sub UpdateBottomCoord(Left As Single, Top As Single, Width As Single, Height As Single)
    '-1 to Skip parameter
    Dim strEnd As String
    
    Select Case mFlowChart.UnitScale
    Case vbCentimeters: strEnd = " cm)"
    Case vbMillimeters: strEnd = " mm)"
    Case vbInches: strEnd = " in)"
    Case vbPixels: strEnd = " px)"
    Case vbCharacters: strEnd = " ch)"
Case Else
        Exit Sub
    End Select
    
    If Left <> -1 Then sbrBottom.Panels(3).Text = "(" & Round(ScaleX(Left, vbTwips, mFlowChart.UnitScale), 2) & ", " & Round(ScaleY(Top, vbTwips, mFlowChart.UnitScale), conDigits) & strEnd
    If Width <> -1 Then sbrBottom.Panels(5).Text = "(" & Round(ScaleX(Width, vbTwips, mFlowChart.UnitScale), 2) & " x " & Round(ScaleY(Height, vbTwips, mFlowChart.UnitScale), conDigits) & strEnd
End Sub



Private Sub UpdateFont()
    Dim lngFlags As Long
    
    If mblnFontBarBusy Then Exit Sub
    
    If Not IsSelected Or mblnAdding Then
        cboFont.Enabled = False
        cboSize.Enabled = False
        tlbFont.Enabled = False
        Exit Sub
    Else
        cboFont.Enabled = True
        cboSize.Enabled = True
        tlbFont.Enabled = True
    End If
    
    mblnFontBarBusy = True
    
    With mSelected.P
        If Len(.FontFace) And cboFont.ListCount > 1 Then
            On Error Resume Next
            cboFont = .FontFace
            If Err Then
                cboFont.ListIndex = -1 'don't select anything
            End If
        Else
            cboFont = "(Default)"
        End If

        On Error Resume Next
        cboSize = .TextSize
        If Err Then
            Err.Clear
            If .TextSize > conFontMin And .TextSize < conFontMax Then
                cboSize.AddItem .TextSize
                cboSize = .TextSize 'add and select proper size
            Else
                cboSize.ListIndex = -1 'do not select anything
            End If
        End If
        
        tlbFont.Buttons!Bold.Value = IIf(.TextBold, tbrPressed, tbrUnpressed)
        tlbFont.Buttons!Italic.Value = IIf(.TextItalic, tbrPressed, tbrUnpressed)
        tlbFont.Buttons!Underline.Value = IIf(.TextUnderline, tbrPressed, tbrUnpressed)
        lngFlags = .TextAlign  'GetTextFlags(frmMain.mSelected)
        tlbFont.Buttons!Left.Value = IIf(lngFlags = DT_LEFT, tbrPressed, tbrUnpressed)
        tlbFont.Buttons!Centre.Value = IIf(lngFlags = DT_CENTER, tbrPressed, tbrUnpressed)
        tlbFont.Buttons!Right.Value = IIf(lngFlags = DT_RIGHT, tbrPressed, tbrUnpressed)
    End With
    
    mblnFontBarBusy = False
End Sub


Private Sub UpdateScrollBars()
    If picBay.Width < picView.ScaleWidth Then
        hsbBar.Max = 0
        hsbBar.Enabled = False
    Else
        hsbBar.Max = ScaleX(picBay.Width - picView.ScaleWidth, vbTwips, conScrollScale)
        hsbBar.LargeChange = ScaleX(picView.ScaleWidth, vbTwips, conScrollScale)
        hsbBar.Enabled = True
    End If
    If picBay.Height < picView.ScaleHeight Then
        vsbBar.Max = 0
        vsbBar.Enabled = False
    Else
        vsbBar.Max = ScaleY(picBay.Height - picView.ScaleHeight, vbTwips, conScrollScale)
        vsbBar.LargeChange = ScaleY(picView.ScaleHeight, vbTwips, conScrollScale)
        vsbBar.Enabled = True
    End If
End Sub

Public Sub UpgradeVersion(ByVal Ver As Single)
    If mFlowChart.Version < Ver Then
        mFlowChart.Version = Ver
        MsgBox "The file version has been updated to version " & Round(Ver, 1) & " to accomodate newer features.", vbInformation
    End If
End Sub



Private Sub cboFont_Click()
    If IsSelected And Not mblnFontBarBusy Then
        If cboFont = "(Default)" Then
            mSelected.P.FontFace = "" 'default font
        Else
            mSelected.P.FontFace = cboFont
        End If
        If gblnWindowMultis Then
            frmMultis.DoCopyFont True, False, False, False, False, False
        Else
            RedrawSingle mSelected
        End If
    End If
End Sub


Private Sub cboSize_Click()
    If IsSelected And Not mblnFontBarBusy Then
        If Len(cboSize) Then
            mSelected.P.TextSize = cboSize
        Else
            mSelected.P.TextSize = 0@
            Prompt "The default font size will be used."
        End If
        
        If gblnWindowMultis Then
            frmMultis.DoCopyFont False, False, False, False, False, True
        Else
            RedrawSingle mSelected
        End If
    End If
End Sub


Private Sub chkAutoAdd_Click()
    If chkAutoAdd = vbUnchecked Then
        'turn off adding mode
        mblnAdding = False
        mlngAddType = 0
        tlbDrawButtonsOff
        RedrawHandles mSelected, True
    End If
End Sub

Private Sub ctlBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'X and Y coordinates relate to the relative position to the control
    'Edit box off when an item is sized
    If (Button And vbLeftButton) = vbLeftButton And IsSelected And IsEditing Then
        EditBoxOff
        RedrawHandles mSelected
'        tlbChangeUpdate
    End If
    'Sizing
    If (Button And vbLeftButton) = vbLeftButton And Not IsBusy Then
        mblnSizing = True 'turn on sizing
        If gblnWindowMultis Then
            If Not IsSelLine Then
                frmMultis.BeginSize
            End If
        End If
        RedrawHandles Nothing
        'store mouse locations
        If IsSelLine Then
            mMousePos.X1 = ctlBox(Index).Width / 2  'center mouse on
            mMousePos.Y1 = ctlBox(Index).Height / 2  'box for lines
        Else
            mMousePos.X1 = X
            mMousePos.Y1 = Y
        End If
        'get rectangle of selected object
        mSize = GetRect(mSelected)
        If IsSelLine Then 'position line
            linSize.X1 = mSize.X1
            linSize.Y1 = mSize.Y1
            linSize.X2 = mSize.X2
            linSize.Y2 = mSize.Y2
            linSize.Visible = True
        Else 'position shape
            shpMove.Left = mSelected.P.Left
            shpMove.Top = mSelected.P.Top
            'add in one pixel to width & height to correct indented pixel
            shpMove.Width = mSelected.P.Width + conPixel
            shpMove.Height = mSelected.P.Height + conPixel
            shpMove.Visible = True
        End If
        AutoRedrawOn
        picBayPaint
    End If
End Sub

Private Sub ctlBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnSizing Then
        mMousePos.X2 = X
        mMousePos.Y2 = Y
        ctlBoxSizing Index
        'Position the box
        If IsSelLine Then
            linSize.X1 = mSize.X1 'position points
            linSize.Y1 = mSize.Y1
            linSize.X2 = mSize.X2
            linSize.Y2 = mSize.Y2
        Else
            ctlBoxRectToShape shpMove, mSize
'            'reverses the negative values
'            'or transfers x,y coordinates
'            'so that shapes do not have
'            ' -width or -height
'            'Also: add in one pixel to width & height to correct indented pixel
'            shpMove.Visible = False
'            If mSize.x2 - mSize.x1 < 0 Then
'                shpMove.Width = mSize.x1 - mSize.x2 + conPixel
'                shpMove.Left = mSize.x1 + (mSize.x2 - mSize.x1)
'            Else
'                shpMove.Width = mSize.x2 - mSize.x1 + conPixel
'                shpMove.Left = mSize.x1
'            End If
'            If mSize.y2 - mSize.y1 < 0 Then
'                shpMove.Height = mSize.y1 - mSize.y2 + conPixel
'                shpMove.Top = mSize.y1 + (mSize.y2 - mSize.y1)
'            Else
'                shpMove.Height = mSize.y2 - mSize.y1 + conPixel
'                shpMove.Top = mSize.y1
'            End If
        End If
    End If
End Sub


Private Sub ctlBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnSizing Then
        mMousePos.X2 = X
        mMousePos.Y2 = Y
        
        mblnSizing = False 'turn off sizing
        shpMove.Visible = False
        linSize.Visible = False
        ctlBoxSizing Index
        'When sized, picture is of a custom size
        If mSelected.Number = conAddPicture Then
            GetPicture(mSelected).SetCustomSize
        End If
        'undo
        If Not mblnAdding And Not gblnWindowMultis Then
            mUndo.Add mSelected, conUndoSize
        End If
        
        'change Left, Width, Height, and Top of the selected object
        'no pixel adjustments are necessary
        
        Dim blnTopChanged As Boolean, blnBottomChanged As Boolean
        
        If Not IsSelLine Then 'not a FLine, FEndArrowLine, or FMidArrowLine
            ctlBoxRectToShape shpMove, mSize
            
            blnTopChanged = mSelected.P.Top <> shpMove.Top
            blnBottomChanged = (mSelected.P.Top + mSelected.P.Height) <> (shpMove.Top + shpMove.Height)
            
            With mSelected.P
                .Left = shpMove.Left
                .Top = shpMove.Top
                .Width = shpMove.Width
                .Height = shpMove.Height
            End With
            
            If mSelected.Number = conAddCircle Then
                Select Case Index
                Case conN, conS
                    mSelected.P.Width = mSelected.P.Height
                Case conW, conE
                    mSelected.P.Height = mSelected.P.Width
                End Select
            End If
                
            If Not mblnAdding And Not gblnWindowMultis Then
                ConnectY mSelected, mUndo.LastItem.P, mSelected.P, blnTopChanged, blnBottomChanged
            End If
        Else
            blnTopChanged = mSelected.P.Top <> mSize.Y1
            blnBottomChanged = (mSelected.P.Top + mSelected.P.Height) <> mSize.Y2
            
            With mSelected.P
                .Left = mSize.X1
                .Top = mSize.Y1
                .Width = mSize.X2 - mSize.X1
                .Height = mSize.Y2 - mSize.Y1
            End With
            If Not mblnAdding And Not gblnWindowMultis Then
                ConnectY mSelected, mUndo.LastItem.P, mSelected.P, blnTopChanged, blnBottomChanged
            End If
        End If

        If gblnWindowMultis Then
            frmMultis.EndSize IsSelLine
        End If
        If Not mblnAdding Then
            'redraws the handles at the new locations
            AutoRedrawOff 'should redraw automatically
            RedrawHandles mSelected
        End If
        mFlowChart.Changed = True
        UpdateBottom
    End If
End Sub

Private Sub Form_Activate()
    If mblnToolRedraw Then
        mblnToolRedraw = False
        Redraw
    End If
End Sub

Private Sub Form_Initialize()
    Set mFlowChart = New FlowChart
    Set mRegistry = New PRegistry
    Set mClipboard = New PClipboard
    Set mUndo = New PUndo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not IsBusy Then
        'font sizing code
        If KeyCode = 188 And IsSelected And (Shift And (vbCtrlMask Or vbShiftMask)) Then '< key
            If Not IsWithoutText Then
                'SetTextFormat Obj:=mSelected, SizeFactor:=(GetTextFormatSizeFactor(mSelected, mFlowChart.FontSize) - 1)
                mSelected.P.TextSize = mSelected.P.TextSize - conSizeFactorMultiplier
                UpgradeVersion 4
                If Not IsEditing Then RedrawSingle mSelected
                mFlowChart.Changed = True
            End If
        ElseIf KeyCode = 190 And IsSelected And (Shift And (vbCtrlMask Or vbShiftMask)) Then '> key
            If Not IsWithoutText Then
                'SetTextFormat Obj:=mSelected, SizeFactor:=(GetTextFormatSizeFactor(mSelected, mFlowChart.FontSize) + 1)
                mSelected.P.TextSize = mSelected.P.TextSize + conSizeFactorMultiplier
                UpgradeVersion 4
                If Not IsEditing Then RedrawSingle mSelected
                mFlowChart.Changed = True
            End If
        ElseIf (Shift And vbCtrlMask) = vbCtrlMask And IsSelected And TypeOf ActiveControl Is Picture Then
            'moves an object
            Dim sngAdjustment As Single
            If (Shift And vbShiftMask) = vbShiftMask Then
                sngAdjustment = Screen.TwipsPerPixelX 'move 1 pixel value
            Else
                sngAdjustment = msngGrid 'move one grid value
            End If
            
            Select Case KeyCode
            Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
                Select Case KeyCode
                Case vbKeyUp
                    mSelected.P.Top = mSelected.P.Top - sngAdjustment
                Case vbKeyDown
                    mSelected.P.Top = mSelected.P.Top + sngAdjustment
                Case vbKeyLeft
                    mSelected.P.Left = mSelected.P.Left - sngAdjustment
                Case vbKeyRight
                    mSelected.P.Left = mSelected.P.Left + sngAdjustment
                End Select
                If ctlBox(conHMin).Visible Then RedrawHandles Nothing
                CheckObjBounds mSelected, picBay.ScaleWidth, picBay.ScaleHeight
                
                UpdateBottomCoord mSelected.P.Left, mSelected.P.Top, -1, -1
                RedrawSingle mSelected
                'picBay.Cls: picBayPaint
            End Select
        ElseIf (vsbBar.Enabled Or hsbBar.Enabled) And TypeOf ActiveControl Is Picture Then
            'move around the screen
            Select Case KeyCode
            Case vbKeyUp 'up and down
                vsbBar.Value = IIf(vsbBar.Value > vsbBar.SmallChange, vsbBar.Value - vsbBar.SmallChange, 0)
            Case vbKeyDown
                vsbBar.Value = IIf(vsbBar.Value < (vsbBar.Max - vsbBar.SmallChange), vsbBar.Value + vsbBar.SmallChange, vsbBar.Max)
            Case vbKeyPageDown
                vsbBar.Value = IIf(vsbBar.Value < (vsbBar.Max - vsbBar.LargeChange), vsbBar.Value + vsbBar.LargeChange, vsbBar.Max)
            Case vbKeyPageUp
                vsbBar.Value = IIf(vsbBar.Value > vsbBar.LargeChange, vsbBar.Value - vsbBar.LargeChange, 0)
            Case vbKeyLeft 'left and right
                hsbBar.Value = IIf(hsbBar.Value > hsbBar.SmallChange, hsbBar.Value - hsbBar.SmallChange, 0)
            Case vbKeyRight
                hsbBar.Value = IIf(hsbBar.Value < (hsbBar.Max - hsbBar.SmallChange), hsbBar.Value + hsbBar.SmallChange, hsbBar.Max)
            Case 192 'key with "`~" above Tab key
                TabSelectItem (Shift And vbShiftMask)
            'more cases on KeyUp, and prior to this code
            End Select
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'edit on type, 22 August 2002
    If ((65 <= KeyAscii And KeyAscii <= 90) Or _
    (48 <= KeyAscii And KeyAscii <= 57) Or _
    (97 <= KeyAscii And KeyAscii <= 122)) And _
    Not IsInMode Then
        EditBoxOn
        If txtEdit.Visible Then
            txtEdit.SelStart = Len(txtEdit)
            txtEdit.SelText = Chr(KeyAscii)
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyCancel Then
        If mblnSizing Or mblnMoving Or mblnAdding Then
            'ResetForDraw vars
            mblnSizing = False
            mblnMoving = False
            If mblnAdding Then 'adding,deselect buttons
                tlbDraw.Buttons(mlngAddType).Value = tbrUnpressed
                mblnAdding = False
            End If
            mblnMouseMovedGrid = False
            mblnMouseMovedSensitive = False
            'ResetForDraw others
            linSize.Visible = False
            shpMove.Visible = False
            AutoRedrawOff
            picBayPaint
            'sometimes there is still an item selected
            If IsSelected Then
                RedrawHandles mSelected
            End If
'        ElseIf IsEditing Then
'            'ResetForDraw text box
'            txtEdit.Visible = False
'            RedrawHandles mSelected
        End If
    ElseIf KeyCode = 93 Then
        PopupMenu mnuPopup, vbPopupMenuLeftButton Or vbPopupMenuRightButton, 0, 0
    ElseIf Not IsBusy Then
        Select Case KeyCode
            Case vbKeyHome
                hsbBar.Value = 0
            Case vbKeyEnd
                hsbBar.Value = hsbBar.Max
            Case vbKeySeparator 'enter key on the numeric keypad
            Case vbKeyF5
                'Force redraw.  Force redraw is
                'not the same as Redraw().
                picBay.Cls
                If Not IsBusy Then
                    picBay.AutoRedraw = False
                End If
                picBayPaint True
                RedrawHandles mSelected
            'This is to redraw with gridlines.
            Case 188, 190 '<, >
                If IsSelected And (Shift And (vbCtrlMask Or vbShiftMask)) Then
                    If (Not IsEditing) And (Not IsWithoutText) Then Redraw
                    'If gblnWindowFont Then frmFont.Update 'refreshes data in font toolbar
                End If
        End Select
        'redraw after moving object
        If ((Shift And vbCtrlMask) = vbCtrlMask) And IsSelected And TypeOf ActiveControl Is Picture Then
            Select Case KeyCode
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
                    Redraw
            End Select
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i           As Integer
    
    On Error Resume Next
    
    'Registry settings
    If mRegistry.ReadRegistry <> 0 Then
        MsgBox "Error reading the registry.", vbExclamation
    End If
    
    'Printer
    Printer.ScaleMode = vbTwips
    
    'Interface
    'Load handle boxes
    For i = conHMin To conHMax
        Load ctlBox(i)
    Next i
    
    tlbDrawInit
    
    'set up these properties
    mnuFormatGridSnap.Checked = mRegistry.GridSnap
    mnuFormatGridShow.Checked = mRegistry.GridShow
    tlbFont.Visible = mRegistry.FontBar
    'sbrBottom.Visible = mRegistry.StatusBar
    
    'load arrow sizes
    mnuPLineSize(0).Caption = "25%"
    Load mnuPLineSize(1): mnuPLineSize(1).Caption = "50%"
    Load mnuPLineSize(2): mnuPLineSize(2).Caption = "75%"
    Load mnuPLineSize(3): mnuPLineSize(3).Caption = "100%"
    Load mnuPLineSize(4): mnuPLineSize(4).Caption = "150%"
    
    If mRegistry.FontBar Then LoadFont
    
    'View settings - must go at front
    SetScrollBars   'size scroll bars
    Form_Resize     'move scroll bars
    SetView 100, conFontName, conFontSize
    
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim objPicture As StdPicture
'    Dim bytReturn() As Byte
    Dim objAdded As FlowItem
    Dim objPic   As FPicture
    Dim F        As Long 'file stream
    
    If Data.GetFormat(vbCFDIB) Then 'bitmap
        Effect = vbDropEffectNone
    ElseIf Data.GetFormat(vbCFText) Then  'text
        Effect = vbDropEffectCopy
        Set objAdded = New FText
        mFlowChart.Add objAdded
        'SetText objAdded, Data.GetData(vbCFText)
        objAdded.P.Text = Data.GetData(vbCFText)
        TextAutoSize objAdded, False, False
        Set mSelected = objAdded
        Redraw
        Set objAdded = Nothing
    ElseIf Data.GetFormat(vbCFFiles) Then  'files
        If Data.Files.Count = 1 Then
            Select Case LCase$(Right$(Data.Files(1), 4))
            Case "." & conDefaultExt  'flow chart file
                Effect = vbDropEffectCopy
                If QuerySave Then
                    MousePointer = vbHourglass
                    OpenFile Data.Files(1)
                    MousePointer = vbDefault
                End If
            Case ".bmp", ".ico", ".cur", ".wmf", ".emf", ".jpg", ".gif"
                Effect = vbDropEffectCopy
                Set objAdded = mFlowChart.Add(New FPicture)
                Set objPic = objAdded
                objAdded.P.Text = Data.Files(1)
                objPic.LoadPicture
                objPic.SetDefaultSize picBay
                Set mSelected = objAdded
                Redraw
                Set objAdded = Nothing
                Set objPic = Nothing
            Case ".txt"
                Effect = vbDropEffectCopy
                Set objAdded = New FText
                mFlowChart.Add objAdded
                
                On Error GoTo Handler
                F = FreeFile()
                Open Data.Files(1) For Input As #F
                If LOF(F) <= 16384 Then ' 2 ^ 14 or 16 Kb
                    'SetText objAdded, Input(LOF(F), #F)
                    objAdded.P.Text = Input(LOF(F), #F)
                Else
                    MsgBox "The text file is too large to be added to this document (by drag-drop operation).", vbExclamation, "Drag Drop"
                End If
                Close #F
                TextAutoSize objAdded, False, False
                Set mSelected = objAdded
                Redraw
                
                Set objAdded = Nothing
            Case Else
                Effect = vbDropEffectNone
            End Select
        Else
            Effect = vbDropEffectNone
        End If
    Else
        Effect = vbDropEffectNone
    End If
    
    SelectionChanged
    Exit Sub
Handler:
    MsgBox "Error on drag drop operation.", vbExclamation, "Drag Drop"
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If Data.GetFormat(vbCFDIB) Then 'bitmap
        'Effect = vbDropEffectCopy
        Effect = vbDropEffectNone
    ElseIf Data.GetFormat(vbCFText) Then  'text
        Effect = vbDropEffectCopy
    ElseIf Data.GetFormat(vbCFFiles) Then  'files
        If Data.Files.Count = 1 Then
            Select Case LCase$(Right$(Data.Files(1), 4))
                Case "." & conDefaultExt  'flow chart file
                    Effect = vbDropEffectCopy
                Case ".bmp", ".ico", ".cur", ".wmf", ".emf", ".jpg", ".gif"
                    Effect = vbDropEffectCopy
                Case ".txt"
                    Effect = vbDropEffectCopy
                Case Else
                    Effect = vbDropEffectNone
            End Select
        Else
            Effect = vbDropEffectNone
        End If
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub Form_Paint()
    Line (hsbBar.Left + hsbBar.Width, vsbBar.Top + vsbBar.Height)-(ScaleWidth, ScaleHeight), BackColor, BF
    PaintPicture ilsTools.ListImages![small_icon].Picture, hsbBar.Left + hsbBar.Width, vsbBar.Top + vsbBar.Height, 16 * Screen.TwipsPerPixelX, 16 * Screen.TwipsPerPixelY
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not QuerySave Then
        'could not save, cancel
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    Dim lngTop As Long
    Dim lngHeight As Long
    
    On Error GoTo Handler
    If WindowState <> vbMinimized Then
        
    
        lngHeight = ScaleHeight - tlbDraw.Height - sbrBottom.Height - tlbColour.Height
        lngTop = tlbDraw.Height
        
        If tlbFont.Visible Then
            lngHeight = lngHeight - tlbFont.Height
            lngTop = lngTop + tlbFont.Height
        End If
        
        sbrBottom.Move picSelection.Width, ScaleHeight - sbrBottom.Height, ScaleWidth - picSelection.Width
        
        picView.Move picSelection.Width, lngTop, ScaleWidth - vsbBar.Width - picSelection.Width, lngHeight
        tlbColour.Move picSelection.Width, ScaleHeight - sbrBottom.Height - tlbColour.Height
        
        vsbBar.Move ScaleWidth - vsbBar.Width, picView.Top, vsbBar.Width, picView.Height
        hsbBar.Move tlbColour.Left + tlbColour.Width, picView.Top + picView.Height, picView.Width - tlbColour.Width
        
        UpdateScrollBars
    End If
Handler:
End Sub

Private Sub Form_Terminate()
    Set mUndo = Nothing
    Set mcolSelections = Nothing
    Set mFlowChart = Nothing
    Set mRegistry = Nothing
    Set mClipboard = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If gblnWindowMultis Then Unload frmMultis
    Set frmMultis = Nothing
    'Save settings
    mRegistry.GridSnap = mnuFormatGridSnap.Checked
    mRegistry.GridShow = mnuFormatGridShow.Checked
    mRegistry.FontBar = tlbFont.Visible
    'mRegistry.StatusBar = sbrBottom.Visible
    
    If mRegistry.SaveRegistry <> 0 Then
        MsgBox "Error saving the registry.", vbExclamation
    End If
End Sub

Private Sub hsbBar_Change()
    picBay.Left = -ScaleX(hsbBar.Value, conScrollScale, vbTwips)
End Sub




Private Sub mnuEditCopy_Click()
    If IsEditing Then
        mClipboard.Text = txtEdit.SelText
    ElseIf IsMultipleSelected Then
        frmMultis.DoCopy
    Else
        Set mClipboard.FlowItem = mSelected
    End If
End Sub

Private Sub mnuEditCut_Click()
    If IsEditing Then
        mClipboard.Text = txtEdit.SelText
        txtEdit.SelText = ""
    ElseIf IsMultipleSelected Then
        frmMultis.DoCut
    Else
        mUndo.Add mSelected, conUndoCut
        mClipboard.SetFreeFlowItem mSelected
        mFlowChart.RemoveObj mSelected
        Set mSelected = Nothing
        SelectionChanged
        Redraw
    End If
End Sub

Private Sub mnuEditPaste_Click()
    Dim objItem As FlowItem
    Dim objPaste As FlowItem
    
    'there is also mnuPopupPasteHere
    If IsEditing Then
        txtEdit.SelText = mClipboard.Text
    ElseIf mClipboard.IsItem Then
        Dim objAdd As FlowItem
        
        Set objAdd = mFlowChart.Add(mClipboard.FlowItem)
        'if tempx & y values filled, use those
        If mClipboard.TempX <> 0 And mClipboard.TempY <> 0 Then
            objAdd.P.Left = Grid(mClipboard.TempX)
            objAdd.P.Top = Grid(mClipboard.TempY)
        End If
        Set mSelected = objAdd
        CheckObjBounds objAdd, picBay.ScaleWidth, picBay.ScaleHeight
        SelectionChanged
        Redraw
    ElseIf mClipboard.IsSelection Then
        frmMultis.DoPaste
    End If
End Sub


Private Sub mnuEditSelAll_Click()
    Dim objItem As FlowItem
    
    Load frmMultis
    With frmMultis
        .Show vbModeless, Me
        .Clear
        For Each objItem In mFlowChart
            .AddItem2 objItem
        Next objItem
        Set mSelected = mFlowChart(1)
        .Update
        Redraw
    End With
End Sub

Private Sub mnuEditUndo_Click()
    mUndo.DoUndoOp mFlowChart
    SelectionChanged
    Redraw 'refresh display
    mFlowChart.Changed = True
End Sub



Private Sub mnuFilePreview_Click()
    Dim intZoom As Integer
    
    intZoom = mFlowChart.ZoomPercent
    picBay.Visible = False
    Load frmForm
    Set frmForm.mFlowChart = mFlowChart
    frmForm.Caption = "Print Preview"
    frmForm.SetZoom 50
    frmForm.Show vbModal, Me
    mFlowChart.ZoomPercent = intZoom
    picBay.Visible = True
End Sub





Private Sub mnuFormatExportWord_Click()
    Prompt "Exporting to MS Word..."
    ExportWord mFlowChart, Me
    Prompt
End Sub

Private Sub mnuFormatObjProp_Click()
    If Not IsBusy Then
        Load frmObject
        frmObject.LoadObj mFlowChart, mSelected
        frmObject.Show vbModal, Me
        Unload frmObject
        UpdateBottom
    End If
End Sub

Private Sub mnuFormatOpt_Click()
    frmProp.LoadProperties mFlowChart, mRegistry
    frmProp.Show vbModal, Me
    UpdateBottom
End Sub



Private Sub mnuHelp_Click()
    Prompt
End Sub

Private Sub mnuHelpOInfo_Click()
    Const conYN = "Yes/No"
    
    Load frmVal
    
    frmVal.AddItem "Changed:", Format(mFlowChart.Changed, conYN)
    frmVal.AddItem "Auto-redraw:", Format(picBay.AutoRedraw, "On/Off")
    frmVal.AddItem "Editing:", Format(IsEditing, conYN)
    frmVal.AddItem "Adding:", Format(mblnAdding, conYN)
    frmVal.AddItem "Moving:", Format(mblnMoving, conYN)
    frmVal.AddItem "Moved/Sized:", Format(mblnMovedSized, conYN)
    frmVal.AddItem "Sizing:", Format(mblnSizing, conYN)
    frmVal.AddItem "Window Zoom:", IIf(mFlowChart.ZoomPercent = 0, "100%", mFlowChart.ZoomPercent)
    frmVal.FitToSize
    frmVal.Caption = "Window Properties"
    frmVal.Show vbModal, Me
End Sub



Private Sub mnuPButton_Click()
    mnuPButtonRunMacro.Enabled = Len(mSelected.P.Tag3)
    mnuPButtonDelMacro.Enabled = Len(mSelected.P.Tag3)
End Sub

Private Sub mnuPButtonDelMacro_Click()
    mUndo.Add mSelected, conUndoChangeTag3Macro
    mSelected.P.Tag3 = ""
End Sub

Private Sub mnuPButtonEditMacro_Click()
    Load frmObject
    frmObject.LoadObj mFlowChart, mSelected
    frmObject.fraButton.Font.Bold = True
    frmObject.Show vbModal, Me
    UpdateBottom
End Sub

Private Sub mnuPButtonRunMacro_Click()
    GetButton(mSelected).Click picView, mFlowChart
End Sub

Private Sub mnuPCancel_Click()
    mnuEditCancel_Click
End Sub

Private Sub mnuPGroup_Click()
    If IsMultipleSelected Then
        mnuPGroupGroup.Enabled = frmMultis.GetCount
        mnuPGroupUngroup.Enabled = frmMultis.GetCount And mSelected.P.GroupNo <> 0
        mnuPGroupUnselGroup.Enabled = mSelected.P.GroupNo
        mnuPGroupSelGroup.Enabled = mSelected.P.GroupNo
        mnuPGroupSelOne.Enabled = True
    Else
        mnuPGroupGroup.Enabled = False
        mnuPGroupUngroup.Enabled = False
        mnuPGroupUnselGroup.Enabled = False
        mnuPGroupSelGroup.Enabled = False
        mnuPGroupSelOne.Enabled = False
    End If
End Sub

Private Sub mnuPGroupGroup_Click()
    frmMultis.mnuToolsGroup_Click
End Sub


Private Sub mnuPGroupSelGroup_Click()
    Load frmMultis
    frmMultis.Clear
    frmMultis.AddGroup mSelected.P.GroupNo
End Sub

Private Sub mnuPGroupSelOne_Click()
    Unload frmMultis
End Sub

Private Sub mnuPGroupUngroup_Click()
    frmMultis.mnuToolsUngroup_Click
End Sub

Private Sub mnuPGroupUnselGroup_Click()
    frmMultis.RemoveGroup mSelected.P.GroupNo
End Sub

Private Sub mnuPLine_Click()
    With mSelected.P
        mnuPLineNormal.Checked = Not .ArrowEngg
        mnuPLineEng.Checked = .ArrowEngg
        
        mnuPLineSolid.Checked = .FillStyle = vbFSSolid
        mnuPLineHollow.Checked = .FillStyle <> vbFSSolid
        
        mnuPLineNoArrow.Checked = mSelected.Number = conAddLine
        mnuPLineFlowArrow.Checked = mSelected.Number = conAddMidArrowLine
        mnuPLinePointingArrow.Checked = mSelected.Number = conAddEndArrowLine
    End With
End Sub

Private Sub mnuPLineEng_Click()
    mUndo.Add mSelected, conUndoChangeArrowType
    mSelected.P.ArrowEngg = True
    Redraw
End Sub


Private Sub mnuPLineFlowArrow_Click()
    mnuShapeItem_Click conAddMidArrowLine
End Sub


Private Sub mnuPLineHollow_Click()
    mUndo.Add mSelected, conUndoChangeArrowSolid
    mSelected.P.FillStyle = vbFSTransparent
    Redraw
End Sub


Private Sub mnuPLineNoArrow_Click()
    mnuShapeItem_Click conAddLine
End Sub

Private Sub mnuPLineNormal_Click()
    mUndo.Add mSelected, conUndoChangeArrowType
    mSelected.P.ArrowEngg = False
    Redraw
End Sub


Private Sub mnuPLinePointingArrow_Click()
    mnuShapeItem_Click conAddEndArrowLine
End Sub


Private Sub mnuPLineSize_Click(Index As Integer)
    mUndo.Add mSelected, conUndoChangeArrowSize
    mSelected.P.ArrowSize = Val(mnuPLineSize(Index).Caption)
    Redraw
End Sub


Private Sub mnuPLineSolid_Click()
    mUndo.Add mSelected, conUndoChangeArrowSolid
    mSelected.P.FillStyle = vbFSSolid
    Redraw
End Sub


Private Sub mnuPopup_Click()
    'popup menu mirrors edit menu
    mnuEdit_Click 'copies the mnuEdit functions
    mnuPopupCut.Enabled = mnuEditCut.Enabled
    mnuPopupCopy.Enabled = mnuEditCopy.Enabled
    mnuPopupPasteHere.Enabled = (mnuEditPaste.Enabled And mClipboard.IsItem)
    mnuPopupEdit.Enabled = mnuEditEdit.Enabled
    mnuPopupBringFront.Enabled = mnuEditToFront.Enabled
    mnuPopupSendBack.Enabled = mnuEditToBack.Enabled
    mnuPCancel.Enabled = mnuEditCancel.Enabled
    mnuPProp.Enabled = IsSelected
End Sub


Private Sub mnuPopupBringFront_Click()
    mnuEditToFront_Click
End Sub


Private Sub mnuPopupCopy_Click()
    mnuEditCopy_Click
End Sub

Private Sub mnuPopupCut_Click()
    mnuEditCut_Click
End Sub

Private Sub mnuPopupEdit_Click()
    mnuEditEdit_Click
End Sub

Private Sub mnuPopupPasteHere_Click()
    'Dim objAdd As FlowItem
    
    'pastes objects to grid
'    Set objAdd = mFlowChart.Add(mClipboard.FlowItem)
'    objAdd.P.Left = Grid(mClipboard.TempX)
'    objAdd.P.Top = Grid(mClipboard.TempY)
'    Set mSelected = objAdd
'    SelectionChanged
'    Redraw
'    mClipboard.TempClear

    mnuEditPaste_Click
    mSelected.P.Left = Grid(mClipboard.TempX)
    mSelected.P.Top = Grid(mClipboard.TempY)
    mClipboard.TempClear
    Redraw
End Sub

Private Sub mnuPopupSendBack_Click()
    mnuEditToBack_Click
End Sub

Private Sub mnuPPic_Click()
    mnuPPicRelative.Checked = (InStr(1, mSelected.P.Text, "\") = 0)
    mnuPPicAbs.Checked = InStr(1, mSelected.P.Text, "\")
    mnuPPicReload.Enabled = Len(mSelected.P.Text) 'filename
End Sub

Private Sub mnuPPicAbs_Click()
    Dim lngLast As Long
    
    mUndo.Add mSelected, conUndoObject
    lngLast = InStrRev(mSelected.P.Text, "\")
    If lngLast = 0 Then
        mSelected.P.Text = mFlowChart.GetPath & mSelected.P.Text
    End If
    mnuPPicReload_Click
End Sub

Private Sub mnuPPicChange_Click()
    EditBoxOnPic
End Sub

Private Sub mnuPPicRatio_Click()
    Dim objPic As FPicture
    
    'For pictures only
    Set objPic = mSelected
    If objPic.IsLoaded() Then
        Load frmZoom
        frmZoom.Zoom = objPic.EstimatedZoom(picBay)
        frmZoom.Caption = "Format Picture"
        frmZoom.Show 1, Me
        If frmZoom.Changed Then
            With mSelected
                mUndo.Add mSelected, conUndoPicZoom
                objPic.SetCustomSize
                .P.Width = objPic.GetDefWidth(picBay) * (frmZoom.Zoom / 100)
                .P.Height = objPic.GetDefHeight(picBay) * (frmZoom.Zoom / 100)
            End With
            'version changed
            UpgradeVersion 4
            mFlowChart.Changed = True
            Redraw
        End If
        Unload frmZoom
    Else
        MsgBox "No picture has been loaded.", vbInformation
    End If
    Set objPic = Nothing
End Sub


Private Sub mnuPPicRelative_Click()
    Dim lngLast As Long
    
    mUndo.Add mSelected, conUndoObject
    lngLast = InStrRev(mSelected.P.Text, "\")
    If lngLast > 0 Then
        mSelected.P.Text = Mid$(mSelected.P.Text, lngLast + 1)
    End If
    mnuPPicReload_Click
End Sub

Private Sub mnuPPicReload_Click()
    mSelected.Refresh mFlowChart, picView
    Redraw
    UpdateBottom
End Sub

Private Sub mnuPPicReset_Click()
    Dim objPic As FPicture
    
    If IsSelected Then
        If TypeOf mSelected Is FPicture Then
            mUndo.Add mSelected, conUndoPicReset
            Set objPic = mSelected
            objPic.LoadPicture
            objPic.SetDefaultSize picBay
            mFlowChart.Changed = True
            Redraw
            Set objPic = Nothing
        End If
    End If
End Sub

Private Sub mnuPProp_Click()
    mnuFormatObjProp_Click
End Sub

Private Sub mnuPText_Click()
    mnuPTextRemove.Enabled = Len(mSelected.P.Tag3) <> 0
    
End Sub

Private Sub mnuPTextAdd_Click()
    Load frmObject
    frmObject.LoadObj mFlowChart, mSelected
    frmObject.fraField.Font.Bold = True
    frmObject.Show vbModal, Me
    UpdateBottom
End Sub


Private Sub mnuPTextRemove_Click()
    mUndo.Add mSelected, conUndoChangeTag3Field
    mSelected.P.Tag3 = ""
    mSelected.P.CanEdit = True
End Sub

Private Sub mnuShape_Click()
'Form information
'   mnuShape        &Shape
'   mnuShapeItem(1) <blank>
    Dim i As Long
    Const conUpper = conAddPicture

    'disable or enable all items, which were
    'generated at startup.  These items
    'mirror the items on the toolbar.
    If IsSelected And Not IsInMode And Not IsMultipleSelected Then
        For i = 1 To conUpper
            mnuShapeItem(i).Enabled = True
            mnuShapeItem(i).Checked = False
        Next i
        'mnuShapeItem(conAddPicture).Enabled = False
        If mSelected.Number > 0 And mSelected.Number <= conUpper Then
            mnuShapeItem(mSelected.Number).Checked = True
        End If
        mnuShapeButton.Enabled = True
        mnuShapeButton.Checked = (mSelected.Number = conAddButton)
        mnuShapeArea.Enabled = True
        mnuShapeArea.Checked = (mSelected.Number = conAddExtra1)
    Else
        For i = 1 To conUpper
            mnuShapeItem(i).Enabled = False
            mnuShapeItem(i).Checked = False
        Next i
        mnuShapeButton.Enabled = False
        mnuShapeButton.Checked = False
        mnuShapeArea.Enabled = False
        mnuShapeArea.Checked = False
    End If
    
    Prompt "Use this menu to change the shape of an object."
End Sub

Private Sub mnuShapeArea_Click()
    Dim objArea As FlowItem
    Dim objAny  As FlowItem
    
      
    For Each objAny In mFlowChart
        If objAny.Number = conAddExtra1 Then
            Select Case MsgBox("A previous Area object is defined.  Do you wish to remove the old area and replace it with this one?", vbQuestion Or vbYesNoCancel)
            Case vbYes
                mFlowChart.RemoveObj objAny
            Case vbCancel
                Exit Sub
            End Select
        End If
    Next objAny
    
    Set objArea = mFlowChart.AddBefore(Duplicate(, conAddExtra1), Nothing)
    objArea.P.Left = 0
    objArea.P.Top = 0
    objArea.P.Width = mSelected.P.Left + mSelected.P.Width
    objArea.P.Height = mSelected.P.Top + mSelected.P.Height
    
    mFlowChart.RemoveObj mSelected
    
    Set mSelected = objArea
    Redraw
End Sub

Private Sub mnuShapeButton_Click()
    Dim objButton As FlowItem
    
    Set objButton = New FButton
    objButton.P.Left = mSelected.P.Left
    objButton.P.Top = mSelected.P.Top
    objButton.P.Width = mSelected.P.Width
    objButton.P.Height = mSelected.P.Height
    objButton.P.Text = InputBox("Type in the caption you wish to have for the button.  Then, another window will show and fill in the data in the Buttons frame.", "Add Button")
    If Len(objButton.P.Text) Then
        mFlowChart.AddBefore objButton, mSelected
        mFlowChart.RemoveObj mSelected
        Set mSelected = objButton
        mnuFormatObjProp_Click
        'mnuToolsChangeMacro_Click
        Redraw
    Else 'cancelled or no Caption
        Set objButton = Nothing
    End If
End Sub

'mnuShapeItem index starts at 1
Private Sub mnuShapeItem_Click(Index As Integer)
    Dim objDup As FlowItem 'pointer to new item
    Dim objOrg As FlowItem 'pointer to original item
    Dim blnLine As Boolean

    Set objOrg = mSelected 'hold original item
    blnLine = IsSelLine
    Set objDup = Duplicate(mSelected, Index) 'add new item
    mFlowChart.AddBefore objDup, objOrg 'to flow chart
    
    If Not IsObjLine(objDup) And blnLine Then
        'original item is a line, but the new item is not
        If objDup.P.Width < 0 Then
            objDup.P.Left = objDup.P.Left + objDup.P.Width
            objDup.P.Width = -objDup.P.Width
        End If
        If objDup.P.Height < 0 Then
            objDup.P.Top = objDup.P.Top + objDup.P.Height
            objDup.P.Height = -objDup.P.Height
        End If
    End If
    
    Set mSelected = objDup
    
    If objDup.Number = conAddPicture Then 'load picture if needs be
        EditBoxOnPic
        If Not GetPicture(objDup).IsLoaded Then
            Set mSelected = objOrg
            mFlowChart.RemoveObj objDup
            Set objDup = Nothing
            SelectionChanged
            Redraw
            Exit Sub
        End If
    End If
    
    mFlowChart.RemoveObj objOrg
    
    Set mSelected = objDup
    Set objDup = Nothing
    Set objOrg = Nothing
    
    SelectionChanged
    Redraw
End Sub

Private Sub mnuTextFont_Click()
    On Error Resume Next
    With dlgFile
        .Flags = cdlCFBoth Or cdlCFLimitSize Or cdlCFWYSIWYG Or cdlCFScalableOnly Or cdlCFForceFontExist
        If Len(mSelected.P.FontFace) Then
            .FontName = mSelected.P.FontFace
        Else
            .FontName = mFlowChart.FontName
        End If
        .FontSize = mSelected.P.TextSize
        .FontBold = mSelected.P.TextBold
        .FontItalic = mSelected.P.TextItalic
        .FontUnderline = mSelected.P.TextUnderline
        .Min = conFontMin
        .Max = conFontMax
        .ShowFont
        If Err = cdlCancel Then 'cancelled
            Exit Sub
        ElseIf Err <> 0 Then
            MsgBox "Error showing Font dialog box.", vbExclamation, "Font"
        Else
            mUndo.Add mSelected, conUndoFontChanged
            'if version less than 6 and font names the same
            If mFlowChart.FontName = .FontName And mFlowChart.Version < 6 Then
                mSelected.P.FontFace = "" 'use global font
            Else
                mSelected.P.FontFace = .FontName 'set the custom font name, version 6 absolute font
            End If
            mSelected.P.TextSize = .FontSize
            mSelected.P.TextBold = .FontBold
            mSelected.P.TextItalic = .FontItalic
            mSelected.P.TextUnderline = .FontUnderline
            If Not IsEditing Then
                Redraw
            Else
                SetEditFont mSelected
            End If
            'If gblnWindowFont Then frmFont.Update 'refreshes data in font toolbar
            'version changed
            UpgradeVersion 5 'a version 5.0 feature
            mFlowChart.Changed = True
        End If
    End With
End Sub

Private Sub mnuTextSpellAll_Click()
    SpellCheck True
End Sub

Private Sub mnuTextSpellThis_Click()
    SpellCheck False
End Sub

Private Sub mnuTextUnderline_Click()
    mSelected.P.TextUnderline = Not mnuTextUnderline.Checked
    If Not IsEditing Then
        Redraw
    Else
        SetEditFont mSelected
    End If
    'If gblnWindowFont Then frmFont.Update 'refreshes data in font toolbar
    mFlowChart.Changed = True
    UpgradeVersion 6
End Sub

Private Sub mnuToolsAutoFitText_Click()
    If Len(mSelected.P.Text) Then
        mUndo.Add mSelected, conUndoAutoSize
        If InStr(1, mSelected.P.Text, vbNewLine, vbBinaryCompare) = 0 And _
            mSelected.P.Height < picBay.TextHeight(mSelected.P.Text) * 2 Then
            'single line
            TextAutoSize mSelected, False, True
        Else
            'there is a return character
            TextAutoSize mSelected, True, True
        End If
        ConnectY mSelected, mUndo.LastItem.P, mSelected.P, False, True
        Redraw
    End If
End Sub

Private Sub mnuToolsDuplicate_Click()
    Dim objDup As FlowItem
    'Duplicates an object and adds it to the Flow Chart
    'collection.
    If IsMultipleSelected Then
        frmMultis.DoDuplicate
    ElseIf IsSelected Then
        Set objDup = Duplicate(mSelected)
        'objDup.P.Left = objDup.P.Left + objDup.P.Width / 2
        'objDup.P.Top = objDup.P.Top + objDup.P.Height / 2
        mFlowChart.Add objDup 'add it to the file
        Set mSelected = objDup
        mFlowChart.Changed = True
        SelectionChanged
        Redraw
        Set objDup = Nothing
    End If
End Sub

Private Sub mnuFile_Click()
    mnuFilePrint.Enabled = Not IsBusy
    Prompt
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    If QuerySave Then 'saved
        Set mFlowChart = New FlowChart
        mClipboard.ClearLink
        ResetForDraw
        ResetForNewOpen
        SetView
        Recaption
    End If
End Sub

Private Sub mnuFileOpen_Click()
    Dim objOpen As FileDlg
    
    On Error Resume Next
    If QuerySave Then 'saved
        Set objOpen = New FileDlg
        objOpen.Initialize hwnd, conDefaultExt, conDefaultFilter, mFlowChart.FileName, True, False, cdlOFNHideReadOnly
        Select Case objOpen.ShowOpen()
        Case CDERR_CANCELLED
            Exit Sub
        Case CDERR_OK
            OpenFile objOpen.FileName
        Case Else
            MsgBox "Unable to show Open dialog box.", vbExclamation
            Exit Sub
        End Select
    
'        With dlgFile
'            .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
'            .Filter = conDefaultFilter
'            .FilterIndex = 1
'            .ShowOpen
'            If Err = cdlCancel Then 'cancelled
'                Exit Sub
'            ElseIf Err <> 0 Then 'error
'                MsgBox "Unable to show Open dialog box.  (" & Err & ")", vbExclamation
'                .FileName = ""
'                .InitDir = ""
'            Else 'good
'                OpenFile .FileName
'            End If
'        End With
    End If
End Sub

Private Sub mnuFilePrint_Click()
    Dim lngLower As Long, lngUpper As Long
    
    On Error Resume Next
    Printer.KillDoc 'release printer object handle
    Load frmPrint
    
    mFlowChart.Layers.Requery
    mFlowChart.Layers.MinMaxLayer lngLower, lngUpper
    
    If frmPrint.Initialize(conPrintAll Or conPrintRange Or conPrintSelection, conPrintSelection, lngLower, lngUpper, lngLower, lngUpper) Then
        If frmPrint.ShowModal(Me) Then
            'restore previous orientation because that can ONLY be changed through Print Setup, not Print
            mFlowChart.Header1PDevName = Printer.DeviceName 'user okay this printer
            
            mFlowChart.PrintFile frmPrint.PrintRange, frmPrint.FromPage, frmPrint.ToPage, frmPrint.PrintOnSamePage
            mFlowChart.Changed = True
            SetView
        End If
    Else
        MsgBox "No printers are installed.", vbExclamation
    End If
    mFlowChart.Layers.CloseQuery
    Unload frmPrint
    Set frmPrint = Nothing
    
'    dlgFile.Flags = cdlPDNoPageNums Or cdlPDNoSelection
'    dlgFile.Orientation = IIf(mFlowChart.Orientation = vbPRORPortrait, cdlPortrait, cdlLandscape)
'    dlgFile.Min = 1 'ResetForDraw these values
'    dlgFile.Max = 1
'    Screen.MousePointer = vbHourglass
'    dlgFile.ShowPrinter
'    If Err = 0 Then
'        If gblnWindowMultis Then Unload frmMultis
'
'        Printer.TrackDefault = True 'restate this again
'        Printer.Copies = dlgFile.Copies
'        Printer.Orientation = mFlowChart.Orientation
'        SetView
'
'        mFlowChart.PrintFile Printer
'    ElseIf Err = cdlCancel Then
'        'ignore
'    Else
'        MsgBox "Problem showing the printer dialog box.", vbExclamation, "Print"
'    End If
'    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFilePrintSetup_Click()
    On Error Resume Next
    Printer.KillDoc 'release printer object handle
    Load frmPrint
    If frmPrint.InitializeSetup(mFlowChart.Orientation) Then
        If frmPrint.ShowModal(Me) Then
            mFlowChart.SaveSettingsFromPrinter
            mFlowChart.Changed = True
            SetView
        End If
    Else
        MsgBox "No printers are installed.", vbExclamation
    End If
    Unload frmPrint
    Set frmPrint = Nothing
    
'    dlgFile.Flags = cdlPDPrintSetup
'    dlgFile.Orientation = IIf(mFlowChart.Orientation = vbPRORPortrait, cdlPortrait, cdlLandscape)
'    dlgFile.Min = 1 'ResetForDraw these values
'    dlgFile.Max = 1
'    Screen.MousePointer = vbHourglass
'    dlgFile.ShowPrinter
'    If Err = 0 Then
'        Printer.TrackDefault = True 'restate this again
'        Printer.Copies = dlgFile.Copies
'        mFlowChart.Orientation = dlgFile.Orientation
'        Printer.Orientation = dlgFile.Orientation 'cdlPortrait == vbPRORPortrait
'        mFlowChart.PrintFile Printer, True
'        SetView 'resize the view area
'    ElseIf Err = cdlCancel Then
'        'ignore
'    Else
'        MsgBox "Problem showing the printer setup dialog box.", vbExclamation, "Print Setup"
'    End If
'    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFileSave_Click(Index As Integer)
    On Error Resume Next
    Dim lngError As Long
    Dim objSave  As FileDlg
        
    If Index = 1 Or Len(mFlowChart.FileName) = 0 Then
        Set objSave = New FileDlg
        objSave.Initialize hwnd, conDefaultExt, conDefaultFilter, mFlowChart.FileName, False, True, cdlOFNHideReadOnly
        Select Case objSave.ShowSave()
        Case CDERR_CANCELLED
            Exit Sub
        Case CDERR_OK
            mFlowChart.FileName = objSave.FileName
        Case Else
            MsgBox "Unable to show Save dialog box.", vbExclamation
            Exit Sub
        End Select
'        With dlgFile
'            .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
'            .Filter = conDefaultFilter
'            .ShowSave
'            If Err = cdlCancel Then 'cancelled
'                Exit Sub
'            ElseIf Err <> 0 Then 'error
'                MsgBox "Unable to show Save dialog box.", vbExclamation
'                Exit Sub
'            Else 'good
'                mFlowChart.Filename = .Filename
'            End If
'        End With
    End If
    If IsEditing Then
        'must turn off editing to save values
        EditBoxOff
    End If
    'for version 2 and greater
    With mFlowChart
        'for version 4 and greater
        If hsbBar.Max > 0 Then .ScrollX = hsbBar.Value / hsbBar.Max * conScrollParts Else .ScrollX = 0
        If vsbBar.Max > 0 Then .ScrollY = vsbBar.Value / vsbBar.Max * conScrollParts Else .ScrollY = 0
    End With
    
    Recaption
TryAgain:
    
    sbrBottom.SimpleText = "Saving..."
    sbrBottom.Style = sbrSimple

    'save routine
    lngError = mFlowChart.Save(mFlowChart.FileName)
    If lngError = conFail Then
        MsgBox "Problem saving the flow chart.  This file version cannot be saved.", vbExclamation
    ElseIf lngError Then 'return error messages
        Err.Raise lngError 'error must be replicated because Error$() does not return all error messages
        If MsgBox("Error saving the flow chart.  " & Err.Description & " (" & lngError & ")", vbExclamation Or vbRetryCancel, "Save Flow Chart") = vbRetry Then
            Err.Clear
            GoTo TryAgain 'no code is processed after this
        Else
            Err.Clear
        End If
'        mFlowChart.Changed = True 'so save will ask again
'    Else
'        mFlowChart.Changed = False
    End If

    sbrBottom.Style = sbrNormal
    sbrBottom.SimpleText = ""
End Sub


Private Sub mnuFormat_Click()
    mnuFormatAlignOne.Enabled = IsSelected And Not IsInMode
    mnuFormatAlignAll.Enabled = Not IsInMode
'    mnuFormatFont.Enabled = Not IsInMode
    mnuFormatGridShow.Enabled = Not IsInMode
    mnuFormatZoom.Enabled = Not IsInMode
    mnuFormatObjProp.Enabled = IsSelected And Not IsInMode
    
    Prompt
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub mnuHelpPrintInfo_Click()
    On Error GoTo Handler
    
    Load frmVal
    With Printer
        frmVal.AddItem "Number of copies:", .Copies
        frmVal.AddItem "Device name:", .DeviceName
        frmVal.AddItem "Driver name:", .DriverName
        frmVal.AddItem "Orientation:", Choose(.Orientation, "Portrait", "Landscape") & " (" & .Orientation & ")"
        frmVal.AddItem "Paper size:", .PaperSize
        frmVal.AddItem "Port:", .Port
        frmVal.AddItem "Print quality:", IIf(.PrintQuality < 0, Choose(-.PrintQuality, "Draft", "Low", "Medium", "High"), .PrintQuality & " dpi")
        frmVal.AddItem "Zoom:", IIf(.Zoom = 0, "100%", .Zoom)
        frmVal.AddItem "", ""
        frmVal.AddItem "FROM FILE", ""
        With frmVal.GetLastItem()
            .Font.Bold = True
            .Font.Underline = True
        End With
    End With
    With mFlowChart
        frmVal.AddItem "Printer Driver:", .Header1PDevName
        frmVal.AddItem "Orientation:", Choose(.Orientation, "Portrait", "Landscape") & " (" & .Orientation & ")"
        frmVal.AddItem "Paper Size:", .PaperSize
        frmVal.FitToSize
        frmVal.Caption = "Printer Properties"
        frmVal.Show vbModal, Me
    End With
    Exit Sub
Handler:
    MsgBox "Error accessing printer object.  " & Err.Description & " (" & Err.Number & ").", vbExclamation, "Printer Info"
    Unload frmVal 'release window
End Sub

Private Sub mnuEdit_Click()
    Dim blnTextMode As Boolean 'Cut,Copy,Paste
    
    mnuEditDel.Enabled = IsSelected And Not IsInMode 'only deletes selected ones
    mnuEditCancel.Enabled = IsSelected And IsEditing And Not (mblnMoving Or mblnSizing Or mblnAdding)
    mnuEditUndo.Enabled = Not IsInMode And Not (mUndo.Last Is Nothing)
    mnuEditEdit.Checked = IsEditing
    mnuEditSelAll.Enabled = Not IsBusy And mFlowChart.Count > 0
    
    If IsEditing And (Not mblnAdding Or mblnMoving Or mblnSizing) Then
        'For text
        mnuEditWhat.Caption = "For text"
        mnuEditCut.Enabled = txtEdit.SelLength
        mnuEditCopy.Enabled = txtEdit.SelLength
        mnuEditPaste.Enabled = mClipboard.IsText
        blnTextMode = True
    End If
    If IsSelected And Not IsInMode Then 'drawing order
        If Not blnTextMode Then
            'for mSelected
            If IsMultipleSelected Then
                mnuEditWhat.Caption = "For objects"
            Else
                mnuEditWhat.Caption = "For an object"
            End If
            mnuEditCut.Enabled = True
            mnuEditCopy.Enabled = True
            mnuEditPaste.Enabled = mClipboard.IsItem Xor mClipboard.IsSelection
        End If
        mnuEditToFront.Enabled = (mSelected.P.DrawOrder < conTop)
        mnuEditToBack.Enabled = (mSelected.P.DrawOrder > conBottom)
        mnuEditEdit.Enabled = mSelected.P.CanEdit
    Else
        If Not blnTextMode Then
            'for nothing
            mnuEditWhat.Caption = "No edit"
            mnuEditCut.Enabled = False
            mnuEditCopy.Enabled = False
            If IsBusy Then
                mnuEditPaste.Enabled = False
            Else
                mnuEditPaste.Enabled = mClipboard.IsItem Xor mClipboard.IsSelection
            End If
        End If
        mnuEditToFront.Enabled = False
        mnuEditToBack.Enabled = False
        mnuEditEdit.Enabled = blnTextMode
    End If
    
    Prompt
End Sub



Private Sub mnuEditCancel_Click()
    'ResetForDraw text box
    txtEdit.Visible = False
    RedrawHandles mSelected
End Sub

Private Sub mnuEditDel_Click()
    If gblnWindowMultis Then
        frmMultis.mnuToolsDelete_Click
    Else
        'undo data
        mUndo.Add mSelected, conUndoDelete
        'delete
        mFlowChart.RemoveObj mSelected
        Set mSelected = Nothing
        SelectionChanged
        Redraw
        RedrawHandles Nothing
    End If
End Sub

Private Sub mnuEditEdit_Click()
    If IsEditing Then
        EditBoxOff
    Else
        EditBoxOn
    End If
End Sub

Private Sub mnuFormatExport_Click()
'    Dim strFilename As String 'filename to retain for open/save
    Dim objItem     As FlowItem
    Dim intZoom     As Integer
    Dim objExport   As FileDlg
    
    On Error Resume Next
'    With dlgFile
'        strFilename = .Filename
'        .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
'        .Filter = "Bitmap Files (*.bmp)|*.bmp"
'        .DialogTitle = "Save Picture"
'        .Filename = ""
'        .ShowSave
        Set objExport = New FileDlg
        objExport.Title = "Export As Bitmap"
        objExport.Initialize hwnd, "bmp", "Bitmap Files (*.bmp)|*.bmp", "", False, True, cdlOFNHideReadOnly
        Select Case objExport.ShowSave()
        'If Err = 0 Then
        Case CDERR_OK
            Set picExport.Font = picBay.Font
            picExport.ForeColor = picBay.ForeColor
            picExport.BackColor = picBay.BackColor
            'set zoom
            intZoom = mFlowChart.ZoomPercent
            mFlowChart.ZoomPercent = 100
            
            If mFlowChart.PrinterError Then
                picExport.Move 0, 0, mFlowChart.PScaleWidth, mFlowChart.PScaleHeight
            Else
                picExport.Move 0, 0, Printer.ScaleWidth, Printer.ScaleHeight
            End If
            picExport.AutoRedraw = True
            
            If Err = 0 Then
                mFlowChart.DrawFile picExport, False, True, 0, 0, 0, 0
'                'bottom
'                For Each objItem In mFlowChart
'                    If objItem.P.DrawOrder = conBottom Then
'                        objItem.Draw picExport, mFlowChart
'                    End If
'                Next objItem
'
'                'middle
'                For Each objItem In mFlowChart
'                    If objItem.P.DrawOrder = conMiddle Then
'                        objItem.Draw picExport, mFlowChart
'                    End If
'                Next objItem
'
'                'top
'                For Each objItem In mFlowChart
'                    If objItem.P.DrawOrder = conTop Then
'                        objItem.Draw picExport, mFlowChart
'                    End If
'                Next objItem
            Else
                MsgBox "Error creating an image of the document.  " & Err.Description & " (" & Err.Number & ")", vbExclamation, "Save Picture"
                Err.Clear
            End If
            
            SavePicture picExport.Image, objExport.FileName
            
            picExport.Cls
            picExport.AutoRedraw = False
            picExport.Move 0, 0, 15, 15 'make it really small to save memory
            Set picExport.Font = New StdFont 'release picBay font
            mFlowChart.ZoomPercent = intZoom
'        ElseIf Err = cdlCancel Then
            'do nothing
        Case CDERR_CANCELLED
        Case Else
        'Else
            MsgBox "Error showing the Export dialog box.", vbExclamation
        End Select 'If
'        .DialogTitle = ""
'        .Filename = strFilename
'    End With
End Sub



Private Sub mnuToolsPic_Click()
    Load frmObject
    frmObject.fraPic2.Font.Bold = True
    mnuFormatObjProp_Click
End Sub

Private Sub mnuToolsRefreshSingle_Click()
    mSelected.Refresh mFlowChart, picBay
    picBay.Refresh
    If tlbFont.Visible Then UpdateFont
    UpdateBottom
End Sub





Private Sub mnuToolsSplit_Click()
    Dim sngCX As Single 'center x
    Dim sngCY As Single 'center y
    Dim tRect As Rect
    
    tRect = GetRect(mSelected)
    sngCX = mSelected.P.CenterX
    sngCY = mSelected.P.CenterY
    
    'add new one to NW
    mFlowChart.AddParam New FLine, tRect.X1, tRect.Y1, sngCX - tRect.X1, sngCY - tRect.Y1, ""
    'adjust orginal at SE
    mSelected.P.Left = sngCX
    mSelected.P.Top = sngCY
    mSelected.P.Width = tRect.X2 - sngCX
    mSelected.P.Height = tRect.Y2 - sngCY
    Redraw
End Sub

Private Sub mnuEditToBack_Click()
    mSelected.P.DrawOrder = mSelected.P.DrawOrder - 1
    Redraw
    mFlowChart.Changed = True
End Sub

Private Sub mnuEditToFront_Click()
    mSelected.P.DrawOrder = mSelected.P.DrawOrder + 1
    Redraw
    mFlowChart.Changed = True
End Sub


Private Sub mnuText_Click()
    Dim can      As Boolean
    Dim intFactor As Long
    
    If IsSelected And Not (mblnMoving Or mblnSizing Or mblnAdding) Then  'not busy, but editing's ok
        If IsSelLine Or IsWithoutText Then
            can = False
        Else
            can = True
        End If
    End If
    
    If IsSelected Then
        mnuTextLeft.Checked = (mSelected.P.TextAlign = DT_LEFT)
        mnuTextCentre.Checked = (mSelected.P.TextAlign = DT_CENTER)
        mnuTextRight.Checked = (mSelected.P.TextAlign = DT_RIGHT)

        mnuTextBold.Checked = mSelected.P.TextBold 'GetTextFormatBold(mSelected) Or mSelected.P.TextBold
        mnuTextItalic.Checked = mSelected.P.TextItalic 'GetTextFormatItalic(mSelected) Or mSelected.P.TextItalic
        mnuTextUnderline.Checked = mSelected.P.TextUnderline
        
    Else
        mnuTextLeft.Checked = False
        mnuTextCentre.Checked = False
        mnuTextRight.Checked = False
    End If
    
    mnuTextLeft.Enabled = can
    mnuTextCentre.Enabled = can
    mnuTextRight.Enabled = can
    mnuTextBold.Enabled = can
    mnuTextItalic.Enabled = can
    mnuTextUnderline.Enabled = can
    mnuTextFont.Enabled = can
    'mnuTextSize.Enabled = can
    mnuTextSpell.Enabled = Not IsBusy And Not IsMultipleSelected
    mnuTextSpellThis.Enabled = Not IsMultipleSelected And IsSelected And Not IsBusy And mFlowChart.Count > 0
    
    Prompt "Font changes may be done using the Font Bar."
End Sub



Private Sub mnuTextBold_Click()
'    If Len(mSelected.Tag1) Then 'old style
'        SetTextFormat Obj:=mSelected, Bold:=Not mnuTextBold.Checked
'        UpgradeVersion 4
'    Else 'new style used
        mSelected.P.TextBold = Not mnuTextBold.Checked
'        UpgradeVersion 6
'    End If
    
    If Not IsEditing Then
        Redraw
    Else
        SetEditFont mSelected
    End If
    'If gblnWindowFont Then frmFont.Update 'refreshes data in font toolbar
    mFlowChart.Changed = True
End Sub

Private Sub mnuTextCentre_Click()
    'SetTextFlags mSelected, DT_CENTER
    mSelected.P.TextAlign = DT_CENTER
    mlngDefaultTextFlags = DT_CENTER
    UpgradeVersion 3
    If Not IsEditing Then
        Redraw
    End If
    'If gblnWindowFont Then frmFont.Update 'refreshes data in font toolbar
    mFlowChart.Changed = True
End Sub

Private Sub mnuTextItalic_Click()
'    If Len(mSelected.Tag1) Then 'old sytle
'        SetTextFormat Obj:=mSelected, Italic:=Not mnuTextItalic.Checked
'        UpgradeVersion 4
'    Else
        mSelected.P.TextItalic = Not mnuTextItalic.Checked
'        UpgradeVersion 6
'    End If
    If Not IsEditing Then
        Redraw
    Else
        SetEditFont mSelected
    End If
    'If gblnWindowFont Then frmFont.Update 'refreshes data in font toolbar
    mFlowChart.Changed = True
End Sub

Private Sub mnuTextLeft_Click()
    'SetTextFlags mSelected, DT_LEFT
    mSelected.P.TextAlign = DT_LEFT
    mlngDefaultTextFlags = DT_LEFT
    If Not IsEditing Then
        Redraw
    End If
    'If gblnWindowFont Then frmFont.Update 'refreshes data in font toolbar
    mFlowChart.Changed = True
End Sub



Private Sub mnuTextRight_Click()
    'SetTextFlags mSelected, DT_RIGHT
    mSelected.P.TextAlign = DT_RIGHT
    mlngDefaultTextFlags = DT_RIGHT
    UpgradeVersion 3
    If Not IsEditing Then
        Redraw
    End If
    'If gblnWindowFont Then frmFont.Update 'refreshes data in font toolbar
    mFlowChart.Changed = True
End Sub




Private Sub mnuTools_Click()
    'Dim objPic As FPicture
    If IsSelected And Not IsInMode Then  'drawing order
        mnuToolsRefreshSingle.Enabled = True
        mnuToolsAutoFitText.Enabled = Not (IsWithoutText Or IsSelLine Or (TypeOf mSelected Is FDecision))
        mnuToolsDuplicate.Enabled = Not IsMultipleSelected
        mnuToolsSplit.Enabled = IsSelLine And Not IsMultipleSelected
        If TypeOf mSelected Is FPicture Then
            'Set objPic = mSelected
            'mnuToolsReset.Enabled = objPic.IsDefaultSize(picBay)
            'mnuToolsPicZoom.Enabled = True
            mnuToolsPic.Enabled = True
            'Set objPic = Nothing
        Else
            'mnuToolsPicZoom.Enabled = False
            'mnuToolsReset.Enabled = False
            mnuToolsPic.Enabled = False
        End If

    Else
        mnuToolsRefreshSingle.Enabled = False
        mnuToolsAutoFitText.Enabled = False
        mnuToolsDuplicate.Enabled = False
        mnuToolsSplit.Enabled = False
        'mnuToolsReset.Enabled = False
        mnuToolsPic.Enabled = False
        'mnuToolsPicZoom.Enabled = False

    End If
    
    Prompt "To change unlisted functions, use Format > Object."
End Sub

Private Sub mnuFormatAlignAll_Click()
    Dim objItem As FlowItem
    
    Screen.MousePointer = vbHourglass
    For Each objItem In mFlowChart
        Align objItem
    Next objItem
    'and update
    Redraw
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFormatAlignOne_Click()
    If gblnWindowMultis Then
        If frmMain.Visible Then
            frmMultis.DoAlign 'calls back this form
        Else
            GoTo Skip
        End If
    Else
Skip:
        mUndo.Add mSelected, conUndoAlignment
        Align mSelected
    End If
    'and update
    Redraw
End Sub


Private Sub mnuFormatGridSnap_Click()
    mnuFormatGridSnap.Checked = Not mnuFormatGridSnap.Checked
End Sub



Private Sub mnuFormatGridShow_Click()
    mnuFormatGridShow.Checked = Not mnuFormatGridShow.Checked
    RedrawGrid
    Redraw
End Sub







Private Sub mnuFormatZoom_Click()
    Load frmZoom
    frmZoom.Zoom = mFlowChart.ZoomPercent
    frmZoom.Show 1, Me
    If frmZoom.Changed Then
        RedrawHandles Nothing
        SetView frmZoom.Zoom
        'version changed
        UpgradeVersion 2
        mFlowChart.Changed = True
    End If
    Unload frmZoom
End Sub

Private Sub mnuWindow_Click()
    mnuWindowFontBar.Checked = tlbFont.Visible
End Sub

Private Sub mnuWindowFontBar_Click()
    tlbFont.Visible = Not tlbFont.Visible
    If tlbFont.Visible And cboSize.ListCount = 0 Then
        UpgradeVersion 5
        LoadFont
    ElseIf tlbFont.Visible Then
        mblnFontBarBusy = False
        LoadFont
    End If
    Form_Resize
End Sub



Private Sub mnuWindowLayer_Click()
    UpgradeVersion 8
    frmLayer.LoadData mFlowChart
    frmLayer.Show vbModeless, Me
End Sub

Private Sub mnuWindowMultis_Click()
    Load frmMultis
    If Not frmMultis.Visible Then
        frmMultis.Show vbModeless, Me
    End If
End Sub

Private Sub mnuWindowShift_Click()
    frmShift.Show vbModeless, Me
End Sub

'Private Sub mnuWindowColourBox_Click()
'    UpgradeVersion 6
'    frmColour.Show 0, Me
'    frmColour.Update
'End Sub

'Private Sub mnuWindowFontBox_Click()
'    On Error GoTo Handler
'    UpgradeVersion 5
'    frmFont.Show 0, Me
'    frmFont.Update
'Handler:
'End Sub


Private Sub mUndo_UndoItemChanged(Item As IUndoInterface)
    Dim objSingle As PUndoItem
    
    If Item Is Nothing Then
        mnuEditUndo.Caption = "&Undo"
    Else
        If Item.IsMultiple Then
            mnuEditUndo.Caption = "&Undo Multiple"
        Else
            Set objSingle = Item
            mnuEditUndo.Caption = "&Undo " & objSingle.theActionName
        End If
    End If
End Sub

Private Sub picBay_DblClick()
    EditBoxOn
End Sub





Private Sub picBay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPrevious As FlowItem
    Dim objHit      As FlowItem

    mblnHitItem = False
    
    If (Button And vbLeftButton) = vbLeftButton Then
        mblnMouseDown = True
        'ResetForDraw some variables
        mblnMouseMovedGrid = False
        mblnMouseMovedSensitive = False
        mblnMovedSized = False
        'Selection
        If Not IsDrawing Then  'excludes editing
            'if user presses Ctrl, hold onto the selected item
            If (Shift And vbCtrlMask) = vbCtrlMask Then Set objPrevious = mSelected
            'get the item that was clicked
            Set objHit = Test(X, Y, False)
            If Not (objHit Is mSelected) Or (objHit Is Nothing) Then
                mblnHitItem = True
                Set mSelected = objHit
            End If
            If IsSelected Then 'if it is still there then
                If (Shift And vbCtrlMask) = vbCtrlMask Then 'multiple selection
                    'CTRL key, clicked on another item or group
                    mnuWindowMultis_Click
                    frmMultis.SwitchItemGroup mSelected
                    If Not (objPrevious Is objHit) And Not (objPrevious Is Nothing) Then frmMultis.HoldItem objPrevious
                Else
                    If mSelected.P.GroupNo = 0 Then
                        'clicked on ungrouped object
                        SelectionChanged
                    Else
                        'clicked on grouped object
                        mnuWindowMultis_Click
                        If frmMultis.GetIndex(mSelected) = 0 Then frmMultis.Clear 'clear group
                        frmMultis.AddGroup mSelected.P.GroupNo 'add this to group
                    End If
                End If
            Else 'user clicked on white space, clear all selections
                SelectionChanged
            End If
        End If
        
        Set objPrevious = Nothing
        Set objHit = Nothing
        
        'Adding
        If Not IsBusy And mblnAdding Then
            If gblnWindowMultis Then Unload frmMultis
            'adding & sizing code
            'Selected item is imaginary - doesn't exist on paper
            If mlngAddType = conAddEndArrowLine Or mlngAddType = conAddLine Or mlngAddType = conAddMidArrowLine Then
                Set mSelected = New FLine 'use line outline
            Else
                Set mSelected = New FlowItem 'use shape outline
            End If
            With mSelected.P 'position the added item
                .Left = Grid(X)
                .Top = Grid(Y)
                .Width = 0
                .Height = 0
            End With
            
            ctlBox_MouseDown conSE, vbLeftButton, 0, ctlBox(conSE).Width / 2, ctlBox(conSE).Width / 2     'start sizing

        'Selecting
        ElseIf Not IsSelected And Not IsBusy Then
            mblnSelecting = True
            mblnSelected = False
            shpMove.Left = X
            shpMove.Top = Y
            shpMove.Width = 0
            shpMove.Height = 0
            mMousePos.X1 = X
            mMousePos.Y1 = Y
        'Moving
        ElseIf IsSelected And Not IsBusy Then
            'moving code
            mMousePos.X1 = X 'set mouse position
            mMousePos.Y1 = Y
            mblnMoving = True 'set flag
            If IsSelLine Then
                linSize.X1 = mSelected.P.Left
                linSize.Y1 = mSelected.P.Top
                linSize.X2 = mSelected.P.Left + mSelected.P.Width
                linSize.Y2 = mSelected.P.Top + mSelected.P.Height
            Else
                shpMove.Left = mSelected.P.Left
                shpMove.Top = mSelected.P.Top
                shpMove.Width = mSelected.P.Width
                shpMove.Height = mSelected.P.Height
            End If
        'Edit off
        ElseIf IsSelected And IsEditing Then
            EditBoxOff
            RedrawHandles mSelected
        End If
    ElseIf (Button And vbRightButton) Then
        'Item Selection
        'IF selected one is right-click, keep it
        'ELSE find the next one to hit and select
        Set objPrevious = Test(X, Y, False)
        If Not objPrevious Is mSelected Then
            Set mSelected = objPrevious
        End If
    End If
End Sub


Private Sub picBay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnMouseDown Then
        mMousePos.X2 = X
        mMousePos.Y2 = Y
                
        If Abs(mMousePos.X2 - mMousePos.X1) >= msngGrid / 2 Or _
           Abs(mMousePos.Y2 - mMousePos.Y1) >= msngGrid / 2 Then
            mblnMouseMovedSensitive = True
        End If
        If Abs(mMousePos.X2 - mMousePos.X1) > msngGrid Or _
           Abs(mMousePos.Y2 - mMousePos.Y1) > msngGrid Then
            mblnMouseMovedGrid = True
        End If
    End If
    
    If mblnMoving Then
        If mblnMouseMovedSensitive Or (Shift = vbShiftMask) Then
            If Not mblnMovedSized Then  'if not auto redraw
                mblnMovedSized = True
                RedrawHandles Nothing
                AutoRedrawOn    'auto redraw
                picBayPaint     'and draw
                picBay.MousePointer = vbCustom
                If gblnWindowMultis Then frmMultis.BeginMove
            End If
            picBayMoving
            If IsSelLine Then
                linSize.X1 = shpMove.Left
                linSize.Y1 = shpMove.Top
                linSize.X2 = shpMove.Left + mSelected.P.Width
                linSize.Y2 = shpMove.Top + mSelected.P.Height
                linSize.Visible = True
            Else
                shpMove.Visible = True
            End If
        End If
    ElseIf mblnAdding And IsSelected Then
        'must go first to override mblnSizing
        ctlBox_MouseMove conSE, vbLeftButton, 0, X - mSelected.P.Left + ctlBox(conSE).Width / 2, Y - mSelected.P.Top + ctlBox(conSE).Width / 2
    ElseIf mblnSizing Then
        'mouse lost, end sizing code
        mblnSizing = False
        linSize.Visible = False 'hide shapes
        shpMove.Visible = False
        RedrawHandles mSelected 'order of these is important
        AutoRedrawOff
    ElseIf mblnSelecting Then
        If mblnMouseMovedGrid And Not mblnSelected Then
            mblnSelected = True
            shpMove.Visible = True
            picBay.MousePointer = vbCrosshair

            AutoRedrawOn    'auto redraw
            If picBay.AutoRedraw Then picBayPaint     'and draw
        End If
        If mblnSelected Then
            ctlBoxRectToShape shpMove, mMousePos
        End If
    End If
End Sub

Private Sub picBay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPrevious As FlowItem
    Dim objTemp     As FlowItem
        
    If mblnSelecting Then
        'get the item that was clicked
        Set objPrevious = Test(X, Y, True)
        If Not objPrevious Is mSelected Then
            SetSelected objPrevious
        End If
    End If
    
    If mblnMoving Then
        mMousePos.X2 = X
        mMousePos.Y2 = Y
        mblnMoving = False
        shpMove.Visible = False
        linSize.Visible = False
        If mblnMovedSized Then
            mblnMovedSized = False
            picBayMoving
            
            If (Shift And vbCtrlMask) = vbCtrlMask And mblnMouseMovedSensitive And gblnWindowMultis Then
                If frmMultis.mCollection.Count = 1 Then 'only one item selected, duplicate it
                    Unload frmMultis
                    Set objTemp = Duplicate(mSelected)  'create another of the same object
                    mFlowChart.Add objTemp              'and add to flow chart
                Else
                    Set objTemp = mSelected
                End If
            Else
                Set objTemp = mSelected
            End If
            
            'save the old values of mSelected, if that is what is being moved
            If mSelected Is objTemp Then
                If Not gblnWindowMultis Then
                    mUndo.Add mSelected, conUndoMove
                End If
            End If
            
            With objTemp.P
                'then move the shape
                .Left = shpMove.Left
                .Top = shpMove.Top
            End With
            
            AutoRedrawOff
            Set mSelected = objTemp 'redraws the handles
            mFlowChart.Changed = True
            
            If Not ((Shift And vbCtrlMask) = vbCtrlMask) Then 'not duplication, multiple move
                If gblnWindowMultis Then frmMultis.EndMove False
            Else 'duplication, cancel multiple move
                If gblnWindowMultis Then frmMultis.EndMove True  'cancel the move op
            End If
        Else
            If gblnWindowMultis Then frmMultis.EndMove True
            If mSelected.Number = conAddButton Then 'button clicked
                GetButton(mSelected).Click picBay, mFlowChart
            End If
        End If
    ElseIf mblnAdding And IsSelected Then
        If (Button And vbRightButton) = vbRightButton Then
            mblnAdding = False
        Else
            ctlBox_MouseUp conSE, vbLeftButton, 0, X - mSelected.P.Left + ctlBox(conSE).Width / 2, Y - mSelected.P.Top + ctlBox(conSE).Width / 2
            
            If picBayAdd(mSelected) Then  'add item
                If mblnMovedSized Then mblnMovedSized = False
            Else
                'resume adding mode
                ctlBox_MouseDown conSE, vbLeftButton, 0, ctlBox(conSE).Width / 2, ctlBox(conSE).Width / 2     'start sizing
                Exit Sub
            End If
        End If
    ElseIf mblnSelecting Then
        mMousePos.X2 = X
        mMousePos.Y2 = Y
        If mMousePos.X2 < mMousePos.X1 Then
            mMousePos.X2 = mMousePos.X1
            mMousePos.X1 = X
        End If
        If mMousePos.Y2 < mMousePos.Y1 Then
            mMousePos.Y2 = mMousePos.Y1
            mMousePos.Y1 = Y
        End If
        shpMove.Visible = False
        mblnSelecting = False
        If mblnSelected Then
            picBaySelect mMousePos
        End If
        If picBay.AutoRedraw Then AutoRedrawOff
    ElseIf (Button And vbRightButton) Then
        If mblnAdding Then
            mblnAdding = False
        Else
            mClipboard.TempX = X 'allows pasting of objects at absolute locations
            mClipboard.TempY = Y
            If IsSelected Then
                If Shift = vbCtrlMask Then
                    PopupMenu mnuText, vbPopupMenuLeftButton Or vbPopupMenuRightButton
                Else
                    PopupMenu mnuPopup, vbPopupMenuLeftButton Or vbPopupMenuRightButton
                End If
            Else
                PopupMenu mnuPopup, vbPopupMenuLeftButton Or vbPopupMenuRightButton
            End If
            'since PopupMenu holds up the entry code,
            'this code will be processed after
            'the popup menu disappears
            mClipboard.TempClear
        End If
    ElseIf gblnWindowMultis And Not IsSelected Then
        Unload frmMultis
    End If

    mblnMouseDown = False
End Sub

Private Sub picBay_Paint()
    picBayPaint
End Sub







Private Sub picSelection_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngParts As Long
    Dim lngItem  As Long
    Dim objSel   As FlowItem
    Const conSelectA = "Select a point to reveal layers."
    
    If mcolSelections Is Nothing Then
        sbrBottom.SimpleText = conSelectA
        sbrBottom.Style = sbrSimple
        Exit Sub
    End If
    
    'allow user to change selection of object on side bar
    If mcolSelections.Count > 0 Then
        lngParts = picSelection.ScaleHeight / mcolSelections.Count
        lngItem = Y \ lngParts + 1
        
        If 0 < lngItem And lngItem <= mcolSelections.Count Then
            Set objSel = mcolSelections(lngItem)
        Else 'item out of range - skip
            Exit Sub
        End If
    End If
    
    If (Button And vbLeftButton) = vbLeftButton Then
        Set mSelected = objSel
        If IsSelected Then
            If mSelected.P.GroupNo Then
                mnuWindowMultis_Click
                If frmMultis.GetIndex(mSelected) = 0 Then frmMultis.Clear
                frmMultis.AddGroup mSelected.P.GroupNo
            Else
                SelectionChanged
            End If
        Else
            sbrBottom.SimpleText = conSelectA
            sbrBottom.Style = sbrSimple
        End If
    ElseIf (Button And vbRightButton) = vbRightButton Then
        Dim strCaption As String
    
        strCaption = "Caption of " & objSel.Description
        If objSel.P.GroupNo Then
            strCaption = strCaption & " in group " & objSel.P.GroupNo
        End If
        strCaption = strCaption & " : " & objSel.P.Text
        sbrBottom.SimpleText = strCaption
        sbrBottom.Style = sbrSimple
    End If
End Sub

Private Sub picSelection_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Or sbrBottom.Style = sbrSimple Then
        sbrBottom.Style = sbrNormal
        sbrBottom.SimpleText = ""
    End If

End Sub


Private Sub picSelection_Paint()
    Dim lngParts As Long
    Dim lngItem  As Long
    
    picSelection.Cls
    
    If mcolSelections Is Nothing Then Exit Sub
    
    If mcolSelections.Count > 0 Then
        lngParts = picSelection.ScaleHeight / mcolSelections.Count
        For lngItem = 1 To mcolSelections.Count
            If mcolSelections(lngItem) Is mSelected Then
                'highlight box
                picSelection.Line (0, lngParts * (lngItem - 1))-Step(picSelection.ScaleWidth, lngParts), vbHighlight, BF
                picSelection.ForeColor = vbHighlightText
            Else 'line
                picSelection.Line (0, lngParts * (lngItem - 1))-Step(picSelection.ScaleWidth, 0), vbButtonText
                picSelection.ForeColor = vbButtonText
            End If
            'number this section
            picSelection.CurrentX = 0: picSelection.CurrentY = lngParts * (lngItem - 1)
            picSelection.Print lngItem
        Next lngItem
    End If
End Sub


Private Sub timMsg_Timer()
    sbrBottom.Style = sbrNormal
    sbrBottom.SimpleText = ""
End Sub

Private Sub tlbColour_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objForm As IToolbar
    
    If Not IsSelected Then
        Prompt "Select an object before using this."
        Exit Sub
    End If
    
    Select Case Button.Index
    Case 1
        Set objForm = frmCPick
        frmCPick.SetColourPart conPartText
'        On Error Resume Next
'        With dlgFile
'            .CancelError = True
'            .Color = mSelected.P.TextColour
'            .Flags = cdlCCFullOpen Or cdlCCRGBInit
'            .ShowColor
'            If Err = 0 Then
'                mUndo.Add mSelected, conUndoTextColourChange
'                mSelected.P.TextColour = .Color
'                ToolBoxDone
'            End If
'        End With
    
    Case 2
        Set objForm = frmCPick
        frmCPick.SetColourPart conPartFore
        '        On Error Resume Next
'        With dlgFile
'            .CancelError = True
'            .Color = mSelected.P.ForeColour
'            .Flags = cdlCCFullOpen Or cdlCCRGBInit
'            .ShowColor
'            If Err = 0 Then
'                If IsSelected Then
'                    mUndo.Add mSelected, conUndoForeColourChange
'                    mSelected.P.ForeColour = .Color
'                    ToolBoxDone
'                End If
'            End If
'        End With
        
    Case 3
        Set objForm = frmStyleThickness
    Case 4
        Set objForm = frmStyleLine
    Case 5
        Set objForm = frmCPick
        frmCPick.SetColourPart conPartBack
'        On Error Resume Next
'        With dlgFile
'            .CancelError = True
'            .Color = mSelected.P.BackColour
'            .Flags = cdlCCFullOpen Or cdlCCRGBInit
'            .ShowColor
'            If Err = 0 Then
'                If IsSelected Then
'                    mUndo.Add mSelected, conUndoBackColourChange
'                    'change fill style automatically
'                    If mSelected.P.FillStyle = vbFSTransparent Then
'                        mSelected.P.FillStyle = vbFSSolid
'                    End If
'                    mSelected.P.BackColour = .Color
'                    ToolBoxDone
'                End If
'            End If
'        End With
    
    Case 6
        Set objForm = frmStyleBack
    
    Case 8 'properties
        mnuFormatObjProp_Click
    End Select
    
    If Not objForm Is Nothing Then
        objForm.Move Left + tlbColour.Left + picSelection.Width, Top + Height - tlbColour.Height - objForm.Height - sbrBottom.Height
        objForm.UpdateToForm mSelected, mFlowChart
        objForm.Show vbModeless, Me
    End If
End Sub

Private Sub tlbDraw_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 And Not IsBusy Then
        mblnAdding = False
        mlngAddType = 0
        tlbDrawButtonsOff
        RedrawHandles mSelected, True
    ElseIf Not IsBusy Then
        'go ahead and go to add mode
        mlngAddType = Button.Index - 1 '6 and 7 are lines
        Set mSelected = Nothing
        SelectionChanged
        mblnAdding = True
    ElseIf IsEditing And Not (mblnMoving Or mblnSizing) Then
        EditBoxOff
        RedrawHandles mSelected, True
        mlngAddType = Button.Index - 1 '6, 7, 8 are lines
        mblnAdding = True
    Else
        'disallow adding of items on busy
        Button.Value = tbrUnpressed
        Beep
    End If
End Sub





Private Sub tlbFont_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Not IsSelected Then Exit Sub

    With mSelected.P
        .TextBold = tlbFont.Buttons!Bold.Value 'chkBold
        .TextItalic = tlbFont.Buttons!Italic.Value 'chkItalic
        .TextUnderline = tlbFont.Buttons!Underline.Value 'chkUnderline
        .TextAlign = IIf(tlbFont.Buttons!Left.Value, DT_LEFT, 0) Or IIf(tlbFont.Buttons!Centre.Value, DT_CENTER, 0) Or IIf(tlbFont.Buttons!Right.Value, DT_RIGHT, 0)
    End With
    If gblnWindowMultis Then
        frmMultis.DoCopyFont False, tlbFont.Buttons!Left.Index <= Button.Index And Button.Index <= tlbFont.Buttons!Right.Index, tlbFont.Buttons!Bold.Index = Button.Index, tlbFont.Buttons!Italic.Index = Button.Index, tlbFont.Buttons!Underline.Index = Button.Index, False
    Else
        RedrawSingle mSelected
    End If
End Sub



Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    'The two columns of keys below Scroll Lock and Pause/Break
    'end editing.
    Select Case KeyCode
        Case vbKeyF11, vbKeyF12
            KeyCode = 0
            EditBoxOff
            RedrawHandles mSelected
        Case vbKeyTab 'cannot display
            KeyCode = 0
    End Select
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And mFlowChart.Version < 4 Then 'before V4, cannot return
        KeyAscii = 0
        EditBoxOff
        RedrawHandles mSelected
    ElseIf KeyAscii = vbKeyTab Then
        KeyAscii = 0 'cannot tab
    End If
End Sub


Private Sub vsbBar_Change()
    picBay.Top = -ScaleY(vsbBar.Value, conScrollScale, vbTwips)
End Sub


Public Sub ToolBoxChanged(UndoAction As PUndoAction, Optional Multiple As Boolean)
    If UndoAction = conUndoMultiple Or Multiple Then
        mUndo.AddMany frmMultis.mCollection, UndoAction
    Else
        mUndo.Add mSelected, UndoAction
    End If
    mFlowChart.Changed = True
End Sub

