VERSION 5.00
Begin VB.Form frmProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "RichPad_Prop.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboFormat 
      Height          =   315
      ItemData        =   "RichPad_Prop.frx":000C
      Left            =   1200
      List            =   "RichPad_Prop.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CheckBox chkLogFile 
      Caption         =   "Log File"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CheckBox chkWebFile 
      Caption         =   "Web File"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdFolder 
      Caption         =   "Open Containing Folder"
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CheckBox chkNotViewer 
      Caption         =   "Modifiable"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CheckBox chkSaved 
      Caption         =   "Saved"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtFormat 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtSize 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label l4 
      Caption         =   "Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label l3 
      Caption         =   "Format:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label l2 
      Caption         =   "Size:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label l1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 26 Oct 2001

Public CboFormatIndex As Integer

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cboFormat_Click()
    CboFormatIndex = cboFormat.ListIndex
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub




Private Sub cmdFolder_Click()
    Dim r As Long
    
    MousePointer = vbHourglass
    r = ShellExecute(hwnd, vbNullString, CStr(txtPath.Text) & vbNullChar, vbNullString, vbNullString, vbNormalFocus)
    If r <= 32 Then 'not successful ShellExecute
        MsgBox "Problem opening folder.", vbExclamation, "Properties"
        cmdFolder.Enabled = False
    End If
    MousePointer = vbDefault
End Sub

Private Sub cmdOpen_Click()
    Dim r As Long
    
    MousePointer = vbHourglass
    r = ShellExecute(hwnd, vbNullString, CStr(txtPath.Text & txtFile.Text) & vbNullChar, vbNullString, vbNullString, vbNormalFocus)
    If r <= 32 Then 'not successful ShellExecute
        MsgBox "Problem opening file.", vbExclamation, "Properties"
        cmdOpen.Enabled = False
    End If
    MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    CboFormatIndex = 0
End Sub

Private Sub txtFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        txtFile.SelStart = 0
        txtFile.SelLength = Len(txtFile)
    End If
End Sub


Private Sub txtFormat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        txtFormat.SelStart = 0
        txtFormat.SelLength = Len(txtFormat)
    End If
End Sub


Private Sub txtPath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        txtPath.SelStart = 0
        txtPath.SelLength = Len(txtPath)
    End If
End Sub


Private Sub txtSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        txtSize.SelStart = 0
        txtSize.SelLength = Len(txtSize)
    End If
End Sub


