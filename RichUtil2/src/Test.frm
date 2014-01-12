VERSION 5.00
Object = "{0385C402-FCB3-41CE-A18E-84CA98F7078B}#7.3#0"; "RichUtil2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Toggle Check"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Toggle TypeToSearch"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "Test.frx":0000
      Left            =   120
      List            =   "Test.frx":000D
      TabIndex        =   8
      Top             =   3240
      Width           =   4215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Toggle Enabled"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Remove 2 to 6"
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Font"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sign control"
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add column"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin RichUtil2.EListBox EListBox1 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5106
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LargeChange     =   13
      RowDivisions    =   -1  'True
      ColumnDivisions =   -1  'True
      CheckBoxes      =   -1  'True
      EditMaxLength   =   5
      TokenChar       =   "#"
      Columns         =   4
      BeginProperty Column {D12F0A11-F007-4652-AA4B-77531B860774} 
         Count           =   4
         I2              =   20
      EndProperty
      ListCount       =   4
      BeginProperty List {7C24CA0C-2284-437D-8D20-9B3A455154D8} 
         Count           =   4
         I1              =   "Column1#Column2#Column3#column4"
         I2              =   "Richard###"
         I3              =   "List###"
         I4              =   "Box###"
      EndProperty
      BeginProperty ItemData {85727F3B-550A-4946-857F-8BECE14841A5} 
         Count           =   4
         I2              =   1
         I3              =   2
         I4              =   3
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   4560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3240
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill list"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private Sub Command2_Click()
    EListBox1.Columns = EListBox1.Columns + 1
End Sub

Private Sub Command3_Click()
    'added May 23, 2004
    EListBox1.OcxHandshake &H1234
    EListBox1.OcxHandshake &H5678
    EListBox1.OcxHandshake &H2832
    EListBox1.OcxHandshake &H2600
End Sub

Private Sub Command4_Click()
    EListBox1.Font.Name = "Arial"
    EListBox1.Font.Size = 12
    EListBox1.Refresh
End Sub

Private Sub Command5_Click()
    EListBox1.RemoveRows 2, 4
End Sub

Private Sub Command6_Click()
    EListBox1.Enabled = Not EListBox1.Enabled
    List1.Enabled = EListBox1.Enabled
End Sub

Private Sub Command7_Click()
    EListBox1.TypeToSearch = Not EListBox1.TypeToSearch
End Sub

Private Sub Command8_Click()
    EListBox1.Selected(EListBox1.ListIndex) = Not EListBox1.Selected(EListBox1.ListIndex)
End Sub

Private Sub EListBox1_BeforeEdit(Row As Long, Col As Long, Cancel As Boolean)
    EListBox1.Locked = Col = 2
End Sub

Private Sub EListBox1_CornerMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "CornerMouseDown "; X, Y
End Sub

Private Sub EListBox1_ItemChecked(ByVal ListIndex As Long, ByVal Checked As Boolean, Cancel As Boolean)
    Debug.Print "ItemChecked "; ListIndex; " to state "; Checked
End Sub

Private Sub EListBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        EListBox1.RemoveItem EListBox1.ListIndex
    End If

End Sub

Private Sub EListBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "MouseDown "; X, Y
End Sub

Private Sub EListBox1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then Debug.Print "MouseMove "; X, Y
End Sub

Private Sub EListBox1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "MouseUp "; X, Y
End Sub

Private Sub EListBox1_RowColClick(ByVal Row As Long, ByVal Col As Long)
    Debug.Print "RowColClick "; Row, Col
    
End Sub

Private Sub EListBox1_RowColDblClick(ByVal Row As Long, ByVal Col As Long)
    Debug.Print "RowColDblClick "; Row, Col
    
    EListBox1.EditBoxOn Row, Col
End Sub

Private Sub EListBox1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, ByVal EditText As String, Cancel As Boolean)
    If Len(EditText) = 0 Then Cancel = True
End Sub


Private Sub Command1_Click()
    EListBox1.AddItem "Hi"
    EListBox1.AddItem3 "Bye" & vbTab & "Wow" & vbTab & "Better"
    EListBox1.AddItem2 "C1", 0, "C2", "Column 3"
    EListBox1.ReadyForResize 4000
    
    Dim intI As Integer
    Dim lngJ As Long
    Dim lngK As Long
    Dim lngT1 As Single
    Dim lngT2 As Single
    
    lngT1 = Timer
    
    EListBox1.AddMode = True
    'EListBox1.ReadyForResize 1000 * 26
    For lngJ = 1 To 4000
        'For intI = Asc("A") To Asc("Z")
            lngK = lngK + 1
            EListBox1.AddItem2 String(Int(Rnd * 20) + 1, Chr(intI)), 0, lngJ, lngK, Asc(LCase(Chr(intI)))
        'Next intI
    Next lngJ
    EListBox1.AddMode = False
    lngT2 = Timer
    
    Debug.Print lngT2 - lngT1
    'lst2.AddItem "Hi"
End Sub


Private Sub Form_Load()
    EListBox1.AddItem3 "Col A#Col B#Col C#Col D", "#"
End Sub

Private Sub Form_Paint()
    Dim tRect As RECT
    
    tRect.Right = 10
    tRect.Bottom = 10
    DrawFocusRect hdc, tRect
End Sub


Private Sub Form_Resize()
    EListBox1.Top = 0
    EListBox1.Height = ScaleHeight
End Sub

