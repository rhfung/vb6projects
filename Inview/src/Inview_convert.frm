VERSION 5.00
Object = "{B0F710F8-474B-4DE4-BE20-2C604C6801EE}#9.2#0"; "RICHUTIL1.OCX"
Begin VB.Form frmConvert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Convert Data"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Inview_convert.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtR 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   3255
   End
   Begin RichUtil1.EListBox lstData 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3413
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
      RowDivisions    =   -1  'True
      ColumnDivisions =   -1  'True
      Columns         =   2
      BeginProperty Column {276D682E-2140-4AB5-8D2C-F105DB576FFF} 
         Count           =   2
      EndProperty
   End
   Begin VB.Label lblOther 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Result:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Select the data type contained at the current location."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private File As BinFile
Private Location As Long

Private Type StrWithHeader
    S As String
End Type
Public Sub Initialize(FilePtr As BinFile, Loc As Long)
    Set File = FilePtr
    Location = Loc
    Label1 = "Select the data type contained at index " & Loc & "."
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConvert_Click()
    On Error GoTo Handler
    Dim F As Long
    F = File.FS
    
    lblOther = "BUSY"
    lblOther.Refresh
    
    Select Case lstData.ItemData(lstData.ListIndex)
    Case vbBoolean
        Dim bln As Boolean
        Get #F, Location, bln
        txtR = Format$(bln, "True/False")
    Case vbInteger
        Dim i As Integer
        Get #F, Location, i
        txtR = i
    Case vbLong
        Dim lng As Long
        Get #F, Location, lng
        txtR = lng
    Case vbSingle
        Dim sng As Single
        Get #F, Location, sng
        txtR = sng
    Case vbDouble
        Dim dbl As Double
        Get #F, Location, dbl
        txtR = dbl
    Case vbCurrency
        Dim cur As Currency
        Get #F, Location, cur
        txtR = FormatCurrency(cur)
    Case vbDate
        Dim dte As Date
        Get #F, Location, dte
        txtR = FormatDateTime(dte, vbGeneralDate)
    Case vbString
        Dim Str As StrWithHeader
        Get #F, Location, Str
        txtR = Str.S
        lblOther = "Len=" & Len(Str.S)
        Exit Sub
    End Select
    lblOther = ""
    Exit Sub
Handler:
    txtR = "(Invalid data)"
    lblOther = "Invalid"
    Exit Sub
End Sub

Private Sub Form_Load()
'Byte 1 byte 0 to 255
'Boolean 2 bytes True or False
'Integer 2 bytes -32,768 to 32,767
'Long
'(long integer) 4 bytes -2,147,483,648 to 2,147,483,647
'Single
'(single-precision floating-point) 4 bytes -3.402823E38 to -1.401298E-45 for negative values; 1.401298E-45 to 3.402823E38 for positive values
'Double
'(double-precision floating-point) 8 bytes -1.79769313486232E308 to
'-4.94065645841247E-324 for negative values; 4.94065645841247E-324 to 1.79769313486232E308 for positive values
'Currency
'(scaled integer) 8 bytes -922,337,203,685,477.5808 to 922,337,203,685,477.5807
'Decimal 14 bytes +/-79,228,162,514,264,337,593,543,950,335 with no decimal point;
'+/-7.9228162514264337593543950335 with 28 places to the right of the decimal; smallest non-zero number is
'+/-0.0000000000000000000000000001
'Date 8 bytes January 1, 100 to December 31, 9999
    
    lstData.AddItem2 "Boolean", VbVarType.vbBoolean, "2 bytes"
    lstData.AddItem2 "Integer (WORD)", VbVarType.vbInteger, "2 bytes"
    lstData.AddItem2 "Long (DWORD)", VbVarType.vbLong, "4 bytes"
    lstData.AddItem2 "Single", VbVarType.vbSingle, "4 bytes"
    lstData.AddItem2 "Double", VbVarType.vbDouble, "8 bytes"
    lstData.AddItem2 "Currency", VbVarType.vbCurrency, "8 bytes"
    lstData.AddItem2 "Date", VbVarType.vbDate, "8 bytes"
    lstData.AddItem2 "VB Struct (Type) String", VbVarType.vbString, "2 byte prefix + length"
End Sub


Private Sub lstData_Click(ListIndex As Long)
    cmdConvert.Enabled = ListIndex > 0
End Sub





Private Sub lstData_DblClick()
    If cmdConvert.Enabled Then cmdConvert_Click
End Sub


