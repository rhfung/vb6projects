VERSION 5.00
Begin VB.Form frmFindByte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find and Replace"
   ClientHeight    =   3630
   ClientLeft      =   30
   ClientTop       =   520
   ClientWidth     =   5560
   Icon            =   "Inview_FindByte.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   370
      Left            =   3120
      TabIndex        =   7
      Top             =   3120
      Width           =   1210
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   370
      Left            =   840
      TabIndex        =   6
      Top             =   3120
      Width           =   1330
   End
   Begin VB.TextBox txtReplace 
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   4930
   End
   Begin VB.CheckBox chkReplace 
      Caption         =   "Enter in the byte(s) to &replace with:"
      Height          =   250
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   5290
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5170
   End
   Begin VB.Label lblReplace 
      Height          =   250
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   4930
   End
   Begin VB.Label lblFind 
      Caption         =   "Preview"
      Height          =   250
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5170
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the decimal value of byte(s) to &find OR the characters OR the characters in quotation marks ALL seperated by spaces:"
      Height          =   440
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5170
   End
End
Attribute VB_Name = "frmFindByte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CheckForm(FindPart As Boolean, ReplacePart As Boolean, strInput As String, strOutput As String) As Boolean
    Dim blnOk As Boolean
    Dim lower   As Long
    Dim lngElement As Long
    Dim lngLenSequence As Long
    'Dim strInput As String
    Dim vntInput As Variant
    Dim bytInput() As Byte
    'Dim strOutput As String
    Dim vntOutput As Variant
    Dim bytOutput() As Byte
    Dim blnFlag   As Boolean
    Dim strEl       As String
    
    blnOk = True
    
    
    On Error GoTo BadConversionFind
    
    vntInput = Split(txtFind.Text, " ")
    lower = LBound(vntInput)
    lngLenSequence = UBound(vntInput) - LBound(vntInput) + 1
    If lngLenSequence < 1 Then
        lblFind = "Preview: nothing to find"
        lblReplace = "Preview:"
        cmdFind.Enabled = False
        Exit Function
    End If
    
    If FindPart Then
        'validate input
        
        ReDim bytInput(lower To UBound(vntInput))
        
        
        For lngElement = lower To UBound(vntInput)
            strEl = vntInput(lngElement)
            If IsNumeric(strEl) Then
                bytInput(lngElement) = CInt(strEl)
            ElseIf Len(strEl) = 1 Then
                bytInput(lngElement) = Asc(strEl)
            ElseIf Len(strEl) = 3 And InQuotationMarks(strEl) Then
                bytInput(lngElement) = Asc(RemoveQuotationMarks(strEl))
            Else
            
                lblFind = "Preview: element " & (lngElement - lower + 1) & " is bad"
                blnOk = False
                Exit For
            End If
            
        Next lngElement
        
        If lngLenSequence < 1 Then
            blnOk = False
            lblFind = "Preview: nothing to find"
        End If
        
        If blnOk Then 'show sequence
            For lngElement = lower To UBound(bytInput)
                strInput = strInput & bytInput(lngElement) & " "
            Next lngElement
            strInput = RTrim$(strInput)
            lblFind = "Preview: " & strInput
        End If
    End If

    On Error GoTo BadConversionReplace
    If Not blnOk And ReplacePart Then
        lblReplace = "Preview: fix Find part first"
    ElseIf ReplacePart Then
        ReDim bytOutput(lower To UBound(vntInput))
        
        'validate output
        
        vntOutput = Split(txtReplace.Text, " ")
        If UBound(vntOutput) <> UBound(vntInput) Or LBound(vntOutput) <> LBound(vntInput) Then
            blnOk = False
            lblReplace = "Preview: length of replace byte sequence must match find"
        End If
        
        If blnOk Then
            For lngElement = lower To UBound(vntOutput)
                strEl = vntOutput(lngElement)
                If IsNumeric(strEl) Then
                    bytOutput(lngElement) = CInt(strEl)
                ElseIf Len(strEl) = 1 Then
                    bytOutput(lngElement) = Asc(strEl)
                ElseIf Len(strEl) = 3 And InQuotationMarks(strEl) Then
                    bytOutput(lngElement) = Asc(RemoveQuotationMarks(strEl))
                Else
                    lblReplace = "Preview: element " & (lngElement - lower + 1) & " is bad"
                    blnOk = False
                    Exit For
                End If
           

            Next lngElement
        End If
        
        If blnOk Then 'show sequence
            For lngElement = lower To UBound(bytOutput)
                strOutput = strOutput & bytOutput(lngElement) & " "
            Next lngElement
            strOutput = RTrim$(strOutput)
            lblReplace = "Preview: " & strOutput
        End If

    End If
    
    CheckForm = blnOk
    cmdFind.Enabled = blnOk
    Exit Function
BadConversionFind:
    lblFind = "Preview: invalid byte value [0,255]"
    CheckForm = False
    cmdFind.Enabled = False
    Exit Function
BadConversionReplace:
    lblReplace = "Preview: invalid byte value [0,255]"
    CheckForm = False
    cmdFind.Enabled = False
    Exit Function
End Function

Private Function InQuotationMarks(ByVal Str As String) As Boolean
    InQuotationMarks = (Left$(Str, 1) = """" And Right$(Str, 1) = """" And Len(Str) > 1) Or _
    (Left$(Str, 1) = "'" And Right$(Str, 1) = "'" And Len(Str) > 1)
End Function

    
Private Function RemoveQuotationMarks(ByVal Str As String) As String
    If InQuotationMarks(Str) Then
        RemoveQuotationMarks = Mid$(Str, 2, Len(Str) - 2)
    Else
        RemoveQuotationMarks = Str
    End If
End Function


Private Sub chkReplace_Click()
    If chkReplace Then
        cmdFind.Caption = "Replace"
    Else
        cmdFind.Caption = "Find"
        lblReplace = ""
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String, strReplace$
    
    If CheckForm(True, chkReplace, strFind, strReplace) Then
        frmMain.DoFindReplace strFind, strReplace, chkReplace
        Unload Me
    End If
End Sub

Private Sub txtFind_Change()
    If Len(txtFind) = 0 Then cmdFind.Enabled = False Else cmdFind.Enabled = True
End Sub

Private Sub txtFind_Validate(Cancel As Boolean)
    CheckForm True, False, "", ""
End Sub


Private Sub txtReplace_Change()
    If Len(txtReplace) > 0 And chkReplace <> vbChecked Then
        chkReplace = vbChecked
    ElseIf Len(txtReplace) = 0 And chkReplace = vbChecked Then
        chkReplace = vbUnchecked
    End If
    
End Sub


Private Sub txtReplace_Validate(Cancel As Boolean)
    CheckForm False, chkReplace, "", ""
End Sub


