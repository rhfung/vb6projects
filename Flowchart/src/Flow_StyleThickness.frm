VERSION 5.00
Begin VB.Form frmStyleThickness 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Thickness"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   1605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsbSample 
      Height          =   255
      Left            =   120
      Max             =   6
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblThin 
      Alignment       =   2  'Center
      Caption         =   "thin"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line linSample 
      X1              =   120
      X2              =   1440
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "frmStyleThickness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IToolbar
'date: December 28, 2002
'For flowchart

Private mItem       As FlowItem
Private mDefaults   As FlowChart

Private Sub ItemChanged()
    If Not mItem Is Nothing Then
        frmMain.ToolBoxChanged conUndoObjectFormatLineWidth  'undone
        mItem.P.LineWidth = hsbSample
        frmMain.ToolBoxDone False, False, True, False, False, False, False, False
    End If
End Sub

Public Sub UpdateToForm(Item As FlowItem, Defaults As FlowChart)
    Set mItem = Item
    Set mDefaults = Defaults
    
    If Not mItem Is Nothing Then
        On Error Resume Next
        hsbSample = mItem.P.LineWidth
        hsbSample_Change
        
        If Err Then
            Err.Clear
            linSample.Visible = False
            lblThin = "other"
            lblThin.Visible = True
        End If
        SetEnabled Me, True
    Else
        SetEnabled Me, False
    End If
End Sub



Private Sub Form_Load()
    gblnWindowThick = True
    UpdateToForm Nothing, Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gblnWindowThick = False
End Sub

Private Property Get IToolbar_Height() As Single
    IToolbar_Height = Height
End Property

Private Sub IToolbar_Move(Left As Single, Top As Single)
    Move Left, Top
End Sub


Private Sub IToolbar_Show(Mode As FormShowConstants, Form As Form)
    Show Mode, Form
End Sub


Private Sub IToolbar_UpdateToForm(Item As FlowItem, Defaults As FlowChart)
    UpdateToForm Item, Defaults
End Sub

Private Sub hsbSample_Change()
    If hsbSample <> 0 Then
        linSample.BorderWidth = hsbSample
        linSample.Visible = True
        lblThin.Visible = False
    Else '0
        linSample.BorderWidth = 1
        linSample.Visible = False
        lblThin = "thin"
        lblThin.Visible = True
    End If
    
    ItemChanged
End Sub


