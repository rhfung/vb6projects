VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PUndoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'March 23, 2001
Implements IUndoInterface

'original
Public theActionWas As PUndoAction
Public P            As Properties

'local variable(s) to hold property value(s)
Private mvarPtr As FlowItem  'local copy

Public Property Set Ptr(ByVal vData As FlowItem)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Ptr = Form1
    Set mvarPtr = vData
End Property

Public Property Get Ptr() As FlowItem
    Set Ptr = mvarPtr
End Property

Public Property Get theActionName() As String
    Select Case theActionWas
    Case conUndoSize
        theActionName = "Size"
    Case conUndoMove
        theActionName = "Move"
    Case conUndoTextChange
        theActionName = "Text"
    Case conUndoDelete
        theActionName = "Delete"
    Case conUndoAutoSize
        theActionName = "Auto-Size"
    Case conUndoAlignment
        theActionName = "Alignment"
    Case conUndoPicReset
        theActionName = "Reset Picture"
    Case conUndoPicZoom
        theActionName = "Picture Zoom"
    Case conUndoFontChanged
        theActionName = "Font Format"
    Case conUndoForeColourChange
        theActionName = "Line Colour"
    Case conUndoBackColourChange
        theActionName = "Fill Colour"
    Case conUndoObjectFormatFillStyle
        theActionName = "Fill Style"
    Case conUndoObjectFormatLineStyle
        theActionName = "Line Style"
    Case conUndoObjectFormatLineWidth
        theActionName = "Line Width"
    Case conUndoTextColourChange
        theActionName = "Text Colour"
    Case conUndoCut
        theActionName = "Object Cut"
    Case conUndoChangeTag3Field
        theActionName = "Change Field"
    Case conUndoChangeTag3Macro
        theActionName = "Change Macro"
    Case conUndoChangeArrowSize
        theActionName = "Arrow Size"
    Case conUndoChangeArrowType, conUndoChangeArrowSolid
        theActionName = "Arrow Type"
    Case conUndoChangeName
        theActionName = "Name"
    Case conUndoObject
        theActionName = "Object"
    Case conUndoMultiple
        theActionName = "Multiple"
    End Select
End Property


Private Sub Class_Initialize()
    Set P = New Properties
End Sub

Private Sub Class_Terminate()
    Set P = Nothing
End Sub

Private Property Get IUndoInterface_IsMultiple() As Boolean
    IUndoInterface_IsMultiple = False
End Property


