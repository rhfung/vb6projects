VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'26 Jan 2002

Enum FOrder
    conBottom = 1
    conMiddle
    conTop
End Enum

Public Left     As Long 'Single
Public Top      As Long 'Single
Public Width    As Long 'Single
Public Height   As Long 'Single
Public Text     As String
Public DrawOrder As FOrder
Public Tag1     As String 'Tag1 Text: text properties (V5 or less) / Pic: picture properties (all versions)
Public FontFace As String 'Tag2 Text: custom font name
Public Tag3     As String 'Tag3 FText: type <format>; FButton: open <file> & shell <file>
'Expansion tags
Public Name     As String 'Tag4
Public ArrowEngg As Boolean 'Arrow for engineer Tag10    As Boolean
Public ArrowSize As Long 'Size of arrow in % 'Tag11    As Long
Public BackColour   As Long 'fill colour of selected shapes
Public ForeColour   As Long 'line colour
Public GroupNo      As Long
'Font/Text properties
Public TextBold     As Boolean
Public TextSize     As Currency
Public TextItalic   As Boolean
Public TextAlign    As Long 'Dt_left, Dt_center, Dt_right
Public TextUnderline As Boolean
Public TextColour   As Long
'Draw styles
Public LineStyle    As Integer
Public FillStyle    As Integer
'Layering and 3 dimensions
Public ZBase        As Single
Public ZHeight      As Single
Public Layer        As Long
'Run-time properties are not saved
Public CanEdit      As Boolean
Public Selected     As Boolean
Public Enabled      As Boolean
Public EmulationNumber As FAddType 'not used normally

Private mlngLineWidth  As Long
Public Function GetLineWidthPx(ByVal Dest As Object) As Integer
    'DrawStyleConstants
    'DrawWidth property measured in twips
    If LineStyle = vbSolid And mlngLineWidth > 0 Then
        GetLineWidthPx = Dest.ScaleX(mlngLineWidth * 15, vbTwips, vbPixels)
    Else
        GetLineWidthPx = 1
    End If
End Function

Public Property Get LineWidth() As Long
    LineWidth = mlngLineWidth 'DrawWidth property measured in pixels
End Property

Public Property Let LineWidth(ByVal Val As Long)
    '0=1 pixel; otherwise, match to screen pixels
    If Val < 0 Then 'DrawWidth property mesasured in pixels
        mlngLineWidth = 0
    ElseIf Val > 6 Then
        mlngLineWidth = 6
    Else
        mlngLineWidth = Val
    End If
End Property

Private Sub Class_Initialize()
    DrawOrder = conMiddle
    LineStyle = vbSolid
    FillStyle = vbFSTransparent
    BackColour = vbWhite
    ForeColour = vbBlack
    TextColour = vbBlack
    LineWidth = 1
    CanEdit = True
    ArrowSize = 100
    ArrowEngg = False
    'ArrowSolid = True
    Enabled = True
End Sub


Public Property Get CenterX() As Single
    CenterX = Left + Width / 2
End Property


Public Property Get CenterY() As Single
    CenterY = Top + Height / 2
End Property

