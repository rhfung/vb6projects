VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "My Windows - File List"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'/// My Windows - File List
Option Explicit
'// Note: The class checks the Options flag, but
'// the StartDir function does not check to see
'// if NonModalAllowed for the busy dialog box.

Public Enum FileListConstants
    conFileListNothing = 0
    conFileListTakeBreaks = 1
    conFileListRaiseEvent = 2
    conFileListBusyDialogBox = 4
End Enum

Public FileAttributes As Long 'in/out
Public CancelOp     As Boolean 'in/out
'Public Options      As FileListConstants

Private m_colFiles   As Collection 'internal data
Private m_colFolders As Collection
Private m_strPathName As String 'out
Private m_lngCount    As Long 'out
Private m_frmBusy     As frmBusy
Private m_Options     As FileListConstants

Public Event Progress(ByVal NumOfFiles As Long, StopOp As Boolean)


Public Property Get Options() As FileListConstants
Attribute Options.VB_Description = "Retuns/sets the options for the file list, in binary flags (or/and allowed).   Cannot set conFileListBusyDialogBox if modeless forms are not allowed."
'/// Retuns/sets the options for the file list, in binary flags (or/and allowed).
'/// Cannot set conFileListBusyDialogBox if modeless forms are not allowed.
    Options = m_Options
End Property


Public Property Let Options(ByVal pOptions As FileListConstants)
    m_Options = pOptions
    If (m_Options And conFileListBusyDialogBox) = conFileListBusyDialogBox And Not App.NonModalAllowed Then
        m_Options = m_Options - (m_Options And conFileListBusyDialogBox) 'remove busy dialog box flag
    End If
End Property





Public Property Get Count() As Long
Attribute Count.VB_Description = "Total count of files and folders."
'/// Total count of files and folders.
    Count = m_lngCount
End Property

Public Sub DoBusy(Optional Caption As String)
Attribute DoBusy.VB_Description = "Raises DoEvents(), the Progress event, or shows the busy dialog box depending on Options."
'/// Raises DoEvents(), the Progress event, or shows the busy dialog box depending on Options.
    'see also StartDir
    If (Options And conFileListTakeBreaks) = conFileListTakeBreaks Then DoEvents
    If (Options And conFileListRaiseEvent) = conFileListRaiseEvent Then RaiseEvent Progress(Count, CancelOp)
    If (Options And conFileListBusyDialogBox) = conFileListBusyDialogBox Then
        On Error Resume Next
        If m_frmBusy Is Nothing Then
            Set m_frmBusy = New frmBusy
        End If
        m_frmBusy.Label1 = Caption
        If App.NonModalAllowed Then
            m_frmBusy.Show
            DoEvents
        End If
    End If
End Sub

Public Property Get FileCol() As CollectionString
Attribute FileCol.VB_Description = "Returns collection for Visual Basic."
'/// Returns collection for Visual Basic.
    Dim objCol As New CollectionString
    Set objCol.Col = m_colFiles
    Set FileCol = objCol
End Property

Public Property Get FileCount() As Long
Attribute FileCount.VB_Description = "Count of files accessed using FileItem()."
'/// Count of files accessed using FileItem().
    FileCount = m_colFiles.Count
End Property

Public Property Get FileItem(ByVal Index As Long) As String
Attribute FileItem.VB_Description = "One-based index."
'/// One-based index.
    FileItem = m_colFiles(Index)
End Property

Public Sub Finished()
Attribute Finished.VB_Description = "Call this function to close the Busy dialog box when the conFileListBusyDialogBox flag is set."
'/// Call this function to close the Busy dialog box when the conFileListBusyDialogBox flag is set.
    'close busy dialog box
    If Not m_frmBusy Is Nothing Then
        Unload m_frmBusy
        Set m_frmBusy = Nothing
    End If
End Sub

Public Property Get FolderCol() As CollectionString
Attribute FolderCol.VB_Description = "Returns collection for Visual Basic."
'/// Returns collection for Visual Basic.
    Dim objCol As New CollectionString
    Set objCol.Col = m_colFolders
    Set FolderCol = objCol
End Property


Public Property Get FolderCount() As Long
Attribute FolderCount.VB_Description = "Count of folders accessed using FolderItem()."
'/// Count of folders accessed using FolderItem().
    FolderCount = m_colFolders.Count
End Property


Public Property Get FolderItem(ByVal Index As Long) As String
Attribute FolderItem.VB_Description = "One-based index."
'/// One-based index.
    FolderItem = m_colFolders(Index)
End Property


Public Property Get Pathname() As String
Attribute Pathname.VB_Description = "Returns formatted pathname with backslash."
'/// Returns formatted pathname with backslash.
    Pathname = m_strPathName
End Property

Public Function StartDir(ByVal Find As String) As Long
Attribute StartDir.VB_Description = "Path supports the use of multiple character (*) and single character (?) wildcards to specify multiple files."
'/// Path supports the use of multiple character (*) and single character (?) wildcards to specify multiple files.
    Dim strDir As String, strLast As String
    Dim lngError As Long 'allow 9 errors before stopping operation
    Dim lngLastErr As Long

    On Error Resume Next
    CancelOp = False
    
    If (Options And conFileListBusyDialogBox) = conFileListBusyDialogBox And m_frmBusy Is Nothing Then
        Set m_frmBusy = New frmBusy
        Set m_frmBusy.m_objFileList = Me
    End If
    
    m_strPathName = FormatPath(GetPath(Find))
    strDir = Dir(Find, FileAttributes)
    
    Set m_colFiles = New Collection
    Set m_colFolders = New Collection
    
    Do While Len(strDir) And strDir <> strLast And Not CancelOp And lngError < 10
        If strDir <> "." And strDir <> ".." Then
            m_lngCount = m_lngCount + 1
            
            If (GetAttr(m_strPathName & strDir) And vbDirectory) <> vbDirectory Then
                m_colFiles.Add strDir
            Else
                m_colFolders.Add strDir
            End If
        End If
        
        If Err Then
            lngLastErr = Err
            lngError = lngError + 1
            Err.Clear
        End If
        
        strLast = strDir 'prevent from getting stuck on a file
        strDir = Dir
        If m_lngCount Mod 50 = 0 Then GoSub DoBusyProcedure
    Loop
    
    StartDir = lngLastErr
    
    'close busy dialog box
    If Not m_frmBusy Is Nothing Then
        Unload m_frmBusy
        Set m_frmBusy = Nothing
    End If
    
    Exit Function
    
DoBusyProcedure:
    If (Options And conFileListTakeBreaks) = conFileListTakeBreaks Then DoEvents
    If (Options And conFileListRaiseEvent) = conFileListRaiseEvent Then RaiseEvent Progress(Count, CancelOp)
    If (Options And conFileListBusyDialogBox) = conFileListBusyDialogBox Then
        On Error Resume Next
        m_frmBusy.Show
        DoEvents
    End If
    
    Return
End Function

Public Property Let TakeBreaks(ByVal pVal As Boolean)
Attribute TakeBreaks.VB_Description = "Determines if DoEvents() should be raised.  Changes the Options flag."
'/// Determines if DoEvents() should be raised.  Changes the Options flag.
    If pVal Then
        'add take breaks
        Options = Options Or conFileListTakeBreaks
    Else
        'remove take breaks
        Options = Options - (Options And conFileListTakeBreaks)
    End If
End Property

Private Sub Class_Initialize()
    FileAttributes = vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory
    Options = conFileListRaiseEvent
End Sub

Private Sub Class_Terminate()
    Finished
End Sub

