Attribute VB_Name = "OX_FileSystem"
'-------------------------------------------------------------------------
'OX163文件夹、文件创建与控制----------------------------------------------
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
'剪贴板控制API------------------------------------------------------------
'-------------------------------------------------------------------------
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal Format As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal Flags As Long, ByVal lent As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As Long, ByVal pSource As Long, ByVal lent As Long)

Private Const CF_UNICODETEXT = &HD&
Private Const GMEM_MOVEABLE = &O2&
Private Const GMEM_ZEROINIT = &O40&
'-------------------------------------------------------------------------
'创建文件（含有Unicode字符亦可）------------------------------------------
'-------------------------------------------------------------------------
Public Function OX_GreatFile(ByVal OX_GreatFileName As String) As Boolean
On Error GoTo OX_GreatFileErr
    Dim ADO_Stream As Object
    
    Dim OX_GreatFile_retry As Boolean
    OX_GreatFile_retry = False

OX_GreatFileRetry:
    Set ADO_Stream = CreateObject("ADODB.Stream")
    With ADO_Stream
        .Type = 1
        .Open
        .SaveToFile OX_GreatFileName, 2
        .Close
    End With
    Set ADO_Stream = Nothing
    
    If OX_Dirfile(OX_GreatFileName) = False And OX_GreatFile_retry = False Then
        GoTo OX_GreatFileRetry
    ElseIf OX_Dirfile(OX_GreatFileName) = False Then
        GoTo OX_GreatFileErr
    End If
    
    OX_GreatFile = True
    Exit Function
    
OX_GreatFileErr:
    Err.Clear
    OX_GreatFile = False
End Function

'-------------------------------------------------------------------------
'判断文件是否存在---------------------------------------------------------
'-------------------------------------------------------------------------
Public Function OX_Dirfile(ByVal OX_FileName As String) As Boolean
On Error GoTo OX_DirfileErr
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    OX_Dirfile = fso.FileExists(OX_FileName)
    Set fso = Nothing
    Exit Function
    
OX_DirfileErr:
    Err.Clear
    OX_Dirfile = False
End Function

'-------------------------------------------------------------------------
'删除文件-----------------------------------------------------------------
'-------------------------------------------------------------------------
Public Function OX_Delfile(ByVal OX_FileName As String) As Boolean
On Error GoTo OX_DelfileErr
OX_Delfile = False
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile OX_FileName, True
    Set fso = Nothing
    OX_Delfile = Not OX_Dirfile(OX_FileName)
    Exit Function
    
OX_DelfileErr:
    Err.Clear
    OX_Delfile = False
End Function

'-------------------------------------------------------------------------
'判断文件夹是否存在-------------------------------------------------------
'-------------------------------------------------------------------------
Public Function OX_DirFolder(ByVal OX_FolderName As String) As Boolean
On Error GoTo OX_DirFolderErr
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    OX_DirFolder = fso.FolderExists(OX_FolderName)
    Set fso = Nothing
    Exit Function
    
OX_DirFolderErr:
    Err.Clear
    OX_DirFolder = False
End Function

'-------------------------------------------------------------------------
'创建文件夹---------------------------------------------------------------
'-------------------------------------------------------------------------
Public Function OX_CreateFolder(ByVal OX_FolderName As String) As Boolean
On Error GoTo OX_CreateFolderErr

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    OX_CreateFolder = fso.CreateFolder(OX_FolderName)
    Set fso = Nothing
    Exit Function

OX_CreateFolderErr:
    Err.Clear
    OX_CreateFolder = False
End Function

'-------------------------------------------------------------------------
'创建url文件--------------------------------------------------------------
'-------------------------------------------------------------------------
Public Sub OX_CreateUrlIniFile(ByVal OX_UrlIniFileName As String)
    If Dir(App_path & "\url\" & OX_UrlIniFileName) = "" Then
        If Dir(App_path & "\url", vbDirectory) = "" Then MkDir App_path & "\url"
        WriteUnicodeIni "maincenter", "url", OX_UrlIniFileName, App_path & "\url\" & OX_UrlIniFileName
    End If
End Sub

'-------------------------------------------------------------------------
'复制unicode字符到剪贴板--------------------------------------------------
'-------------------------------------------------------------------------
Sub SetClipboardText(ClipboardText)
On Error GoTo OX_SetClipboardTextErr
    If sysSet.Unicode_File = 0 Then
        Dim hMem As Long, pMem As Long, StringToCopy As String
        StringToCopy = fix_Unicode_FileName(ClipboardText)
        Clipboard.Clear
        Call OpenClipboard(Form1.hWnd)
        Call EmptyClipboard
        hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(StringToCopy) + 1)
        pMem = GlobalLock(hMem)
        Call RtlMoveMemory(pMem, StrPtr(StringToCopy), LenB(StringToCopy) + 1)
        Call GlobalUnlock(hMem)
        Call SetClipboardData(CF_UNICODETEXT, hMem)
        Call CloseClipboard
        Exit Sub
    End If
OX_SetClipboardTextErr:
        Clipboard.Clear
        Clipboard.SetText ClipboardText
End Sub
















