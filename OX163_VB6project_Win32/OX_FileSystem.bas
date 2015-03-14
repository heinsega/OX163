Attribute VB_Name = "OX_FileSystem"
'-------------------------------------------------------------------------
'-----------------OX163文件夹、文件创建与控制模块-------------------------
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
'剪贴板控制API------------------------------------------------------------
'-------------------------------------------------------------------------
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal Format As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal lent As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As Long, ByVal pSource As Long, ByVal lent As Long)

Private Const CF_UNICODETEXT = &HD&
Private Const GMEM_MOVEABLE = &O2&
Private Const GMEM_ZEROINIT = &O40&

'取得文件夹短路径（支持unnicode字符）-------------------------------------
Private Declare Function GetShortPathNameW Lib "Kernel32.dll" (ByVal sLongPath As Long, ByVal sShortPath As Long, ByVal maxLen As Integer) As Integer
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
'取得文件夹短路径（支持unnicode字符）-------------------------------------
'-------------------------------------------------------------------------
Public Function GetShortName(ByVal sLongFileName As String) As String
    On Error Resume Next
    'Unicode API mode-----------------------------------------------------------------
    GetShortName = Space(255)
    Dim GetShortName_slength As Integer
    
    GetShortName_slength = GetShortPathNameW(StrPtr(sLongFileName), StrPtr(GetShortName), 255)
    GetShortName = Left(GetShortName, GetShortName_slength)
    If Right(GetShortName, 1) = "\" Or Right(GetShortName, 1) = "/" Then GetShortName = Left(GetShortName, GetShortName_slength - 1)
    
    If GetShortName = "" Then GetShortName = sLongFileName
    
    'Scripting.FileSystemObject mode---------------------------------------
    '    GetShortName = ""
    '    Dim GetShortName_Fso
    '
    '    Set GetShortName_Fso = CreateObject("Scripting.FileSystemObject")
    '
    '    Err.Clear
    '    GetShortName = GetShortName_Fso.GetFile(sLongFileName).ShortPath
    '
    '    If Err.Number <> 0 Then
    '        Err.Clear
    '        GetShortName = GetShortName_Fso.GetFolder(sLongFileName).ShortPath
    '    End If
    '
    '    Set GetShortName_Fso = Nothing
    '
    '    If GetShortName = "" Then GetShortName = sLongFileName
End Function
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
        .Type = 1 '1-二进制 2-文本
        .Open
        .SaveToFile OX_GreatFileName, 2 '1-不允许覆盖 2-覆盖写入
        .Close
    End With
    Set ADO_Stream = Nothing
    
    If OX_Dirfile(OX_GreatFileName) = False And OX_GreatFile_retry = False Then
        OX_GreatFile_retry = True
        GoTo OX_GreatFileRetry
    ElseIf OX_Dirfile(OX_GreatFileName) = False Then
        GoTo OX_GreatFileErr
    End If
    
    OX_GreatFile = True
    Exit Function
    
OX_GreatFileErr:
    err.Clear
    OX_GreatFile = False
End Function
'-------------------------------------------------------------------------
'创建自定义字符集的文本文件-----------------------------------------------
'-------------------------------------------------------------------------
Public Function OX_GreatTxtFile(OX_GreatTxtFileName As String, TxtFile_Char As String, TxtFileCharset As String) As Boolean
    On Error GoTo OX_GreatFileErr
    
    Dim ADO_Stream As Object
    Dim OX_GreatTxtFile_retry As Boolean
    OX_GreatTxtFile_retry = False
    
OX_GreatFileRetry:
    Set ADO_Stream = CreateObject("ADODB.Stream")
    With ADO_Stream
        .Type = 2 '1-二进制 2-文本
        .Open
        .Charset = TxtFileCharset
        .WriteText TxtFile_Char
        .SaveToFile OX_GreatTxtFileName, 2 '1-不允许覆盖 2-覆盖写入
        .Close
    End With
    Set ADO_Stream = Nothing
    
    If OX_Dirfile(OX_GreatTxtFileName) = False And OX_GreatTxtFile_retry = False Then
        OX_GreatTxtFile_retry = True
        GoTo OX_GreatFileRetry
    ElseIf OX_Dirfile(OX_GreatTxtFileName) = False Then
        GoTo OX_GreatFileErr
    End If
    
    OX_GreatTxtFile = True
    Exit Function
    
OX_GreatFileErr:
    err.Clear
    OX_GreatTxtFile = False
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
    err.Clear
    OX_Dirfile = False
End Function

'-------------------------------------------------------------------------
'删除文件-----------------------------------------------------------------
'-------------------------------------------------------------------------
Public Function OX_DelFile(ByVal OX_FileName As String) As Boolean
    On Error GoTo OX_DelfileErr
    OX_DelFile = False
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile OX_FileName, True
    Set fso = Nothing
    OX_DelFile = Not OX_Dirfile(OX_FileName)
    Exit Function
    
OX_DelfileErr:
    err.Clear
    OX_DelFile = False
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
    err.Clear
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
    err.Clear
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
        StringToCopy = fix_Unicode_FileName(ClipboardText) & vbNullChar
        Clipboard.Clear
        Call OpenClipboard(Form1.hwnd)
        Call EmptyClipboard
        hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(StringToCopy))
        pMem = GlobalLock(hMem)
        Call RtlMoveMemory(pMem, StrPtr(StringToCopy), LenB(StringToCopy))
        Call GlobalUnlock(hMem)
        Call SetClipboardData(CF_UNICODETEXT, hMem)
        Call CloseClipboard
        Exit Sub
    End If
OX_SetClipboardTextErr:
    Clipboard.Clear
    Clipboard.SetText ClipboardText
End Sub

Function GetClipboardText() As String
    On Error GoTo OX_GetClipboardTextErr
    If sysSet.Unicode_File = 0 And IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        Dim hMem As Long, pMem As Long, StringToCopy As String, nSize As Long
        Call OpenClipboard(Form1.hwnd)
        hMem = GetClipboardData(CF_UNICODETEXT)
        pMem = GlobalLock(hMem)
        nSize = GlobalSize(hMem)
        StringToCopy = String(nSize, 0)
        RtlMoveMemory ByVal StrPtr(StringToCopy), ByVal pMem, ByVal nSize
        Call GlobalUnlock(hMem)
        GetClipboardText = Left(StringToCopy, InStr(StringToCopy, Chr(0)) - 1)
        Call CloseClipboard
        Exit Function
    End If
OX_GetClipboardTextErr:
    GetClipboardText = Clipboard.GetText
End Function















