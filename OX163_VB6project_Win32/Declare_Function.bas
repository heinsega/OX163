Attribute VB_Name = "Declare_Function"
'-------------------------------------------------------------------------
'-----------------OX163调用API操作、控制、传递模块------------------------
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
'Sleep方式让程序循环使释放CPU占用-----------------------------------------
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'-------------------------------------------------------------------------
'使用XP风格---------------------------------------------------------------
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

'-------------------------------------------------------------------------
'程序窗口重在最前面-------------------------------------------------------
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40

'-------------------------------------------------------------------------
'系统文件夹---------------------------------------------------------------
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryW" (ByVal lpBuffer As Long, ByVal nSize As Long) As Long

'-------------------------------------------------------------------------
'读取ini配置--------------------------------------------------------------
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long
'写入ini配置
'Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lpFileName As Long) As Long

'-------------------------------------------------------------------------
'得到路径是否可写---------------------------------------------------------
Public Declare Function GetFileAttributesAPI Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

'-------------------------------------------------------------------------
'进程优先级---------------------------------------------------------------
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
'Const IDLE_PRIORITY_CLASS = &H40
'Const BELOW_NORMAL_PRIORITY_CLASS = &H4000
'Const NORMAL_PRIORITY_CLASS = &H20
'Const ABOVE_NORMAL_PRIORITY_CLASS = &H8000
'Const HIGH_PRIORITY_CLASS = &H80
'Const REALTIME_PRIORITY_CLASS = &H100


'-------------------------------------------------------------------------
'激活toolbar buttonMenu---------------------------------------------------
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As Long, ByVal lpsz2 As Long) As Long

Private Const TBN_FIRST = -700&
Private Const TBN_DROPDOWN = (TBN_FIRST - 10)

Private Const WM_USER1 = &H400
Private Const TB_GETBUTTON As Long = (WM_USER1 + 23)
Private Const WM_NOTIFY As Long = &H4E&

Private Type OX_ToolBar_Button
   OTBButton_Bitmap As Long
   OTBButton_idCommand As Long
   OTBButton_fsState As Byte
   OTBButton_fsStyle As Byte
   OTBButton_bReserved1 As Byte
   OTBButton_bReserved2 As Byte
   OTBButton_dwData As Long
   OTBButton_iString As Long
End Type

Private Type OX_NMHDR
   O_NMHDR_hwndFrom As Long
   O_NMHDR_idFrom As Long
   O_NMHDR_Code As Long
End Type

Private Type OX_NMTOOLBAR
    O_NMTB_hdr As OX_NMHDR
    O_NMTB_Item As Long
    O_NMTB_Btn As OX_ToolBar_Button
    O_NMTB_cchText As Long
    O_NMTB_lpszString As Long
End Type


'-------------------------------------------------------------------------
'API将返回Windows的Non - Unicode设定--------------------------------------
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

'-------------------------------------------------------------------------
'调用shell保存路径--------------------------------------------------------
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderW" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListW" (ByVal pidl As Long, ByVal pszPath As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
'Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Const MAX_PATH = 1024

'可以参看“调用shell保存路径 BROWSEINFO structure (Windows) 参数.doc”
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_EDITBOX = &H10

Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_VALIDATEFAILED As Long = 3

Private Const WM_USER = &H400

Private Const BFFM_SETSTATUSTEXT As Long = (WM_USER + 100)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTION As Long = (WM_USER + 102)

Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40

Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

'-------------------------------------------------------------------------
'最小化系统托盘-----------------------------------------------------------
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconW" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'-------------------------------------------------------------------------
'打开IE-------------------------------------------------------------------
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteW" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'-------------------------------------------------------------------------
'解Gzip压缩数组-----------------------------------------------------------
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long) '已经定义
Private Declare Function InitDecompression Lib "gzip.dll" () As Long
Private Declare Function CreateDecompression Lib "gzip.dll" (ByRef context As Long, ByVal flags As Long) As Long
Private Declare Function DestroyDecompression Lib "gzip.dll" (ByRef context As Long) As Long
Private Declare Function Decompress Lib "gzip.dll" (ByVal context As Long, inBytes As Any, ByVal input_size As Long, outBytes As Any, ByVal output_size As Long, ByRef input_used As Long, ByRef output_used As Long) As Long
'Private Const OFFSET As Long = &H8

'-------------------------------------------------------------------------
'API代替CommonDialog参数定义---------------------------------start--------
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameW" (ByRef pOpenfilename As OPENFILENAME) As Long
Private Const OFN_FILEMUSTEXIST = &H1000
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As Long
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'API代替CommonDialog参数定义---------------------------------end----------

'-------------------------------------------------------------------------
'API调用unicode文件对话框CommonDialog---------------start-----------------
'-------------------------------------------------------------------------
Public Function ShowOpenFileDialog(InitialDir As String, DialogTitle As String, Filter As String, ByVal FrmhWnd As Long) As String
    
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long, lReturn_str As String
    
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = FrmhWnd
    OpenFile.hInstance = App.hInstance
    Dim sFile As String
    sFile = Space$(1024)
    
    OpenFile.lpstrFilter = StrConv(Replace(Filter, "|", Chr$(0)), vbUnicode)
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = StrPtr(sFile)
    OpenFile.nMaxFile = Len(sFile)
    OpenFile.lpstrFileTitle = Space$(512)
    OpenFile.nMaxFileTitle = Len(OpenFile.lpstrFileTitle)
    OpenFile.lpstrInitialDir = StrConv(InitialDir, vbUnicode)
    OpenFile.lpstrTitle = StrConv(DialogTitle, vbUnicode)
    OpenFile.flags = OFN_FILEMUSTEXIST
    lReturn = GetOpenFileName(OpenFile)
    
    If lReturn = 0 Then  'lReturn is always 0 even when a file is selected!!
        ShowOpenFileDialog = "" 'Cancel Button Pressed
    Else
        Dim byte_temp(1 To 1024) As Byte
        CopyMemory byte_temp(1), ByVal OpenFile.lpstrFile, 1024
        
        lReturn_str = byte_temp
        lReturn_str = Trim(lReturn_str)
        If InStr(lReturn_str, vbNullChar) > 0 Then lReturn_str = Left(lReturn_str, InStr(lReturn_str, vbNullChar) - 1)
        ShowOpenFileDialog = lReturn_str
    End If
    
End Function

Public Function ShowSaveFileDialog(InitialDir As String, DialogTitle As String, Filter As String, file_type As String, ByVal FrmhWnd As Long) As String
    On Error Resume Next
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long, lReturn_str As String, file_type_split
    
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = FrmhWnd
    OpenFile.hInstance = App.hInstance
    Dim sFile As String
    sFile = Space$(1024)
    
    OpenFile.lpstrFilter = StrConv(Replace(Filter, "|", Chr$(0)), vbUnicode)
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = StrPtr(sFile)
    OpenFile.nMaxFile = Len(sFile)
    OpenFile.lpstrFileTitle = Space$(512)
    OpenFile.nMaxFileTitle = Len(OpenFile.lpstrFileTitle)
    OpenFile.lpstrInitialDir = StrConv(InitialDir, vbUnicode)
    OpenFile.lpstrTitle = StrConv(DialogTitle, vbUnicode)
    OpenFile.flags = OFN_FILEMUSTEXIST
    lReturn = GetSaveFileName(OpenFile)
    
    If lReturn = 0 Then  'lReturn is always 0 even when a file is selected!!
        ShowSaveFileDialog = "" 'Cancel Button Pressed
    Else
        
        Dim byte_temp(1 To 1024) As Byte
        CopyMemory byte_temp(1), ByVal OpenFile.lpstrFile, 1024
        
        lReturn_str = byte_temp
        lReturn_str = Trim(lReturn_str)
        If InStr(lReturn_str, vbNullChar) > 0 Then lReturn_str = Left(lReturn_str, InStr(lReturn_str, vbNullChar) - 1)
        
        '检查文件后缀名
        file_type_split = Split(file_type, "|")
        Filter = ""
        lReturn = OpenFile.nFilterIndex - 1
        Filter = file_type_split(lReturn)
        If LCase(Right(lReturn_str, Len(Filter))) <> LCase(Filter) Then lReturn_str = lReturn_str & Filter
        
        ShowSaveFileDialog = lReturn_str
    End If
    
End Function

'API调用unicode文件对话框CommonDialog---------------end-------------------

'-------------------------------------------------------------------------
'---------------------调用shell选择保存目录---------start-----------------
'-------------------------------------------------------------------------
Public Function GetFolder(ByVal title As String, ByVal start As String, ByVal newfolder As Boolean) As String
    Dim BI As BROWSEINFO, pidl As Long, lpSelPath As Long
    Dim spath  As String
    'fill in the info it needs
    With BI
        .hOwner = GetForegroundWindow
        .pidlRoot = 0
        .lpszTitle = StrConv(title, vbUnicode)
        .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_EDITBOX
        If newfolder = True Then .ulFlags = BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE + BIF_EDITBOX
        lpSelPath = LocalAlloc(LPTR, LenB(start) + 1)
        CopyMemory ByVal lpSelPath, ByVal start, LenB(start) + 1
        .lParam = lpSelPath
    End With
    
    'get the idlist long from the returned folder
    pidl = SHBrowseForFolder(BI)
    
    'do then if they clicked ok
    If pidl Then
        spath = Space(MAX_PATH)
        If SHGetPathFromIDList(pidl, StrPtr(spath)) Then
            'next line is the returned folder
            If InStr(spath, vbNullChar) > 0 Then spath = Left(spath, InStr(spath, vbNullChar) - 1)
            GetFolder = GetShortName(spath)
        End If
        'Call CoTaskMemFree(pidl)
    Else
        'user clicked cancel
    End If
    
    Call LocalFree(lpSelPath)
    
End Function

'this seems to happen before the box comes up and when a folder is clicked on within it
Public Function BrowseCallbackProcStr(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Dim spath As String, bFlag As Long
    
    Select Case uMsg
    Case BFFM_INITIALIZED
        'browse has been initialized, set the start folder
        Call SendMessage(hwnd, BFFM_SETSELECTION, 1, ByVal lpData)
    Case BFFM_SELCHANGED
        spath = Space$(MAX_PATH)
        If SHGetPathFromIDList(lParam, StrPtr(spath)) Then
            spath = Trim(spath) & vbNullChar
            Call SendMessage(hwnd, BFFM_SETSTATUSTEXTW, 0, StrPtr(spath))
        End If
    End Select
    
End Function

Public Function FARPROC(pfn As Long) As Long
    FARPROC = pfn
End Function
'---------------------调用shell选择保存目录------end----------------------

Public Sub OX_ShowButtonMenu(OX_Toolbar As MSComctlLib.Toolbar, ByVal ButtonIndex As Long) 'buttonindex从0开始
    Dim OX_Button As OX_ToolBar_Button
    Dim OX_Notify As OX_NMTOOLBAR
    Dim lResult As Long
    Dim m_hwnd As Long
    Dim lCommandId As Long
    m_hwnd = FindWindowEx(OX_Toolbar.hwnd, 0, StrPtr("msvb_lib_toolbar"), StrPtr(vbNullString))
    lResult = SendMessage(m_hwnd, TB_GETBUTTON, ButtonIndex, OX_Button)
    lCommandId = OX_Button.OTBButton_idCommand
    With OX_Notify
        .O_NMTB_hdr.O_NMHDR_Code = TBN_DROPDOWN
        .O_NMTB_hdr.O_NMHDR_hwndFrom = m_hwnd
        .O_NMTB_Item = lCommandId
    End With
    lResult = SendMessage(OX_Toolbar.hwnd, WM_NOTIFY, 0, OX_Notify)
End Sub

'OX163启动函数------------------------------------------------------------
'XP风格-----start---------------------------------------------------------
'-------------------------------------------------------------------------
Public Function InitCommonControlsVB() As Boolean
    On Error Resume Next
    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = &H200
    End With
    InitCommonControlsEx iccex
    InitCommonControlsVB = (err.Number = 0)
    On Error GoTo 0
End Function

Sub Main()
    InitCommonControlsVB
    OX_POPMENU_X = 5 * Screen.TwipsPerPixelX
    OX_POPMENU_Y = 5 * Screen.TwipsPerPixelY
    start_ox163.Show
End Sub

'-------------XP风格--------end-------------------------------------------

'-------------------------------------------------------------------------
'取得系统文件夹-----------------------------------------------------------
'-------------------------------------------------------------------------
Public Function GetSysDir() As String
    Dim strBuf As String
    Dim lngBuf As Long
    
    strBuf = Space$(1024)
    lngBuf = 1024
    
    lngBuf = GetSystemDirectory(StrPtr(strBuf), lngBuf)
    strBuf = Trim(strBuf)
    'If InStr(strBuf, vbNullChar) > 0 Then strBuf = left(strBuf, InStr(strBuf, vbNullChar) - 1)
    strBuf = Left(strBuf, lngBuf)
    GetSysDir = GetShortName(strBuf)
End Function

'-------------------------------------------------------------------------
'INI文件读取保存----------------------------------------------------------
'-------------------------------------------------------------------------
'建立或检查unicode编码的ini配置文件
Public Function GreatUnicodeIniFile(ByVal ini_path As String) As Boolean
    On Error GoTo GreatUnicodeIniFileErr
    Dim ReturnEncoding As String, GUIF_Object As Object, GUIF_file As Object
    If Dir(GetShortName(ini_path)) <> "" Then
        ReturnEncoding = GetEncoding(ini_path)
        
        If ReturnEncoding = "Unicode" Then
            'Unicode处理
            GreatUnicodeIniFile = True
            Exit Function
        ElseIf ReturnEncoding = "UTF-8" Then
            'UTF-8处理
            Set GUIF_Object = CreateObject("ADODB.Stream") '建立一个流对象
            With GUIF_Object
                .Type = 2
                .mode = 3
                .Charset = "UTF-8"
                .Open
                .LoadFromFile GetShortName(ini_path)
                ReturnEncoding = .ReadText
                .Close
            End With
            Set GUIF_Object = Nothing
        ElseIf ReturnEncoding = "UnicodeBigEndian" Then
            'Unicode-BE处理
            ReturnEncoding = load_normal_file(GetShortName(ini_path), -1)
        Else
            'ANSI处理
            ReturnEncoding = load_normal_file(GetShortName(ini_path), 0)
        End If
        Set GUIF_Object = CreateObject("Scripting.FileSystemObject")
        Set GUIF_file = GUIF_Object.CreateTextFile(ini_path, True, True)
        GUIF_file.Write ReturnEncoding
        GUIF_file.Close
        
    Else
        Set GUIF_Object = CreateObject("Scripting.FileSystemObject")
        Set GUIF_file = GUIF_Object.CreateTextFile(ini_path, True, True)
        GUIF_file.Write ""
        GUIF_file.Close
    End If
    
    GreatUnicodeIniFile = True
    
GreatUnicodeIniFileErr:
    GreatUnicodeIniFile = False
End Function
'以下两个函数,读/写ini文件,固定节点setting,in_key为写入/读取的主键
'仅仅针对是非值
'Y：yes,N：no,E：error
Public Function GetIniTF(ByVal AppName As String, ByVal In_Key As String, Optional ByVal ini_path As String) As Integer
    'GetIniTF("maincenter", "openfloder")
    On Error GoTo GetIniTFErr
    GetIniTF = True
    Dim GetStr As String, dwSize As Long, str_null As String
    
    GetStr = String(256, 0)
    ini_path = App_path & "\OX163setup.ini" & vbNullChar
    AppName = AppName & vbNullChar
    In_Key = In_Key & vbNullChar
    str_null = "" & vbNullChar
    dwSize = GetPrivateProfileString(StrPtr(AppName), StrPtr(In_Key), StrPtr(str_null), StrPtr(GetStr), 256, StrPtr(ini_path))
    GetStr = Left(GetStr, dwSize)
    If CBool(GetStr) = True Then
        GetIniTF = True
        GetStr = ""
    Else
         GetIniTF = False
    End If
    Exit Function
GetIniTFErr:
    OX_Global_Err_Num = err.Number
    err.Clear
    GetIniTF = "error"
    GetStr = ""
End Function

Public Sub WriteIniTF(ByVal AppName As String, ByVal In_Key As String, ByVal In_Data As Boolean, Optional ByVal ini_path As String)
    On Error GoTo WriteIniTFErr
    Dim WriteIniTF_Cstr_tf As String
    If In_Data = True Then
        WriteIniTF_Cstr_tf = "True"
    Else
        WriteIniTF_Cstr_tf = "False"
    End If
    
    Dim WIS_lp As Long
    AppName = AppName & vbNullChar
    In_Key = In_Key & vbNullChar
    WriteIniTF_Cstr_tf = WriteIniTF_Cstr_tf & vbNullChar
    ini_path = App_path & "\OX163setup.ini" & vbNullChar
    GreatUnicodeIniFile App_path & "\OX163setup.ini"
    WIS_lp = WritePrivateProfileString(StrPtr(AppName), StrPtr(In_Key), StrPtr(WriteIniTF_Cstr_tf), StrPtr(ini_path))
    Exit Sub
WriteIniTFErr:
    OX_Global_Err_Num = err.Number
    err.Clear
End Sub

'以下两个函数,读/写ini文件,不固定节点,in_key为写入/读取的主键
'针对字符串值
'空值表示出错
Public Function GetIniStr(ByVal AppName As String, ByVal In_Key As String, Optional ByVal ini_path As String) As String
    On Error GoTo GetIniStrErr
    If Trim(In_Key) = "" Then
        GoTo GetIniStrErr
    End If
    Dim GetStr As String, dwSize As Long, str_null As String
    GetStr = Space$(2048)
    ini_path = App_path & "\OX163setup.ini" & vbNullChar
    AppName = AppName & vbNullChar
    In_Key = In_Key & vbNullChar
    str_null = "" & vbNullChar
    dwSize = GetPrivateProfileString(StrPtr(AppName), StrPtr(In_Key), StrPtr(str_null), StrPtr(GetStr), 2048, StrPtr(ini_path))
    GetStr = Left(GetStr, dwSize)
    If GetStr = "" Then
        GoTo GetIniStrErr
    Else
        GetIniStr = GetStr
        GetStr = ""
    End If
    Exit Function
GetIniStrErr:
    OX_Global_Err_Num = err.Number
    err.Clear
    GetIniStr = ""
    GetStr = ""
End Function

Public Sub WriteIniStr(ByVal AppName As String, ByVal In_Key As String, ByVal In_Data As String, Optional ByVal ini_path As String)
    On Error GoTo WriteIniStrErr
    If Trim(In_Key) = "" Or Trim(AppName) = "" Then
        GoTo WriteIniStrErr
    Else
        Dim WIS_lp As Long
        AppName = AppName & vbNullChar
        In_Key = In_Key & vbNullChar
        In_Data = In_Data & vbNullChar
        ini_path = App_path & "\OX163setup.ini" & vbNullChar
        GreatUnicodeIniFile App_path & "\OX163setup.ini"
        WIS_lp = WritePrivateProfileString(StrPtr(AppName), StrPtr(In_Key), StrPtr(In_Data), StrPtr(ini_path))
    End If
    Exit Sub
WriteIniStrErr:
    OX_Global_Err_Num = err.Number
    err.Clear
End Sub

'-------------------------------------------------------------------------
'保存和读取url历史记录配置文件--------------------------------------------

Public Function GetUnicodeIniStr(ByVal AppName As String, ByVal In_Key As String, ByVal ini_path As String) As String
    On Error GoTo GetIniStrErr
    If Trim(In_Key) = "" Then
        GoTo GetIniStrErr
    End If
    Dim GetStr As String, dwSize As Long, str_null As String
    GetStr = Space$(2048)
    ini_path = ini_path & vbNullChar
    AppName = AppName & vbNullChar
    In_Key = In_Key & vbNullChar
    str_null = "" & vbNullChar
    dwSize = GetPrivateProfileString(StrPtr(AppName), StrPtr(In_Key), StrPtr(""), StrPtr(GetStr), 2048, StrPtr(ini_path))
    GetStr = Left(GetStr, dwSize)
    If GetStr = "" Then
        GoTo GetIniStrErr
    Else
        GetUnicodeIniStr = GetStr
        GetStr = ""
    End If
    Exit Function
GetIniStrErr:
    err.Clear
    GetUnicodeIniStr = ""
    GetStr = ""
End Function

Public Sub WriteUnicodeIni(ByVal AppName As String, ByVal In_Key As String, ByVal In_Data As String, ByVal ini_path As String)
    On Error GoTo WriteIniStrErr
    If Trim(In_Key) = "" Or Trim(AppName) = "" Then
        GoTo WriteIniStrErr
    Else
        Dim WIS_lp As Long
        AppName = AppName & vbNullChar
        In_Key = In_Key & vbNullChar
        In_Data = In_Data & vbNullChar
        GreatUnicodeIniFile ini_path
        ini_path = ini_path & vbNullChar
        WIS_lp = WritePrivateProfileString(StrPtr(AppName), StrPtr(In_Key), StrPtr(In_Data), StrPtr(ini_path))
        
    End If
    Exit Sub
WriteIniStrErr:
    err.Clear
End Sub

'-------------------------------------------------------------------------
'-------------------------------解Gzip压缩数组----------------------------

Public Function UnCompressGzipByte(ByteArray() As Byte) As Variant
    Dim BufferSize  As Long
    Dim buffer()    As Byte
    Dim lReturn    As Long
    
    Dim outUsed    As Long
    Dim inUsed      As Long
    
    '创建解压缩后的缓存
    CopyMemory BufferSize, ByteArray(0), &H8 'OFFSET
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
    ReDim buffer(BufferSize) As Byte
    
    '创建解压缩进程
    Dim contextHandle As Long: InitDecompression
    CreateDecompression contextHandle, 1  '创建
    
    '解压缩数据
    lReturn = Decompress(ByVal contextHandle, ByteArray(0), UBound(ByteArray) + 1, buffer(0), BufferSize, inUsed, outUsed)
    
    DestroyDecompression contextHandle
    
    '删除重复的数据
    ReDim Preserve ByteArray(0 To outUsed - 1)
    CopyMemory ByteArray(0), buffer(0), outUsed
End Function

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-----------------返回Windows的Non - Unicode设定组------------------------
Public Function GetOSLCID() As Integer
    Dim sysLCID As Long
    
    'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Nls\Language
    'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Nls\Locale
    
    '说明:
    'GetSystemDefaultLCID API将返回Windows的Non - Unicode设定 StrConv的最后一个参数有可能用到
    '
    '例如:
    'LOCALE_ILANGUAGE: 0804
    'LOCALE_SLANGUAGE: Chinese (PRC)
    'LOCALE_SENGLANGUAGE: Chinese
    'LOCALE_SABBREVLANGNAME: CHS
    'LOCALE_SNATIVELANGNAME: 中文 (简体)
    'LOCALE_ICOUNTRY: 86
    'LOCALE_SCOUNTRY: People 's Republic of China
    'LOCALE_SENGCOUNTRY: People 's Republic of China
    'LOCALE_SABBREVCTRYNAME: CHN
    'LOCALE_SNATIVECTRYNAME: 中华人民共和国
    'LOCALE_IDEFAULTLANGUAGE: 0804
    'LOCALE_IDEFAULTCOUNTRY: 86
    'LOCALE_IDEFAULTCODEPAGE: 936
    
    sysLCID = GetSystemDefaultLCID
    '&H409(us-英文)&H809(gb-英文)  &HC09(au-英文) &H1009(ca-英文)
    '&H004(zh-中文)&H404(tw-Big5)&H804(cn-GBK/GB)&HC04(hk-Big5) &H1004(sg-GBK)
    '&H804=2052 &H404=1028 &H409=1033 &H809=2057
    
    'eslDefault 0
    'eslHongKong = &HC04     '3076 SAR   950 ZHH
    'eslMacau = &H1404       '5124 SAR   950 ZHM
    'eslChinese = &H804      '2052       936 CHS
    'eslSingapore = &H1004   '4100       936 ZHI
    'eslTaiwan = &H404       '1028       950 CHT
    'eslEnglish = &H409      '1033
    'eslJapanese = &H411     '1041 Japan 932 JAPAN
    'eslKorean = &H412       '1042 Korea Unicode only KOREA
    
    
    If sysLCID = &H804 Or sysLCID = &H4 Or sysLCID = &H1004 Then
        GetOSLCID = 1 '中文简体
    ElseIf sysLCID = &H404 Or sysLCID = &HC04 Then
        GetOSLCID = 2 '中文繁体
    ElseIf sysLCID = &H411 Then
        GetOSLCID = 3 '日文
    ElseIf sysLCID = &H412 Then
        GetOSLCID = 4 '韩文
    ElseIf sysLCID Mod 16 = 9 Then
        GetOSLCID = 0 '英文
    Else
        GetOSLCID = 0
    End If
    
End Function

