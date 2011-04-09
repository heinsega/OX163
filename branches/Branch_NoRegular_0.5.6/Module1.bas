Attribute VB_Name = "Declare_Function"
'使用XP风格
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

'程序窗口重在最前面-------------------------------------
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40

'系统文件夹-----------------------------------------
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'读取ini配置-----------------------------------------
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'写入ini配置
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'得到路径是否可写-----------------------------------------
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

'进程优先级-----------------------------------------
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
'Const IDLE_PRIORITY_CLASS = &H40
'Const BELOW_NORMAL_PRIORITY_CLASS = &H4000
'Const NORMAL_PRIORITY_CLASS = &H20
'Const ABOVE_NORMAL_PRIORITY_CLASS = &H8000
'Const HIGH_PRIORITY_CLASS = &H80
'Const REALTIME_PRIORITY_CLASS = &H100

'调用shell保存路径-----------------------------------------
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Const MAX_PATH = 500

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_NEWDIALOGSTYLE = &H40

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

'----------------------------------最小化系统托盘---------------------------------------------------

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconW" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'----------------------------------打开IE-----------------------------------------------------------
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'----------------------------------InternetCookie---------------------------------------------------

Public Declare Function InternetGetCookie Lib "wininet.dll" Alias "InternetGetCookieA" (ByVal lpszUrlName As String, ByVal lpszCookieName As String, ByVal lpszCookieData As String, ByRef lpdwSize As Long) As Long


'-------------------------------解Gzip压缩数组------------------------------------------
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long) '已经定义
Private Declare Function InitDecompression Lib "gzip.dll" () As Long
Private Declare Function CreateDecompression Lib "gzip.dll" (ByRef context As Long, ByVal flags As Long) As Long
Private Declare Function DestroyDecompression Lib "gzip.dll" (ByRef context As Long) As Long
Private Declare Function Decompress Lib "gzip.dll" (ByVal context As Long, inBytes As Any, ByVal input_size As Long, outBytes As Any, ByVal output_size As Long, ByRef input_used As Long, ByRef output_used As Long) As Long
'Private Const OFFSET As Long = &H8

''----------------------------------代替CommonDialog--------------------------------------------
'
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

'API调用unicode文件对话框CommonDialog

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
        If Right(lReturn_str, 1) = Chr(0) Then lReturn_str = Left(lReturn_str, Len(lReturn_str) - 1)
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
        If Right(lReturn_str, 1) = Chr(0) Then lReturn_str = Left(lReturn_str, Len(lReturn_str) - 1)
        
        '检查文件后缀名
        file_type_split = Split(file_type, "|")
        Filter = ""
        lReturn = OpenFile.nFilterIndex - 1
        Filter = file_type_split(lReturn)
        If LCase(Right(lReturn_str, Len(Filter))) <> LCase(Filter) Then lReturn_str = lReturn_str & Filter
        
        ShowSaveFileDialog = lReturn_str
    End If
    
End Function

'-------------XP风格-----start-----------------------------------------------------------

Public Function InitCommonControlsVB() As Boolean
    On Error Resume Next
    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = &H200
    End With
    InitCommonControlsEx iccex
    InitCommonControlsVB = (Err.Number = 0)
    On Error GoTo 0
End Function


Sub Main()
    InitCommonControlsVB
    
    Dim CurrentProcesshWnd As Long
    CurrentProcesshWnd = GetCurrentProcess
    Call SetPriorityClass(CurrentProcesshWnd, &H8000)
        
    start_ox163.Show
End Sub


'-------------XP风格--------end--------------------------------------------------------

'get cookies--------------------------------------------------------------------
Public Function GetCookie(ByVal InternetGetCookie_url) As String
    Dim buf_Cookies As String * 256, ret As Long, cLen As Long
    cLen = 256
    ret = InternetGetCookie(InternetGetCookie_url, "", buf_Cookies, cLen)
    GetCookie = Left(buf_Cookies, cLen)
End Function

Public Function GetShortName(ByVal sLongFileName As String) As String
    On Error Resume Next
    GetShortName = ""
    Dim GetShortName_Fso
    
    Set GetShortName_Fso = CreateObject("Scripting.FileSystemObject")
    
    Err.Number = 0
    GetShortName = GetShortName_Fso.GetFile(sLongFileName).ShortPath
    
    If Err.Number <> 0 Then
        Err.Number = 0
        GetShortName = GetShortName_Fso.GetFolder(sLongFileName).ShortPath
    End If
    
    Set GetShortName_Fso = Nothing
    
    If GetShortName = "" Then GetShortName = sLongFileName
End Function



Public Function GetSysDir() As String
    Dim strBuf As String
    Dim lngBuf As Long
    
    strBuf = Space(255)
    lngBuf = 255
    
    lngBuf = GetSystemDirectory(ByVal strBuf, lngBuf)
    
    GetSysDir = Left$(strBuf, lngBuf)
End Function

'------------------------------------------------------------

'以下两个函数,读/写ini文件,固定节点setting,in_key为写入/读取的主键
'仅仅针对是非值
'Y：yes,N：no,E：error

Public Function GetIniTF(ByVal AppName As String, ByVal In_Key As String) As Boolean
    On Error GoTo GetIniTFErr
    GetIniTF = True
    Dim GetStr As String
    GetStr = VBA.String(128, 0)
    GetPrivateProfileString AppName, In_Key, "", GetStr, 256, App_path & "\OX163setup.ini"
    GetStr = VBA.Replace(GetStr, VBA.Chr(0), "")
    If CBool(GetStr) = True Then
        GetIniTF = True
        GetStr = ""
    Else
        GoTo GetIniTFErr
    End If
    Exit Function
GetIniTFErr:
    Err.Clear
    GetIniTF = False
    GetStr = ""
End Function

Public Sub WriteIniTF(ByVal AppName As String, ByVal In_Key As String, ByVal In_Data As Boolean)
    On Error GoTo WriteIniTFErr
    If In_Data = True Then
        WritePrivateProfileString AppName, In_Key, "True", App_path & "\OX163setup.ini"
    Else
        WritePrivateProfileString AppName, In_Key, "False", App_path & "\OX163setup.ini"
    End If
    Exit Sub
WriteIniTFErr:
    Err.Clear
End Sub


'以下两个函数,读/写ini文件,不固定节点,in_key为写入/读取的主键
'针对字符串值
'空值表示出错
Public Function GetIniStr(ByVal AppName As String, ByVal In_Key As String) As String
    On Error GoTo GetIniStrErr
    If VBA.Trim(In_Key) = "" Then
        GoTo GetIniStrErr
    End If
    Dim GetStr As String
    GetStr = VBA.String(128, 0)
    GetPrivateProfileString AppName, In_Key, "", GetStr, 256, App_path & "\OX163setup.ini"
    GetStr = VBA.Replace(GetStr, VBA.Chr(0), "")
    If GetStr = "" Then
        GoTo GetIniStrErr
    Else
        GetIniStr = GetStr
        GetStr = ""
    End If
    Exit Function
GetIniStrErr:
    Err.Clear
    GetIniStr = ""
    GetStr = ""
End Function

Public Sub WriteIniStr(ByVal AppName As String, ByVal In_Key As String, ByVal In_Data As String)
    On Error GoTo WriteIniStrErr
    If VBA.Trim(In_Key) = "" Or VBA.Trim(AppName) = "" Then
        GoTo WriteIniStrErr
    Else
        WritePrivateProfileString AppName, In_Key, In_Data, App_path & "\OX163setup.ini"
    End If
    Exit Sub
WriteIniStrErr:
    Err.Clear
End Sub

'可选路径
Public Function GetUrlStr(ByVal AppName As String, ByVal In_Key As String, ByVal url_str_path As String) As String
    On Error GoTo GetIniStrErr
    If VBA.Trim(In_Key) = "" Then
        GoTo GetIniStrErr
    End If
    Dim GetStr As String
    GetStr = VBA.String(128, 0)
    GetPrivateProfileString AppName, In_Key, "", GetStr, 256, url_str_path
    GetStr = VBA.Replace(GetStr, VBA.Chr(0), "")
    If GetStr = "" Then
        GoTo GetIniStrErr
    Else
        GetUrlStr = GetStr
        GetStr = ""
    End If
    Exit Function
GetIniStrErr:
    Err.Clear
    GetUrlStr = ""
    GetStr = ""
End Function

Public Sub WriteUrlStr(ByVal AppName As String, ByVal In_Key As String, ByVal In_Data As String, ByVal url_str_path As String)
    On Error GoTo WriteIniStrErr
    If VBA.Trim(In_Key) = "" Or VBA.Trim(AppName) = "" Then
        GoTo WriteIniStrErr
    Else
        WritePrivateProfileString AppName, In_Key, In_Data, url_str_path
    End If
    Exit Sub
WriteIniStrErr:
    Err.Clear
End Sub

Public Function URLEncode(ByVal vstrIn)
    On Error Resume Next
    strReturn = ""
    vstrIn = Replace(vstrIn, "%", "%25")
    Dim i
    For i = 1 To Len(vstrIn)
        ThisChr = Mid(vstrIn, i, 1)
        If Abs(Asc(ThisChr)) < &HFF Then
            strReturn = strReturn & ThisChr
        Else
            innerCode = Asc(ThisChr)
            If innerCode < 0 Then
                innerCode = innerCode + &H10000
            End If
            Hight8 = (innerCode And &HFF00) \ &HFF
            Low8 = innerCode And &HFF
            strReturn = strReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
        End If
    Next
    strReturn = Replace(strReturn, "!", "%21")
    strReturn = Replace(strReturn, Chr(34), "%22")
    strReturn = Replace(strReturn, "#", "%20")
    strReturn = Replace(strReturn, "$", "%24")
    strReturn = Replace(strReturn, "&", "%26")
    strReturn = Replace(strReturn, "'", "%27")
    strReturn = Replace(strReturn, "(", "%28")
    strReturn = Replace(strReturn, ")", "%29")
    strReturn = Replace(strReturn, "*", "%2A")
    strReturn = Replace(strReturn, "+", "%2B")
    strReturn = Replace(strReturn, ",", "%2C")
    strReturn = Replace(strReturn, ".", "%2E")
    strReturn = Replace(strReturn, "/", "%2F")
    strReturn = Replace(strReturn, ":", "%3A")
    strReturn = Replace(strReturn, ";", "%3B")
    strReturn = Replace(strReturn, "<", "%3C")
    strReturn = Replace(strReturn, "=", "%3D")
    strReturn = Replace(strReturn, ">", "%3E")
    strReturn = Replace(strReturn, "?", "%3F")
    strReturn = Replace(strReturn, "@", "%40")
    strReturn = Replace(strReturn, "[", "%5B")
    strReturn = Replace(strReturn, "\", "%5C")
    strReturn = Replace(strReturn, "]", "%5D")
    strReturn = Replace(strReturn, "^", "%5E")
    strReturn = Replace(strReturn, "`", "%60")
    strReturn = Replace(strReturn, "{", "%7B")
    strReturn = Replace(strReturn, "|", "%7C")
    strReturn = Replace(strReturn, "}", "%7D")
    strReturn = Replace(strReturn, "~", "%7E")
    strReturn = Replace(strReturn, Chr(32), "%20")
    URLEncode = strReturn
End Function

'Public Function URLDecode(strURL As String) As String
'    Dim strChar As String
'    Dim strText As String
'    Dim strTemp As String
'    Dim strRet As String
'    Dim LngNum As Long
'    Dim I As Integer
'    For I = 1 To Len(strURL)
'        strChar = Mid(strURL, I, 1)
'        Select Case strChar
'          Case "+"
'            strText = strText & " "
'          Case "%"
'            strTemp = Mid(strURL, I + 1, 2) '暂时取2位
'            LngNum = Val("&H" & strTemp)
'            '>127即为汉字
'            If LngNum < 128 Then
'                strRet = Chr(LngNum)
'                I = I + 2
'            Else
'                strTemp = strTemp & Mid(strURL, I + 4, 2)
'                strRet = Chr(Val("&H" & strTemp))
'                I = I + 5
'            End If
'            strText = strText & strRet
'          Case Else
'            strText = strText & strChar
'        End Select
'    Next
'    URLDecode = strText
'End Function

'Public Function URLDecode(strURL)
'On Error Resume Next
'    Dim i
'
'    If InStr(strURL, "%") = 0 Then
'        URLDecode = strURL
'        Exit Function
'    End If
'
'    For i = 1 To Len(strURL)
'        If Mid(strURL, i, 1) = "%" Then
'            If Eval("&H" & Mid(strURL, i + 1, 2)) > 127 Then
'                URLDecode = URLDecode & Chr(Eval("&H" & Mid(strURL, i + 1, 2) & Mid(strURL, i + 4, 2)))
'                i = i + 5
'            Else
'                URLDecode = URLDecode & Chr(Eval("&H" & Mid(strURL, i + 1, 2)))
'                i = i + 2
'            End If
'        Else
'            URLDecode = URLDecode & Mid(strURL, i, 1)
'        End If
'    Next
'End Function

Public Function UTF8EncodeURI(ByVal szInput)
    On Error Resume Next
    Dim wch, uch, szRet
    Dim x
    Dim nAsc, nAsc2, nAsc3
    
    If szInput = "" Then
        UTF8EncodeURI = szInput
        Exit Function
    End If
    
    For x = 1 To Len(szInput)
        wch = Mid(szInput, x, 1)
        nAsc = AscW(wch)
        
        If nAsc < 0 Then nAsc = nAsc + 65536
        
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
    
    UTF8EncodeURI = szRet
End Function

'---------------------------------------------------------------

Public Function GetEncoding(ByVal fileName) As String
    On Error GoTo Err
    
    Dim fBytes(1) As Byte, freeNum As Integer
    freeNum = FreeFile
    
    Open fileName For Binary Access Read As #freeNum
    Get #freeNum, , fBytes(0)
    Get #freeNum, , fBytes(1)
    Close #freeNum
    
    If fBytes(0) = &HFF And fBytes(1) = &HFE Then GetEncoding = "Unicode"
    If fBytes(0) = &HFE And fBytes(1) = &HFF Then GetEncoding = "UnicodeBigEndian"
    If fBytes(0) = &HEF And fBytes(1) = &HBB Then GetEncoding = "UTF8"
Err:
End Function

Public Sub FileToUTF8(fileName As String)
    Dim fBytes() As Byte, uniString As String, freeNum As Integer
    Dim ADO_Stream As Object
    
    freeNum = FreeFile
    
    ReDim fBytes(FileLen(fileName))
    Open fileName For Binary Access Read As #freeNum
    Get #freeNum, , fBytes
    Close #freeNum
    
    uniString = StrConv(fBytes, vbUnicode)
    
    Set ADO_Stream = CreateObject("ADODB.Stream")
    With ADO_Stream
        .Type = 2
        .mode = 3
        .Charset = "utf-8"
        .Open
        .WriteText uniString
        .SaveToFile fileName, 2
        .Close
    End With
    Set ADO_Stream = Nothing
End Sub
'---------------------------------------------------------------

Public Sub Proxy_set()
    '-------------------------------------------------------------------------
    Form1.fast_down.Proxy = ""
    Form1.fast_down.username = ""
    Form1.fast_down.password = ""
    Form1.Proxy_img(0).Visible = False
    Form1.Proxy_img(1).Visible = False
    Form1.Proxy_img(2).Visible = False
    
    Select Case sysSet.proxy_A_type
    Case 1
        Form1.fast_down.AccessType = icDirect
    Case 2
        sysSet.proxy_A = Trim(Replace(Replace(sysSet.proxy_A, Chr(10), ""), Chr(13), ""))
        If Len(sysSet.proxy_A) > 4 Then
            Form1.fast_down.AccessType = icNamedProxy
            Form1.fast_down.Proxy = sysSet.proxy_A
            Form1.Proxy_img(1).Visible = True
            sysSet.proxy_A_user = Trim(Replace(Replace(sysSet.proxy_A_user, Chr(10), ""), Chr(13), ""))
            sysSet.proxy_A_pw = Trim(Replace(Replace(sysSet.proxy_A_pw, Chr(10), ""), Chr(13), ""))
            If Len(sysSet.proxy_A_user) > 0 Then Form1.fast_down.username = sysSet.proxy_A_user
            If Len(sysSet.proxy_A_pw) > 0 Then Form1.fast_down.password = sysSet.proxy_A_pw
        Else
            Form1.fast_down.AccessType = icUseDefault
        End If
        
    Case Else
        Form1.fast_down.AccessType = icUseDefault
    End Select
    
    '-------------------------------------------------------------------------
    Form1.Inet1.Proxy = ""
    Form1.Inet1.username = ""
    Form1.Inet1.password = ""
    Form1.check_header.Proxy = ""
    Form1.check_header.username = ""
    Form1.check_header.password = ""
    
    Select Case sysSet.proxy_B_type
    Case 1
        Form1.Inet1.AccessType = icDirect
        Form1.check_header.AccessType = icDirect
    Case 2
        sysSet.proxy_B = Trim(Replace(Replace(sysSet.proxy_B, Chr(10), ""), Chr(13), ""))
        If Len(sysSet.proxy_B) > 4 Then
            Form1.Inet1.AccessType = icNamedProxy
            Form1.Inet1.Proxy = sysSet.proxy_B
            Form1.check_header.AccessType = icNamedProxy
            Form1.check_header.Proxy = sysSet.proxy_B
            Form1.Proxy_img(2).Visible = True
            sysSet.proxy_B_user = Trim(Replace(Replace(sysSet.proxy_B_user, Chr(10), ""), Chr(13), ""))
            sysSet.proxy_B_pw = Trim(Replace(Replace(sysSet.proxy_B_pw, Chr(10), ""), Chr(13), ""))
            If Len(sysSet.proxy_B_user) > 0 Then Form1.Inet1.username = sysSet.proxy_B_user: Form1.check_header.username = sysSet.proxy_B_user
            If Len(sysSet.proxy_B_pw) > 0 Then Form1.Inet1.password = sysSet.proxy_B_pw: Form1.check_header.password = sysSet.proxy_B_pw
        Else
            Form1.Inet1.AccessType = icUseDefault
            Form1.check_header.AccessType = icUseDefault
        End If
        
    Case Else
        Form1.Inet1.AccessType = icUseDefault
        Form1.check_header.AccessType = icUseDefault
    End Select
    
    If Form1.Proxy_img(1).Visible = True And Form1.Proxy_img(2).Visible = True Then
        Form1.Proxy_img(0).Visible = True
        Form1.Proxy_img(1).Visible = False
        Form1.Proxy_img(2).Visible = False
    End If
    
End Sub


'---------------------调用shell选择保存目录---------start------------------------------------------
Public Function GetFolder(ByVal title As String, ByVal start As String, ByVal newfolder As Boolean) As String
    Dim BI As BROWSEINFO, pidl As Long, lpSelPath As Long
    Dim spath As String * MAX_PATH
    
    'fill in the info it needs
    With BI
        .hOwner = GetForegroundWindow
        .pidlRoot = 0
        .lpszTitle = title
        .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
        .ulFlags = BIF_RETURNONLYFSDIRS
        If newfolder = True Then .ulFlags = BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE
        lpSelPath = LocalAlloc(LPTR, LenB(start) + 1)
        CopyMemory ByVal lpSelPath, ByVal start, LenB(start) + 1
        .lParam = lpSelPath
    End With
    
    'get the idlist long from the returned folder
    pidl = SHBrowseForFolder(BI)
    
    'do then if they clicked ok
    If pidl Then
        If SHGetPathFromIDList(pidl, spath) Then
            'next line is the returned folder
            GetFolder = Left$(spath, InStr(spath, vbNullChar) - 1)
        End If
        Call CoTaskMemFree(pidl)
    Else
        'user clicked cancel
    End If
    
    Call LocalFree(lpSelPath)
    
End Function

'this seems to happen before the box comes up and when a folder is clicked on within it
Public Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Dim spath As String, bFlag As Long
    
    spath = Space$(MAX_PATH)
    
    Select Case uMsg
    Case BFFM_INITIALIZED
        'browse has been initialized, set the start folder
        Call SendMessage(hWnd, BFFM_SETSELECTION, 1, ByVal lpData)
    Case BFFM_SELCHANGED
        If SHGetPathFromIDList(lParam, spath) Then
            spath = Left(spath, InStr(1, spath, Chr(0)) - 1)
        End If
    End Select
    
End Function

Public Function FARPROC(pfn As Long) As Long
    FARPROC = pfn
End Function
'---------------------调用shell选择保存目录------end---------------------------------------------


'-------------------------------解Gzip压缩数组------------------------------------------
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
