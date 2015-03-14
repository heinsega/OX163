Attribute VB_Name = "OX_CookiesCtrl"
'-------------------------------------------------------------------------
'-----------------------------OX163 IE控制模块----------------------------
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
'InternetCookie-----------------------------------------------------------
Public Declare Function InternetGetCookie Lib "wininet.dll" Alias "InternetGetCookieA" (ByVal lpszUrlName As String, ByVal lpszCookieName As String, ByVal lpszCookieData As String, ByRef lpdwSize As Long) As Boolean
Public Declare Function InternetSetCookie Lib "wininet.dll" Alias "InternetSetCookieA" (ByVal lpszUrlName As String, ByVal lpszCookieName As String, ByVal lpszCookieData As String) As Boolean


'-------------------------------------------------------------------------
'get&set cookies and Referer info-----------------------------------------
'-------------------------------------------------------------------------
Public Function GetCookie(ByVal InternetGetCookie_URL) As String
    Dim buf_Cookies As String * 5000, cLen As Long
    cLen = 5000
    Call InternetGetCookie(InternetGetCookie_URL, vbNullString, buf_Cookies, cLen)
    GetCookie = Left(buf_Cookies, cLen)
End Function

Public Function SetCookie(ByVal InternetSetCookie_URL, ByVal InternetSetCookieString) As Boolean
    On Error Resume Next
    InternetSetCookie_URL = GetUrlRoot(InternetSetCookie_URL)
    If InternetSetCookie_URL = "" Then
        SetCookie = False
        Exit Function
    End If
    
    Dim SetCookie_temp, split_str, i
    Call CleanCookie(InternetSetCookie_URL, "all") 'all全部，非all清除特定以“;”号间隔
    
    If InStr(LCase(InternetSetCookieString), "cookie:") = 1 Then InternetSetCookieString = Mid(InternetSetCookieString, 8)
    split_str = Split(InternetSetCookieString, ";")
    For i = 0 To UBound(split_str)
        Call InternetSetCookie(InternetSetCookie_URL, vbNullString, Trim(split_str(i)))
    Next
    SetCookie = True
End Function

'全部或指定cookie值赋空,=""
Public Function CleanCookie(ByVal InternetCleanCookie_URL As String, ByVal InternetCleanCookieString As String) As Boolean
    On Error Resume Next
    Dim SetCookie_temp As String
    Dim split_str, i
    
    InternetCleanCookie_URL = GetUrlRoot(InternetCleanCookie_URL)
    If InternetCleanCookie_URL = "" Then
        CleanCookie = False
        Exit Function
    End If
    
    If InternetCleanCookieString = "all" Then
        SetCookie_temp = GetCookie(InternetCleanCookie_URL)
        split_str = Split(SetCookie_temp, ";")
        
        For i = 0 To UBound(split_str)
            split_str(i) = Trim(split_str(i))
            If InStr(split_str(i), "=") > 1 Then
                InternetSetCookie InternetCleanCookie_URL, vbNullString, Trim(Mid(split_str(i), 1, InStr(split_str(i), "=")))
            End If
        Next
        
    Else
        split_str = Split(InternetCleanCookieString, ";")
        For i = 0 To UBound(split_str)
            split_str(i) = Trim(split_str(i))
            If Len(split_str(i)) <> "" Then
                InternetSetCookie InternetCleanCookie_URL, split_str(i), ""
            End If
        Next
        
    End If
    CleanCookie = True
End Function

'取得GetUrlRoot_Url根目录地址
Public Function GetUrlRoot(ByVal GetUrlRoot_URL As String) As String
    On Error Resume Next
    If InStr(GetUrlRoot_URL, "//") < 1 Then
        GetUrlRoot = ""
        Exit Function
    End If
    GetUrlRoot_URL = Trim(GetUrlRoot_URL)
    GetUrlRoot = Mid(GetUrlRoot_URL, InStr(GetUrlRoot_URL, "//") + 2)
    GetUrlRoot_URL = Mid(GetUrlRoot_URL, 1, InStr(GetUrlRoot_URL, "//") + 1)
    If InStr(GetUrlRoot, "/") > 1 Then
        GetUrlRoot = Mid(GetUrlRoot, 1, InStr(GetUrlRoot, "/"))
    End If
    GetUrlRoot = GetUrlRoot_URL & GetUrlRoot
End Function

Public Function OX_Set_Referer(ByVal Referer_info, ByVal Referer_URL) As String
    On Error GoTo Referer_error
    Dim split_str, split_Referer
    
    Referer_URL = Trim(Referer_URL)
    split_Referer = Split(Referer_info, vbCrLf)
    
    '检查带有OX163自定义格式的html头信息
    Select Case Left(split_Referer(0), Len("parent"))
        
    Case "me" '引用页即下载地址
        split_Referer(0) = "Referer: " & Referer_URL
        
    Case "dir" '引用页为链接自身目录地址
        split_Referer(0) = Left(Referer_URL, InStrRev(Referer_URL, "/"))
        split_Referer(0) = "Referer: " & split_Referer(0)
        
    Case "root" '引用页为链接根目录地址
        split_Referer(0) = Mid(Referer_URL, 1, InStr(Referer_URL, "//") + 1)
        split_str = Mid(Referer_URL, InStr(Referer_URL, "//") + 2)
        split_str = Split(split_str, "/")
        split_Referer(0) = "Referer: " & split_Referer(0) & split_str(0) & "/"
        
    Case "parent" '引用页下载地址n级父路径,（如parent2）：根目录root下2级目录地址，http://moe.imouto.org/data/f9/
        Dim Referer_num 'n
        Referer_num = Right(Referer_info, Len(Referer_info) - 6)
        
        If IsNumeric(Referer_num) Then
            Referer_num = Int(Referer_num)
            split_Referer(0) = Mid(Referer_URL, 1, InStr(Referer_URL, "//") + 1)
            split_str = Mid(Referer_URL, InStr(Referer_URL, "//") + 2)
            split_str = Split(split_str, "/")
            
            If Referer_num < 1 Or Referer_num > UBound(split_str) - 1 Then
                split_Referer(0) = "Referer: " & Referer_URL
            Else
                split_Referer(0) = "Referer: " & split_Referer(0) & split_str(0)
                For i = 1 To Referer_num
                    split_Referer(0) = split_Referer(0) & "/" & split_str(i)
                Next
                split_Referer(0) = split_Referer(0) & "/"
            End If
        Else
            split_Referer(0) = "Referer: " & Referer_URL
        End If
        '其他判断
    Case Else
        If InStr(LCase(split_Referer(0)), "http://") = 1 Then
            split_Referer(0) = "Referer: " & split_Referer(0)
        End If
    End Select
    
    '格式化每行html头信息,分离cookies并设置信息
    For i = 0 To UBound(split_Referer)
        split_Referer(i) = Trim(split_Referer(i))
        If Left(LCase(split_Referer(i)), 7) = "cookie:" Then
            'set cookies
            split_Referer(i) = Trim(Mid(split_Referer(i), 8))
            Call SetCookie(Referer_URL, split_Referer(i))
            split_Referer(i) = ""
        ElseIf split_Referer(i) = "" Or InStr(split_Referer(i), ":") < 2 Or Trim(Mid(split_Referer(i), InStr(split_Referer(i), ":") + 1)) = "" Then
            split_Referer(i) = ""
        Else
            split_Referer(i) = split_Referer(i) & vbCrLf
        End If
    Next
    OX_Set_Referer = Join(split_Referer, "")
    
    Do While Right(OX_Set_Referer, 2) = vbCrLf
        OX_Set_Referer = Left(OX_Set_Referer, Len(OX_Set_Referer) - 2)
    Loop
    
    If InStr(LCase(OX_Set_Referer), "user-agent:") <> 1 And InStr(LCase(OX_Set_Referer), vbCrLf & "user-agent:") < 1 Then OX_Set_Referer = OX_Set_Referer & vbCrLf & "User-Agent: " & sysSet.Customize_UA
    If sysSet.Cache_no_cache = 1 Then OX_Set_Referer = OX_Set_Referer & vbCrLf & "Pragma: no-cache"
    If sysSet.Cache_no_store = 1 Then OX_Set_Referer = OX_Set_Referer & vbCrLf & "Cache-Control: no-store"
    Exit Function
    
Referer_error:
    OX_Set_Referer = sysSet.OX_HTTP_Head
End Function
