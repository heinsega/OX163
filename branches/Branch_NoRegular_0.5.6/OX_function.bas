Attribute VB_Name = "OX_function"
'-------------------------------------------------------------------------
'debug调试用函数----------------------------------------------------------
Public Sub OX_Debug_File(ByVal Debug_file_String As String)
    If Dir(App_path & "\debug", vbDirectory) = "" Then
        MkDir App_path & "\debug"
    End If
    Dim FileNumber
    FileNumber = FreeFile ' 取得未使用的文件号。
    Open App_path & "\debug\OX163_Debug_File(" & Now() & ").txt" For Output As #FileNumber   ' 创建文件名。
    Write #FileNumber, Debug_file_String ' 输出文本至文件中。
    Close #FileNumber   ' 关闭文件。
End Sub

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'--------------------------------OX163常用函数----------------------------
'-------------------------------------------------------------------------
'加载文本文件（可自定义字符集）-------------------------------------------
Public Function load_normal_file(file_name, unicode_charset) As String
    On Error Resume Next
    Dim fileline As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(file_name, 1, False, unicode_charset)
    load_normal_file = file.Readall
    file.Close
    Set file = Nothing
    Set fso = Nothing
End Function

'-------------------------------------------------------------------------
'加载ANSI脚本（现在默认使用）---------------------------------------------
'-------------------------------------------------------------------------
Public Function load_Script(file_name) As String
    On Error Resume Next
    Dim fileline As String
    Dim fso As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(file_name, 1, False, 0)
    load_Script = file.Readall
    file.Close
    Set fso = Nothing
    
    'Open file_name For Input As #5
    'Do While Not EOF(5)
    'Line Input #5, fileline
    'load_script = load_script + fileline & vbCrLf
    'DoEvents
    'Loop
    'Close #5
    'load_script = Left$(load_script, Len(load_script) - 2)
End Function

'-------------------------------------------------------------------------
'格式化图片尺寸文本（1920*1080）------------------------------------------
'-------------------------------------------------------------------------
Public Function fix_Pix(ByVal pix_str) As String
    fix_Pix = ""
    pix_str = Split(pix_str, "x")
    For i = 0 To UBound(pix_str)
        DoEvents
        fix_Pix = fix_Pix & Format$(Int(pix_str(i)), "0000") & "x"
    Next i
    fix_Pix = Mid$(fix_Pix, 1, Len(fix_Pix) - 1)
End Function

'-------------------------------------------------------------------------
'将检查是否为文件名-------------------------------------------------------
'-------------------------------------------------------------------------
Public Function is_fileName(ByVal file_name As String) As Boolean
    is_fileName = True
    If InStr(file_name, Chr(92)) > 0 Then is_fileName = False: Exit Function
    If InStr(file_name, Chr(47)) > 0 Then is_fileName = False: Exit Function
    If InStr(file_name, Chr(34)) > 0 Then is_fileName = False: Exit Function
    If InStr(file_name, Chr(63)) > 0 Then is_fileName = False: Exit Function
    If InStr(file_name, Chr(58)) > 0 Then is_fileName = False: Exit Function
    If InStr(file_name, Chr(42)) > 0 Then is_fileName = False: Exit Function
    If InStr(file_name, Chr(60)) > 0 Then is_fileName = False: Exit Function
    If InStr(file_name, Chr(62)) > 0 Then is_fileName = False: Exit Function
    If InStr(file_name, Chr(124)) > 0 Then is_fileName = False: Exit Function
    
    If Left(file_name, 1) = "." Then is_fileName = False: Exit Function
    If Right(file_name, 1) = "." Then is_fileName = False: Exit Function
End Function

'-------------------------------------------------------------------------
'将url格式化为可保存的文件名格式------------------------------------------
'-------------------------------------------------------------------------
Public Function rename_URL(ByVal old_url)
    '＼／＂？：＊＜＞｜
    '\/"?:*<>|
    If IsNull(old_url) Or IsEmpty(old_url) Then
        rename_URL = ""
        Exit Function
    End If
    If Left(old_url, 1) = "." Then old_url = Mid$(old_url, 2)
    
    code_E = Array("＼", "／", Chr(-23646), "？", "：", "＊", "＜", "＞", "｜")
    code_F = Array(Chr(92), Chr(47), Chr(34), Chr(63), Chr(58), Chr(42), Chr(60), Chr(62), Chr(124))
    
    rename_URL = old_url
    
    For i = 0 To 8
        rename_URL = Replace(rename_URL, code_F(i), code_E(i))
    Next
    
End Function

'-------------------------------------------------------------------------
'将url文件名格式格式化为正常url-------------------------------------------
'-------------------------------------------------------------------------
Public Function rename_URLfile(ByVal old_url)
    If IsNull(old_url) Or IsEmpty(old_url) Then
        rename_URLfile = ""
        Exit Function
    End If
    If Left(old_url, 1) = "." Then old_url = Mid$(old_url, 2)
    
    code_E = Array("＼", "／", Chr(-23646), "？", "：", "＊", "＜", "＞", "｜")
    code_F = Array(Chr(92), Chr(47), Chr(34), Chr(63), Chr(58), Chr(42), Chr(60), Chr(62), Chr(124))
    
    rename_URLfile = old_url
    
    For i = 0 To 8
        rename_URLfile = Replace(rename_URLfile, code_E(i), code_F(i))
    Next
    
End Function

'-------------------------------------------------------------------------
'取得文件编码格式---------------------------------------------------------
'-------------------------------------------------------------------------
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
    If fBytes(0) = &HEF And fBytes(1) = &HBB Then GetEncoding = "UTF-8"
Err:
End Function

'-------------------------------------------------------------------------
'文本转换保存为UTF-8格式（暂未使用）--------------------------------------
'-------------------------------------------------------------------------
Public Sub FileTo_UTF8File(fileName As String)
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
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'加载OX163脚本默认函数----------------------------------------------------
Public Sub OX_load_Script_Code(sourceScriptInfo As ScriptInfo, sourceScriptApp As ScriptControl)
    On Error Resume Next
    Dim OX_load_Script_Code_STR As String
    If LCase(Trim(sourceScriptInfo.Language)) = "vbscript" Then
        sourceScriptInfo.Language = "vbscript"
        sourceScriptApp.Language = "vbscript"
        OX_load_Script_Code_STR = in_Script_Code.OX163_vbs_var & load_Script(App_path & "\include\" & sourceScriptInfo.fileName) & in_Script_Code.OX163_vbs_fn
    Else
        sourceScriptInfo.Language = "javascript"
        sourceScriptApp.Language = "javascript"
        OX_load_Script_Code_STR = in_Script_Code.OX163_js_var & load_Script(App_path & "\include\" & sourceScriptInfo.fileName) & in_Script_Code.OX163_js_fn
    End If
    Call sourceScriptApp.AddCode(OX_load_Script_Code_STR)
End Sub

Public Sub load_in_Script_Code()
    On Error Resume Next
    in_Script_Code.OX163_vbs_var = ""
    If Dir(App_path & "\include\OX163_vbs_var.vbs") <> "" Then
    in_Script_Code.OX163_vbs_var = vbCrLf & load_Script(App_path & "\include\OX163_vbs_var.vbs") & vbCrLf
    Else
    in_Script_Code.OX163_vbs_var = vbCrLf & "Dim OX163_urlpage_Referer,OX163_urlpage_Cookies" & vbCrLf
    End If
    
    in_Script_Code.OX163_vbs_fn = ""
    If Dir(App_path & "\include\OX163_vbs_fn.vbs") <> "" Then
    in_Script_Code.OX163_vbs_fn = vbCrLf & load_Script(App_path & "\include\OX163_vbs_fn.vbs") & vbCrLf
    Else
    in_Script_Code.OX163_vbs_fn = vbCrLf & "Function set_urlpagecookies(byVal set_str)" & vbCrLf & "On Error Resume Next" & vbCrLf & "OX163_urlpage_Cookies = set_str" & vbCrLf & "End Function" & vbCrLf
    End If
    
    in_Script_Code.OX163_js_var = ""
    If Dir(App_path & "\include\OX163_js_var.vbs") <> "" Then
    in_Script_Code.OX163_js_var = vbCrLf & load_Script(App_path & "\include\OX163_js_var.vbs") & vbCrLf
    Else
    in_Script_Code.OX163_js_var = vbCrLf & "var OX163_urlpage_Referer='';var OX163_urlpage_Cookies='';" & vbCrLf
    End If
    
    in_Script_Code.OX163_js_fn = ""
    If Dir(App_path & "\include\OX163_js_fn.vbs") <> "" Then
    in_Script_Code.OX163_js_fn = vbCrLf & load_Script(App_path & "\include\OX163_js_fn.vbs") & vbCrLf
    Else
    in_Script_Code.OX163_js_fn = vbCrLf & "function set_urlpagecookies(set_str){OX163_urlpage_Cookies=set_str;}" & vbCrLf
    End If
    
    OX163_WebBrowser_scriptCode = ""
    If Dir(App_path & "\include\OX163_Web_Browser_ctrl.vbs") <> "" Then
        OX163_WebBrowser_scriptCode = load_Script(App_path & "\include\OX163_Web_Browser_ctrl.vbs")
        OX163_WebBrowser_scriptCode = Trim(OX163_WebBrowser_scriptCode)
    End If
End Sub

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'参数调整后重新设置代理服务器设置-----------------------------------------
'-------------------------------------------------------------------------
Public Sub Proxy_set()
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
