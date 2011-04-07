Attribute VB_Name = "OX_function"
'debug调试用函数---------------------------------------------------
Public Sub OX_Debug_File(ByVal Debug_file_String As String)
    Dim FileNumber
    FileNumber = FreeFile ' 取得未使用的文件号。
    Open "C:\OX163_Debug_File(" & Now() & ").txt" For Output As #FileNumber   ' 创建文件名。
    Write #FileNumber, Debug_file_String ' 输出文本至文件中。
    Close #FileNumber   ' 关闭文件。
End Sub

'加载OX163脚本默认函数----------------------------------------------------
Public Sub OX_load_Script_Code(sourceScriptInfo As ScriptInfo, sourceScriptApp As ScriptControl)
    On Error Resume Next
    Dim OX_load_Script_Code_STR As String
    If LCase(Trim(sourceScriptInfo.Language)) = "vbscript" Then
        sourceScriptInfo.Language = "vbscript"
        sourceScriptApp.Language = "vbscript"
        OX_load_Script_Code_STR = in_Script_Code.OX163_vbs_var & load_Script(App.Path & "\include\" & sourceScriptInfo.fileName) & in_Script_Code.OX163_vbs_fn
    Else
        sourceScriptInfo.Language = "javascript"
        sourceScriptApp.Language = "javascript"
        OX_load_Script_Code_STR = in_Script_Code.OX163_js_var & load_Script(App.Path & "\include\" & sourceScriptInfo.fileName) & in_Script_Code.OX163_js_fn
    End If
    Call sourceScriptApp.AddCode(OX_load_Script_Code_STR)
End Sub

Public Sub load_in_Script_Code()
    On Error Resume Next
    in_Script_Code.OX163_vbs_var = ""
    If Dir(App.Path & "\include\OX163_vbs_var.vbs") <> "" Then
    in_Script_Code.OX163_vbs_var = vbCrLf & load_Script(App.Path & "\include\OX163_vbs_var.vbs") & vbCrLf
    Else
    in_Script_Code.OX163_vbs_var = vbCrLf & "Dim OX163_urlpage_Referer,OX163_urlpage_Cookies" & vbCrLf
    End If
    
    in_Script_Code.OX163_vbs_fn = ""
    If Dir(App.Path & "\include\OX163_vbs_fn.vbs") <> "" Then
    in_Script_Code.OX163_vbs_fn = vbCrLf & load_Script(App.Path & "\include\OX163_vbs_fn.vbs") & vbCrLf
    Else
    in_Script_Code.OX163_vbs_fn = vbCrLf & "Function set_urlpagecookies(byVal set_str)" & vbCrLf & "On Error Resume Next" & vbCrLf & "OX163_urlpage_Cookies = set_str" & vbCrLf & "End Function" & vbCrLf
    End If
    
    in_Script_Code.OX163_js_var = ""
    If Dir(App.Path & "\include\OX163_js_var.vbs") <> "" Then
    in_Script_Code.OX163_js_var = vbCrLf & load_Script(App.Path & "\include\OX163_js_var.vbs") & vbCrLf
    Else
    in_Script_Code.OX163_js_var = vbCrLf & "var OX163_urlpage_Referer='';var OX163_urlpage_Cookies='';" & vbCrLf
    End If
    
    in_Script_Code.OX163_js_fn = ""
    If Dir(App.Path & "\include\OX163_js_fn.vbs") <> "" Then
    in_Script_Code.OX163_js_fn = vbCrLf & load_Script(App.Path & "\include\OX163_js_fn.vbs") & vbCrLf
    Else
    in_Script_Code.OX163_js_fn = vbCrLf & "function set_urlpagecookies(set_str){OX163_urlpage_Cookies=set_str;}" & vbCrLf
    End If
    
    OX163_WebBrowser_scriptCode = ""
    If Dir(App.Path & "\include\OX163_Web_Browser_ctrl.vbs") <> "" Then
        OX163_WebBrowser_scriptCode = load_Script(App.Path & "\include\OX163_Web_Browser_ctrl.vbs")
        OX163_WebBrowser_scriptCode = Trim(OX163_WebBrowser_scriptCode)
    End If
End Sub

'OX163常用函数----------------------------------------------------------
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

Public Function load_Script(file_name) As String
    On Error Resume Next
    Dim fileline As String
    
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


Public Function check_Include(ByVal url_str As String) As String
    On Error Resume Next
    
    check_Include = ""
    If Dir(App.Path & "\include\include.txt") = "" Then Exit Function
    
    Dim include_str, include_str1
    
    include_str = load_Script(App.Path & "\include\include.txt")
    If include_str = "" Then Exit Function
    
    include_str = Split(Trim$(include_str), vbCrLf)
    
    For i = 0 To UBound(include_str)
        DoEvents
        If include_str(i) <> "" Then
            include_str1 = Split(include_str(i), "|")
            
            If UBound(include_str1) < 4 Then GoTo next_i
            If Dir(App.Path & "\include\" & include_str1(0)) = "" Then GoTo next_i
            If LCase$(include_str1(1)) <> "vbscript" And LCase$(include_str1(1)) <> "javascript" Then GoTo next_i
            If include_str1(2) = "" Then GoTo next_i
            If LCase$(include_str1(3)) <> "photo" And LCase$(include_str1(3)) <> "album" Then GoTo next_i
            If include_str1(4) = "" Then GoTo next_i
            
            'url_str(输入的网址)
            'include_str1(4)(带有?*等符号的规范网址)
            
            If LCase(url_str) Like LCase(include_str1(4)) Then
                check_Include = include_str1(0) & "|" & include_str1(1) & "|" & include_str1(2) & "|" & include_str1(3) & "|" & url_str
                Exit Function
            End If
            
        End If
        
next_i:
        
    Next i
End Function

Public Function fix_Pix(ByVal pix_str)
    fix_Pix = ""
    pix_str = Split(pix_str, "x")
    For i = 0 To UBound(pix_str)
        DoEvents
        fix_Pix = fix_Pix & Format$(Int(pix_str(i)), "0000") & "x"
    Next i
    fix_Pix = Mid$(fix_Pix, 1, Len(fix_Pix) - 1)
End Function

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

