Attribute VB_Name = "Transcoding"
Option Explicit

'返回内部字符串---------------script_app.Run不需要用到，script_app.Eval需要用来规格化文本------------------------------------
Public Function OX_CInternal(ByVal sourceString As String, ByVal sourceType As String) As String
    On Error Resume Next
    Select Case sourceType
    Case "vbscript"
        sourceString = Replace$(sourceString, """", """""")             'Chr(34), Chr(34) & Chr(34))
        sourceString = Replace$(sourceString, vbCr, """ & vbCr & """)   'Chr(10), Chr(34) & " & Chr(10) & " & Chr(34))
        sourceString = Replace$(sourceString, vbLf, """ & vbLf & """)   'Chr(13), Chr(34) & " & Chr(13) & " & Chr(34))
    Case Else
        sourceString = Replace$(sourceString, """", "\""")                                  'Chr(34), "\" & Chr(34))
        sourceString = Replace$(sourceString, vbLf, """ + String.fromCharCode(10) + """)    'Chr(10), Chr(34) & "+String.fromCharCode(10)+" & Chr(34))
        sourceString = Replace$(sourceString, vbCr, """ + String.fromCharCode(13) + """)    'Chr(13), Chr(34) & "+String.fromCharCode(13)+" & Chr(34))
    End Select
    OX_CInternal = sourceString
End Function

'过滤指定关键字集------------------------------------------------------------------------------------------
Public Function OX_FilterKeywords(ByVal sourceString As String, ByVal keywords As String) As String
    On Error Resume Next
    Dim script_code_replace
    Dim i As Long
    If keywords <> "0" Then
        script_code_replace = Split(keywords, ",")
        For i = 0 To UBound(script_code_replace)
            DoEvents
            If IsNumeric(script_code_replace(i)) Then
                sourceString = Replace$(sourceString, Chr(Int(script_code_replace(i))), "")
            Else
                sourceString = Replace$(sourceString, script_code_replace(i), "")
            End If
        Next i
    End If
    OX_FilterKeywords = sourceString
End Function
'还原OX163自定义字符----------------------------------------------------------------------------------------
Public Function OX_PrivateChr(ByVal sourceString As String) As String
sourceString = Replace(sourceString, "&for_ox163_replace_vbcrlf&", vbCrLf)
sourceString = Replace(sourceString, "&for_ox163_replace_vline&", "|")
OX_PrivateChr = sourceString
End Function

'网页JS代码中unicode转换ascii函数“\u”开头字符，163相册中用到
Public Function unicode2asc(ByVal old_str)
    Dim unicode_tf As Boolean, i As Long, j As Long
    Dim unicode_number As Long
    Dim check_str
    old_str = Replace$(old_str, "\/", "/")
    If InStr(old_str, "\u") < 1 Then unicode2asc = old_str: Exit Function
    check_str = Split(old_str, "\u")
    For i = 0 To UBound(check_str)
        DoEvents
        unicode_tf = False
        If i = 0 And InStr(old_str, "\u") > 1 Then GoTo end_last
        If Len(check_str(i)) > 3 Then
            unicode_tf = True
            For j = 1 To 4
                If InStr("0123456789abcdefABCDEF", Mid$(check_str(i), j, 1)) < 1 Then unicode_tf = False: GoTo end_last
            Next j
            old_str = Left(check_str(i), 4)
            unicode_number = "&H" & old_str
            check_str(i) = ChrW(unicode_number) & Right(check_str(i), Len(check_str(i)) - 4)
        End If
end_last:
        If unicode_tf = True Then
            unicode2asc = unicode2asc & check_str(i)
        ElseIf i = 0 Then
            unicode2asc = check_str(i)
        Else
            unicode2asc = unicode2asc & "\u" & check_str(i)
        End If
    Next i
End Function

'网页字符转换为常规字符
Public Function fix_Code(ByVal old_str As String) As String
    '&lt;   - <
    old_str = Replace$(old_str, "&lt;", "<")
    '&gt;   - >
    old_str = Replace$(old_str, "&gt;", ">")
    '&quot; - "
    old_str = Replace$(old_str, "&quot;", Chr(34))
    '&#0039; - '
    old_str = Replace$(old_str, "&#0039;", "'")
    '&#039; - '
    old_str = Replace$(old_str, "&#039;", "'")
    '&#39; - '
    old_str = Replace$(old_str, "&#39;", "'")
    '&amp;  - &
    fix_Code = Replace$(old_str, "&amp;", "&")
End Function

'修正文件名，去除不可用的字符
Public Function reName_Str(ByVal old_Name As String) As String
Dim i As Long
    reName_Str = Replace$(old_Name, Chr(92), "_")
    reName_Str = Replace$(reName_Str, Chr(47), "_")
    reName_Str = Replace$(reName_Str, Chr(34), "_")
    reName_Str = Replace$(reName_Str, Chr(58), "_")
    reName_Str = Replace$(reName_Str, Chr(42), "_")
    reName_Str = Replace$(reName_Str, Chr(60), "[")
    reName_Str = Replace$(reName_Str, Chr(62), "]")
    reName_Str = Replace$(reName_Str, Chr(124), "_")
    For i = 1 To Len(reName_Str)
        If Asc(Mid(reName_Str, i, 1)) = 63 Then reName_Str = Replace(reName_Str, Mid(reName_Str, i, 1), "_")
    Next
    If Left(reName_Str, 1) = "." Then reName_Str = "_" & Mid$(reName_Str, 2)
    If Right(reName_Str, 1) = "." Then reName_Str = Mid$(reName_Str, 1, Len(reName_Str) - 1) & "_"
End Function
