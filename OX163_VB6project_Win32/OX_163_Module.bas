Attribute VB_Name = "OX_163_Module"
'-------------------------------------------------------------------------
'-----------------------OX163网易相册函数模块-----------------------------
'-------------------------------------------------------------------------

Public Function is_username(ByVal username As String) As Boolean
    is_username = True
    If Len(username) > 2 And Len(username) < 50 Then
        For i = 1 To Len(username)
            If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.-_@", Mid(username, i, 1)) < 1 Then is_username = False: Exit Function
        Next i
    Else
        is_username = False
    End If
End Function

Public Function format_username(ByVal format_username_str As String, ByRef username_format As Byte) As String
    '@popo.163.com
    'http://photo.163.com/yq-32@163.com.popo/#m=0&p=1
    '@188.com
    'http://photo.163.com/sally71@188/#m=0&p=1
    '@yeah.net
    'http://photo.163.com/jin1219@yeah/#m=0&p=1
    '@126.com
    'http://photo.163.com/takabe@126/#m=0&p=1
    Dim format_username_word1
    Dim format_username_word2
    Const FU_key1 = "@126|@yeah|@188|@163.com.popo"
    Const FU_key2 = "@126.com|@yeah.net|@188.com|@popo.163.com"
    
    Select Case username_format
    Case 1
        'username_format=1 转换为登录ID,sally71@188--->sally71@188.com
        format_username_word1 = Split(FU_key1, "|")
        format_username_word2 = Split(FU_key2, "|")
        For i = 0 To UBound(format_username_word1)
            If Right(LCase(format_username_str), Len(format_username_word1(i))) = format_username_word1(i) Then
                format_username_str = Left(format_username_str, Len(format_username_str) - Len(format_username_word1(i))) & format_username_word2(i)
                format_username = format_username_str
                Exit Function
            End If
        Next
        
    Case 2
        'username_format=1 转换为地址ID,sally71@188.com--->sally71@188
        format_username_word1 = Split(FU_key1, "|")
        format_username_word2 = Split(FU_key2, "|")
        For i = 0 To UBound(format_username_word1)
            If Right(LCase(format_username_str), Len(format_username_word2(i))) = format_username_word2(i) Then
                format_username_str = Left(format_username_str, Len(format_username_str) - Len(format_username_word2(i))) & format_username_word1(i)
                format_username = format_username_str
                Exit Function
            End If
        Next
        
    End Select
    
    format_username = format_username_str
End Function

