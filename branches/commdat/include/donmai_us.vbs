'2009-10-04 163.shanhaijing.net
Dim page_counter, url_parent
Dim tags, page, url_instr
Dim login_TF
Function return_download_url(ByVal url_str)
'http://miezaru.donmai.us/post?tags=emma
'http://miezaru.donmai.us/post?page=5&tags=miko
'http://danbooru.donmai.us/post/index
'http://danbooru.donmai.us/post?page=3
'http://danbooru.donmai.us/post/index?tags=tagme
'http://danbooru.donmai.us/post/index?tags=morishima_haruka&page=2
On Error Resume Next
dim temp_url
login_TF=0

temp_url=url_str
url_parent = Mid(url_str, 1, InStr(LCase(url_str), ".donmai.us/") + 9)
url_str = Mid(url_str, InStr(LCase(url_str), ".donmai.us/") + 10)
tags = ""
If InStr(LCase(url_str), "/post/show/") = 1 Then
url_str=Mid(url_str,1,instrrev(url_str,"/"))
'http://danbooru.donmai.us/post/show/
return_download_url = "inet|10,13|" & url_parent & url_str & "|" & temp_url

Else

If InStr(LCase(url_str), "tags=") > 0 Then
    tags = Mid(url_str, InStr(1, url_str, "tags=", 1) + 5)
    If InStr(1, tags, "&", 1) > 0 Then tags = Mid(tags, 1, InStr(1, tags, "&", 1) - 1)
    If InStr(1, tags, " ", 1) > 0 Then tags = Mid(tags, 1, InStr(1, tags, " ", 1) - 1)
End If

If tags <> "" Then url_str = "/post/index?tags=" & tags
page_counter = 0
page = 1
return_download_url = "inet|10,13,34|" & url_parent & url_str & "|" & url_parent & url_str
End If
End Function
'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list = ""
If login_TF=0 and Len(html_str)<5 Then
	MsgBox "您可能需要登陆donmai.us" & vbcrlf & "请使用内置浏览器登陆(右侧第二个按钮，IE页面图案)" & vbcrlf & "或使用IE类浏览器登陆",vbokonly
	Exit Function
End If
login_TF=1
url_str = html_str
If InStr(LCase(html_str), "<span class=thumb") > 0 Then
html_str = Mid(html_str, InStr(LCase(html_str), "<span class=thumb") + 17)

Dim split_str, add_temp, folder_name
split_str = Split(html_str, "<span class=thumb")
    For split_i = 0 To UBound(split_str)
    add_temp = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "id=") + 3)
    'id
    add_temp = Mid(add_temp, 1, InStr(1, add_temp, ">", 1) - 1)
    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "<a href=") + 8)
    'url
    split_str(split_i) = url_parent & Mid(split_str(split_i), 1, InStr(split_str(split_i), " ") - 1)
    add_temp = add_temp & "_" & Mid(split_str(split_i), InStrRev(split_str(split_i), "/") + 1)
    'folder_name
    folder_name = "donmai.us"
    If tags <> "" Then folder_name = tags
    return_albums_list = return_albums_list & "0|1|" & split_str(split_i) & "|" & folder_name & "|" & add_temp & vbCrLf
    Next
End If

If InStr(LCase(url_str), "<div id=paginator>") > 0 And tags <> "" Then
    If page_counter = 0 Then
    url_str = Mid(url_str, InStr(LCase(url_str), "<div id=paginator>") + 18)
    url_str = Mid(url_str, InStr(LCase(url_str), "<a href=") + 8)
    url_str = Mid(url_str, 1, InStr(LCase(url_str), "</div>") - 1)
    split_str = Split(url_str, "<a href=", -1, 1)
    url_str = Mid(split_str(UBound(split_str) - 1), InStr(split_str(UBound(split_str) - 1), ">") + 1)
    url_str = Mid(url_str, 1, InStr(url_str, "<") - 1)
        If IsNumeric(url_str) Then
            page_counter = Int(url_str)
        Else
            page_counter = 1
        End If
    End If

    If page < page_counter And tags <> "" Then
    page = page + 1
    return_albums_list = return_albums_list & page_counter & "|inet|10,13,34|" & url_parent & "/post/index?tags=" & tags & "&page=" & page
    Else
    return_albums_list = return_albums_list & "0"
    
    End If
Else
return_albums_list = return_albums_list & "0"
End If
End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim pic_alt
If InStr(LCase(html_str), "<li>size:") > 0 Then
    '<img alt=
    'http://danbooru.donmai.us/post/show/245889/        'censored-gloves-nakoruru-nude-pubic_hair-pussy-rib
    'ID
    url_str = Mid(url_str, 1, InStrRev(url_str, "/") - 1)
    url_str = "p" & Mid(url_str, InStrRev(url_str, "/") + 1)
    'alt
    pic_alt = Mid(html_str, InStr(LCase(html_str), "<div id=""note-container"">"))
    url_str = url_str & "_" & Mid(pic_alt, InStr(LCase(pic_alt), "<img alt=")+10)
    If Len(url_str)>180 Then url_str=Left(url_str,179) & "~"
    url_str = Mid(url_str, 1, InStr(url_str, Chr(34)) - 1)
    'url
    html_str = Mid(html_str,InStr(LCase(html_str), "<li>size:"))
    html_str = Mid(html_str,InStr(LCase(html_str), "<a href=")+9)
    html_str = Mid(html_str, 1, InStr(html_str, Chr(34)) - 1)
    url_str = Replace(Trim(url_str), " ", "-") & Mid(html_str, InStrRev(html_str, "."))
    return_download_list = "|" & html_str & "|" & url_str & "|" & vbCrLf & "0"
ElseIf InStr(LCase(html_str), "<param name=""movie""") > 0 Then
    '<param name=movie value=http://danbooru.donmai.us/data/94990da8859881113e661645f32191cf.swf>
    '<embed src=http://danbooru.donmai.us/data/94990da8859881113e661645f32191cf.swf width=800 height=600 allowScriptAccess=never></embed>
    'ID
    url_str = Mid(url_str, 1, InStrRev(url_str, "/") - 1)
    url_str = "p" & Mid(url_str, InStrRev(url_str, "/") + 1)
    'alt
    pic_alt = Mid(html_str, InStr(LCase(html_str), "<input id=""post_old_tags"""))
    url_str = url_str & "_" & Mid(pic_alt, InStr(LCase(pic_alt), "value=""")+7)
    If Len(url_str)>180 Then url_str=Left(url_str,179) & "~"
    url_str = Mid(url_str, 1, InStr(url_str, Chr(34)) - 1)
    'url
    html_str = Mid(html_str,InStr(LCase(html_str), "<param name=""movie"""))
    html_str = Mid(html_str,InStr(LCase(html_str), "value=""")+7)
    html_str = Mid(html_str, 1, InStr(html_str, Chr(34)) - 1)
    url_str = Replace(Trim(url_str), " ", "-") & Mid(html_str, InStrRev(html_str, "."))
    return_download_list = "|" & html_str & "|" & url_str & "|" & vbCrLf & "0"
Else
return_download_list = "0"
End If

End Function