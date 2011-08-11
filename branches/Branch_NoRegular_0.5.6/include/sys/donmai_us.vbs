'2011-8-11 163.shanhaijing.net
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
'http://hijiribe.donmai.us/post/index?page=5
On Error Resume Next
dim temp_url
login_TF=0

If InStr(LCase(url_str), "donmai.us/pool/show/") >1 Then
    url_str=Mid(url_str, InStr(lcase(url_str),"donmai.us/pool/show/")+len("donmai.us/pool/show/"))
    If InStr(LCase(url_str), "/") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "/") - 1)
    If InStr(LCase(url_str), "?") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "?") - 1)
    If InStr(LCase(url_str), "#") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "#") - 1)
    url_str="http://danbooru.donmai.us/post?tags=pool%3A" & url_str
End If

temp_url=url_str
url_instr=url_str
url_parent = Mid(url_str, 1, InStr(LCase(url_str), ".donmai.us/") + 9)
url_str = Mid(url_str, InStr(LCase(url_str), ".donmai.us/") + 10)

tags = ""
If InStr(LCase(url_str), "/post/show/") = 1 Then
url_instr=Mid(url_str,1,instrrev(url_str,"/"))
'http://danbooru.donmai.us/post/show/
return_download_url = "inet|10,13|" & url_parent & url_instr & "|" & temp_url

Else

tags=""
If InStr(LCase(url_str), "?tags=") > 0 or InStr(LCase(url_str), "&tags=") > 0 Then
	url_str=Mid(url_str,InStr(LCase(url_str),"&tags=")+6)
	url_str=Mid(url_str,InStr(LCase(url_str),"?tags=")+6)
	url_str=Mid(url_str,1,InStr(url_str,"&")-1)
	tags=url_str
End If
page_counter = 0
page=1
If InStr(LCase(temp_url),"&page=") Or InStr(LCase(temp_url),"?page=") Then
	If InStr(LCase(temp_url),"&page=") Then temp_url=Mid(temp_url,InStr(LCase(temp_url),"&page=")+6)
	If InStr(LCase(temp_url),"?page=") Then temp_url=Mid(temp_url,InStr(LCase(temp_url),"?page=")+6)
	If InStr(temp_url,"&") Then temp_url=Mid(temp_url,1,InStr(temp_url,"&")-1)
	If IsNumeric(temp_url) Then
		If Int(temp_url)>1 Then
			If MsgBox("本页为第" & temp_url & "页" & vbcrlf & "是否从第1页开始？", vbYesNo, "问题")=vbyes Then
				page=1
				url_instr=format_page(url_instr)
			Else
				page=Int(temp_url)
			End If
		End If
	End If
Else
	page=1
End If

return_download_url = "inet|10,13|" & url_instr & "|" & url_parent
url_instr=""
End If
End Function

'--------------------------------------------------------
Function return_image_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim key_str, split_str, add_temp, file_url
return_image_list = ""
url_str = html_str
key_str="Post.register({""preview_url"":"""
If InStr(LCase(html_str), LCase(key_str)) > 0 Then
	html_str = Mid(html_str, InStr(LCase(html_str), LCase(key_str)) + len(key_str))
	
	split_str = Split(html_str, key_str)
	
	For split_i = 0 To UBound(split_str)
			'tags
			html_str=""	
			key_str=",""tags"":"""
	    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + len(key_str))
	    html_str=Mid(split_str(split_i),1,InStr(split_str(split_i),chr(34))-1)
	    html_str=replace(html_str,"|","&#124;")
	    html_str=replace(html_str,"\\","\")
	    
			'file_url
			file_url=""
			key_str=",""file_url"":"""
	    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + len(key_str))
	    file_url=Mid(split_str(split_i),1,InStr(split_str(split_i),chr(34))-1)
	    
			'ID
			add_temp=""
			key_str=",""id"":"
	    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + len(key_str))
	    add_temp=Mid(split_str(split_i),1,InStr(split_str(split_i),"}")-1)
			If IsNumeric(add_temp)=false Then add_temp=""
			
			'file name
			split_str(split_i)="p" & add_temp & "_" & Trim(html_str)
	    If Len(split_str(split_i))>100 Then split_str(split_i)=Left(split_str(split_i),99) & "~"
	    split_str(split_i) = Replace(split_str(split_i), " ", "-") & Mid(file_url, InStrRev(file_url, "."))	    

	    return_image_list = return_image_list & "|" & file_url & "|" & split_str(split_i) & "|" & html_str & vbCrLf
	Next
End If

If InStr(LCase(url_str), "<div id=""paginator"">") > 0 or page<page_counter Then
    If page_counter = 0 Then
	    url_str = Mid(url_str, InStr(LCase(url_str), "<div id=""paginator"">") + 20)
	    url_str = Mid(url_str, InStr(LCase(url_str), "<a href=""") + 9)
	    url_str = Mid(url_str, 1, InStr(LCase(url_str), "</div>") - 1)
	    split_str = Split(url_str, "<a href=""", -1, 1)
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
	    return_image_list = return_image_list & page_counter & "|inet|10,13|" & url_parent & "/post/index?tags=" & tags & "&page=" & page
    ElseIf page < page_counter Then
	    page = page + 1
	    return_image_list = return_image_list & page_counter & "|inet|10,13|" & url_parent & "/post/index?page=" & page
    Else
    return_image_list = return_image_list & "0"
    End If
Else
		return_image_list = return_image_list & "0"
End If
End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim pic_alt

If login_TF=0 and Len(html_str)<5 Then
	MsgBox "您可能需要登陆donmai.us" & vbcrlf & "请使用内置浏览器登陆(右侧第二个按钮，IE页面图案)" & vbcrlf & "或使用IE8及以下类浏览器登陆",vbokonly
	Exit Function
End If
login_TF=1

If url_instr<>"" Then
	If InStr(LCase(html_str), "<li>size:") > 0 Then
	    '<img alt=
	    'http://danbooru.donmai.us/post/show/245889/        'censored-gloves-nakoruru-nude-pubic_hair-pussy-rib
	    'ID
	    url_str = Mid(url_str, 1, InStrRev(url_str, "/") - 1)
	    url_str = "p" & Mid(url_str, InStrRev(url_str, "/") + 1)
	    'alt
	    pic_alt = Mid(html_str, InStr(LCase(html_str), "<div id=""note-container"">"))
	    url_str = url_str & "_" & Mid(pic_alt, InStr(LCase(pic_alt), "<img alt=")+10)
	    url_str = replace(url_str,"|","%7C")	    
	    url_str = replace(url_str,"\\","\")
	    If Len(url_str)>100 Then url_str=Left(url_str,99) & "~"
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
	    url_str = replace(url_str,"|","%7C")	    
	    url_str = replace(url_str,"\\","\")
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
Else
	return_download_list=return_image_list(html_str,url_str)
End If
End Function
'-----------------------------------------------------------------------------
Function format_page(url_str)
format_page=url_str
Dim temp_str(2)
If instr(lcase(url_str),"?page=")>0 or instr(lcase(url_str),"&page=")>0 Then
	If instr(lcase(url_str),"?page=")>0 Then
		temp_str(0)=mid(url_str,1,instr(lcase(url_str),"?page=")+5)
		temp_str(1)=mid(url_str,InStr(lcase(url_str),"?page=")+1)
	ElseIf instr(lcase(url_str),"&page=")>0 Then
		temp_str(0)=mid(url_str,1,InStr(lcase(url_str),"&page=")+5)
		temp_str(1)=mid(url_str,InStr(lcase(url_str),"&page=")+1)
	End If
	If instr(temp_str(1),"&")>0 Then
		temp_str(1)=mid(url_str,instr(temp_str(1),"&"))
	Else
		temp_str(1)=""
	End If
	format_page=temp_str(0) & "1" & temp_str(1)
End if
End Function