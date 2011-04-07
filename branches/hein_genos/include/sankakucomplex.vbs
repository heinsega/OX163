'2010-2-25 163.shanhaijing.net
Dim page_counter
Dim tags, page, url_instr, pool
Dim retry_time, retry_url

Function return_download_url(ByVal url_str)
'http://chan.sankakucomplex.com/
'http://chan.sankakucomplex.com/post/
'http://chan.sankakucomplex.com/?page=3
''http://chan.sankakucomplex.com/post/index
''http://chan.sankakucomplex.com/post?page=3

'http://chan.sankakucomplex.com/post/index?tags=tagme
'http://chan.sankakucomplex.com/?page=2&tags=chibi
''http://chan.sankakucomplex.com/post?tags=emma
''http://chan.sankakucomplex.com/post?page=5&tags=miko

'http://chan.sankakucomplex.com/pool/show/596?page=3

'http://chan.sankakucomplex.com/wiki/show?title=park_sung-woo
'http://chan.sankakucomplex.com/wiki/show?page=3&title=park_sung-woo

'http://chan.sankakucomplex.com/post/show/9506/cg-d-o-_-publisher-eigoukaiki-eroge-ino-tagme
On Error Resume Next
tags=""
retry_url=""
retry_time=0
page_counter=0
page=1

If InStr(LCase(url_str), "http://chan.sankakucomplex.com/post/show/") = 1 Then
	pool="post"
	return_download_url = "inet|10,13|" & url_str
	retry_url=return_download_url
	Exit Function
End If

Dim page_str
page_str=""

If InStr(LCase(url_str), "tags=") > 0 Then
    
    If InStr(LCase(url_str), "?page=") > 10 Then
    	page_str = Mid(url_str, InStr(LCase(url_str), "?page=") + 6)
    	If InStr(page_str, "&") > 0 Then page_str = Mid(page_str, 1, InStr(page_str, "&") - 1)
    	If InStr(page_str, " ") > 0 Then page_str = Mid(page_str, 1, InStr(page_str, " ") - 1)
    ElseIf InStr(LCase(url_str), "&page=") > 10 Then
    	page_str = Mid(url_str, InStr(LCase(url_str), "&page=") + 6)
    	If InStr(page_str, "&") > 0 Then page_str = Mid(page_str, 1, InStr(page_str, "&") - 1)
    	If InStr(page_str, " ") > 0 Then page_str = Mid(page_str, 1, InStr(page_str, " ") - 1)
    End If
    
    tags = Mid(url_str, InStr(LCase(url_str), "tags=") + 5)
    url_instr = Mid(url_str,1, InStr(LCase(url_str), "?") -1)
    If InStr(tags, "&") > 0 Then tags = Mid(tags, 1, InStr(tags, "&") - 1)
    If InStr(tags, " ") > 0 Then tags = Mid(tags, 1, InStr(tags, " ") - 1)
    retry_url=""
    
End If

If tags <> "" Then url_str = url_instr & "?tags=" & tags

pool=""

If InStr(LCase(url_str), "http://chan.sankakucomplex.com/pool/show/") =1 Then
    pool="pool"
    If InStr(LCase(url_str), "?page=") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "?page=") - 1)
    url_instr=url_str
ElseIf InStr(LCase(url_str), "http://chan.sankakucomplex.com/wiki/show?") =1 Then
    pool="wiki"
    url_instr=Mid(url_str, InStr(url_str, "?") - 1)
    If InStr(LCase(url_instr), "?title=") > 0 Then
    	url_instr = Mid(url_instr, InStr(LCase(url_instr), "?title="))
    	If InStr(url_instr, "&")>0 Then url_instr = Mid(url_instr,1,InStr(url_instr, "&")-1)
    ElseIf InStr(LCase(url_instr), "&title=") > 0 Then
    	url_instr = "?" & Mid(url_instr, InStr(LCase(url_instr), "&title=")+1)
    	If InStr(url_instr, "&")>0 Then url_instr = Mid(url_instr,1,InStr(url_instr, "&")-1)
    End If
    url_str="http://chan.sankakucomplex.com/wiki/show" & url_instr
    url_instr=url_str
End If

retry_url=""
return_download_url = "inet|10,13|" & url_str
retry_url=return_download_url

If page_str<>"" and IsNumeric(page_str)=true Then
	If MsgBox("您输入的网页地址不是从第一页开始的，" & vbCrLf & "是否从第一页开始下载？" & vbCrLf & vbCrLf & "[YES]从第一页开始" & vbCrLf & "[NO]从当前页开始", vbYesNo, "询问") = vbNo Then
		If Int(page_str)>1 Then
			page=Int(page_str)
			return_download_url = return_download_url & "&page=" & page
		End If
	End If
End If

End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""

If pool="post" Then
	'http://chan.sankakucomplex.com/post/show/9506/cg-d-o-_-publisher-eigoukaiki-eroge-ino-tagme
	Dim pic_alt
	If InStr(LCase(html_str), "<li>size:") > 0 Then
	    retry_time=0
	    '<img alt=
	    'http://danbooru.donmai.us/post/show/245889/        'censored-gloves-nakoruru-nude-pubic_hair-pussy-rib
	    'ID
	    url_str = Mid(url_str, 1, InStrRev(url_str, "/") - 1)
	    url_str = "p" & Mid(url_str, InStrRev(url_str, "/") + 1)
	    'alt
	    pic_alt = Mid(html_str, InStr(LCase(html_str), "<div id=""note-container"">"))
	    url_str = url_str & "_" & Mid(pic_alt, InStr(LCase(pic_alt), "<img alt=")+10)
	    url_str = Trim(Mid(url_str, 1, InStr(url_str, Chr(34)) - 1))
	    If Len(url_str)>180 Then url_str=Left(url_str, 179) & "~"
	    'url
	    html_str = Mid(html_str,InStr(LCase(html_str), "<li>size:"))
	    html_str = Mid(html_str,InStr(LCase(html_str), "<a href=""#""")+11)
	    html_str = Mid(html_str,InStr(LCase(html_str), "<a href=""#""")+11)
	    html_str = Mid(html_str,InStr(LCase(html_str), "<a href=")+9)
	    html_str = Mid(html_str, 1, InStr(html_str, Chr(34)) - 1)
	    url_str = url_str & Mid(html_str, InStrRev(html_str, "."))
	    return_download_list = "|" & html_str & "|" & url_str & "|" & vbCrLf & "0"
	ElseIf retry_time<5 Then
		retry_time=retry_time+1
		return_download_list = "2|" & retry_url
	Else	
	return_download_list = "0"
	End If
	
	Exit Function
End If


url_str=html_str

If InStr(LCase(html_str), "post.register({""") > 0 Then

retry_time=0
html_str = Mid(html_str, InStr(LCase(html_str), "post.register({""") + 16)

Dim split_str,url_temp
split_str = Split(html_str, "post.register({""", -1, 1)

    For split_i = 0 To UBound(split_str)
    html_str=Mid(split_str(split_i), InStr(LCase(split_str(split_i)), """tags"":""") +8)
    url_temp=html_str
    'Tags
    html_str =Trim(Mid(html_str,1, InStr(html_str, Chr(34)) -1))
    url_temp=Mid(url_temp, InStr(LCase(url_temp), ",""id"":") +6)
    url_temp="p" & Mid(url_temp,1, InStr(url_temp, ",") -1) & "_"
    If IsNumeric(url_temp)=false Then url_temp=""
    html_str=url_temp & html_str
    If Len(html_str) > 180 Then html_str = Left(html_str, 179) & "~"
    
    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), """file_url"":""") +12)
    split_str(split_i) = Mid(split_str(split_i),1, InStr(split_str(split_i),Chr(34))-1)
    'url
    split_str(split_i)=replace(split_str(split_i),"\/","/")
    
    'name
    html_str=html_str & unescape(Mid(split_str(split_i),instrrev(split_str(split_i),".")))
    
    return_download_list = return_download_list & "|" & split_str(split_i) & "|" & html_str & "|" & vbCrLf
    Next
    
ElseIf retry_time<5 Then
	
    retry_time=retry_time+1
    return_download_list = "2|" & retry_url
    Exit Function
    
End If



If InStr(LCase(url_str), "<div id=""paginator""") > 0 and (tags<>"" or pool<>"") Then
	If page_counter=0 Then
	url_str = Mid(url_str, InStr(LCase(url_str), "<div id=""paginator""") + 20)
	url_str = Mid(url_str, InStr(LCase(url_str), "<a href=""") + 9)
	url_str = Mid(url_str,1, InStr(LCase(url_str), "</div>") -1)
	split_str=Split(url_str, "<a href=""", -1, 1)
	url_str=Mid(split_str(UBound(split_str)-1), InStr(split_str(UBound(split_str)-1), ">") + 1)
	url_str=Mid(url_str,1, InStr(url_str, "<") -1)
		If IsNumeric(url_str) Then
			page_counter=Int(url_str)		
		Else
			page_counter=1
		End If
	End If

	If page<page_counter and tags<>"" Then
	page=page+1
	return_download_list = return_download_list & page_counter & "|inet|10,13|" & url_instr & "?page=" & page & "&tags=" & tags
	'ElseIf page<page_counter and pool="pool" Then
	'page=page+1
	'return_download_list = return_download_list & page_counter & "|inet|10,13|" & url_instr & "?page=" & page
	'ElseIf page<page_counter and pool="wiki" Then
	'page=page+1
	'return_download_list = return_download_list & page_counter & "|inet|10,13|" & url_instr & "&page=" & page
	Else
	return_download_list = return_download_list & "0"	
	End If
Else
return_download_list = return_download_list & "0"
End If
End Function
'------------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list = ""
url_str = html_str
If InStr(LCase(html_str), "<span class=""thumb""") > 0 Then
html_str = Mid(html_str, InStr(LCase(html_str), "<span class=""thumb""") + 19)

Dim split_str, add_temp, folder_name
split_str = Split(html_str, "<span class=""thumb""")
    For split_i = 0 To UBound(split_str)
    add_temp = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "id=""") + 4)
    'id
    add_temp = Mid(add_temp, 1, InStr(add_temp, Chr(34)) - 1)
    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "<a href=""") + 9)
    'url
    split_str(split_i) = "http://chan.sankakucomplex.com" & Mid(split_str(split_i), 1, InStr(split_str(split_i), Chr(34)) - 1)
    
    add_temp = add_temp & "_" & Mid(split_str(split_i), InStrRev(split_str(split_i), "/") + 1)
    
    'folder_name
	If pool="pool" Then
	'http://chan.sankakucomplex.com/pool/show/596?page=3
	folder_name="pool_" & Mid(url_instr,InStr(LCase(url_instr),"/show/")+6)
	folder_name=Mid(folder_name,1,InStr(folder_name,"?")-1)
	ElseIf pool="wiki" Then
	'http://chan.sankakucomplex.com/wiki/show?title=park_sung-woo
	folder_name="wiki_" & Mid(url_instr,InStr(LCase(url_instr),"?title=")+7)
	folder_name=Mid(folder_name,1,InStr(folder_name,"&")-1)
	Else
	folder_name = url_instr
	End If
	
    return_albums_list = return_albums_list & "0|1|" & split_str(split_i) & "|" & folder_name & "|" & add_temp & vbCrLf
    Next
End If

If InStr(LCase(url_str), "<div id=""paginator"">") > 0 and pool<>"" Then
	If page_counter=0 Then
	url_str = Mid(url_str, InStr(LCase(url_str), "<div id=""paginator"">") + 20)
	url_str = Mid(url_str, InStr(LCase(url_str), "<a href=""") + 9)
	url_str = Mid(url_str,1, InStr(LCase(url_str), "</div>") -1)
	split_str=Split(url_str, "<a href=""", -1, 1)
	
	url_str=Mid(split_str(UBound(split_str)-1), InStr(split_str(UBound(split_str)-1), ">") + 1)
	url_str=Mid(url_str,1, InStr(url_str, "<") -1)
		If IsNumeric(url_str) Then
			page_counter=Int(url_str)		
		Else
			page_counter=1
		End If
	End If

	If page<page_counter and pool="pool" Then
	page=page+1
	return_albums_list = return_albums_list & page_counter & "|inet|10,13|" & url_instr & "?page=" & page
	ElseIf page<page_counter and pool="wiki" Then
	page=page+1
	return_albums_list = return_albums_list & page_counter & "|inet|10,13|" & url_instr & "&page=" & page
	Else
	return_albums_list = return_albums_list & "0"	
	End If
	
Else
return_albums_list = return_albums_list & "0"
End If

End Function