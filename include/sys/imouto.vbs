'2014-2-7 163.shanhaijing.net
Dim page_counter
Dim tags, page, url_instr, pool, url_head
'----------------------------------------------------
Sub get_url_head(ByVal url_str)
On Error Resume Next
	url_head="https://yande.re/"
	If left(url_str,Len("http://yande.re/"))="http://yande.re/" Then
		url_head="http://yande.re/"
	Else
		url_head="https://yande.re/"
	End If
End Sub
'----------------------------------------------------
Function return_download_url(ByVal url_str)
On Error Resume Next
'http://yande.re/post?tags=emma
'http://yande.re/post?page=5&tags=miko
'http://yande.re/post/index
'http://yande.re/post?page=3
'http://yande.re/post/index?tags=tagme
'http://yande.re/pool/show/596?page=3
'http://yande.re/wiki/show?title=park_sung-woo
'http://yande.re/wiki/show?page=3&title=park_sung-woo
tags=""
Call get_url_head(url_str)
Dim page_tmp
If InStr(LCase(url_str), url_head & "pool/show/") =1 Then
		'http://yande.re/pool/show/2218
		'http://yande.re/post?tags=pool%3A2218
		pool="pool"
		return_download_url = "inet|10,13|" & url_str & "|" & url_str & vbcrlf & "User-Agent: Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0)"
		OX163_urlpage_Referer = "User-Agent: Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0)"
		Exit Function
		'使用tag来获取pool数据（不使用该方法）
    url_str=Mid(url_str, len(url_head & "pool/show/")+1)
    If InStr(LCase(url_str), "/") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "/") - 1)
    If InStr(LCase(url_str), "?") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "?") - 1)
    If InStr(LCase(url_str), "#") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "#") - 1)
    url_str=url_head & "post?tags=pool%3A" & url_str
End If
page_tmp=url_str
If InStr(LCase(url_str), "tags=") > 0 Then
    tags = Mid(url_str, InStr(LCase(url_str), "tags=") + 5)
    url_instr = Mid(url_str,1, InStr(LCase(url_str), "tags=") -2)
    If InStr(tags, "&") > 0 Then tags = Mid(tags, 1, InStr(tags, "&") - 1)
    If InStr(tags, " ") > 0 Then tags = Mid(tags, 1, InStr(tags, " ") - 1)
    If InStr(LCase(url_instr), "page=") > 0 Then url_instr = Mid(url_instr, 1, InStr(LCase(url_instr), "page=") - 2)
End If
If tags <> "" Then url_str = url_instr & "?tags=" & tags
pool=""
If InStr(LCase(url_str), url_head & "wiki/show?") =1 Then
    pool="wiki"
    url_instr=Mid(url_str, InStr(url_str, "?") - 1)
    If InStr(LCase(url_instr), "?title=") > 0 Then
    	url_instr = Mid(url_instr, InStr(LCase(url_instr), "?title="))
    	If InStr(url_instr, "&")>0 Then url_instr = Mid(url_instr,1,InStr(url_instr, "&")-1)
    ElseIf InStr(LCase(url_instr), "&title=") > 0 Then
    	url_instr = "?" & Mid(url_instr, InStr(LCase(url_instr), "&title=")+1)
    	If InStr(url_instr, "&")>0 Then url_instr = Mid(url_instr,1,InStr(url_instr, "&")-1)
    End If
    url_str=url_head & "wiki/show" & url_instr
    url_instr=url_str
End If
page_counter=0
page=1
If InStr(LCase(page_tmp),"&page=")>len(url_head & "post") Or InStr(LCase(page_tmp),"?page=")>len(url_head & "post") Then
	If InStr(LCase(page_tmp),"&page=") Then page_tmp=Mid(page_tmp,InStr(LCase(page_tmp),"&page=")+6)
	If InStr(LCase(page_tmp),"?page=") Then page_tmp=Mid(page_tmp,InStr(LCase(page_tmp),"?page=")+6)
	If InStr(page_tmp,"&") Then page_tmp=Mid(page_tmp,1,InStr(page_tmp,"&")-1)
	If IsNumeric(page_tmp) Then
		If Int(page_tmp)>1 Then
			If MsgBox("本页为第" & page_tmp & "页" & vbcrlf & "是否从第1页开始？", vbYesNo, "问题")=vbyes Then
				page=1
				If InStr(LCase(url_str),"&page=") Or InStr(LCase(url_str),"?page=") Then url_str=format_page(url_str)
			Else
				page=Int(page_tmp)
				If instr(url_str,"?") Then
					url_str=url_str & "&page=" & page
					Else
					url_str=url_str & "?page=" & page
				End if
			End If
		End If
	End If
Else
	page=1
End If
return_download_url = "inet|10,13|" & url_str & "|" & url_head & vbcrlf & "User-Agent: Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0)"
OX163_urlpage_Referer = "User-Agent: Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0)"
End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""
Call get_url_head(url_str)

Dim key_word,js_block,url_temp,split_str,split_i

If pool="pool" Then
	key_word="Post.register_resp({"
	If instr(LCase(html_str),lcase(key_word))>0 Then
		js_block=mid(html_str,instr(LCase(html_str),lcase(key_word))+len(key_word))
		key_word="],""pools"":["
		js_block=mid(js_block,1,instr(LCase(js_block),lcase(key_word))-1)
		split_str=split(js_block,"},{")
		For split_i = 0 To UBound(split_str)
		js_block=""
		url_str=""
		url_temp=""
		
		'width x height
		key_word="""width"":"
		url_temp=mid(split_str(split_i),instr(LCase(split_str(split_i)),LCase(key_word))+len(key_word))
		url_temp=mid(url_temp,1,instr(url_temp,",")-1)
		If IsNumeric(url_temp)=false Then url_temp=""
		js_block=url_temp
		key_word="""height"":"
		url_temp=mid(split_str(split_i),instr(LCase(split_str(split_i)),LCase(key_word))+len(key_word))
		url_temp=mid(url_temp,1,instr(url_temp,",")-1)
		If IsNumeric(url_temp)=false Then url_temp=""
		If js_block<>"" and url_temp<>"" Then
			url_temp=js_block & " x " & url_temp
		Else
			url_temp=""
		End If
		
		'url
		key_word="""file_url"":"""
		url_str=mid(split_str(split_i),instr(LCase(split_str(split_i)),LCase(key_word))+len(key_word))
		url_str=mid(url_str,1,instr(url_str,chr(34))-1)
		url_str=replace(url_str,"\/","/")
    'name
    js_block=""
    js_block=unescape(Mid(url_str,instrrev(url_str,"/")+1))    
    
		split_str(split_i)=""
		split_str(split_i)="|" & url_str & "|" & js_block & "|" & url_temp & vbCrLf
		Next
		
		return_download_list=join(split_str,"")
	End If
Exit Function
End If
		
    '"id":219298,
    '"tags":"chinadress hong_meiling moneti touhou",
    '"created_at":1343833859,
    '"creator_id":111274,
    '"author":"\u690e\u540d\u6df1\u590f",
    '"change":1127276,
    '"source":"http:\/\/i2.pixiv.net\/img76\/img\/daifuku1285\/29049995.jpg",
    '"score":53,"md5":"d89fd5410e06811c4ed9e12cb28d07b5",
    '"file_size":559972,
    '"file_url":"https:\/\/yande.re\/image\/d89fd5410e06811c4ed9e12cb28d07b5\/yande.re%20219298%20chinadress%20hong_meiling%20moneti%20touhou.jpg",
    '"is_shown_in_index":true,
    '"preview_url":"https:\/\/yande.re\/data\/preview\/d8\/9f\/d89fd5410e06811c4ed9e12cb28d07b5.jpg",
    '"preview_width":116,
    '"preview_height":150,
    '"actual_preview_width":233,"actual_preview_height":300,
    '"sample_url":"https:\/\/yande.re\/sample\/d89fd5410e06811c4ed9e12cb28d07b5\/yande.re%20219298%20sample.jpg",
    '"sample_width":1164,
    '"sample_height":1500,
    '"sample_file_size":430205,
    '"jpeg_url":"https:\/\/yande.re\/image\/d89fd5410e06811c4ed9e12cb28d07b5\/yande.re%20219298%20chinadress%20hong_meiling%20moneti%20touhou.jpg",
    '"jpeg_width":1246,
    '"jpeg_height":1606,
    '"jpeg_file_size":0,
    '"rating":"s","
    'has_children":false,
    '"parent_id":null,
    '"status":"active",
    '"width":1246,
    '"height":1606,
    '"is_held":false,
    '"frames_pending_string":"",
    '"frames_pending":[],
    '"frames_string":"",
    '"frames":[]    

	key_word="Post.register({"
	If instr(LCase(html_str),lcase(key_word))>0 Then
		js_block=mid(html_str,instr(LCase(html_str),lcase(key_word))+len(key_word))
		key_word="</script>"
		js_block=mid(js_block,1,instr(LCase(js_block),lcase(key_word))-1)
		split_str=split(js_block,"Post.register({")
		For split_i = 0 To UBound(split_str)
		js_block=""
		url_str=""
		url_temp=""
		
		'width x height
		key_word="""width"":"
		url_temp=mid(split_str(split_i),instr(LCase(split_str(split_i)),LCase(key_word))+len(key_word))
		url_temp=mid(url_temp,1,instr(url_temp,",")-1)
		If IsNumeric(url_temp)=false Then url_temp=""
		js_block=url_temp
		key_word="""height"":"
		url_temp=mid(split_str(split_i),instr(LCase(split_str(split_i)),LCase(key_word))+len(key_word))
		url_temp=mid(url_temp,1,instr(url_temp,",")-1)
		If IsNumeric(url_temp)=false Then url_temp=""
		If js_block<>"" and url_temp<>"" Then
			url_temp=js_block & " x " & url_temp
		Else
			url_temp=""
		End If
		
		'url
		key_word="""file_url"":"""
		url_str=mid(split_str(split_i),instr(LCase(split_str(split_i)),LCase(key_word))+len(key_word))
		url_str=mid(url_str,1,instr(url_str,chr(34))-1)
		url_str=replace(url_str,"\/","/")
    'name
    js_block=""
    js_block=unescape(Mid(url_str,instrrev(url_str,"/")+1))    
    
		split_str(split_i)=""
		split_str(split_i)="|" & url_str & "|" & js_block & "|" & url_temp & vbCrLf
		Next
		
		return_download_list=join(split_str,"")
End If

url_str=html_str
If InStr(LCase(url_str), "<div id=""paginator"">") > 0 Then
	If page_counter=0 Then
	url_str = Mid(url_str, InStr(LCase(url_str), "<div id=""paginator"">") + 20)
	url_str = Mid(url_str,1,InStr(LCase(url_str), "<a class=""next_page""") - 1)
	url_str = Mid(url_str,1,InStrrev(LCase(url_str), "</a>") - 1)
	url_str = Mid(url_str,InStrrev(LCase(url_str), ">")+1)
		If IsNumeric(url_str) Then
			page_counter=Int(url_str)	
		Else
			page_counter=1
		End If
	End If

	If page<page_counter and tags<>"" Then
		page=page+1
		return_download_list = return_download_list & page_counter & "|inet|10,13|" & url_instr & "?page=" & page & "&tags=" & tags
	ElseIf page<page_counter and pool="pool" Then
		page=page+1
		return_download_list = return_download_list & page_counter & "|inet|10,13|" & url_instr & "?page=" & page
	ElseIf page<page_counter and pool="wiki" Then
		page=page+1
		return_download_list = return_download_list & page_counter & "|inet|10,13|" & url_instr & "&page=" & page
	ElseIf page<page_counter Then
		page=page+1
		return_download_list = return_download_list & page_counter & "|inet|10,13|" & url_instr & url_head & "post?page=" & page
	Else
	return_download_list = return_download_list & "0"	
	End If
Else
return_download_list = return_download_list & "0"
End If
End Function
'-----------------------------------------------
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