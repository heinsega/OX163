'2011-8-14 163.shanhaijing.net
Dim page_counter
Dim tags, page, url_instr, pool, url_head
Dim retry_time, retry_url
Function return_download_url(ByVal url_str)
'idol.sankakucomplex.com
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
'idol.sankakucomplex.com
'chan.sankakucomplex.com
If InStr(LCase(url_str), "http://idol.sankakucomplex.com") = 1 Then
	url_head="http://idol"
	Else
	url_head="http://chan"
End If

If InStr(LCase(url_str), ".sankakucomplex.com/post/show/") = 12 Then
	pool="post"
	return_download_url = "inet|10,13|" & url_str
	retry_url=return_download_url
	Exit Function
End If

Dim page_str
page_str=""
If InStr(LCase(url_str), "?page=") > 10 Then
    page_str = Mid(url_str, InStr(LCase(url_str), "?page=") + 6)
    url_str = Mid(url_str, 1, InStr(LCase(url_str), "?page=") -1)
    If InStr(page_str, "&") > 0 Then
    	url_str =url_str & "?" & Mid(page_str, InStr(page_str, "&")+1)
    	page_str = Mid(page_str, 1, InStr(page_str, "&") - 1)
    End If
ElseIf InStr(LCase(url_str), "&page=") > 10 Then
    page_str = Mid(url_str, InStr(LCase(url_str), "&page=") + 6)
    url_str = Mid(url_str, 1, InStr(LCase(url_str), "&page=") -1)
    If InStr(page_str, "&") > 0 Then
    	url_str =url_str & "&" & Mid(page_str, InStr(page_str, "&")+1)
    	page_str = Mid(page_str, 1, InStr(page_str, "&") - 1)
    End If
End If

retry_url=""
url_instr=url_str
return_download_url = "inet|10,13|" & url_instr
retry_url=return_download_url

If page_str<>"" and IsNumeric(page_str)=true Then
	If MsgBox("您输入的网页地址不是从第一页开始的，" & vbCrLf & "是否从第一页开始下载？" & vbCrLf & vbCrLf & "[YES]从第一页开始" & vbCrLf & "[NO]从当前页开始", vbYesNo, "询问") = vbNo Then
		If Int(page_str)>1 Then
			page=Int(page_str)
			If InStr(LCase(url_instr), "?")>10 Then
				return_download_url = return_download_url & "&page=" & page
			Else
				return_download_url = return_download_url & "?page=" & page
			End If
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
Dim key_str, split_str, add_temp, file_url
key_str="Post.register({"
If InStr(LCase(html_str), LCase(key_str)) > 0 Then	
	retry_time=0
	html_str = Mid(html_str, InStr(LCase(html_str), LCase(key_str)) + len(key_str))
	split_str = Split(html_str, key_str, -1, 1)

    For split_i = 0 To UBound(split_str)
			'tags
			html_str=""	
			key_str=",""tags"":"""
	    html_str=Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + len(key_str))
	    html_str=Mid(html_str,1,InStr(html_str,chr(34))-1)
	    html_str=replace(html_str,"|","&#124;")
	    html_str=replace(html_str,"\\","\")
	    
			'file_url
			file_url=""
			key_str=",""file_url"":"""
	    file_url=Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + len(key_str))
	    file_url=Mid(file_url,1,InStr(file_url,chr(34))-1)
	    
			'ID
			add_temp=""
			key_str=",""id"":"
	    add_temp=Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + len(key_str))
	    If InStr(add_temp,"}") Then add_temp=Mid(add_temp,1,InStr(add_temp,"}")-1)
	    If InStr(add_temp,",") Then add_temp=Mid(add_temp,1,InStr(add_temp,",")-1)
			If IsNumeric(add_temp)=false Then add_temp=""
			
			'file name
			split_str(split_i)="p" & add_temp & "_" & Trim(html_str)
	    If Len(split_str(split_i))>180 Then split_str(split_i)=Left(split_str(split_i),179) & "~"
	    split_str(split_i) = Replace(split_str(split_i), " ", "-") & Mid(file_url, InStrRev(file_url, "."))	    

	    return_download_list = return_download_list & "|" & file_url & "|" & split_str(split_i) & "|" & html_str & vbCrLf
   Next
   
ElseIf retry_time<5 Then	
  retry_time=retry_time+1
  return_download_list = "2|" & retry_url
  Exit Function    
End If

If InStr(LCase(url_str), "<div id=""paginator""") > 0 Then
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

	If page<page_counter Then
		page=page+1
		If InStr(LCase(url_instr), "?")>10 Then
			return_download_list = return_download_list & page_counter & "|inet|10,13|" & url_instr & "&page=" & page
		Else
			return_download_list = return_download_list & page_counter & "|inet|10,13|" & url_instr & "?page=" & page
		End If
	Else
	return_download_list = return_download_list & "0"	
	End If
Else
return_download_list = return_download_list & "0"
End If
End Function