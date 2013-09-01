'2012-2-25 163.shanhaijing.net
Dim url_parent, tags, page, pic_counter, url_instr, retry_time, login_TF

Function return_download_url(ByVal url_str)
'http://miezaru.donmai.us/posts?tags=emma
'http://miezaru.donmai.us/posts?page=5&tags=miko
'http://danbooru.donmai.us/posts?tags=&page=
'http://danbooru.donmai.us/posts?page=3
'http://danbooru.donmai.us/posts/index?tags=tagme
'http://danbooru.donmai.us/posts/index?tags=morishima_haruka&page=2
'http://hijiribe.donmai.us/posts/index?page=5

On Error Resume Next
login_TF=0
tags = ""
url_parent=""
retry_time = 0

'miezaru.donmai.us
'danbooru.donmai.us
'hijiribe.donmai.us
url_parent = Mid(url_str,1,InStr(LCase(url_str), ".donmai.us")+9)
 
'http://danbooru.donmai.us/pools/1446?page=4
If InStr(LCase(url_str), ".donmai.us/pools/") >1 Then
    url_str=Mid(url_str, InStr(lcase(url_str),".donmai.us/pools/")+len(".donmai.us/pools/"))
    If InStr(LCase(url_str), "/") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "/") - 1)
    If InStr(LCase(url_str), "?") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "?") - 1)
    If InStr(LCase(url_str), "#") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "#") - 1)
    url_str=url_parent & "/posts?tags=pool%3A" & Trim(url_str)
End If

url_str = Mid(url_str, len(url_parent)+1)

'http://danbooru.donmai.us/posts/1345051
If InStr(LCase(url_str), "/posts/") = 1 Then
url_instr="/posts/"
return_download_url = "inet|10,13|" & url_parent & url_str & "|" & url_parent
Exit Function
End If

tags=""
If InStr(LCase(url_str), "?tags=") > 0 or InStr(LCase(url_str), "&tags=") > 0 Then
	If InStr(LCase(url_str), "?tags=") > 0 Then url_str=Mid(url_str,InStr(LCase(url_str),"?tags=")+6)
	If InStr(LCase(url_str), "&tags=") > 0 Then url_str=Mid(url_str,InStr(LCase(url_str),"&tags=")+6)
	If InStr(LCase(url_str), "&") > 0 Then url_str=Mid(url_str,1,InStr(url_str,"&")-1)
	tags=url_str
End If

page=""
If InStr(LCase(temp_url),"&page=") Or InStr(LCase(temp_url),"?page=") Then
	If InStr(LCase(temp_url),"&page=") Then temp_url=Mid(temp_url,InStr(LCase(temp_url),"&page=")+6)
	If InStr(LCase(temp_url),"?page=") Then temp_url=Mid(temp_url,InStr(LCase(temp_url),"?page=")+6)
	If InStr(temp_url,"&") Then temp_url=Mid(temp_url,1,InStr(temp_url,"&")-1)
	If Len(temp_url)>0 Then
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
	page=""
End If

If len(page)>0 Then page="&page=" & page
pic_counter=-1

return_download_url = "inet|10,13|" & url_parent & "/post/index.xml?tags=" & tags & page & "&limit=100|" & url_parent
url_instr="1|inet|10,13|" & url_parent & "/post/index.xml?tags=" & tags & page & "&limit=100"

End Function

'--------------------------------------------------------
Function return_image_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim key_str, split_str, add_temp, file_tag, file_url, file_id, split_counter
return_image_list = ""

page=""
key_str="<post preview_url="""
split_counter=-1

If InStr(LCase(html_str), LCase(key_str)) > 0 Then
	
	html_str = Mid(html_str, InStr(LCase(html_str), LCase(key_str)) + len(key_str))	
	split_str = Split(html_str, key_str)
	split_counter=UBound(split_str)+1
	
	For split_i = 0 To UBound(split_str)
			file_tag=""
			file_url=""
			file_id=""
			html_str=""
			'tag
			key_str=""" tags="""
	    file_tag=Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + len(key_str))
	    file_tag=Mid(file_tag,1,InStr(file_tag,chr(34))-1)
	    file_tag=replace(file_tag,"|","&#124;")
	    file_tag=replace(file_tag,"\\","\")
	    
			'file_url
			key_str=""" file_url="""
	    file_url=Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + len(key_str))
	    file_url=Mid(file_url,1,InStr(file_url,chr(34))-1)
	    file_url=url_parent & file_url
	    
			'ID
			key_str=""" id="""
	    file_id=Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + len(key_str))
	    file_id=Mid(file_id,1,InStr(file_id,chr(34))-1)
			If IsNumeric(file_id)=false Then file_id=""
			
			If len(file_id)>0 and len(file_url)>0 Then
				'file name
				html_str="p" & file_id & "_" & Trim(file_tag)
		    If Len(html_str)>180 Then html_str=Left(html_str,179) & "~"
		    html_str = Replace(html_str, " ", "-") & Mid(file_url, InStrRev(file_url, "."))
				'http://danbooru.donmai.us/post/index.xml?tags=pool%3A6117&limit=100&page=a1345100 'a表示该ID前100张
				'http://danbooru.donmai.us/post/index.xml?tags=pool%3A6117&limit=100&page=b1345100 'b表示该ID后100张
		    page="b" & file_id
		    
		    split_str(split_i) = "|" & file_url & "|" & html_str & "|" & file_tag & vbCrLf
	    Else
	    	split_str(split_i)=""
	  	End If
	Next
	return_image_list=join(split_str,"")
End If

If split_counter>99 and len(page)>0 Then
		page="&page=" & page
	  url_instr="1|inet|10,13|" & url_parent & "/post/index.xml?tags=" & tags & page & "&limit=100"
	  return_image_list = return_image_list & url_instr
	  pic_counter=split_counter
	  retry_time=0
ElseIf split_counter=-1 and (pic_counter=-1 or pic_counter>99) and retry_time<3 Then
		return_image_list = url_instr
		retry_time=retry_time+1
Else
		return_image_list = return_image_list & "0"
End If
End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next

If login_TF=0 and Len(html_str)<5 Then
	MsgBox "您可能需要登陆donmai.us" & vbcrlf & "请使用内置浏览器登陆(右侧第二个按钮，IE页面图案)" & vbcrlf & "或使用IE8及以下类浏览器登陆",vbokonly
	Exit Function
End If
login_TF=1

'http://danbooru.donmai.us/posts/245889        'censored-gloves-nakoruru-nude-pubic_hair-pussy-rib
If url_instr="/posts/" Then
	Dim pic_alt
	If InStr(LCase(html_str), "<h1>information</h1>") > 0 Then
			pic_alt=html_str
			html_str=Mid(html_str,InStr(LCase(html_str), "<h1>information</h1>"))
			html_str=Mid(html_str,1,InStr(LCase(html_str), "</section>"))
			
	    'ID
	    url_str = Mid(url_str, InStr(lcas(url_str), "/posts/") +7)
	    If InStr(url_str, "?")>0 Then url_str = Mid(url_str,1,InStr(url_str,"?")-1)
	    If InStr(url_str, "&")>0 Then url_str = Mid(url_str,1,InStr(url_str,"&")-1)
	    If InStr(url_str, "#")>0 Then url_str = Mid(url_str,1,InStr(url_str,"#")-1)
	    url_str=trim(url_str)
	    If IsNumeric(url_str) Then
	    	url_str = "p" & url_str
	    Else
	    	url_str=Mid(html_str,InStr(LCase(html_str), "<li>id:")+len("<li>id:"))
	    	url_str=Trim(Mid(url_str,InStr(url_str, "<")-1))
	    	If IsNumeric(url_str) Then
	    		url_str = "p" & url_str
	    	Else
	    		url_str = "p"
	    	End If
	  	End If
	  	
	    'url
	    html_str = Mid(html_str,InStr(LCase(html_str), "size: <a href=""")+len("size: <a href="""))
	    html_str = Mid(html_str, 1, InStr(html_str, Chr(34)) - 1)
	    html_str = url_parent & html_str
	    
	    'alt
	    pic_alt = Mid(pic_alt,InStr(LCase(pic_alt), "id=""post_tag_string"""))
	    pic_alt = Mid(pic_alt,InStr(pic_alt, ">")+1)
	    pic_alt = Mid(pic_alt,1,InStr(LCase(pic_alt), "</textarea>")-1)
	    pic_alt=replace(pic_alt,"|","&#124;")
	    pic_alt=replace(pic_alt,"\\","\")
	    
			'file name
			url_str=url_str & "_" & Trim(pic_alt)
		  If Len(url_str)>180 Then url_str=Left(url_str,179) & "~"
		  url_str = Replace(url_str, " ", "-") & Mid(html_str, InStrRev(html_str, "."))
		  
	    return_download_list = "|" & html_str & "|" & url_str & "|" & pic_alt & vbCrLf & "0"
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