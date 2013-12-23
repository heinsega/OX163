'2013-12-23 163.shanhaijing.net
Dim start_time
Dim delay_tf
Dim page_url,page,album_info,root_url
Dim retry
Function return_download_url(ByVal url_str)
On Error Resume Next
'http://g.e-hentai.org/g/48158/f8b481edf9/           =====2-m-y/   '只能1-n-n了
'http://g.e-hentai.org/g/48158/f8b481edf9/2-m-y/3    =====page4
'2-m-y=====1->8,2->16,3->24,4->32,5->40,6->48,7->56,8->64,9->72,10->80;n->none,m->normal,l->large;y->Description yes,n->Description no
'http://g.e-hentai.org/g/21601/3199e9d96e/10-m-y/1   =====page2
'code page
'http://g.e-hentai.org/codegen.php?gid=360282&t=fb3a56c95c&s=1-m-y&type=bbcode
'http://g.e-hentai.org/codegen.php?gid=360282&t=fb3a56c95c&s=1-m-y&type=html
'http://g.e-hentai.org/g/360282/fb3a56c95c/

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'tfolder=objFSO.GetSpecialFolder(2)
'tfolder=tfolder & "\..\..\Cookies\*@g.e-hentai*.txt"
'objFSO.Deletefile tfolder ,True
'Set objFSO=nothing
retry=0
delay_tf=0
page=0
start_time=Now()
Dim split_str
'http://exhentai.org
If InStr(LCase(url_str), "http://g.e-hentai.org/g/")=1 Then
	url_str=Mid(url_str,InStr(LCase(url_str),"http://g.e-hentai.org/g/")+Len("http://g.e-hentai.org/g/"))
	split_str=split(url_str,"/")
	page_url="http://g.e-hentai.org/g/" & split_str(0) & "/" & split_str(1) & "/"
	album_info=split_str(0)
	root_url="http://g.e-hentai.org/"
	return_download_url="inet|10,13|" & page_url
ElseIf InStr(LCase(url_str), "http://exhentai.org/g/")=1 Then
	url_str=Mid(url_str,InStr(LCase(url_str),"http://exhentai.org/g/")+Len("http://exhentai.org/g/"))
	split_str=split(url_str,"/")
	page_url="http://exhentai.org/g/" & split_str(0) & "/" & split_str(1) & "/"
	album_info=split_str(0)
	root_url="http://exhentai.org/"
	return_download_url="inet|10,13|" & page_url
Else
	return_download_url="inet|10,13|http://www.163.com/?Delay_5s-利用163页面延迟3秒"' & vbcrlf & "Cookie: lastvisit=1238694351; impcookie=71b644e1b37ee8c8a6052d952804eb54df00d475cb9c72bf3e82e43384c30ab8; Apache=168296599x0.116+1240025775x1022208447; b=%3A%3Amm4l%2Cmm4e%2Cmm4s%2Cmm4p%2Cih53%2Cih56; ut=1%3Aq1YqM1SyqlYqTi1WslJKya%2FJzsxIMa8x0kkxgTINdVJMEaKZMFElHaXc1JJEkN6S4iKQbgszEwOD2tpaAA%3D%3D; __utma=11274144.128020507.1240025776.1241191778.1241714935.4; __utmz=11274144.1240025776.1.1.utmccn=(direct)&for_ox163_replace_vline&utmcsr=(direct)&for_ox163_replace_vline&utmcmd=(none); geo=1%3Aq1YqM1SyqlZKyU1UslIyUNJRSs4vBbKc%2FYDMovQ8IDPYA8gszkwHMi1M0swtjc0NUkzTjJMsExMNTNISU5JNDYwSLQyNUpPTlGprAQ%3D%3D; __utmb=11274144; x=344921-573e35987c794f547f5eeb330dde7dbbabbb0450"
	delay_tf=1
	OX163_urlpage_Referer=""
End If

'突破墙，暂时无用
'OX163_urlpage_Referer="Host: 95.211.21.16" & vbcrlf & "Referer: http://g.e-hentai.org/"' & vbcrlf & "Cookie: lastvisit=1238694351; impcookie=71b644e1b37ee8c8a6052d952804eb54df00d475cb9c72bf3e82e43384c30ab8; Apache=168296599x0.116+1240025775x1022208447; b=%3A%3Amm4l%2Cmm4e%2Cmm4s%2Cmm4p%2Cih53%2Cih56; ut=1%3Aq1YqM1SyqlYqTi1WslJKya%2FJzsxIMa8x0kkxgTINdVJMEaKZMFElHaXc1JJEkN6S4iKQbgszEwOD2tpaAA%3D%3D; __utma=11274144.128020507.1240025776.1241191778.1241714935.4; __utmz=11274144.1240025776.1.1.utmccn=(direct)&for_ox163_replace_vline&utmcsr=(direct)&for_ox163_replace_vline&utmcmd=(none); geo=1%3Aq1YqM1SyqlZKyU1UslIyUNJRSs4vBbKc%2FYDMovQ8IDPYA8gszkwHMi1M0swtjc0NUkzTjJMsExMNTNISU5JNDYwSLQyNUpPTlGprAQ%3D%3D; __utmb=11274144; x=344921-573e35987c794f547f5eeb330dde7dbbabbb0450"
'http://g.e-hentai.org/g/  ->  http://r.e-hentai.org/g/
End Function
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list = ""
Dim Instr_String,counts,regex,matches

If delay_tf=1 Then
	If DateDiff("s", start_time, Now())<30 Then
		return_albums_list="1|inet|10,13|http://www.163.com/?Delay_30s-利用163页面延迟30秒"
	Else
		return_albums_list="1|inet|10,13|" & page_url & "?p=" & page
		delay_tf=0
	End If
	Exit Function
End If

If InStr(LCase(html_str), "<a href=""" & root_url & "s/") > 0 Then
	retry=0
	delay_tf=0
	Instr_String="<h1 id=""gn"">"
	If InStr(LCase(html_str), LCase(Instr_String)) > 0 Then
		url_str=Mid(html_str,InStr(LCase(html_str), LCase(Instr_String))+len(Instr_String))
		url_str=Mid(url_str,1,InStr(LCase(url_str), "</h1>")-1)
		url_str=rename_utf8(Trim(url_str))
	Else
		url_str="E-Hentai Unknow Title Gallery"
	End If

	Set regex = new RegExp
	regex.Global = True
	regex.Pattern = "<a href=""(" & root_url & "s/[0-9A-Za-z]+/" & album_info & "-[0-9]+)"">"
	Set matches = regex.Execute(html_str)
	
	For Each match In matches
		return_albums_list = return_albums_list & "0|1|" & match.SubMatches(0) & "|" & url_str & "|" & url_str & vbCrLf
	Next

	Instr_String="onclick=""return false"">&gt;</a>"
	If InStr(LCase(html_str), LCase(Instr_String)) > 0 Then
		page=page+1
		
		If page mod 10=0 Then
			return_albums_list=return_albums_list & "1|inet|10,13|http://www.163.com/?Delay_30s-利用163页面延迟30秒"
			start_time=Now()
			delay_tf=1
		Else
			return_albums_list=return_albums_list & "1|inet|10,13|" & page_url & "?p=" & page
		End If
		
	Else
		return_albums_list = return_albums_list & "0"
	End If
	
ElseIf retry<0 and html_str<>"" Then
	retry=retry+1
	return_albums_list="1|inet|10,13|" & page_url
Else
return_albums_list = "0"
End If
End Function
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
If DateDiff("s", start_time, Now()) < 3 Then
	return_download_list="1|inet|10,13|http://www.163.com/?Delay_5s-利用163页面延迟3秒"
	Exit Function
ElseIf delay_tf=1 Then
	'OX163_urlpage_Referer="Host: 95.211.21.16" & vbcrlf & "Referer: http://g.e-hentai.org/"
	delay_tf=0
	return_download_list="1|inet|10,13|" & url_str'replace(url_str,"http://g.e-hentai.org","http://r.e-hentai.org")
	Exit Function
End If

Dim file_name,file_type,full_image

If InStr(LCase(html_str),"<div id=""i3"">")>0 Then

	If InStr(LCase(html_str),"<h1>")>0 Then
		file_name=Mid(html_str,InStr(LCase(html_str),"<h1>"))
		file_name=Mid(file_name,InStr(file_name,">")+1)
		file_name=replace(Mid(file_name,1,InStr(file_name,"<")-1),"|","_") & "_"
		If Len(file_name)>100 Then file_name=Left(file_name,100) & "_"
	Else
		file_name=""
	End If
	
	'pic_url
	full_image=""
	'是否有Download original选项
	'full_image=Mid(html_str,InStrrev(LCase(html_str),"<div class=""if"">"))
	'full_image=Get_full_image_url(full_image)
	'---------------------------
	html_str=Mid(html_str,InStr(LCase(html_str),"<div id=""i3"">"))
	html_str=Mid(html_str,InStr(LCase(html_str),"src=""")+5)
	file_type=Mid(html_str,InStr(html_str,Chr(34))+1)
	html_str=Mid(html_str,1,InStr(html_str,Chr(34))-1)
	html_str=replace(html_str,"&amp;","&")
	
	file_type=Mid(file_type,InStr(LCase(file_type),"<div>")+5)
	file_type=Mid(file_type,1,InStr(file_type," ")-1)


	If file_type<>"" Then
		If Mid(file_type,InStrrev(file_type,"."))=Mid(html_str,InStrrev(html_str,".")) Then
			file_name=file_name & file_type
		Else
			file_type=Mid(file_type,1,InStrrev(file_type,".")-1)
			file_name=file_name & file_type & Mid(html_str,instrrev(html_str,"."))
		End If
	Else
		file_name=file_name & Mid(html_str,instrrev(html_str,"/")+1)
	End If

	
	file_name=replace(replace(file_name,Chr(10),""),Chr(13),"")
	file_name=rename_utf8(file_name)
	'Download original image
	If left(full_image,7)="http://" Then html_str=full_image
	'-----------------------
	return_download_list="|" & html_str & "|" & file_name & "|" & file_name & vbCrLf & "0"'&nl=1

ElseIf InStr(LCase(html_str),"<div class=""sni""")>0 Then
		
	If InStr(LCase(html_str),"<h1>")>0 Then
		file_name=Mid(html_str,InStr(LCase(html_str),"<h1>"))
		file_name=Mid(file_name,InStr(file_name,">")+1)
		file_name=replace(Mid(file_name,1,InStr(file_name,"<")-1),"|","_") & "_"
		If Len(file_name)>100 Then file_name=Left(file_name,100) & "_"
	Else
		file_name=""
	End If
	'pic_url
	full_image=""
	'是否有Download original选项
	'full_image=Mid(html_str,InStrrev(LCase(html_str),"<div class=""if"">"))
	'full_image=Get_full_image_url(full_image)
	'---------------------------
	html_str=Mid(html_str,InStr(LCase(html_str),"<div class=""sn"">"))
	html_str=Mid(html_str,InStr(LCase(html_str),"<img src=""")+Len("<img src="""))
	html_str=Mid(html_str,InStr(LCase(html_str),"<img src=""")+Len("<img src="""))
	html_str=Mid(html_str,InStr(LCase(html_str),"<img src=""")+Len("<img src="""))
	html_str=Mid(html_str,InStr(LCase(html_str),"<img src=""")+Len("<img src="""))
	file_type=Mid(html_str,InStr(html_str,Chr(34))+1)
	html_str=Mid(html_str,InStr(LCase(html_str),"<img src=""")+Len("<img src="""))
	html_str=Mid(html_str,1,InStr(html_str,Chr(34))-1)
	html_str=replace(html_str,"&amp;","&")
	
	file_type=Mid(file_type,InStr(LCase(file_type),"<div>")+5)
	file_type=Mid(file_type,1,InStr(file_type," ")-1)
	
	If file_type<>"" Then
		If Mid(file_type,InStrrev(file_type,"."))=Mid(html_str,InStrrev(html_str,".")) Then
			file_name=file_name & file_type
		Else
			file_type=Mid(file_type,1,InStrrev(file_type,".")-1)
			file_name=file_name & file_type & Mid(html_str,instrrev(html_str,"."))
		End If
	Else
		file_name=file_name & Mid(html_str,instrrev(html_str,"/")+1)
	End If
	
	file_name=replace(replace(file_name,Chr(10),""),Chr(13),"")
	file_name=rename_utf8(file_name)
	'Download original image
	If left(full_image,7)="http://" Then html_str=full_image
	'-----------------------
	return_download_list="|" & html_str & "|" & file_name & "|" & file_name & vbCrLf & "0"'&nl=1
	
ElseIf retry<4 and html_str<>"" Then
	retry=retry+1
	return_download_list="1|inet|10,13|" & url_str'replace(url_str,"http://g.e-hentai.org","http://r.e-hentai.org")
Else
	retry=0
	return_download_list="0"
End If
End Function
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
Function Get_full_image_url(byval full_image)
Get_full_image_url=""
If instr(LCase(full_image),""">download original")>0 Then
	full_image=Mid(full_image,1,instr(LCase(full_image),""">download original")-1)
	full_image=Mid(full_image,instrrev(full_image,chr(34))+1)
	If left(full_image,7)="http://" Then
		full_image=replace(full_image,"&amp;","&")
		Get_full_image_url=full_image
	End If
Else
	Get_full_image_url=""
End If
End Function


Function rename_utf8(byval utf8_Str)
If Len(utf8_Str)=0 Then Exit Function
For i=1 to Len(utf8_Str)
	If  Asc(Mid(utf8_Str,i,1))=63 Then utf8_Str=replace(utf8_Str,Mid(utf8_Str,i,1),"_")
Next
rename_utf8=utf8_Str
End Function