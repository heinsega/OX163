'2010-3-2 163.shanhaijing.net
Dim start_time,counts
Dim delay_tf
Dim page_url,page
Dim retry
Function return_download_url(ByVal url_str)
On Error Resume Next

'http://g.e-hentai.org/g/48158/f8b481edf9/           =====2-m-y/   '只能1-n-n了
'http://g.e-hentai.org/g/48158/f8b481edf9/2-m-y/3    =====page4
'2-m-y=====1->8,2->16,3->24,4->32,5->40,6->48,7->56,8->64,9->72,10->80;n->none,m->normal,l->large;y->Description yes,n->Description no
'http://g.e-hentai.org/g/21601/3199e9d96e/10-m-y/1   =====page2
'http://g.e-hentai.org/s/a8aa8c7cefdb94bfb2c894272bc33f65def46405-728761-2039-1600-jpg/48158-1

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'tfolder=objFSO.GetSpecialFolder(2)
'tfolder=tfolder & "\..\..\Cookies\*@g.e-hentai*.txt"
'objFSO.Deletefile tfolder ,True
'Set objFSO=nothing

counts=0
retry=0
start_time=Now()
Dim split_str
If InStr(LCase(url_str), "http://g.e-hentai.org/g/")>0 Then
	url_str=Mid(url_str,InStr(LCase(url_str),"http://g.e-hentai.org/g/")+Len("http://g.e-hentai.org/g/"))
	split_str=split(url_str,"/")
	return_download_url="inet|10,13|http://g.e-hentai.org/g/" & split_str(0) & "/" & split_str(1) & "/1-n-n/"
	page_url="http://g.e-hentai.org/g/" & split_str(0) & "/" & split_str(1) & "/1-n-n/"'只能1-n-n了
	page=0
	OX163_urlpage_Referer="Host: 95.211.21.16" & vbcrlf & "Referer: http://g.e-hentai.org/"
Else
	return_download_url="inet|10,13|http://www.163.com/?Delay_5s-利用163页面延迟5秒" & "|http://g.e-hentai.org/"' & vbcrlf & "Cookie: lastvisit=1238694351; impcookie=71b644e1b37ee8c8a6052d952804eb54df00d475cb9c72bf3e82e43384c30ab8; Apache=168296599x0.116+1240025775x1022208447; b=%3A%3Amm4l%2Cmm4e%2Cmm4s%2Cmm4p%2Cih53%2Cih56; ut=1%3Aq1YqM1SyqlYqTi1WslJKya%2FJzsxIMa8x0kkxgTINdVJMEaKZMFElHaXc1JJEkN6S4iKQbgszEwOD2tpaAA%3D%3D; __utma=11274144.128020507.1240025776.1241191778.1241714935.4; __utmz=11274144.1240025776.1.1.utmccn=(direct)&for_ox163_replace_vline&utmcsr=(direct)&for_ox163_replace_vline&utmcmd=(none); geo=1%3Aq1YqM1SyqlZKyU1UslIyUNJRSs4vBbKc%2FYDMovQ8IDPYA8gszkwHMi1M0swtjc0NUkzTjJMsExMNTNISU5JNDYwSLQyNUpPTlGprAQ%3D%3D; __utmb=11274144; x=344921-573e35987c794f547f5eeb330dde7dbbabbb0450"
	delay_tf=1
	OX163_urlpage_Referer=""
End If
'OX163_urlpage_Referer="Host: 95.211.21.16" & vbcrlf & "Referer: http://g.e-hentai.org/"' & vbcrlf & "Cookie: lastvisit=1238694351; impcookie=71b644e1b37ee8c8a6052d952804eb54df00d475cb9c72bf3e82e43384c30ab8; Apache=168296599x0.116+1240025775x1022208447; b=%3A%3Amm4l%2Cmm4e%2Cmm4s%2Cmm4p%2Cih53%2Cih56; ut=1%3Aq1YqM1SyqlYqTi1WslJKya%2FJzsxIMa8x0kkxgTINdVJMEaKZMFElHaXc1JJEkN6S4iKQbgszEwOD2tpaAA%3D%3D; __utma=11274144.128020507.1240025776.1241191778.1241714935.4; __utmz=11274144.1240025776.1.1.utmccn=(direct)&for_ox163_replace_vline&utmcsr=(direct)&for_ox163_replace_vline&utmcmd=(none); geo=1%3Aq1YqM1SyqlZKyU1UslIyUNJRSs4vBbKc%2FYDMovQ8IDPYA8gszkwHMi1M0swtjc0NUkzTjJMsExMNTNISU5JNDYwSLQyNUpPTlGprAQ%3D%3D; __utmb=11274144; x=344921-573e35987c794f547f5eeb330dde7dbbabbb0450"
'Cookie: lastvisit=1238694351; impcookie=71b644e1b37ee8c8a6052d952804eb54df00d475cb9c72bf3e82e43384c30ab8; Apache=168296599x0.116+1240025775x1022208447; b=%3A%3Amm4l%2Cmm4e%2Cmm4s%2Cmm4p%2Cih53%2Cih56; ut=1%3Aq1YqM1SyqlYqTi1WslJKya%2FJzsxIMa8x0kkxgTINdVJMEaKZMFElHaXc1JJEkN6S4iKQbgszEwOD2tpaAA%3D%3D; __utma=11274144.128020507.1240025776.1241191778.1241714935.4; __utmz=11274144.1240025776.1.1.utmccn=(direct)|utmcsr=(direct)|utmcmd=(none); geo=1%3Aq1YqM1SyqlZKyU1UslIyUNJRSs4vBbKc%2FYDMovQ8IDPYA8gszkwHMi1M0swtjc0NUkzTjJMsExMNTNISU5JNDYwSLQyNUpPTlGprAQ%3D%3D; __utmb=11274144; x=344921-573e35987c794f547f5eeb330dde7dbbabbb0450
End Function
'-------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list = ""

If DateDiff("s", start_time, Now()) < 35 and counts>9 Then
	delay_tf=1
	return_albums_list="1|inet|10,13|http://www.163.com/?Delay_30s-利用163页面延迟30秒"
	Exit Function
ElseIf delay_tf=1 Then
	OX163_urlpage_Referer="Host: 95.211.21.16" & vbcrlf & "Referer: http://g.e-hentai.org/"
	counts=0
	delay_tf=0
	return_albums_list="1|inet|10,13|" & page_url & page
	Exit Function
End If

start_time=Now()
If InStr(LCase(html_str), "<a class=""noul""") > 0 Then
	retry=0
	If InStr(LCase(html_str), "<p class=""g4"">") > 0 Then
		url_str=Mid(html_str,InStr(LCase(html_str), "<p class=""g4"">")+14)
		url_str=Mid(url_str,1,InStr(LCase(url_str), "</p>")-1)
		url_str=rename_utf8(url_str)
	Else
		url_str="E-Hentai Unknow Title"
	End If
	
	Dim split_str,page_check,page_check_count
	page_check=Mid(html_str,InStr(LCase(html_str), "<p class=""ip"">")+14)
	page_check=Mid(page_check,1,InStr(LCase(page_check), " images")-1)
	page_check=Mid(page_check,InStr(LCase(page_check), "-")+1)
	page_check_count=Trim(Mid(page_check,InStr(LCase(page_check), "of ")+3))
	page_check=Trim(Mid(page_check,1,InStr(LCase(page_check), " of")-1))
	
	html_str=Mid(html_str,InStr(LCase(html_str), "<td class=""ttd"""))
	html_str=Mid(html_str,InStr(LCase(html_str), "<a class=""noul"" href=""")+22)
	html_str=Mid(html_str,1,InStr(LCase(html_str), "</div>")-1)
		
	split_str=Split(html_str, "<a class=""noul"" href=""", -1, 1)
	
	For split_i = 0 To UBound(split_str)
		'url
		split_str(split_i) = Mid(split_str(split_i), 1, InStr(split_str(split_i), Chr(34)) - 1)	
		return_albums_list = return_albums_list & "0|1|" & split_str(split_i) & "|" & url_str & "|" & url_str & vbCrLf
	Next
	
	If (page_check_count<>page_check) and IsNumeric(page_check_count) and IsNumeric(page_check) Then
		counts=counts+1
		page=page+1
		If counts>9 Then
			return_albums_list = return_albums_list & "1|inet|10,13|http://www.163.com/?Delay_30s-利用163页面延迟30秒"
			OX163_urlpage_Referer=""
		Else
			return_albums_list = return_albums_list & "1|inet|10,13|" & page_url & page
			OX163_urlpage_Referer="Host: 95.211.21.16" & vbcrlf & "Referer: http://g.e-hentai.org/"
		End If
	Else
		return_albums_list = return_albums_list & "0"
	End If
ElseIf retry<4 and html_str<>"" Then
	retry=retry+1
	return_albums_list="1|inet|10,13|" & page_url & page	
Else
return_albums_list = "0"
End If
End Function
'-------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
If DateDiff("s", start_time, Now()) < 0 Then
	return_download_list="1|inet|10,13|http://www.163.com/?Delay_5s-利用163页面延迟5秒"
	Exit Function
ElseIf delay_tf=1 Then
	OX163_urlpage_Referer="Host: 95.211.21.16" & vbcrlf & "Referer: http://g.e-hentai.org/"
	delay_tf=0
	return_download_list="1|inet|10,13|" & url_str
	Exit Function
End If

If InStr(LCase(html_str),"title=""last""")>0 Then
		
	Dim file_name,file_type
	If InStr(LCase(html_str),"<h1 class=""")>0 Then
		file_name=Mid(html_str,InStr(LCase(html_str),"<h1 class="""))
		file_name=Mid(file_name,InStr(file_name,">")+1)
		file_name=replace(Mid(file_name,1,InStr(file_name,"<")-1),"|","_") & "_"
		If Len(file_name)>100 Then file_name=Left(file_name,100) & "_"
	Else
		file_name=""
	End If
	
	'pic_url
	html_str=Mid(html_str,InStr(LCase(html_str),"title=""last"""))
	html_str=Mid(html_str,InStr(LCase(html_str),"<img id="""))
	html_str=Mid(html_str,InStr(LCase(html_str)," src=""")+6)
	file_type=Mid(html_str,InStr(html_str,Chr(34))+1)
	
	html_str=Mid(html_str,1,InStr(html_str,Chr(34))-1)
	
	file_type=Mid(file_type,InStr(LCase(file_type)," alt=""")+6)
	file_type=Mid(file_type,1,InStr(file_type,Chr(34))-1)
	
	If file_type<>"" Then
		If file_type<>"" Then
			If Mid(file_type,InStrrev(file_type,"."))=Mid(html_str,InStrrev(html_str,".")) Then
				file_name=file_name & file_type
			Else
				file_name=file_name & file_type & Mid(html_str,instrrev(html_str,"."))
			End If
		End If
	Else
		'http://rsb.hentaiverse.net/ehg/image.php?f=7767ec5a5347dc720950fc49420ca6189dc89c05-306395-950-1328-jpg&t=348757-8ae15de1553bb3afc9a9a84ca75ce23fa80421f8&n=hanaharu200810_4.jpg
		'http://gu.e-hentai.org/image.php?f=7767ec5a5347dc720950fc49420ca6189dc89c05-306395-950-1328-jpg&t=348757-ed799eecd5330a109cad57d15f04cb3ff32d12b8&n=hanaharu200810_4.jpg
		file_name=file_name & Mid(html_str,instrrev(html_str,"/")+1)
	End If
	
	file_name=replace(replace(file_name,Chr(10),""),Chr(13),"")
	file_name=rename_utf8(file_name)
	
	return_download_list="|" & html_str & "|" & file_name & "|" & file_name & vbCrLf & "0"'&nl=1
	
ElseIf retry<4 and html_str<>"" Then
	retry=retry+1
	return_download_list="1|inet|10,13|" & url_str
Else
	retry=0
	return_download_list="0"
End If
End Function

'----------------------------------------------
Function rename_utf8(byval utf8_Str)
If Len(utf8_Str)=0 Then Exit Function
For i=1 to Len(utf8_Str)
	If  Asc(Mid(utf8_Str,i,1))=63 Then utf8_Str=replace(utf8_Str,Mid(utf8_Str,i,1),"_")
Next
rename_utf8=utf8_Str
End Function