'2008-9-29 http://www.shanhaijing.net/163
Dim pic_counter,tom_album_type,nameid,page

Function return_download_url(byVal url_str)
On Error Resume Next
return_download_url = ""
pic_counter=0
If InStr(1,url_str,"http://games.tom.com/gamepic/",1)>0 Then
	url_str=Mid(LCase(url_str),InStr(LCase(url_str),"http://games.tom.com/gamepic/")+29)
	url_str=Mid(url_str,1,InStr(url_str,"_")-1)
	If IsNumeric(url_str) Then
	    nameid=url_str
	    page=1
	    return_download_url="inet|10,13,34|" & "http://games.tom.com/gamepic/" & url_str & "_1.html"
	    '原始tom pic图片
	    tom_album_type=0
	End If
ElseIf InStr(1,url_str,"http://games.tom.com/manhua/",1)>0 Then
	return_download_url="inet|10,13|" & url_str
	tom_album_type=2
ElseIf InStr(1,url_str,"http://games.tom.com/",1)>0 Then
	'http://games.tom.com/2008-09-22/0082/09697365.html
	'http://games.tom.com/2008-09-22/0G5D/22839063_02.html
	If InStr(url_str,"_")>0 Then
		url_str=Mid(url_str,1,instrrev(url_str,"_")-1) & ".html"
	End If
	nameid=Mid(url_str,1,instrrev(url_str, ".html")-1)
	page=1
	return_download_url="inet|10,13|" & url_str
	'新tom pic图片
	tom_album_type=1
ElseIf InStr(1,url_str,"http://photo.tom.com/personal.php?pid=",1)>0 Then
	'http://photo.tom.com/personal.php?pid=509399332&class=1&mod=2&page=2
	'http://photo.tom.com/personal.php?pid=20027993&class=1&mod=2
	page=1
	url_str=Mid(url_str,InStr(1,url_str,"http://photo.tom.com/personal.php?pid=",1)+38)
	If InStr(url_str,"&")>0 Then url_str=Mid(url_str,1,InStr(url_str,"&")-1)
	url_str="http://photo.tom.com/personal.php?pid=" & url_str & "&class=1&mod=2"
	nameid=url_str
	return_download_url="inet|10,13|" & url_str
	tom_album_type=3
ElseIf InStr(1,url_str,"http://photo.tom.com/pim.php?",1)>0 Then
	'http://photo.tom.com/pim.php?class=1&mod=2
	page=1
	url_str="http://photo.tom.com/pim.php?class=1&mod=2"
	nameid=url_str
	return_download_url="inet|10,13|" & url_str
	tom_album_type=3
ElseIf InStr(1,url_str,"http://photo.tom.com/photo.php?id=",1)>0 Then
	page=1	
	return_download_url="inet|10,13|" & url_str
	If InStr(1,url_str,"&totle=",1)>0 Then
		url_str=Mid(url_str,InStr(1,url_str,"&totle=",1)+7)
		If IsNumeric(url_str) Then
			pic_counter=Int(url_str)
		Else
			pic_counter=1
		End If
	End If
	tom_album_type=4
ElseIf InStr(1,url_str,"http://photo.tom.com/pim_photo.php?id=",1)>0 Then
	page=1	
	return_download_url="inet|10,13|" & url_str
	If InStr(1,url_str,"&totle=",1)>0 Then
		url_str=Mid(url_str,InStr(1,url_str,"&totle=",1)+7)
		If IsNumeric(url_str) Then
			pic_counter=Int(url_str)
		Else
			pic_counter=1
		End If
	End If
	tom_album_type=4
End If
End Function
'----------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim split_str
return_albums_list = ""

Select Case tom_album_type
Case 2
	If InStr(1, html_str, "<div class=""text""><a href=""", 1) > 0 Then		
		html_str=Mid(html_str,InStr(1, html_str, "<div class=""text""><a href=""", 1)+27)
		split_str = Split(html_str, "<div class=""text""><a href=""",-1,1)
		For i = 0 To UBound(split_str)
			'url
			url_str="http://games.tom.com" & Mid(split_str(i),1,InStr(split_str(i),Chr(34))-1)
			
			'name
			split_str(i)=Mid(split_str(i),InStr(split_str(i),">")+1)
			split_str(i)=Trim(Mid(split_str(i),1,InStr(split_str(i),"<")-1))
			
			return_albums_list = return_albums_list & "0||" & url_str & "|" & split_str(i) & vbcrlf
		Next
		return_albums_list = return_albums_list & "0"
	Else
	    return_albums_list = "0"
	End If
Case 3
	If InStr(1, html_str, "<dd class=""col444"">", 1) > 0 Then
		Dim htm_temp
		htm_temp=html_str
		html_str=Mid(html_str,InStr(1, html_str, "<dd class=""col444"">", 1)+19)
		split_str = Split(html_str, "<dd class=""col444"">",-1,1)
		For i = 0 To UBound(split_str)
			split_str(i)=Mid(split_str(i),InStr(1,split_str(i),"<a href=""",1)+9)
			'url
			url_str="http://photo.tom.com/" & Mid(split_str(i),1,InStr(split_str(i),Chr(34))-1)
			
			'name
			split_str(i)=Mid(split_str(i),InStr(1,split_str(i),"专辑名：",1)+Len("专辑名："))
			html_str=Trim(Mid(split_str(i),1,InStr(1,split_str(i),"</a>",1)-1))
			If html_str="" Then html_str="No_Name_Album_" & i
			
			'totle
			split_str(i)=Mid(split_str(i),InStr(1,split_str(i),"<dd>共",1)+Len("<dd>共"))
			split_str(i)=Trim(Mid(split_str(i),1,InStr(1,split_str(i),"张照片",1)-1))
			If IsNumeric(split_str(i)) Then
				split_str(i)=Int(split_str(i))
				If split_str(i)>0 Then
					url_str=url_str & "&totle=" & split_str(i)
					Else
					split_str(i)=1
					url_str=url_str & "&totle=1"
				End If
			Else
				split_str(i)=1
				url_str=url_str & "&totle=1"
			End If
			
			return_albums_list = return_albums_list & "0|" & split_str(i) & "|" & url_str & "|" & html_str & vbcrlf
		Next
		
		page=page+1
		If InStr(1,htm_temp,nameid & "&page=" & page,1)>0 Then
			return_albums_list = return_albums_list & page & "|inet|10,13|" & nameid & "&page=" & page
		End If
	End If
Case 4
	If InStr(1,html_str,"<a href=""personal.php?pid=",1)>0 Then
		html_str=Mid(html_str,InStr(1,html_str,"<a href=""personal.php?pid=",1)+9)
		html_str="http://photo.tom.com/" & Mid(html_str,1,InStr(html_str,Chr(34))-1)
		page=1
		nameid=url_str
		tom_album_type=3
		return_albums_list ="2|inet|10,13|" & html_str
	ElseIf InStr(1,html_str,"<a href=""pim.php",1)>0 Then
		page=1
		nameid=url_str
		tom_album_type=3
		return_albums_list ="2|inet|10,13|http://photo.tom.com/pim.php?class=1&mod=2"
	End If
End Select
End Function
'----------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""

Dim max_page,page_format,split_str,pic_url,pic_type

Select Case tom_album_type
Case 0
	If InStr(1, html_str, "<img src=http://img.games.tom.com/gamepic/upload_dir/", 1) > 0 Then
	html_str = Mid(html_str, InStr(1, html_str, "<td align=right class=list>", 1) + 27)
	
	'总页码
	max_page = Mid(html_str, 1, InStr(1, html_str, "页", 1) - 1)
	max_page = Trim(Mid(max_page, InStr(1, max_page, "共", 1) + 1))
	If IsNumeric(max_page) = False Then max_page = 1
	
	page_format = Len(max_page * 25)
	
	split_str = Split(html_str, "<img src=http://img.games.tom.com/gamepic/upload_dir/")
	
	For i = 1 To UBound(split_str)
	    pic_counter = pic_counter + 1
	    pic_url = Mid(split_str(i), 1, InStr(split_str(i), " ") - 1)
	    pic_url = Replace(pic_url, "_s", "")
	    pic_type = split(pic_url,"/")
	    Do While InStr(pic_type(UBound(pic_type)),".")>0
	    pic_type(UBound(pic_type))=Mid(pic_type(UBound(pic_type)),InStr(pic_type(UBound(pic_type)),".")+1)
	    loop
	    
	    split_str(i) = Mid(split_str(i), InStr(1, split_str(i), "alt=", 1) + 4)
	    split_str(i) = Mid(split_str(i), 1, InStr(1, split_str(i), "></a>", 1) - 1) & "_图片"
	    For j = 1 To page_format - Len(pic_counter)
	    split_str(i) = split_str(i) & "0"
	    Next
	    split_str(i) = split_str(i) & pic_counter
	    
	    return_download_list = return_download_list & pic_type(UBound(pic_type)) & "|http://img.games.tom.com/gamepic/upload_dir/" & pic_url & "|" & split_str(i) & "|" & split_str(i) & vbCrLf
	    
	Next
	
	If Int(page) < Int(max_page) Then
	    page=page+1
	    return_download_list = return_download_list & page & "|inet|10,13,34|http://games.tom.com/gamepic/" & nameid & "_" & page & ".html"
	Else
	    return_download_list = return_download_list & "0"
	End If
	
	Else
	    return_download_list = "0"
	End If
Case 1
	
	html_str=replace(html_str,"<img  src=""","<img src=""")
	html_str=replace(html_str,"<a  href=""","<a href=""")
	If InStr(1, html_str, "<CENTER><a href=""", 1) > 0 Then
	'有大图的照片
		url_str=html_str
		html_str=Mid(html_str,InStr(1, html_str, "<CENTER><a href=""", 1)+17)
		split_str = Split(html_str, "<CENTER><a href=""",-1,1)
		For i = 0 To UBound(split_str)
			'url
			pic_url=Mid(split_str(i),1,InStr(split_str(i),Chr(34))-1)
			pic_url=Mid(pic_url,InStr(pic_url,"=")+1)
			'name
			split_str(i)=Mid(split_str(i),InStr(1,split_str(i),"alt=""",1)+5)
			split_str(i)=Mid(split_str(i),1,InStr(split_str(i),Chr(34))-1) & Mid(pic_url,instrrev(pic_url,"/")+1)
			return_download_list = return_download_list & "|" & pic_url & "|" & split_str(i) & "|" & split_str(i) & vbCrLf
		Next
		
		If InStr(1,url_str,"<div class=""yema"">")>0 and InStr(1,url_str,"<a href=""#"" target=""_self""><img src=""/news_end/images/next.gif""")<1 and page<1000 Then
			page=page+1
			html_str=nameid & "_" & format_page(page) & ".html"
			return_download_list = return_download_list & page & "|inet|10,13|" & html_str
		Else
			return_download_list = return_download_list & "0"
		End If
		
	ElseIf InStr(1, html_str, "<CENTER><img src=""", 1) > 0 Then
	'没大图的照片
		url_str=html_str
		html_str=Mid(html_str,InStr(1, html_str, "<CENTER><img src=""", 1)+18)
		split_str = Split(html_str, "<CENTER><img src=""",-1,1)
		For i = 0 To UBound(split_str)
			'url
			pic_url="http://games.tom.com" & Mid(split_str(i),1,InStr(split_str(i),Chr(34))-1)
			
			'name
			split_str(i)=Mid(split_str(i),InStr(1,split_str(i),"alt=""",1)+5)
			split_str(i)=Mid(split_str(i),1,InStr(split_str(i),Chr(34))-1) & Mid(pic_url,instrrev(pic_url,"/")+1)
			return_download_list = return_download_list & "|" & pic_url & "|" & split_str(i) & "|" & split_str(i) & vbCrLf
		Next
		
		If InStr(1,url_str,"<div class=""yema"">")>0 and InStr(1,url_str,"<a href=""#"" target=""_self""><img src=""/news_end/images/next.gif""")<1 and page<1000 Then
			page=page+1
			html_str=nameid & "_" & format_page(page) & ".html"
			return_download_list = return_download_list & page & "|inet|10,13|" & html_str
		Else
			return_download_list = return_download_list & "0"
		End If
	ElseIf InStr(1, html_str, ">将此页漫画加入收藏夹</a>", 1) > 0 Then
	'漫画
		'url
		html_str=Mid(html_str,InStr(1,html_str,"prev.jpg""",1))
		html_str=Mid(html_str,InStr(1,html_str,"<img src=""",1)+10)
		url_str="http://games.tom.com" & Mid(html_str,1,InStr(html_str,Chr(34))-1)
		
		page_format=html_str
		
		'name
		html_str=Mid(html_str,InStr(1,html_str,"alt=""",1)+5)
		html_str=Mid(html_str,1,InStr(html_str,Chr(34))-1) & "_"
		html_str=html_str & Mid(url_str,instrrev(url_str,"/")+1)
		
		return_download_list = return_download_list & "|" & url_str & "|" & html_str & "|" & html_str & vbCrLf

		page=page+1
		pic_url=replace(nameid & "_" & format_page(page) & ".html","http://games.tom.com","")		
		If InStr(1,page_format,pic_url,1)>0 Then
			return_download_list = return_download_list & page & "|inet|10,13|" & "http://games.tom.com" & pic_url
		End If
	End If
		
Case 4
	If InStr(1,html_str,"<div class=""photoshow",1)>0 Then
		html_str=Mid(html_str,InStr(1,html_str,"<div class=""photoshow",1))
		html_str=Mid(html_str,InStr(1,html_str,"<h2 class=""titlebox"">",1)+21)
		'name
		url_str=Trim(Mid(html_str,1,InStr(html_str,"<")-1))
		
		html_str=Mid(html_str,InStr(1,html_str,"""><img",1)+6)
		pic_url=Mid(html_str,InStr(1,html_str,"src=""",1)+5)
		'url
		pic_url=Mid(pic_url,1,InStr(pic_url,Chr(34))-1)
		pic_url=replace(pic_url,"_w3.",".")
		
		'name
		If url_str="" Then
			url_str=Mid(pic_url,instrrev(pic_url,"/")+1)
		Else
			url_str=url_str & Mid(pic_url,instrrev(pic_url,"."))
		End If
		
		return_download_list = "|" & pic_url & "|" & url_str & "|" & vbCrLf
		
		If page<pic_counter Then
			page=page+1
			html_str=Mid(html_str,InStr(1,html_str,"<a href=""",1)+9)
			'next_url
			html_str="http://photo.tom.com/" & Mid(html_str,1,InStr(html_str,Chr(34))-1) & "&totle=" & pic_counter
			
			return_download_list = return_download_list & pic_counter & "|inet|10,13|" & html_str		
		End If
		
	End If

End Select
End Function

Function format_page(ByVal page_number)
If Len(page_number)<2 Then page_number="0" & page_number
format_page=page_number
End Function