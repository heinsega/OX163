'2011-4-18 http://www.shanhaijing.net/163
Dim sohu_ID,page_num,photo_type,album_type
Function return_download_url(ByVal url_str)
On Error Resume Next
return_download_url = ""
'sohu相册album
'http://pp.sohu.com/setlist.jhtml?method=list&userId=26930220&pageNo=1
'http://pp.sohu.com/user/26930220/setlist
'http://pp.sohu.com/member/yangwang-xk
'sohu相册photo
'http://pp.sohu.com/photosetview-34368955-26930220.html
'http://pp.sohu.com/photoview-293236651-26930220.html
'http://pp.sohu.com/photoview-292591789-26930220.html#292591788
'sohu博客album
'http://anita712.blog.sohu.com/
'http://anita712.blog.sohu.com/album/
'http://anita712.blog.sohu.com/album/setlist
'http://anita712.blog.sohu.com/album/setlist.jhtml?method=list&userId=28734030&pageNo=2
'sohu博客photo
'http://anita712.blog.sohu.com/album/photosetview-33490707-28734030.html
'http://anita712.blog.sohu.com/album/photoview-292673977-28734030.html
'http://anita712.blog.sohu.com/album/photoview-292673977-28734030.html#292673968
sohu_ID=""
page_num=1
If InStr(LCase(url_str), "http://pp.sohu.com/user/") > 0 Then
    url_str = Mid(url_str,InStr(LCase(url_str), "http://pp.sohu.com/user/")+24)
    url_str = Mid(url_str,1,InStr(url_str, "/")-1)
    If IsNumeric(url_str)=true Then
			album_type="http://pp.sohu.com"
    	sohu_ID=url_str
    	return_download_url = "inet|10,13|" & album_type & "/setlist.jhtml?method=list&userId=" & sohu_ID & "&pageNo=" & page_num
    End If
ElseIf InStr(LCase(url_str), "/setlist.jhtml?") > 0 Then
		If InStr(LCase(url_str), "http://pp.sohu.com/") > 0 Then
			album_type="http://pp.sohu.com"
		Else
			'.blog.sohu.com/album/
			album_type=Mid(url_str,1,InStr(LCase(url_str), ".blog.sohu.com")-1) & ".blog.sohu.com/album"
		End If
    url_str = Mid(url_str,InStr(LCase(url_str), "userid=")+7)
    url_str = Mid(url_str,1,InStr(url_str, "&")-1)
    If IsNumeric(url_str)=true Then
    	sohu_ID=url_str
    	return_download_url = "inet|10,13|" & album_type & "/setlist.jhtml?method=list&userId=" & sohu_ID & "&pageNo=" & page_num
    End If
ElseIf InStr(LCase(url_str), "http://pp.sohu.com/member/") > 0 Then
	album_type="http://pp.sohu.com"
	sohu_ID="get_sohu_ID"
	return_download_url = "inet|10,13|" & url_str
ElseIf InStr(LCase(url_str), ".blog.sohu.com") > 0 And InStr(LCase(url_str), ".blog.sohu.com/album/photo") <1 Then
	album_type=Mid(url_str,1,InStr(LCase(url_str), ".blog.sohu.com")-1) & ".blog.sohu.com/album"
	sohu_ID="get_sohu_ID"
	return_download_url = "inet|10,13|" & album_type & "/setlist"
Else
	sohu_ID=url_str
	return_download_url = "inet|10,13|" & url_str
End If
End Function

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list = ""
If sohu_ID="get_sohu_ID" and InStr(html_str, "var _uid=") > 0 Then
	sohu_ID=""
	html_str=Mid(html_str, InStr(html_str, "var _uid=")+9)
	html_str=Mid(html_str, 1,InStr(html_str, ";")-1)
	If IsNumeric(html_str)=true Then
	sohu_ID=html_str
	page_num=1
	return_albums_list = return_albums_list & "1|inet|10,13|" & album_type & "/setlist.jhtml?method=list&userId=" & sohu_ID & "&pageNo=" & page_num
	End If
ElseIf InStr(LCase(html_str), "<div class=""albumcover"">") > 0 Then
	html_str = Mid(html_str, InStr(LCase(html_str), "<div class=""albumcover"">")+24)
	Dim str_split
	str_split=split(html_str,"<div class=""albumCover"">")
	For i=0 to UBound(str_split)
		str_split(i)=Mid(str_split(i),InStr(str_split(i),"<a href='")+9)
		
		'url
		url_str=Mid(str_split(i),1,InStr(str_split(i),"'")-1)

		str_split(i)=Mid(str_split(i),InStr(str_split(i),"title='")+7)

		'name
		html_str=Trim(Mid(str_split(i),1,InStr(str_split(i),"'")-1))
		If html_str="" Then html_str=sohu_ID & "_No_Name_Album"
		
		'pic number
		str_split(i)=Mid(str_split(i),InStr(str_split(i),"<span class=""count"">")+21)
		str_split(i)=Mid(str_split(i),1,InStr(str_split(i),")")-1)
		If IsNumeric(str_split(i))=false Then str_split(i)="0"
		If InStr(LCase(album_type), "http://pp.sohu.com")=1 Then
			return_albums_list = return_albums_list & "0|" & str_split(i) & "|" & album_type & url_str & "|" & html_str & vbcrlf
		Else		
			return_albums_list = return_albums_list & "0|" & str_split(i) & "|" & album_type & Mid(url_str,2) & "|" & html_str & vbcrlf
		End If
	Next	
	If UBound(str_split)<15 Then
		return_albums_list = return_albums_list & "0"
	Else
		page_num=page_num+1
		return_albums_list = return_albums_list & "1|inet|10,13|" & album_type & "/setlist.jhtml?method=list&userId=" & sohu_ID & "&pageNo=" & page_num
	End If
End If
End Function
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list=""

If InStr(LCase(sohu_ID), "http://pp.sohu.com/photosetview-")=1 Or InStr(LCase(sohu_ID), ".blog.sohu.com/album/photosetview-")>0 Then
	If InStr(LCase(sohu_ID), ".blog.sohu.com/album/")>0 Then
		return_download_list ="1|inet|10,13|" & Mid(sohu_ID,1,InStr(LCase(sohu_ID), ".blog.sohu.com/album/")+20) & "photoview-"
	Else
		return_download_list ="1|inet|10,13|http://pp.sohu.com/photoview-"
	End If
	sohu_ID=""
	html_str=Mid(html_str, InStr(html_str, "var _uid=")+9)
	sohu_ID=Mid(html_str, 1,InStr(html_str, ";")-1)
	html_str=Mid(html_str, InStr(html_str, "var initPhotoList = {"""))
	html_str=Mid(html_str, 1, InStr(html_str,"var initPhotosetList = {"""))
	If InStr(html_str, "[{""id"":")>0 Then
		html_str=Mid(html_str, InStr(html_str, "[{""id"":")+Len("[{""id"":"))
		photo_type=1
	ElseIf InStr(html_str, ",""id"":")>0 Then
		html_str=Mid(html_str, InStr(html_str, ",""id"":")+Len(",""id"":"))
		photo_type=2
	Else
		Exit Function
	End If
	html_str=Mid(html_str, 1,InStr(html_str, ",")-1)
	If IsNumeric(sohu_ID)=true and IsNumeric(html_str)=true Then	
		'http://pp.sohu.com/photoview-293236670-26930220.html#293236670
		'http://anita712.blog.sohu.com/album/photoview-292673977-28734030.html
		return_download_list =return_download_list & html_str & "-" & sohu_ID & ".html"
		sohu_ID=Mid(return_download_list,InStr(LCase(return_download_list),"http://"))
	Else
		return_download_list=""
	End If
ElseIf InStr(LCase(sohu_ID), "http://pp.sohu.com/photoview-")=1 Or InStr(LCase(sohu_ID), ".blog.sohu.com/album/photoview-")>0 Then
	If InStr(html_str,"var initPhotoList = {""")>0 Then
		html_str=Mid(html_str,InStr(html_str,"var initPhotoList = {"""))
		html_str=Mid(html_str,1,InStr(html_str,"};")-1)
		html_str=Mid(html_str,InStr(html_str,"""imageSize"":")+Len("""imageSize"":"))
		Dim photo_split
		photo_split=split(html_str,"""imageSize"":")
		For i=0 to UBound(photo_split)
			'url
			url_str=Mid(photo_split(i),InStr(photo_split(i),"""source"":""")+10)
			url_str=Mid(url_str,1,InStr(url_str,Chr(34))-1)
			
			If url_str="" Then
			url_str=Mid(photo_split(i),InStr(photo_split(i),"""middle"":""")+10)
			url_str=Mid(url_str,1,InStr(url_str,Chr(34))-1)				
			End If
			
			If url_str="" Then
			url_str=Mid(photo_split(i),InStr(photo_split(i),"""small130"":""")+12)
			url_str=Mid(url_str,1,InStr(url_str,Chr(34))-1)				
			End If
			
			'name
			html_str=Mid(photo_split(i),InStr(photo_split(i),"""name"":""")+8)
			html_str=Trim(replace(Mid(html_str,1,InStr(html_str,Chr(34))-1),"|","_"))
			html_str=html_str & Mid(url_str,instrrev(url_str,"."))
			
			'desc
			photo_split(i)=Mid(photo_split(i),InStr(photo_split(i),"""imageSizeDesc"":""")+17)
			photo_split(i)=Mid(photo_split(i),1,InStr(photo_split(i),Chr(34))-1)
						
			'info
			photo_split(i)= "|" & url_str & "|" & html_str & "|" & photo_split(i) & vbcrlf
		Next
		return_download_list=join(photo_split,"") & "0"
	End If
End If
End Function
