'2010-12-27 http://www.shanhaijing.net/163

Function return_download_url(ByVal url_str)
On Error Resume Next
return_download_url = ""
Dim url_str_split
If InStr(1, url_str, "http://anime.xunlei.com/book/", 1) > 0 Then
    'http://anime.xunlei.com/Book/823
    url_str = Trim(Replace(LCase(url_str), "http://anime.xunlei.com/book/", ""))
    url_str = Mid(url_str,1,InStr(url_str,"/")-1)
    If IsNumeric(url_str) Then
    return_download_url = "inet|10,13|http://anime.xunlei.com/book/" & url_str & "/"
    Else
    return_download_url = ""
    End If
ElseIf InStr(1, url_str, "http://images.anime.xunlei.com/book/segment/", 1) > 0 Then
    'http://images.anime.xunlei.com/book/segment/18/17194.html
    'http://images.anime.xunlei.com/book/segment/18/17194.html?page=7#photobox
    url_str = Replace(LCase(url_str), "http://images.anime.xunlei.com/book/segment/", "")
    url_str = Mid(url_str,1,InStr(1, url_str, ".html", 1)-1)
    url_str_split=split(url_str,"/")
    If IsNumeric(url_str_split(0)) and UBound(url_str_split)=1 Then
	If IsNumeric(url_str_split(1)) Then
	return_download_url = "inet|10,13|http://images.anime.xunlei.com/book/segment/" & url_str_split(0) & "/" & url_str_split(1) & ".html|http://images.anime.xunlei.com/"
	End If
    End If
ElseIf InStr(1, url_str, "http://images.anime.xunlei.com/collections/", 1) > 0 Then
    'http://images.anime.xunlei.com/collections/2/1584.html?page=6#photobox
    url_str = Replace(LCase(url_str), "http://images.anime.xunlei.com/collections/", "")
    url_str = Mid(url_str,1,InStr(1, url_str, ".html", 1)-1)
    url_str_split=split(url_str,"/")
    If IsNumeric(url_str_split(0)) and UBound(url_str_split)=1 Then
	If IsNumeric(url_str_split(1)) Then
	return_download_url = "inet|10,13|http://images.anime.xunlei.com/collections/" & url_str_split(0) & "/" & url_str_split(1) & ".html|http://images.anime.xunlei.com/"
	End If
    End If
End If
End Function
'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next

return_albums_list = ""
If InStr(lcase(html_str), "<li style=""height:22px""><a href=") > 0 Then
    Dim comic_Name
    comic_Name = Mid(html_str, InStr(lcase(html_str), "<title>") + 7)
    comic_Name = replace(Mid(comic_Name, 1, InStr(lcase(comic_Name), "</title>") - 1),"漫画","")
    if len(comic_Name)>100 then comic_Name = trim(left(comic_Name,148) & "~~")
    html_str = Mid(html_str, InStr(lcase(html_str), "<li style=""height:22px""><a href=") + 33)
    html_str = Mid(html_str,1, InStr(lcase(html_str), "</div></div>") - 1)
    
    album_list = Split(html_str, "<li style=""height:22px""><a href=" & Chr(34), -1, 1)
    
    For i =0  To UBound(album_list)
        If InStr(1, album_list(i), "http://images.anime.xunlei.com/book/segment/", 1) > 0 Then
        'url
        album_list(i)=Mid(album_list(i),InStr(1, album_list(i), "http://images.anime.xunlei.com/book/segment/", 1))
        html_str=Mid(album_list(i), 1, InStr(album_list(i), Chr(34)) - 1) & "|"
        'name
        album_list(i) = Mid(album_list(i), InStr(album_list(i), ">") + 1)
        url_str = comic_Name & "_" & Mid(album_list(i), 1, InStr(1, album_list(i), "</a>", 1) - 1)
        'pic_number
        album_list(i) = Mid(album_list(i), InStr(1, album_list(i), "</a>(", 1) + 5)
        album_list(i) = Mid(album_list(i),1, InStr(1, album_list(i), "页)", 1) - 1)
        If IsNumeric(album_list(i))=false Then album_list(i)=""
        return_albums_list = return_albums_list & "0|" & album_list(i) & "|" & html_str & url_str & "|" & url_str & vbcrlf
        End If
    Next
    return_albums_list = return_albums_list & "0"
Else
    return_albums_list = "0"
End If
End Function
'----------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list=""

Dim comic_Name,pic_type,page_num
'comic 'http://images.mh.xunlei.com/origin/
If InStr(html_str, "var images_id_arr = new Array();") <1 and InStr(1, html_str, "var images_arr = new Array();", 1) > 0 Then
    comic_Name = Mid(html_str, InStr(1, html_str, "<title>", 1) + 7)
    comic_Name = Mid(comic_Name, 1, InStr(1, comic_Name, "</title>", 1) - 1)
    comic_Name = Mid(comic_Name, 1, InStr(comic_Name, "在线漫画") - 1) & "第"
   
    html_str = Mid(html_str, InStr(1, html_str, "var images_arr = new Array();", 1))
    html_str = Mid(html_str, InStr(1, html_str, "images_arr[", 1)+11)
    html_str = Mid(html_str, 1, InStr(1, html_str, "var page", 1))
    pic_type = split(html_str,"images_arr[")

For i=0 to UBound(pic_type)
	page_num=Mid(pic_type(i),1,InStr(pic_type(i), "]")-1)
	If IsNumeric(page_num)=false Then page_num=0
	url_str=comic_Name
	    For j=1 to (4-Len(page_num))
		url_str = url_str & "0"
	    Next
	    url_str = url_str & page_num & "页"
	'http://images.mh.xunlei.com/origin/1029/11345737d625844567cf30111aca4e04ee0775e93fc5b9.jpg
	'1029/11345737d625844567cf30111aca4e04ee0775e93fc5b9.jpg  
	pic_type(i)=Mid(pic_type(i),InStr(pic_type(i), "'")+1)
	pic_type(i)="http://images.mh.xunlei.com/origin/" & Mid(pic_type(i),1,InStr(pic_type(i), "'")-1)
	return_download_list = return_download_list & "|" & pic_type(i) & "|" & url_str & Mid(pic_type(i),instrrev(pic_type(i),".")) & "|" & url_str & vbcrlf
Next

    return_download_list = return_download_list & "0"
    
'picture
ElseIf InStr(1, html_str, "var images_id_arr = new Array();", 1) > 0 and InStr(1, html_str, "var images_arr = new Array();", 1) > 0 Then
    comic_Name = Mid(html_str, InStr(1, html_str, "<title>", 1) + 7)
    comic_Name = Mid(comic_Name, 1, InStr(1, comic_Name, "</title>", 1) - 1)
    comic_Name = Mid(comic_Name, 1, InStr(1, comic_Name, "动漫频道", 1) - 1) & "第"

    html_str = Mid(html_str, InStr(1, html_str, "var images_id_arr = new Array();", 1))
    html_str = Mid(html_str, InStr(1, html_str, "images_arr[", 1)+11)
    html_str = Mid(html_str, 1, InStr(1, html_str, "var page", 1))
    pic_type = split(html_str,"images_arr[")

For i=0 to UBound(pic_type)
	page_num=Mid(pic_type(i),1,InStr(pic_type(i), "]")-1)
	If IsNumeric(page_num)=false Then page_num=0
	url_str=comic_Name
	    For j=1 to (4-Len(page_num))
		url_str = url_str & "0"
	    Next
	    url_str = url_str & page_num & "页"
	'http://images.anime.xunlei.com/collection/origin/2/124617OTuSi9OpjajnGrfteiLQui8buMEBDZfrt2UXx6Fp.jpg
	'2/124617OTuSi9OpjajnGrfteiLQui8buMEBDZfrt2UXx6Fp.jpg
	pic_type(i)=Mid(pic_type(i),InStr(pic_type(i), "'")+1)
	pic_type(i)="http://images.anime.xunlei.com/collection/origin/" & Mid(pic_type(i),1,InStr(pic_type(i), "'")-1)
	return_download_list = return_download_list & "|" & pic_type(i) & "|" & url_str & Mid(pic_type(i),instrrev(pic_type(i),".")) & "|" & url_str & vbcrlf
Next
    return_download_list = return_download_list & "0"
End If
End Function