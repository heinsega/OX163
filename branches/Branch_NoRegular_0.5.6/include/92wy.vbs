'2010-12-27 http://www.shanhaijing.net/163
Dim page_counter,page,nameid,unpic

Function return_download_url(ByVal url_str)
On Error Resume Next
'http://comic.92wy.com/go/show_551_9890_81.htm
url_str=LCase(url_str)
If InStr(url_str, "http://comic.92wy.com/go/show_") > 0 Then
    url_str = Replace(url_str, "http://comic.92wy.com/go/show_", "")
    url_str = Replace(url_str, ".htm", "")
    Dim str_temp
    str_temp=split(url_str,"_")
    If UBound(str_temp)>1 Then
	If IsNumeric(str_temp(0)) And IsNumeric(str_temp(1)) Then
	page_counter = 0
	unpic=0
	Page = 1
	nameid = str_temp(0) & "_" & str_temp(1) & "_"
	return_download_url = "inet|10,13|http://comic.92wy.com/go/show_" & nameid & "1.htm|http://comic.92wy.com/"
	Else
	return_download_url = ""
	End If
    Else
    return_download_url = ""
    End If
ElseIf InStr(url_str, "http://comic.92wy.com/go/info_") > 0 Then
    url_str = Replace(url_str, "http://comic.92wy.com/go/info_", "")
    url_str = Replace(url_str, ".htm", "")
    If IsNumeric(url_str) Then return_download_url="inet|10,13|http://comic.92wy.com/go/info_" & url_str & ".htm|http://comic.92wy.com/"
End If
End Function
'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim html_temp,url_temp
url_temp=url_str
return_albums_list=""
If InStr(LCase(html_str), "<a href=""show_") > 0 Then
    Dim comic_Name
    comic_Name = Mid(html_str, InStr(LCase(html_str), "<h5>") +4)
    comic_Name = Mid(comic_Name, InStr(LCase(comic_Name), "<") +1)
    comic_Name = Mid(comic_Name, InStr(LCase(comic_Name), ">") +1)
    comic_Name = Mid(comic_Name, 1, InStr(LCase(comic_Name), "<") - 1)
    comic_Name = Trim(replace(comic_Name,"&nbsp;"," "))
    html_str = Mid(html_str, InStr(LCase(html_str), "<a href=""show_")+Len("<a href=""show_"))
    
    album_list = Split(html_str, "<a href=""show_")

    For i = 0 To UBound(album_list)
        return_albums_list = return_albums_list & "0||http://comic.92wy.com/go/show_" & Mid(album_list(i), 1, InStr( LCase(album_list(i)),"""") - 1) & "|"
        album_list(i) = Mid(album_list(i), InStr(album_list(i), ">") + 1)
        album_list(i) = Mid(album_list(i), 1, InStr(LCase(album_list(i)), "</a>") - 1)
        return_albums_list = return_albums_list & Replace(comic_Name & "_" & album_list(i), "|", "_") & "|" & comic_Name & "_" & album_list(i) & vbCrLf
    Next
    
    return_albums_list = return_albums_list & "0"
Else
    return_albums_list = "0"
End If

End Function
'----------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""
If page_counter < 1 Then
url_str = Mid(html_str,1,InStr(html_str, "×îÄ©Ò³"))
page_counter = Mid(url_str, InStrRev(url_str, "_")+1)
page_counter = Mid(page_counter,1, InStr(page_counter, ".")-1)
If Not IsNumeric(page_counter) Then page_counter = 1
page_counter=Int(page_counter)
End If

If page <= page_counter and html_str<>"" Then
If InStr(LCase(html_str), "<img src=""http://comicpic.92wy.com/pics/") > 0 Then
	
    html_str = Mid(html_str, InStr(LCase(html_str), "<acronym title=""") + 16)
    comic_Name = replace(Trim(Mid(html_str, 1, InStr(LCase(html_str), """") - 1))," ","_")
    
    html_str = Mid(html_str, InStr(LCase(html_str), "<img src=""http://comicpic.92wy.com/pics/") + 10)
    html_str = Mid(html_str, 1, InStr(LCase(html_str), """") - 1)
    
    Dim split_str
    For i=1 to (Len(page_counter)-Len(page-unpic))
	split_str = split_str & "0"
    Next
    
    split_str = split_str & (page-unpic) & "Ò³.jpg"
    comic_Name = comic_Name & "_µÚ" & split_str
    return_download_list = "|" & html_str & "|" & comic_Name & "|" & comic_Name & vbCrLf
Else
unpic=unpic+1
End If

    If page < page_counter and html_str<>"" Then
        page = page + 1
        return_download_list = return_download_list & page_counter & "|inet|10,13|http://comic.92wy.com/go/show_" & nameid & page & ".htm"
        Else
        return_download_list = return_download_list & "0"
    End If

Else
    return_download_list = "0"
End If
End Function