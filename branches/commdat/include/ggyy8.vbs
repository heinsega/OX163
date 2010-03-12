'2009-6-2 163.shanhaijing.net

Dim checked_url

Function return_download_url(ByVal url_str)
Dim split_temp
If InStr(LCase(url_str),".html")>1 Then
	url_str=Mid(url_str,1,InStr(url_str,".html")+5) & ".js"
End If
return_download_url="inet|10,13,34|" & url_str & "|http://www.ggyy8.cc/"

checked_url=url_str
End Function
'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list = ""
If InStr(1, html_str, "<td  Height=20><a  target=_blank  href=", 1) > 0 Then

url_str="http://www.ggyy8.cc/"
	
Dim album_list,url_temp
html_str = Mid(html_str, InStr(1, html_str, "<td  Height=20><a  target=_blank  href=", 1) + 39)
    
album_list = Split(html_str, "<td  Height=20><a  target=_blank  href=", -1, 1)

For i = 0 To UBound(album_list)
        If InStr(1, album_list(i), ".html", 1) > 0 Then
        'url
        url_temp = Mid(album_list(i),1,InStr(1, album_list(i), ".html", 1)+4)
        url_temp = Mid(url_temp,InStr(1, url_temp, "/html/", 1))
        url_temp = url_str & url_temp
        album_list(i)=Mid(album_list(i),InStr(album_list(i), ">")+1)
        album_list(i)=replace(Mid(album_list(i),1,InStr(1,album_list(i), "</a>",1)-1),"&nbsp;"," ")
        
        return_albums_list = return_albums_list & "0||" & url_temp & "|" & album_list(i) & "|" & album_list(i) & vbcrlf
	End If
Next
return_albums_list = return_albums_list & "0"

Else	
return_albums_list = return_albums_list & "0"
End If
End Function
'----------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list=""

If InStr(LCase(html_str), "<img src=")>0 Then

html_str=Mid(html_str,InStr(LCase(html_str), "<img src="))

Dim comic_name,split_str
split_str=Split(html_str, "@#@")

For i = 0 To UBound(split_str)
	split_str(i)=Mid(split_str(i),InStr(LCase(split_str(i)), "<img src=")+9)
	split_str(i)=Mid(split_str(i),1,InStr(split_str(i), ">")-1)
	'http://images.ggyy8.cn/comic/C/cxywshz/vol_002/002003.png
	comic_name=replace(split_str(i),"http://","")
	comic_name=Mid(comic_name,InStr(comic_name, "/")+1)
	comic_name=Mid(comic_name,InStr(comic_name, "/")+1)
	comic_name=Mid(comic_name,InStr(comic_name, "/")+1)
	comic_name=replace(comic_name,"/","_")
	return_download_list=return_download_list & "|" & split_str(i) & "|" & comic_name & "|" & vbcrlf
next

return_download_list = return_download_list & "0"


Else
    return_download_list = "0"
End If
End Function
