'2011-5-22 http://www.shanhaijing.net/163
Dim root_url
Function return_download_url(ByVal url_str)
On Error Resume Next
return_download_url = ""
url_str=replace(url_str,".sky-fire.com",".sfacg.com")
root_url="sfacg"'sfacg comic.sky-fire.com

If InStr(1, url_str, "http://comic." & root_url & ".com/HTML/", 1) > 0 Then
    'http://comic." & root_url & ".com/HTML/Naruto/
    'http://comic." & root_url & ".com/HTML/Naruto
    return_download_url = "inet|10,13|" & url_str
Else
    return_download_url = Mid(url_str,1,InStr(LCase(url_str),".com/")+4)
    url_str = FixAllComicUrl(url_str)
    return_download_url = "inet|10,13|" & return_download_url & "Utility/" & url_str & ".js|" & return_download_url
End If
End Function
'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next

return_albums_list = ""

If InStr(1, html_str, "<ul class=""serialise_list", 1) > 0 Then
    Dim comic_Name,album_list
    comic_Name = Mid(html_str, InStr(LCase(html_str), "<title>") + 7)
    comic_Name = Mid(comic_Name, 1, InStr(comic_Name, "</title>") - 1)
    comic_Name = Mid(comic_Name, 1, InStr(comic_Name, ",") - 1)
    
    html_str = Mid(html_str, InStr(LCase(html_str), "<ul class=""serialise_list"))
    html_str = Mid(html_str, InStr(LCase(html_str), "<li><a href=""")+13)
    album_list = Split(html_str, "<li><a href=""")

    For i = 0 To UBound(album_list)
        'url
        html_str=Mid(album_list(i),1,InStr(album_list(i), Chr(34))-1) & "|"
        'name
        album_list(i) = Mid(album_list(i), InStr(album_list(i), ">")+1)
        album_list(i) = comic_Name & "_" & Mid(album_list(i),1, InStr(LCase(album_list(i)), "</a>")-1)
        
        return_albums_list = return_albums_list & "0||" & html_str & album_list(i) & "|" & album_list(i) & vbcrlf
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
Dim comic_name
comic_name=""
If InStr(LCase(html_str), "var comicname = """)>0 Then
url_str=FixAllComicUrl(url_str)
comic_name=Mid(html_str,InStr(LCase(html_str), "var comicname = """)+17)
comic_name=Mid(comic_name,1,InStr(comic_name, Chr(34))-1) & "_vol" & url_str & "_"
End If
If InStr(LCase(html_str), ";picay[0] = """) > 0 Then
Dim split_str,pic_name
	html_str=Mid(html_str,InStr(LCase(html_str), ";picay[0] = """)+7)
	html_str=replace(html_str,""";picAy[",""";picay[")
	split_str=split(html_str,""";picay[")
	
	For i=0 to UBound(split_str)
	split_str(i)=Mid(split_str(i),InStr(split_str(i),Chr(34))+1)
	split_str(i)=Mid(split_str(i),1,InStr(split_str(i),Chr(34))-1)
	pic_name=Mid(split_str(i),instrrev(split_str(i),"/")+1)
	return_download_list = return_download_list & "|" & split_str(i) & "|" & comic_name & pic_name & "|" & vbcrlf
	Next
return_download_list = return_download_list & "0"
Else
    return_download_list = "0"
End If
End Function
'-------------------------------------------------------------------
Function FixAllComicUrl(url_str)
    'http://pic." & root_url & ".com/AllComic/Browser.html?c=4&v=379&p=1
    'http://hotpic." & root_url & ".com/AllComic/Browser.html?c=4&v=379&p=1
    'http://pic." & root_url & ".com/Utility/4/379.js
    
    'http://coldpic." & root_url & ".com/AllComic/Browser.html?c=454&v=080&p=1
    'http://coldpic." & root_url & ".com/Utility/454/080.js    
    'http://coldpic." & root_url & ".com/AllComic/Browser.html?c=456&v=tbp1&t=TBP&p=1    
    'http://coldpic." & root_url & ".com/Utility/456/TBP/tbp1.js
    
    'http://pic3." & root_url & ".com/Utility/331/080.js
    'http://pic3." & root_url & ".com/AllComic/331/081/    
On Error Resume Next
    Dim temp_str,temp_url
    url_str = Mid(url_str, InStr(LCase(url_str), "." & root_url & ".com/allcomic/")+Len("." & root_url & ".com/allcomic/"))
    If InStr(LCase(url_str),"browser.html")>0 Then
	    url_str = Mid(url_str, InStr(LCase(url_str), "?c=")+3)
	    temp_str=Mid(url_str,1, InStr(LCase(url_str), "&v=")-1) & "/"
	    url_str = Mid(url_str, InStr(1, url_str, "&v=", 1)+3)
	    If InStr(LCase(url_str),"&t=")>1 Then
	    	temp_str=temp_str & Mid(url_str,InStr(LCase(url_str), "&t=")+3)
	    	temp_str=Mid(temp_str,1,InStr(temp_str, "&")-1) & "/"
	    End If
	    url_str = temp_str & Mid(url_str,1, InStr(1, url_str, "&", 1)-1)
    Else
	    temp_url=split(url_str,"/")
	    temp_str=0
	    If InStr(temp_url(UBound(temp_url)),"?")>0 or temp_url(UBound(temp_url))="" Then temp_url(UBound(temp_url))="":temp_str=1
	    If temp_str=0 Then
	    	temp_str=join(temp_url,"/")
	    	url_str=temp_str
	    Else
	    	temp_str=join(temp_url,"/")
	    	url_str=Mid(temp_str,1,Len(temp_str)-1)    
	    End If
    End If
FixAllComicUrl=url_str
End Function