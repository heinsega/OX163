'2011-3-14 http://www.shanhaijing.net/163
Dim UserId,AlbumID

Function return_download_url(byVal url_str)
On Error Resume Next
return_download_url = ""
'albums
'http://photo.tom.com/user/545362173.html
'http://photo.tom.com/userindex.php
'photos
'http://photo.tom.com/album/545362173/6067.html
'http://photo.tom.com/picture/545362173/11593.html
UserId=""

'albums
If InStr(LCase(url_str),"http://photo.tom.com/user/")=1 Then
	url_str=Mid(url_str,InStr(LCase(url_str),"http://photo.tom.com/user/")+Len("http://photo.tom.com/user/"))
	url_str=Mid(url_str,1,InStr(url_str,".")-1)
	If IsNumeric(url_str) Then
		UserId=url_str
		'http://photo.tom.com/ajax/userindexalbums.php?userid=20027993&order=1&time=0.8056439607329862
		return_download_url="inet|10,13|http://photo.tom.com/ajax/userindexalbums.php?userid=" & UserId & "&order=1&time=" & Timer() & "|http://photo.tom.com/"
	End If
ElseIf InStr(LCase(url_str),"http://photo.tom.com/userindex.php")=1 Then
	UserId="ower"
	return_download_url="inet|10,13|" & url_str & "|http://photo.tom.com/"
Else
	'photo
	If InStr(LCase(url_str),"http://photo.tom.com/album/")=1 Then
		AlbumID="album"
		UserId=Mid(url_str,InStr(LCase(url_str),"http://photo.tom.com/album/")+Len("http://photo.tom.com/album/"))
		UserId=Mid(UserId,1,InStr(UserId,"/")-1)
		If IsNumeric(UserId) Then
			return_download_url="inet|10,13|" & url_str & "|http://photo.tom.com/"
		End If
	ElseIf InStr(LCase(url_str),"http://photo.tom.com/picture/")=1 Then
		AlbumID="picture"
		return_download_url="inet|10,13|" & url_str & "|http://photo.tom.com/"
	End If
End If
End Function
'----------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim split_str,split_i,photo_count
return_albums_list = ""
If UserId="ower" and InStr(html_str,"http://photo.tom.com/user/")>1 Then
	html_str=Mid(html_str,InStr(html_str,"http://photo.tom.com/user/"))
	html_str=Mid(html_str,1,InStr(html_str,"""")-1)
	html_str=Mid(html_str,1,InStr(html_str,"'")-1)
	
	url_str=Mid(html_str,InStr(LCase(html_str),"http://photo.tom.com/user/")+Len("http://photo.tom.com/user/"))
	url_str=Mid(url_str,1,InStr(url_str,".")-1)
	If IsNumeric(url_str) Then
		UserId=url_str
		'http://photo.tom.com/ajax/userindexalbums.php?userid=20027993&order=1&time=0.8056439607329862
		return_albums_list="1|inet|10,13|http://photo.tom.com/ajax/userindexalbums.php?userid=" & UserId & "&order=1&time=" & Timer() & "|http://photo.tom.com/"
	End If
	
ElseIf IsNumeric(UserId) and InStr(html_str,"<li id='album_")>1 Then
	html_str=Mid(html_str,InStr(html_str,"<li id='album_")+Len("<li id='album_"))
	split_str=split(html_str,"<li id='album_")
	
	For split_i = 0 To UBound(split_str)
		html_str=""
		url_str=""
		photo_count=""
		
		'album_id
		url_str = Mid(split_str(split_i),1,InStr(split_str(split_i), "'")-1)
		
		split_str(split_i) = Mid(split_str(split_i),InStr(split_str(split_i), "<a href="))
		split_str(split_i) = Mid(split_str(split_i),InStr(split_str(split_i), ">")+1)
		
		'name
		html_str= Mid(split_str(split_i),1,InStr(split_str(split_i), "</a>")-1)
		If Len(html_str)>50 Then html_str=Left(html_str,49) & "~"
		If html_str="" Then html_str="UserId-" & UserId & "(album_" & url_str & ")"
		html_str=replace(html_str,"|","_")
		url_str=""
		
		split_str(split_i) = Mid(split_str(split_i),InStr(split_str(split_i), "</a>")+1)
		split_str(split_i) = Mid(split_str(split_i),InStr(split_str(split_i), "<span>(")+Len("<span>("))
		
		'photo_count
		photo_count = Mid(split_str(split_i),1,InStr(split_str(split_i), ")")-1)
		If IsNumeric(photo_count)=False Then photo_count=""
		
		'url
		url_str=Mid(split_str(split_i),InStr(split_str(split_i), "<a href=")+Len("<a href="))
		url_str=Mid(url_str,1,InStr(url_str, " ")-1)
		'/album/545362173/6069.html
		'http://photo.tom.com/album/545362173/6067.html
		url_str="http://photo.tom.com" & url_str
		
		'description
		split_str(split_i) = Mid(split_str(split_i),InStr(split_str(split_i), "<p class='txt'>")+Len("<p class='txt'>"))
		split_str(split_i) = Mid(split_str(split_i),1,InStr(split_str(split_i), "</p>")-1)	
		
		return_albums_list = return_albums_list & "0|" & photo_count & "|" & url_str & "|" & html_str & "|" & description & vbCrLf
	Next
	return_albums_list=return_albums_list & "0"
Else
return_albums_list = ""
End If


End Function
'----------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""

If AlbumID="album" and InStr(html_str,"<li id=""li_")>1 and InStr(html_str,"http://photo.tom.com/picture/" & UserId & "/")>1 Then
	html_str=Mid(html_str,InStr(html_str,"<li id=""li_"))
	html_str=Mid(html_str,InStr(html_str,"http://photo.tom.com/picture/" & UserId & "/"))
	html_str=Mid(html_str,1,InStr(html_str,".html")-1)
	url_str=Mid(html_str,InStr(html_str,"http://photo.tom.com/picture/" & UserId & "/")+Len("http://photo.tom.com/picture/" & UserId & "/"))
	html_str=html_str & ".html"
	If IsNumeric(url_str) Then
		AlbumID="picture"
		return_download_list = "1|inet|10,13|" & html_str
	End If
ElseIf AlbumID="picture" and InStr(LCase(html_str),"var imagelist = {")>1 Then
	
	Dim split_str,split_i,file_type
	
	html_str=Mid(html_str,InStr(LCase(html_str),"var imagelist = {"))
	html_str=Mid(html_str,1,InStr(html_str,"}]},")-1)
	
	html_str=Mid(html_str,InStr(html_str,":[{")+3)
	split_str=split(html_str,"},{")
	
	For split_i = 0 To UBound(split_str)
		html_str=""
		url_str=""
		split_str(split_i)=Mid(split_str(split_i),InStr(split_str(split_i),"""title"":""")+Len("""title"":"""))
		
		'name
		html_str=Mid(split_str(split_i),1,InStr(split_str(split_i),"""")-1)
		html_str=Trim(html_str)
		If Len(html_str)>80 Then html_str=Left(html_str,79) & "~"
		If html_str="" Then html_str="No_Title_Photo"
		
		split_str(split_i)=Mid(split_str(split_i),InStr(split_str(split_i),"""description"":""")+Len("""description"":"""))
		'description
		url_str=Mid(split_str(split_i),1,InStr(split_str(split_i),"""")-1)
		
		'url
		split_str(split_i)=Mid(split_str(split_i),InStr(split_str(split_i),"""original"":""")+Len("""original"":"""))
		split_str(split_i)=Mid(split_str(split_i),1,InStr(split_str(split_i),"""")-1)
		
		'file_type
		file_type=Mid(split_str(split_i),instrrev(split_str(split_i),"."))
		If Right(html_str,Len(file_type))<>file_type Then html_str = html_str & file_type
		
		'rnd time
		split_str(split_i)=split_str(split_i) & "?ntime=" & Timer()

		split_str(split_i) = "|" & split_str(split_i) & "|" & html_str & "|" & url_str & vbCrLf
	Next
	
	return_download_list=join(split_str,"") & "0"	
End If

End Function