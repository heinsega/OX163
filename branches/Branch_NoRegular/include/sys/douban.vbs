'2011-10-27 163.shanhaijing.net
Dim retry_time,page,album_ID,html_type

Function return_download_url(ByVal url_str)
On Error Resume Next
return_download_url=""
retry_time=0
If instr(lcase(url_str),"http://www.douban.com/photos/")=1 Then
'http://www.douban.com/photos/album/11028188/?start=18
	If instr(lcase(url_str),"http://www.douban.com/photos/album/")=1 Then
		url_str=mid(url_str,instr(lcase(url_str),"http://www.douban.com/photos/album/")+len("http://www.douban.com/photos/album/"))
		If instr(url_str,"/")>1 Then url_str=mid(url_str,1,instr(url_str,"/")-1)
		If instr(url_str,"?")>1 Then url_str=mid(url_str,1,instr(url_str,"?")-1)
		If instr(url_str,"#")>1 Then url_str=mid(url_str,1,instr(url_str,"#")-1)
		If IsNumeric(url_str) and len(url_str)>0 Then
			album_ID="http://www.douban.com/photos/album/" & url_str & "/"
			html_type="album"
			page=1
			return_download_url = "inet|10,13|" & album_ID
		End If
	ElseIf instr(lcase(url_str),"http://www.douban.com/photos/photo/")=1 Then
'http://www.douban.com/photos/photo/125705371/
'http://www.douban.com/photos/photo/125716028/#next_photo
		album_ID=url_str
		html_type="photo"
		page=0
		return_download_url = "inet|10,13|" & album_ID
	End If
	
ElseIf instr(lcase(url_str),"http://www.douban.com/people/")=1 Then
'http://www.douban.com/people/royzhong/
'http://www.douban.com/people/royzhong/notes
'http://www.douban.com/people/royzhong/photos
'http://www.douban.com/people/royzhong/photos?start=32
	url_str=mid(url_str,instr(lcase(url_str),"http://www.douban.com/people/")+len("http://www.douban.com/people/"))
	If instr(url_str,"/")>1 Then url_str=mid(url_str,1,instr(url_str,"/")-1)
	If instr(url_str,"?")>1 Then url_str=mid(url_str,1,instr(url_str,"?")-1)
	If instr(url_str,"#")>1 Then url_str=mid(url_str,1,instr(url_str,"#")-1)
	album_ID="http://www.douban.com/people/" & url_str & "/photos"
	page=1
	html_type="people"
	return_download_url = "inet|10,13|" & album_ID
End If
End Function
'--------------------------------------------------------
Function Get_album_ID(ByVal url_str)
On Error Resume Next
Get_album_ID=""
If instr(lcase(url_str),"http://www.douban.com/photos/album/")=1 Then
	url_str=mid(url_str,instr(lcase(url_str),"http://www.douban.com/photos/album/")+len("http://www.douban.com/photos/album/"))
	If instr(url_str,"/")>1 Then url_str=mid(url_str,1,instr(url_str,"/")-1)
	If instr(url_str,"?")>1 Then url_str=mid(url_str,1,instr(url_str,"?")-1)
	If instr(url_str,"#")>1 Then url_str=mid(url_str,1,instr(url_str,"#")-1)
	If IsNumeric(url_str) and len(url_str)>0 Then	Get_album_ID=url_str
End If
End Function
'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list=""
Dim key_word,split_str,album_title(3)
If page>0 and html_type="people" and InStr(LCase(html_str),"<a class=""album_photo"" href=""")>0 Then
	retry_time=0
	url_str=html_str
	key_word="<div class=""albumlst"">"
	html_str=mid(html_str,InStr(LCase(html_str),LCase(key_word))+len(key_word))
	split_str=split(html_str,key_word)
	
	For i=0 to ubound(split_str)
		key_word=""
		pic_title=""
		'album ID
		key_word=Mid(split_str(i),InStr(split_str(i),"<a class=""album_photo"" href=""")+len("<a class=""album_photo"" href="""))
		key_word=Mid(key_word,1,InStr(key_word,"""")-1)
		key_word=Get_album_ID(key_word)
		
		If IsNumeric(key_word) and len(key_word)>0 Then
			'title
			album_title(0)=Mid(split_str(i),InStr(split_str(i),"<div class=""pl2"">")+25)
			album_title(0)=Mid(album_title(0),InStr(album_title(0),">")+1)
			album_title(0)=Mid(album_title(0),1,InStr(album_title(0),"</a>")-1)
			album_title(0)=Trim(replace(album_title(0),"|","£¸"))
			album_title(0)="AID" & key_word  & "-" & album_title(0)
			
			'desc
			album_title(1)=Mid(split_str(i),InStr(split_str(i),"<div class=""albumlst_descri"">")+len("<div class=""albumlst_descri"">"))
			album_title(1)=Mid(album_title(1),1,InStr(album_title(1),"</div>")-1)
			
			'pic number
			album_title(2)=Mid(split_str(i),InStr(split_str(i),"<span class=""pl"">")+len("<span class=""pl"">"))
			album_title(2)=Mid(album_title(2),1,InStr(album_title(2),"’≈’’∆¨")-1)
			album_title(2)=Trim(album_title(2))
			If IsNumeric(album_title(2))=false Then album_title(2)=""
			
			return_albums_list=return_albums_list & "0|" & album_title(2) & "|http://www.douban.com/photos/album/" & key_word & "/|" & album_title(0) & "|" & album_title(1) & vbCrLf
		End If
	Next
	
	key_word="<link rel=""next"" href="""
	If InStr(LCase(url_str),LCase(key_word))>0 Then
		url_str=mid(url_str,InStr(LCase(url_str),LCase(key_word))+len(key_word))		
		url_str=mid(url_str,1,InStr(LCase(url_str),"""")-1)
		page=page+1
		album_ID=url_str
		return_albums_list=return_albums_list & "1|inet|10,13|" & url_str
	End If
	
'ElseIf Then
	'retry_time=retry_time+1
	'return_download_list="1|inet|10,13|" & album_ID
End If
End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
'http://otho.douban.com/view/photo/thumb/xvvoA0chC9-RG7NT0eh9pw/x1159718333.jpg
'http://otho.douban.com/view/photo/photo/afk-0HXr3gp8fXp3ehM_8g/x1159718333.jpg
'http://img3.douban.com/view/photo/thumb/public/p1159718333.jpg
'http://img3.douban.com/view/photo/photo/public/p1159718333.jpg
return_download_list = ""
Dim key_word,split_str,pic_title
If page>0 and html_type="album" and InStr(LCase(html_str),"<div class=""photo_wrap"">")>0 Then
	retry_time=0
	url_str=html_str
	key_word="<div class=""photo_wrap"">"
	html_str=mid(html_str,InStr(LCase(html_str),LCase(key_word))+len(key_word))
	split_str=split(html_str,key_word)
	For i=0 to ubound(split_str)
		key_word=""
		pic_title=""
		'title
		pic_title=Mid(split_str(i),InStr(split_str(i),"title=""")+len("title="""))
		pic_title=Mid(pic_title,1,InStr(pic_title,"""")-1)
		'url
		key_word=Mid(split_str(i),InStr(split_str(i),"<img src=""")+len("<img src="""))
		key_word=Mid(key_word,1,InStr(key_word,"""")-1)
		key_word=replace(key_word,"/thumb/","/photo/")
		if left(key_word,len("http://img1."))<>"http://img1." then key_word="http://img1." & mid(key_word,instr(key_word,".")+1)
		If len(key_word)>0 Then
			'file name
			split_str(i)=mid(key_word,instrrev(key_word,"/")+1)
			return_download_list=return_download_list & "|" & key_word & "|" & split_str(i) & "|" & pic_title & vbcrlf
		End If
	Next
	
	key_word="<link rel=""next"" href="""
	If InStr(LCase(url_str),LCase(key_word))>0 Then
		url_str=mid(url_str,InStr(LCase(url_str),LCase(key_word))+len(key_word))		
		url_str=mid(url_str,1,InStr(LCase(url_str),"""")-1)
		page=page+1
		album_ID=url_str
		return_download_list=return_download_list & "1|inet|10,13|" & url_str
	End If
	
ElseIf page=0 and html_type="photo" and InStr(LCase(html_str),"<span class='rr'>")>0 Then
	retry_time=0
	page=1
	html_type="album"
	html_str=Mid(html_str,InStr(LCase(html_str),"<span class='rr'>"))
	html_str=Mid(html_str,InStr(LCase(html_str),"<a href=""")+9)
	html_str=Mid(html_str,1,InStr(html_str,"""")-1)
	album_ID=html_str
	return_download_list="1|inet|10,13|" & album_ID	
	
ElseIf retry_time<4 Then
	retry_time=retry_time+1
	return_download_list="1|inet|10,13|" & album_ID
End If
End Function