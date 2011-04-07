'2009-10-22 http://www.shanhaijing.net/163
Dim pic_counter,yxnpc_type,nameid,page

Function return_download_url(byVal url_str)
On Error Resume Next
return_download_url = ""
pic_counter=0
If InStr(1,url_str,"http://www.yxnpc.com/manhua/",1)>0 Then
	return_download_url="inet|10,13|" & url_str
ElseIf InStr(1,url_str,"http://www.yxnpc.com/",1)>0 Then
	'http://www.yxnpc.com/2008-09-22/0082/09697365.html
	'http://www.yxnpc.com/2008-09-22/0G5D/22839063_02.html
	url_str=replace(url_str,"#","")
	If InStr(url_str,"_")>0 Then
		url_str=Mid(url_str,1,instrrev(url_str,"_")-1) & ".html"
	End If
	
	nameid=Mid(url_str,1,instrrev(url_str, ".html")-1)
	page=1
	return_download_url="inet|10,13|" & url_str
	'yxnpc pic图片
End If
End Function
'----------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim split_str
return_albums_list = ""
	If InStr(1, html_str, "<div class=""mhlist"">", 1) > 0 Then		
		html_str=Mid(html_str,InStr(1, html_str, "<div class=""mhlist"">", 1)+20)
		html_str=Mid(html_str,InStr(1, html_str, "<li><a href=""", 1)+13)
		split_str = Split(html_str, "<li><a href=""",-1,1)
		For i = 0 To UBound(split_str)
			'url
			url_str="http://www.yxnpc.com" & Mid(split_str(i),1,InStr(split_str(i),Chr(34))-1)
			
			'name
			split_str(i)=Mid(split_str(i),InStr(split_str(i),">")+1)
			split_str(i)=Trim(Mid(split_str(i),1,InStr(split_str(i),"<")-1))
			
			return_albums_list = return_albums_list & "0||" & url_str & "|" & split_str(i) & vbcrlf
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

Dim split_str,pic_url

	html_str=replace(html_str,"<img alt=""","<img alt=""")
	html_str=replace(html_str,"<a  href=""","<a href=""")
	If InStr(1, html_str, "http://www.yxnpc.com/script/showpic.php?picfile=", 1) > 0 Then
	'有大图的照片
		url_str=Mid(html_str,InStr(LCase(html_str), "<title>")+7)
		url_str=Mid(url_str,1,InStr(LCase(url_str), "</title>")-1)
		'图片title
		url_str=Trim(Mid(url_str,1,instrrev(LCase(url_str), "-")-1))
		
		html_str=Mid(html_str,InStr(1, html_str, "http://www.yxnpc.com/script/showpic.php?picfile=", 1)+Len("http://www.yxnpc.com/script/showpic.php?picfile="))
		split_str = Split(html_str, "http://www.yxnpc.com/script/showpic.php?picfile=",-1,1)
		For i = 0 To UBound(split_str)
			'url http://www.yxnpc.com/uldf/2009/0608/BJZ00686/meiomfsdlk2/042.jpg
			pic_url="http://www.yxnpc.com" & Mid(split_str(i),1,InStr(split_str(i),Chr(34))-1)
			
			'name
			split_str(i)=Trim(Mid(split_str(i),InStr(1,split_str(i),"alt=""",1)+5))
			If split_str(i)<>"" Then
				split_str(i)=Mid(split_str(i),1,InStr(split_str(i),Chr(34))-1) & "_" & Mid(pic_url,instrrev(pic_url,"/")+1)
			Else
				split_str(i)=url_str & "_" & Mid(pic_url,instrrev(pic_url,"/")+1)
			End If
			
			return_download_list = return_download_list & "|" & pic_url & "|" & split_str(i) & "|" & split_str(i) & vbCrLf
		Next
		
		page=page+1
		html_str=nameid & "_" & page & ".html"
		return_download_list = return_download_list & page & "|inet|10,13|" & html_str

		
	ElseIf InStr(1, html_str, ">将此页漫画加入收藏夹</a>", 1) > 0 Then
	'漫画
		url_str=Mid(html_str,InStr(LCase(html_str),"<title>")+7)
		url_str=Trim(Mid(url_str,1,InStr(url_str,"</title>")-1))
		url_str=Trim(Mid(url_str,1,InStr(url_str,"-")-1))
		url_str=Trim(Mid(url_str,1,InStr(url_str,"(")-1))
		url_str=Trim(replace(replace(url_str,Chr(13)),Chr(10)))
		If Len(url_str)>20 Then url_str=Left(url_str,20)
		'url
		html_str=Mid(html_str,InStr(LCase(html_str),"class=""mhpic"""))
		html_str=Mid(html_str,InStr(LCase(html_str),"src=""")+5)
		pic_url="http://www.yxnpc.com" & Mid(html_str,1,InStr(html_str,Chr(34))-1)
		
		'name
		html_str=Mid(html_str,InStr(1,html_str,"yxnpc在线漫画：第",1)+Len("yxnpc在线漫画：第"))
		html_str=Mid(html_str,1,InStr(html_str,"话")-1)
		If IsNumeric(html_str) Then			
			html_str=url_str & "_" & html_str & "_" & Mid(pic_url,instrrev(pic_url,"/")+1)
		Else
			html_str=url_str & "_" & Mid(pic_url,instrrev(pic_url,"/")+1)
		End If
		
		return_download_list = return_download_list & "|" & pic_url & "|" & html_str & "|" & html_str & vbCrLf

		page=page+1
		pic_url=nameid & "_" & page & ".html"	
		return_download_list = return_download_list & page & "|inet|10,13|" & pic_url


	End If

End Function

