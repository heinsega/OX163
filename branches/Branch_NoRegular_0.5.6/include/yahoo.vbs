'2009-4-19 http://www.shanhaijing.net/163
Dim yahoo_ID,photo_split,photo_ID,select_step
Function return_download_url(ByVal url_str)
On Error Resume Next
return_download_url = ""
'http://photos.i.cn.yahoo.com/05100256868/#p4
'http://photos.i.cn.yahoo.com/05100256868/ad30/#doc-body
'http://photos.i.cn.yahoo.com/05100256868/ad30/
'http://photos.i.cn.yahoo.com/05100256868/ad30/5a93.jpg/#doc-body
'http://photos.i.cn.yahoo.com/05100256868/ad30/5a93.jpg/
'http://photos.i.cn.yahoo.com/down-acDCnvAobr8N4GxqchTh3H_HkY8UiNvVMVp659Im?cq=1&aid=ad30&pid=5a93.jpg
Dim url_str_split
If InStr(LCase(url_str), "http://photos.i.cn.yahoo.com/") > 0 Then
    url_str = Mid(url_str,InStr(LCase(url_str), "http://photos.i.cn.yahoo.com/")+29)
    url_str_split=split(url_str,"/")
    If UBound(url_str_split)=1 Then
    	yahoo_ID=url_str_split(0)
    	return_download_url = "inet|10,13,9,34|http://photos.i.cn.yahoo.com/" & url_str_split(0) & "/"
    ElseIf UBound(url_str)>1 Then
    	yahoo_ID=url_str_split(0)
    	select_step=1
    	return_download_url = "inet|10,13,9,34|http://photos.i.cn.yahoo.com/" & url_str_split(0) & "/" & url_str_split(1) & "/"
    End If
End If
End Function
'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list = ""
If InStr(LCase(html_str), "<div class=album_name>") > 0 Then
	html_str = Mid(html_str, InStr(LCase(html_str), "<div class=album_name>")+22)
	Dim str_split
	str_split=split(html_str,"<div class=album_name>")
	For i=0 to UBound(str_split)
		str_split(i)=Mid(str_split(i),InStr(str_split(i),"<h4"))
		str_split(i)=Mid(str_split(i),InStr(str_split(i),">")+1)
		'name
		html_str=rename_utf8(Trim(Mid(str_split(i),1,InStr(str_split(i),"</h4>")-1)))
		If html_str="" Then html_str=yahoo_ID & "_No_Name_Album"
		
		str_split(i)=Mid(str_split(i),InStr(str_split(i),"<a href=")+8)
		'url
		url_str=Mid(str_split(i),1,InStr(str_split(i),">")-1)
		
		'pic number
		str_split(i)=Mid(str_split(i),InStr(str_split(i),"/> 共")+Len("/> 共"))
		str_split(i)=Mid(str_split(i),1,InStr(str_split(i),"张")-1)
		If IsNumeric(str_split(i))=false Then str_split(i)=""
		
	        return_albums_list = return_albums_list & "0|" & str_split(i) & "|" & url_str & "|" & html_str & vbcrlf
	Next
	return_albums_list = return_albums_list & "0"
End If
End Function
'----------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list=""

Select Case select_step
Case 1

	'<a href="/slideshow-acDCnvAobr8N4GxqchTh3H_HkY8UiNvVMVp659Im?cq=1&aid=ad30"><strong>幻灯播放</strong>
	If InStr(html_str,"<a href=/slideshow-") Then
		html_str=Mid(html_str,InStr(html_str,"<a href=/slideshow-")+19)
		html_str=Mid(html_str,1,InStr(html_str,">")-1)
		select_step=2
		yahoo_ID=html_str & "&pid="
		return_download_list = "2|inet|10,13,9,34|http://photos.i.cn.yahoo.com/data-" & html_str
	End If

Case 2

	'http://photos.i.cn.yahoo.com/data-acDCnvAobr8N4GxqchTh3H_HkY8UiNvVMVp659Im?cq=1&aid=ad30
	If InStr(html_str,yahoo_ID)>0 Then
		Dim pic_desc
		html_str=Mid(html_str,1,instrrev(html_str,"||")-1)
		html_str=Mid(html_str,InStr(html_str,"||")+2)
		photo_split=split(html_str,"||")
		For i=0 to UBound(photo_split)
			'middle-url
			url_str=Mid(photo_split(i),1,InStr(photo_split(i),"|")-1)
			
			photo_split(i)=Mid(photo_split(i),InStr(photo_split(i),"|")+1)
			photo_split(i)=Mid(photo_split(i),InStr(photo_split(i),"|")+1)
			
			'6801款|围巾大碎花，有现货|2368.jpg
			'name
			html_str=Mid(photo_split(i),1,InStr(photo_split(i),"|")-1)
			
			photo_split(i)=Mid(photo_split(i),InStr(photo_split(i),"|")+1)
			'desc
			pic_desc=Trim(replace(Mid(photo_split(i),1,InStr(photo_split(i),"|")-1),vbcrlf,""))
			
			photo_split(i)=Trim(Mid(photo_split(i),InStr(photo_split(i),"|")+1))
			
			'name
			html_str=rename_utf8(html_str & "_" & photo_split(i))
			
			'info
			photo_split(i)= "|" & url_str & "|" & html_str & "|" & pic_desc & vbcrlf
		Next
		return_download_list=join(photo_split,"") & "0"
	End If
End Select
End Function
'------------------------------------------------------
Function rename_utf8(byval utf8_Str)
If Len(utf8_Str)=0 Then Exit Function
For i=1 to Len(utf8_Str)
	If  Asc(Mid(utf8_Str,i,1))=63 Then utf8_Str=replace(utf8_Str,Mid(utf8_Str,i,1),"_")
Next
rename_utf8=utf8_Str
End Function