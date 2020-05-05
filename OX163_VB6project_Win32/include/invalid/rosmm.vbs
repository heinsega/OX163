'2012-11-2 163.shanhaijing.net
Dim page_num, page_count, album_url
Function return_download_url(ByVal url_str)
On Error Resume Next
'http://www.rosmm.com/
page_num = 0
page_count = 1
Dim str_temp,split_str
split_str=split(url_str,"/")
If ubound(split_str)>5 Then
	str_temp = Mid(url_str, InStrRev(url_str, "/") + 1)
	If InStr(str_temp, "_") > 0 Then
	    'http://www.rosmm.com/rosimm/2012/09/19/324_6.htm
	    url_str = Mid(url_str, 1, InStrRev(url_str, "_") - 1)
	    url_str = url_str & ".htm"
	End If
	album_url = url_str
	album_url = Mid(album_url, 1, InStrRev(album_url, ".htm") - 1)
Else
	If LCase(url_str)="http://www.rosmm.com/" Then url_str="http://www.rosmm.com/rosimm/"
	album_url = Mid(url_str,1,InStrRev(url_str, "/"))
End If
return_download_url = "inet|10,13|" & url_str
End Function
'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim key_word,split_str,str_temp
return_albums_list = ""
str_temp=html_str

key_word=""">上一页</a></li>"
If page_num=0 and instr(LCase(html_str),key_word)>0 Then
	page_num=1
	key_word="<div class=""page page_l"">"
	html_str=mid(html_str,instr(LCase(html_str),key_word)+len(key_word))
	html_str=mid(html_str,instr(LCase(html_str),"<li><a href=""")+len("<li><a href="""))
	html_str=album_url & mid(html_str,1,instr(LCase(html_str),"""")-1)
	return_albums_list="1|inet|10,13|" & html_str
	Exit Function
ElseIf page_num=0 Then
	page_num=1
End If

html_str=str_temp
key_word="<ul class=""i_pic"">"
If instr(LCase(html_str),key_word)>0 Then
	html_str=mid(html_str,instr(LCase(html_str),key_word))
	key_word="<div class=""page page_l"">"
	html_str=mid(html_str,1,instr(LCase(html_str),key_word)-1)
	
	key_word="<li><a href="""
	html_str=mid(html_str,instr(LCase(html_str),key_word)+len(key_word))	
	split_str=split(html_str,key_word)
	For i=0 to ubound(split_str)
		html_str=""
		url_str=""
		'url
		url_str="http://www.rosmm.com" & mid(split_str(i),1,instr(split_str(i),"""")-1)
		'title
		html_str=mid(split_str(i),InStr(split_str(i),"title=""")+7)
		html_str=mid(html_str,1,InStr(html_str,"""")-1)
		split_str(i)="0||" & url_str & "|" & html_str
	Next
	return_albums_list=join(split_str,vbcrlf) & vbcrlf
End If

html_str=str_temp
key_word=""">下一页</a></li>"
If InStr(LCase(html_str),key_word)>0 Then
	html_str=mid(html_str,1,instr(LCase(html_str),key_word)-1)
	html_str=album_url & Mid(html_str,InStrrev(LCase(html_str),"""")+1)
	return_albums_list= return_albums_list & "1|inet|10,13|" & html_str
Else
	return_albums_list= return_albums_list & "0"
End If

End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim str_temp
str_temp = html_str
return_download_list = ""
If InStr(html_str, ".htm"">尾页</a></li>") > 0 And page_num = 0 Then
    '<li><a href="324_7.htm">尾页</a></li>
    html_str = Mid(html_str, 1, InStrRev(html_str, ".htm"">尾页</a></li>") - 1)
    html_str = Mid(html_str, InStrRev(html_str, "_") + 1)
    If IsNumeric(html_str) Then
        page_num = Int(html_str)
    Else
        page_num = 1
    End If
ElseIf page_num <1 Then
page_num = 1
End If

html_str = str_temp
If InStr(LCase(html_str), "<a href=""/pic/upload/") > 0 Then
    Dim split_str
    '<a href="/pic/upload/2012/11/01/rosimm-387-15.jpg" target="_blank"><img src="/pic/upload/2012/11/01/rosimm-387-15.jpg" alt="ROSI写真 No.387 匿名寫真" /></a>
    html_str = Mid(html_str, InStr(LCase(html_str), "<a href=""/pic/upload/") + Len("<a href=""/pic/upload/"))
    html_str = Mid(html_str, 1, InStr(LCase(html_str), "</p>") - 1)
    url_str = html_str
    split_str = Split(html_str, "<a href=""/pic/upload/")
    For i = 0 To UBound(split_str)
        html_str = ""
        url_str = ""
        html_str = split_str(i)
        'alt="ROSI写真 No.387 匿名寫真"
        html_str = Mid(html_str, InStr(LCase(html_str), "alt=""") + 5)
        html_str = Mid(html_str, 1, InStr(LCase(html_str), """") - 1)
        'url
        url_str = "http://www.rosmm.com/pic/upload/" & Mid(split_str(i), 1, InStr(split_str(i), """") - 1)
        split_str(i) = "|" & url_str & "|" & html_str & "_" & Mid(url_str, InStrRev(url_str, "/") + 1) & "|" & html_str
    Next
    return_download_list = Join(split_str, vbCrLf) & vbCrLf
End If

If page_count < page_num Then
    page_count = page_count + 1
    return_download_list = return_download_list & "1|inet|10,13|" & album_url & "_" & page_count & ".htm"
Else
    return_download_list = return_download_list & "0"
End If

End Function