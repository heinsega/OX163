'2011-8-11 163.shanhaijing.net
Dim page,pid,tags,url_instr,page_retry

Function return_download_url(ByVal url_str)
'http://gelbooru.com/index.php?page=post&s=list&pid=336960
'http://gelbooru.com/index.php?page=post&s=list&tags=all
'http://gelbooru.com/index.php?page=post&s=list
'http://gelbooru.com/index.php?page=post&s=list&tags=nakoruru
'http://gelbooru.com/index.php?page=post&s=list&tags=nakoruru&pid=60
'http://www.gelbooru.com/index.php?page=post&s=list&tags=canaan
On Error Resume Next
tags=""
Dim page_tmp
page_tmp=url_str

If InStr(url_str, "tags=") > 0 Then
    tags = Mid(url_str, InStr(url_str, "tags=") + 5)
    If InStr(tags, "&") > 0 Then tags = Mid(tags, 1, InStr(tags, "&") - 1)
    If InStr(tags, " ") > 0 Then tags = Mid(tags, 1, InStr(tags, " ") - 1)
    If LCase(tags)="all" Then tags=""
End If
If tags <> "" Then url_str = "http://gelbooru.com/index.php?page=post&s=list&tags=" & tags
url_instr=url_str
page_retry=0
pid=0
page=0
If InStr(LCase(page_tmp),"&pid=")>len("gelbooru.com/") Or InStr(LCase(page_tmp),"?pid=")>len("gelbooru.com/") Then
	If InStr(LCase(page_tmp),"&pid=")>len("gelbooru.com/") Then page_tmp=Mid(page_tmp,InStr(LCase(page_tmp),"&pid=")+5)
	If InStr(LCase(page_tmp),"?pid=")>len("gelbooru.com/") Then page_tmp=Mid(page_tmp,InStr(LCase(page_tmp),"?pid=")+5)
	If InStr(LCase(page_tmp),"&")>0 Then page_tmp=Mid(page_tmp,1,InStr(page_tmp,"&")-1)
	If IsNumeric(page_tmp) Then
		If Int(page_tmp)>1 Then
			If MsgBox("本页为第" & Int(page_tmp/25)+1 & "页" & vbcrlf & "是否从第1页开始？", vbYesNo, "问题")=vbyes Then
				page=0
				url_str=format_page(url_str)
			Else
				page=Int(page_tmp)
			End If
		End If
	End If
Else
	page=0
End If
If page>0 Then url_str=url_str & "&pid=" & page
return_download_url = "inet|10,13|" & url_str & "|http://gelbooru.com/"
End Function
'--------------------------------------------------------
Function format_page(url_str)
format_page=url_str
Dim temp_str(2)
If instr(lcase(url_str),"?pid=")>0 or instr(lcase(url_str),"&pid=")>0 Then
	If instr(lcase(url_str),"?pid=")>0 Then
		temp_str(0)=mid(url_str,1,instr(lcase(url_str),"?pid="))
		temp_str(1)=mid(url_str,InStr(lcase(url_str),"?pid=")+1)
	ElseIf instr(lcase(url_str),"&pid=")>0 Then
		temp_str(0)=mid(url_str,1,InStr(lcase(url_str),"&pid="))
		temp_str(1)=mid(url_str,InStr(lcase(url_str),"&pid=")+1)
	End If
	If instr(temp_str(1),"&")>0 Then
		temp_str(1)=mid(url_str,instr(temp_str(1),"&"))
	Else
		temp_str(1)=""
	End If
	format_page=temp_str(0) & "1" & temp_str(1)
End if
End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""
url_str=html_str
If InStr(html_str, " class=""thumb"">") > 0 Then
	
html_str = Mid(html_str, InStr(html_str, " class=""thumb"">") + 15)
html_str = Mid(html_str, InStr(html_str, "<a id=""") + 7)

Dim split_str,url_temp
split_str = Split(html_str, " class=""thumb""><a id=""")

    For split_i = 0 To UBound(split_str)
    html_str=Mid(split_str(split_i),1, InStr(split_str(split_i), Chr(34)) -1) & "_" 'p371667_
    split_str(split_i) = Mid(split_str(split_i), InStr(split_str(split_i), "<img src=""") +10)
    
    'url
    url_temp = Mid(split_str(split_i), 1,InStr(split_str(split_i), "?") -1)
    url_temp = replace(replace(url_temp,".gelbooru.com/thumbs/",".gelbooru.com//images/"),"thumbnail_","")
    
    'Tags
    html_str =html_str & Trim(Mid(split_str(split_i), InStr(split_str(split_i), "alt=""") +5))
    html_str =Trim(Mid(html_str,1, InStr(html_str, """")-1))
    
    split_str(split_i)=html_str
    'name
    If Len(html_str)>180 Then html_str=Left(html_str,179) & "~"
    html_str=html_str & Mid(url_temp,instrrev(url_temp,"."))
    'If instrrev(html_str,".")<instrrev(html_str,"?") and instrrev(html_str,".")>8 Then html_str=Mid(html_str,1,instrrev(html_str,"?")-1)
    
    return_download_list = return_download_list & "|" & url_temp & "|" & html_str & "|" & split_str(split_i) & vbCrLf
    Next
End If

If InStr(url_str, "alt=""last page"">") > 0 and pid=0 Then
	url_str = Mid(url_str, InStr(url_str, "alt=""next"">") + 11)
	url_str = Mid(url_str, InStr(url_str, "&amp;pid=") + 9)
	url_str = Mid(url_str,1, InStr(url_str, Chr(34)) -1)
	If IsNumeric(url_str) Then pid=CLng(url_str)
End If

If page<pid and pid>0 and return_download_list<>"" Then
	page_retry=0
	page=page+25
	url_instr="http://gelbooru.com/index.php?page=post&s=list&tags=" & tags & "&pid=" & page
	return_download_list = return_download_list & pid & "|inet|10,13|" & url_instr
ElseIf page<pid and pid>0 and return_download_list="" and page_retry<5 Then
	page_retry=page_retry+1
	return_download_list = pid & "|inet|10,13|" & url_instr
Else
	return_download_list = return_download_list & "0"
End If
End Function