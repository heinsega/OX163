'2011-8-11 163.shanhaijing.net
Dim page_counter
Dim tags, page, url_instr

Function return_download_url(ByVal url_str)
'http://behoimi.org/post?page=2&tags=kurosawa_jennifer
'http://www.behoimi.org/post?page=2&tags=kurosawa_jennifer
'http://behoimi.org/post
'http://behoimi.org/post?page=3
'http://behoimi.org/post/index?tags=tagme
On Error Resume Next
tags=""
Dim page_tmp
If InStr(LCase(url_str), "behoimi.org/pool/show/") >1 Then
    url_str=Mid(url_str, InStr(lcase(url_str),"behoimi.org/pool/show/")+len("behoimi.org/pool/show/"))
    If InStr(LCase(url_str), "/") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "/") - 1)
    If InStr(LCase(url_str), "?") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "?") - 1)
    If InStr(LCase(url_str), "#") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "#") - 1)
    url_str="http://behoimi.org/post/index?tags=pool%3A" & url_str
End If
page_tmp=url_str
If InStr(LCase(url_str), "tags=") > 0 Then
    tags = Mid(url_str, InStr(LCase(url_str), "tags=") + 5)
    url_instr = Mid(url_str,1, InStr(LCase(url_str), "tags=") -2)
    If InStr(tags, "&") > 0 Then tags = Mid(tags, 1, InStr(tags, "&") - 1)
    If InStr(tags, " ") > 0 Then tags = Mid(tags, 1, InStr(tags, " ") - 1)
    If Left(LCase(url_instr),11)="http://www." Then url_instr = "http://" & Mid(url_instr, 12)
    If InStr(LCase(url_instr), "page=") > 0 Then url_instr = Mid(url_instr, 1, InStr(LCase(url_instr), "page=") - 2)
End If
If tags <> "" Then url_str = url_instr & "?tags=" & tags
page_counter=0
page=1
If InStr(LCase(page_tmp),"&page=")>0 Or InStr(LCase(page_tmp),"?page=")>0 Then
	If InStr(LCase(page_tmp),"&page=") Then page_tmp=Mid(page_tmp,InStr(LCase(page_tmp),"&page=")+6)
	If InStr(LCase(page_tmp),"?page=") Then page_tmp=Mid(page_tmp,InStr(LCase(page_tmp),"?page=")+6)
	If InStr(page_tmp,"&") Then page_tmp=Mid(page_tmp,1,InStr(page_tmp,"&")-1)
	If IsNumeric(page_tmp) Then
		If Int(page_tmp)>1 Then
			If MsgBox("本页为第" & page_tmp & "页" & vbcrlf & "是否从第1页开始？", vbYesNo, "问题")=vbyes Then
				page=1
				If InStr(LCase(url_str),"&page=") Or InStr(LCase(url_str),"?page=") Then url_str=format_page(url_str)
			Else
				page=Int(page_tmp)
				If instr(url_str,"?") Then
					url_str=url_str & "&page=" & page
					Else
					url_str=url_str & "?page=" & page
				End if
			End If
		End If
	End If
Else
	page=1
End If
return_download_url = "inet|10,13|" & url_str & "|http://behoimi.org/" & vbcrlf & "User-Agent: QuickTime/7.6.2 (qtver=7.6.2;os=Windows NT 5.1Service Pack 2)"
OX163_urlpage_Referer="http://behoimi.org/" & vbcrlf & "User-Agent: QuickTime/7.6.2 (qtver=7.6.2;os=Windows NT 5.1Service Pack 2)"
End Function
'-----------------------------------------------
Function format_page(url_str)
format_page=url_str
Dim temp_str(2)
If instr(lcase(url_str),"?page=")>0 or instr(lcase(url_str),"&page=")>0 Then
	If instr(lcase(url_str),"?page=")>0 Then
		temp_str(0)=mid(url_str,1,instr(lcase(url_str),"?page=")+5)
		temp_str(1)=mid(url_str,InStr(lcase(url_str),"?page=")+1)
	ElseIf instr(lcase(url_str),"&page=")>0 Then
		temp_str(0)=mid(url_str,1,InStr(lcase(url_str),"&page=")+5)
		temp_str(1)=mid(url_str,InStr(lcase(url_str),"&page=")+1)
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
Dim key_str, split_str, add_temp, file_url
return_download_list = ""
url_str = html_str
key_str="Post.register({"
If InStr(LCase(html_str), LCase(key_str)) > 0 Then
	html_str = Mid(html_str, InStr(LCase(html_str), LCase(key_str)) + len(key_str))
	
	split_str = Split(html_str, key_str)
	
	For split_i = 0 To UBound(split_str)
			'file_url
			file_url=""
			key_str=",""file_url"":"""
	    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + len(key_str))
	    file_url=Mid(split_str(split_i),1,InStr(split_str(split_i),chr(34))-1)
	    
			'tags
			html_str=""	
			key_str=",""tags"":"""
	    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + len(key_str))
	    html_str=Mid(split_str(split_i),1,InStr(split_str(split_i),chr(34))-1)
	    html_str=replace(html_str,"|","&#124;")
	    html_str=replace(html_str,"\\","\")
    
			'ID
			add_temp=""
			key_str=",""id"":"
	    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + len(key_str))
	    add_temp=Mid(split_str(split_i),1,InStr(split_str(split_i),",")-1)
			If IsNumeric(add_temp)=false Then add_temp=""
			
			'file name
			split_str(split_i)="(behoimi.org)p" & add_temp & "_" & Trim(html_str)
	    If Len(split_str(split_i))>100 Then split_str(split_i)=Left(split_str(split_i),99) & "~"
	    split_str(split_i) = Replace(split_str(split_i), " ", "-") & Mid(file_url, InStrRev(file_url, "."))	    

	    return_download_list = return_download_list & "|" & file_url & "|" & split_str(split_i) & "|" & html_str & vbCrLf
	Next
End If

If InStr(LCase(url_str), "<div id=""paginator"">") > 0 or page<page_counter Then
    If page_counter = 0 Then
	    url_str = Mid(url_str, InStr(LCase(url_str), "<div id=""paginator"">") + 20)
	    url_str = Mid(url_str, InStr(LCase(url_str), "<a href=""") + 9)
	    url_str = Mid(url_str, 1, InStr(LCase(url_str), "</div>") - 1)
	    split_str = Split(url_str, "<a href=""", -1, 1)
	    url_str = Mid(split_str(UBound(split_str) - 1), InStr(split_str(UBound(split_str) - 1), ">") + 1)
	    url_str = Mid(url_str, 1, InStr(url_str, "<") - 1)
	    If IsNumeric(url_str) Then
	    		page_counter = Int(url_str)
	    Else
	        page_counter = 1
	    End If
    End If

    If page < page_counter And tags <> "" Then
	    page = page + 1
	    return_download_list = return_download_list & page_counter & "|inet|10,13|" & "http://www.behoimi.org/post/index?tags=" & tags & "&page=" & page
    ElseIf page < page_counter Then
	    page = page + 1
	    return_download_list = return_download_list & page_counter & "|inet|10,13|" & "http://www.behoimi.org/post/index?page=" & page
    Else
    return_download_list = return_download_list & "0"
    End If
Else
		return_download_list = return_download_list & "0"
End If
End Function