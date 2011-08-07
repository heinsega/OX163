'2011-6-10 163.shanhaijing.net
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
return_download_url = "inet|10,13,34|" & url_str & "|http://behoimi.org/" & vbcrlf & "User-Agent: QuickTime/7.6.2 (qtver=7.6.2;os=Windows NT 5.1Service Pack 2)"
OX163_urlpage_Referer="http://behoimi.org/" & vbcrlf & "User-Agent: QuickTime/7.6.2 (qtver=7.6.2;os=Windows NT 5.1Service Pack 2)"
End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""
url_str=html_str
If InStr(LCase(html_str), "<span class=thumb blacklisted id=") > 0 Then
html_str = Mid(html_str, InStr(LCase(html_str), "<span class=thumb blacklisted id=") + Len("<span class=thumb blacklisted id="))

Dim split_str,url_temp
split_str = Split(html_str, "<span class=thumb blacklisted id=")

    For split_i = 0 To UBound(split_str)
    html_str="behoimi.org-" & Mid(split_str(split_i),1, InStr(split_str(split_i), ">") -1) & "_"
    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "<a href=") +8)
    'Tags
    url_temp=Trim(Mid(split_str(split_i),1,InStr(split_str(split_i), " ") -1))
    html_str =html_str & replace(Mid(url_temp,instrrev(url_temp,"/")+1)," ","-")

    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "<img "))
    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "src=")+4)
    'url
    split_str(split_i) = replace(Mid(split_str(split_i), 1,InStr(split_str(split_i), " ") -1),"/preview/","/")
        
    'name
    html_str=html_str & Mid(split_str(split_i),instrrev(split_str(split_i),"."))
    
    return_download_list = return_download_list & "|" & split_str(split_i) & "|" & html_str & "|" & vbCrLf
    Next
End If

If InStr(LCase(url_str), "<div id=paginator>") > 0 and tags<>"" Then
	If page_counter=0 Then
	url_str = Mid(url_str, InStr(LCase(url_str), "<div id=paginator>") + 18)
	url_str = Mid(url_str, InStr(LCase(url_str), "<a href=") + 8)
	url_str = Mid(url_str,1, InStr(LCase(url_str), "</div>") -1)
	split_str=Split(url_str, "<a href=", -1, 1)
	url_str=Mid(split_str(UBound(split_str)-1), InStr(split_str(UBound(split_str)-1), ">") + 1)
	url_str=Mid(url_str,1, InStr(url_str, "<") -1)
		If IsNumeric(url_str) Then
			page_counter=Int(url_str)		
		Else
			page_counter=1
		End If
	End If

	If page<page_counter and tags<>"" Then
	page=page+1
	return_download_list = return_download_list & page_counter & "|inet|10,13,34|" & url_instr & "?page=" & page & "&tags=" & tags
	Else
	return_download_list = return_download_list & "0"
	
	End If
Else
return_download_list = return_download_list & "0"
End If
End Function