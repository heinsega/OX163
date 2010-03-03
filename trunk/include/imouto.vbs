'2010-2-14 163.shanhaijing.net
Dim page_counter
Dim tags, page, url_instr, pool
Function return_download_url(ByVal url_str)
'http://moe.imouto.org/post?tags=emma
'http://moe.imouto.org/post?page=5&tags=miko
'http://moe.imouto.org/post/index
'http://moe.imouto.org/post?page=3
'http://moe.imouto.org/post/index?tags=tagme
'http://moe.imouto.org/pool/show/596?page=3
'http://moe.imouto.org/wiki/show?title=park_sung-woo
'http://moe.imouto.org/wiki/show?page=3&title=park_sung-woo
On Error Resume Next
tags=""
If InStr(LCase(url_str), "tags=") > 0 Then
    tags = Mid(url_str, InStr(LCase(url_str), "tags=") + 5)
    url_instr = Mid(url_str,1, InStr(LCase(url_str), "tags=") -2)
    If InStr(tags, "&") > 0 Then tags = Mid(tags, 1, InStr(tags, "&") - 1)
    If InStr(tags, " ") > 0 Then tags = Mid(tags, 1, InStr(tags, " ") - 1)
    If InStr(LCase(url_instr), "page=") > 0 Then url_instr = Mid(url_instr, 1, InStr(LCase(url_instr), "page=") - 2)
End If
If tags <> "" Then url_str = url_instr & "?tags=" & tags
pool=""
If InStr(LCase(url_str), "http://moe.imouto.org/pool/show/") =1 Then
    pool="pool"
    If InStr(LCase(url_str), "?page=") > 0 Then url_str = Mid(url_str, 1, InStr(LCase(url_str), "?page=") - 1)
    url_instr=url_str
ElseIf InStr(LCase(url_str), "http://moe.imouto.org/wiki/show?") =1 Then
    pool="wiki"
    url_instr=Mid(url_str, InStr(url_str, "?") - 1)
    If InStr(LCase(url_instr), "?title=") > 0 Then
    	url_instr = Mid(url_instr, InStr(LCase(url_instr), "?title="))
    	If InStr(url_instr, "&")>0 Then url_instr = Mid(url_instr,1,InStr(url_instr, "&")-1)
    ElseIf InStr(LCase(url_instr), "&title=") > 0 Then
    	url_instr = "?" & Mid(url_instr, InStr(LCase(url_instr), "&title=")+1)
    	If InStr(url_instr, "&")>0 Then url_instr = Mid(url_instr,1,InStr(url_instr, "&")-1)
    End If
    url_str="http://moe.imouto.org/wiki/show" & url_instr
    url_instr=url_str
End If
page_counter=0
page=1
return_download_url = "inet|10,13,34|" & url_str & "|http://moe.imouto.org/"

End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""
'moe.e-n-m.net saki.imouto.org http://sheryl.imouto.org
'html_str=replace(html_str,"saki.imouto.org","moe.imouto.org")
'html_str=replace(html_str,"moe.e-n-m.net","moe.imouto.org")
url_str=html_str
If InStr(LCase(html_str), "<span class=thumb") > 0 Then
html_str = Mid(html_str, InStr(LCase(html_str), "<li id=p") + 8)

Dim split_str,url_temp
split_str = Split(html_str, "<li id=p")

    For split_i = 0 To UBound(split_str)
    'html_str="p" & Mid(split_str(split_i),1, InStr(1, split_str(split_i), " ", 1) -1) & "_"
    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "tags:") +5)

    'Tags
    'html_str =html_str & replace(Trim(Mid(split_str(split_i),1, InStr(1, split_str(split_i), "class=", 1) -1))," ","-")

    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "<a class=directlink"))
    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "href=") +5)
    'url
    url_temp = Mid(split_str(split_i), 1,InStr(split_str(split_i), ">") -1)
    
    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "<span>") +6)
    split_str(split_i) = Mid(split_str(split_i), 1,InStr(LCase(split_str(split_i)), "</span>") -1)
    
    'name
    html_str=unescape(Mid(url_temp,instrrev(url_temp,"/")+1))
    
    return_download_list = return_download_list & "|" & url_temp & "|" & html_str & "|" & split_str(split_i) & vbCrLf
    Next
End If

If InStr(LCase(url_str), "<div id=paginator>") > 0 and (tags<>"" or pool<>"") Then
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
	ElseIf page<page_counter and pool="pool" Then
	page=page+1
	return_download_list = return_download_list & page_counter & "|inet|10,13,34|" & url_instr & "?page=" & page
	ElseIf page<page_counter and pool="wiki" Then
	page=page+1
	return_download_list = return_download_list & page_counter & "|inet|10,13,34|" & url_instr & "&page=" & page
	Else
	return_download_list = return_download_list & "0"	
	End If
Else
return_download_list = return_download_list & "0"
End If
End Function