'2011-6-15 163.shanhaijing.net
Dim url_parent, tags, page, page_counter, Next_page

Function return_download_url(ByVal url_str)
'http://rule34.booru.org/index.php?page=post&s=view&id=44
'http://rule34.booru.org/index.php?page=post&s=list&tags=all&pid=768750
'http://rule34.booru.org/index.php?page=pool&s=show&id=5
On Error Resume Next
return_download_url=""

'rule34
url_parent=mid(url_str,1,InStr(LCase(url_str), ".booru.org/")-1)
url_parent=mid(url_parent,InStr(LCase(url_parent), "http://")+7)
	
If InStr(LCase(url_str), "page=post") >1 and InStr(LCase(url_str), "s=list") >1 Then	
	'tags
	If InStr(LCase(url_str), "tags=") >1 Then
		tags=mid(url_str,InStr(LCase(url_str), "tags=")+5)
	  If InStr(url_str, "&") >1 Then tags=mid(tags,1,InStr(tags, "&")-1)
	Else
		tags="all"
	End If

	'page
	If InStr(LCase(url_str), "&pid=") >1 Then
		page=mid(url_str,InStr(LCase(url_str), "&pid=")+5)
	  If IsNumeric(page)=false Then page=0
	Else
		page=0
	End If
	
	If page>24 Then
		If MsgBox("本页为第" & (int(page/25)+1) & "页" & vbcrlf & "是否从第1页开始？", vbYesNo, "问题")=vbyes Then page=0
	Else
	  page=0
	End If
	
	page_counter = 0
	Next_page=0
	return_download_url = "inet|10,13|http://" & url_parent & ".booru.org/index.php?page=post&s=list&tags=" & tags & "&pid=" & page & "|" & "http://" & url_parent & ".booru.org/"
ElseIf InStr(LCase(url_str), "page=pool") >1 Then
	page="pool"
	return_download_url = "inet|10,13|" & url_str & "|" & url_str
End If
OX163_urlpage_Referer="http://" & url_parent & ".booru.org/" & vbcrlf & "User-Agent: Mozilla/4.0 (compatible; MSIE 8.00; Windows XP)"
End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next	
Dim split_str, sid, pic_type

If page="pool" Then
	'pool部分仅仅复制调整tags部分的代码
	html_str = Mid(html_str, InStr(LCase(html_str), "<span class=""thumb"" id=""") + len("<span class=""thumb"" id="""))
	html_str = Mid(html_str, 1, InStr(LCase(html_str), "</div>"))
	split_str = Split(html_str, "<span class=""thumb"" id=""")
  For split_i = 0 To UBound(split_str)
		html_str=""
		url_str=""
		sid=""
		pic_type=""
  	'sid
	  sid = Mid(split_str(split_i), 1, InStr(split_str(split_i), chr(34))-1)	  
	  split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "src=""")+len("src="""))	  	  
	  'url
	  url_str = Mid(split_str(split_i), 1, InStr(split_str(split_i), chr(34))-1)
	  
	  html_str = Mid(url_str,1,InStr(LCase(url_str), "/thumbnail_")-1)
	  url_str = Mid(url_str, InStr(LCase(url_str), "/thumbnail_")+len("/thumbnail_"))
	  'html_str获取"/17"部分
	  html_str = Mid(html_str,InStrrev(LCase(html_str), "/")) & "/"
	  'pic url
		url_str="http://img.booru.org/" & url_parent & "//images" & html_str & url_str	
		If InStr(url_str,"?")>1 Then url_str=mid(url_str,1,InStr(url_str,"?")-1)
		pic_type=Mid(url_str,instrrev(url_str,"."))
		
		'pic name
	  split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "title=""")+len("title="""))
	  split_str(split_i) = Trim(Mid(split_str(split_i), 1, InStr(split_str(split_i), chr(34))-1))
	  
	  If InStr(LCase(split_str(split_i)), " rating:")>0 Then split_str(split_i) = Trim(Mid(split_str(split_i),1,InStr(LCase(split_str(split_i)), " rating:"))-1)
	  If Len(split_str(split_i))>180 Then split_str(split_i)=Left(split_str(split_i),179) & "~"
	  
	  split_str(split_i) = replace(split_str(split_i),"|","_")
	  split_str(split_i) = "(" & url_parent & ".booru.org) " & sid & " - " & split_str(split_i) & pic_type
		split_str(split_i) = "|" & url_str & "|" & split_str(split_i) & "|"
  Next
  
  return_download_list=join(split_str,vbCrLf) & vbCrLf
ElseIf InStr(LCase(html_str), "class=""thumb""><a id=""") > 0 Then
	Next_page=0
	If InStr(LCase(html_str), "alt=""next"">&gt;</a>") > 0 or page<page_counter*25 Then
		Next_page=1
		If InStr(LCase(html_str), "alt=""last page"">&gt;&gt;</a>") > 0 Then
			url_str=Mid(html_str, InStr(LCase(html_str), "alt=""next"">&gt;</a>") + len("alt=""next"">&gt;</a>"))
			url_str=Mid(url_str, InStr(LCase(url_str), "pid=")+4)
			url_str=Mid(url_str,1,InStr(url_str, chr(34))-1)
			If IsNumeric(url_str) Then page_counter=Int(url_str)
		End If
	End If
	html_str = Mid(html_str, InStr(LCase(html_str), "class=""thumb""><a id=""") + len("class=""thumb""><a id="""))
	html_str = Mid(html_str, 1, InStr(LCase(html_str), "</div>"))
	split_str = Split(html_str, "class=""thumb""><a id=""")
  For split_i = 0 To UBound(split_str)
		html_str=""
		url_str=""
		sid=""
		pic_type=""
  	'sid
	  sid = Mid(split_str(split_i), 1, InStr(split_str(split_i), chr(34))-1)
	  
	  split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "<img src=""")+len("<img src="""))
	  	  
	  'url
	  url_str = Mid(split_str(split_i), 1, InStr(split_str(split_i), chr(34))-1)

	  'http://thumbs2.booru.org/r34/769/thumbnail_28d4f9856426439adbab1ad0f0a7fcc84d3dfcb6.png?768815
		'http://img.booru.org/rule34//images/769/28d4f9856426439adbab1ad0f0a7fcc84d3dfcb6.png		
	  'http://thumbs.booru.org/miko/thumbnails//2/thumbnail_f5e3f7b34316300418638b7922301c604cb92ddc.jpg
		'http://img.booru.org/miko//images/2/f5e3f7b34316300418638b7922301c604cb92ddc.jpg		
		'html_str获取"http://thumbs.booru.org/Equi/thumbnails//17"部分
		'url_str获取"/thumbnail_2f7371104749bc169be51da29c56c932653bb849.jpg"部分
	  html_str = Mid(url_str,1,InStr(LCase(url_str), "/thumbnail_")-1)
	  url_str = Mid(url_str, InStr(LCase(url_str), "/thumbnail_")+len("/thumbnail_"))
	  'html_str获取"/17"部分
	  html_str = Mid(html_str,InStrrev(LCase(html_str), "/")) & "/"
	  'pic url
		url_str="http://img.booru.org/" & url_parent & "//images" & html_str & url_str	
		If InStr(url_str,"?")>1 Then url_str=mid(url_str,1,InStr(url_str,"?")-1)
		pic_type=Mid(url_str,instrrev(url_str,"."))
		
		'pic name
	  split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "alt=""")+len("alt="""))
	  split_str(split_i) = Trim(Mid(split_str(split_i), 1, InStr(split_str(split_i), chr(34))-1))
	  If Len(split_str(split_i))>180 Then split_str(split_i)=Left(split_str(split_i),179) & "~"
	  split_str(split_i) = replace(split_str(split_i),"|","_")
	  split_str(split_i) = "(" & url_parent & ".booru.org) " & sid & " - " & split_str(split_i) & pic_type
		split_str(split_i) = "|" & url_str & "|" & split_str(split_i) & "|"
  Next
  return_download_list=join(split_str,vbCrLf) & vbCrLf
  If Next_page=1 or page<page_counter Then
  	page=page+25
  	return_download_list=return_download_list & "1|inet|10,13|http://" & url_parent & ".booru.org/index.php?page=post&s=list&tags=" & tags & "&pid=" & page
	End If
Else
return_download_list = "0"
End If

End Function