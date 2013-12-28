'2013-12-28 163.shanhaijing.net
Dim url_parent, tags, page, page_counter, Next_page, retry_counter

Function return_download_url(ByVal url_str)
'http://rule34.booru.org/index.php?page=post&s=view&id=44
'http://rule34.booru.org/index.php?page=post&s=list&tags=all&pid=768750
'http://rule34.booru.org/index.php?page=pool&s=show&id=5
'http://rule34.xxx/index.php?page=dapi&s=post&q=index&tags=nakoruru&pid=0
On Error Resume Next
return_download_url=""

url_parent=mid(url_str,InStr(LCase(url_str), "http://")+7)
url_parent=mid(url_parent,1,InStr(LCase(url_parent), "/")-1)
	
If InStr(LCase(url_str), "page=pool") >1 Then
	page="pool"
	return_download_url = "inet|10,13|" & url_str & "|" & url_str
Else
	'tags
	If InStr(LCase(url_str), "tags=") >1 Then
		tags=mid(url_str,InStr(LCase(url_str), "tags=")+5)
		If InStr(tags, "&") >1 Then tags=mid(tags,1,InStr(tags, "&")-1)
	Else
		tags=""
	End If
	If Trim(tags)="" Then
		tags=""
		Else
		tags="&tags=" & Trim(tags)
	End If
	
	page=0
	page_counter=0
	retry_counter=0
	return_download_url = "inet|10,13|http://" & url_parent & "/index.php?page=dapi&s=post&q=index" & tags & "|" & "http://" & url_parent & "/"
End If
return_download_url=return_download_url & vbcrlf & "User-Agent: Mozilla/4.0 (compatible; MSIE 8.00; Windows XP)"
OX163_urlpage_Referer="http://" & url_parent & vbcrlf & "User-Agent: Mozilla/4.0 (compatible; MSIE 8.00; Windows XP)"
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

ElseIf InStr(LCase(html_str), "<posts count=""") > 0 Then
	retry_counter=0
	Dim key_word
	key_word="<posts count="""
	url_str=Mid(html_str, InStr(LCase(html_str), key_word) + len(key_word))
	url_str=Mid(url_str,1,InStr(url_str, chr(34))-1)
	If IsNumeric(url_str) Then page_counter=Int(url_str)

	html_str = Mid(html_str, InStr(LCase(html_str), key_word) + len(key_word))
	html_str = Mid(html_str, InStr(LCase(html_str), "<post ") + len("<post "))
	html_str = Mid(html_str, 1, InStr(LCase(html_str), "</posts>")-1)
	split_str = Split(html_str, "/><post ")
  For split_i = 0 To UBound(split_str)
		html_str=""
		url_str=""
		sid=""
		pic_type=""
				
  	'sid
  	key_word=""" id="""
	  sid = Mid(split_str(split_i), InStr(split_str(split_i), key_word) + len(key_word))
	  sid = Mid(sid,1,InStr(sid, chr(34))-1)
	  	  	  
	  'file_url
  	key_word=""" file_url="""
	  url_str = Mid(split_str(split_i), InStr(split_str(split_i), key_word) + len(key_word))
	  url_str = Mid(url_str,1,InStr(url_str, chr(34))-1)	  
	  
		pic_type=Mid(url_str,instrrev(url_str,"."))
		
		'pic name
  	key_word=""" tags="""
	  html_str = Mid(split_str(split_i), InStr(split_str(split_i), key_word) + len(key_word))
	  html_str = Mid(html_str,1,InStr(html_str, chr(34))-1)
	  
	  If Len(html_str)>180 Then html_str=Left(html_str,179) & "~"
	  html_str = replace(html_str,"|","_")
	  html_str = "(" & url_parent & ") " & sid & " - " & html_str & pic_type
		split_str(split_i) = "|" & url_str & "|" & html_str & "|"
  Next
  
  return_download_list=join(split_str,vbCrLf) & vbCrLf
  If (page+1)*100<page_counter Then
  	page=page+1
  	return_download_list=return_download_list & "1|inet|10,13|http://" & url_parent & "/index.php?page=dapi&s=post&q=index" & tags & "&pid=" & page
	End If
ElseIf retry_counter<3 Then
	retry_counter=retry_counter+1
	return_download_list="1|inet|10,13|http://" & url_parent & "/index.php?page=dapi&s=post&q=index" & tags & "&pid=" & page
Else
	return_download_list = "0"
End If

End Function