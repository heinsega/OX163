'2024-1-8 163.shanhaijing.net
Dim http_type, url_parent, tags, page, Next_page, page_counter, retry_counter

Function return_download_url(ByVal url_str)
Dim XML_TF
'/index.php?page=post&s=view&id=44
'/index.php?page=post&s=list&tags=all&pid=768750
'/index.php?page=pool&s=show&id=5
'http://rule34.booru.org/index.php?page=post&s=view&id=44
'http://rule34.booru.org/index.php?page=post&s=list&tags=all&pid=768750
'http://rule34.booru.org/index.php?page=pool&s=show&id=5
'http://rule34.xxx/index.php?page=dapi&s=post&q=index&tags=nakoruru&pid=0
On Error Resume Next
return_download_url=""

http_type="http://"
If InStr(LCase(url_str), "https://")>0 Then http_type="https://"

url_str=http_type & mid(url_str,InStr(LCase(url_str), http_type)+len(http_type))

url_parent=mid(url_str,InStr(LCase(url_str), LCase(http_type))+len(http_type))

If InStr(LCase(url_parent), "/index.php?")>0 Then
	url_parent=mid(url_parent,1,InStr(LCase(url_parent), "/index.php?")-1)
ElseIf InStr(LCase(url_parent), "/?")>0 Then
	url_parent=mid(url_parent,1,InStr(LCase(url_parent), "/?")-1)
Else
	url_parent=mid(url_parent,1,InStrrev(LCase(url_parent), "/")-1)			
End If

If InStr(LCase(url_str), "page=pool") >1 Then
	page="pool"
	return_download_url = "inet|10,13|" & url_str & "|" & url_str
Else  'If InStr(LCase(url_str), "page=post") >1 and InStr(LCase(url_str), "s=list") >1 Then
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

	'page
	If InStr(LCase(url_str), "&pid=") >1 Then
		page=mid(url_str,InStr(LCase(url_str), "&pid=")+5)
	  If IsNumeric(page)=false Then page=0
	Else
		page=0
	End If
	
	XML_TF=0
	If Specific_web_site(url_parent)=0 Then
		If MsgBox("�Ƿ���ʹ�ø����ٵ�XMLģʽ�б�ͼƬ?" & vbcrlf & "(XMLģʽֻ֧�ִӵ�һҳ��ʼ�б�ĳЩ��վ��Ҫ��XMLģʽ�²�����ȷ���ͼƬ���ӱ��硰rule34��)", vbYesNo, "����")=vbyes Then
			page=0
			XML_TF=1
		End If
	End If
		
	If page>0 Then
		If MsgBox("��ҳ���ǵ�1ҳ" & vbcrlf & "�Ƿ�ӵ�1ҳ��ʼ��", vbYesNo, "����")=vbyes Then page=0
	Else
	  page=0
	End If	
	Next_page=0
	retry_counter=0
	
	If XML_TF=1 Then
	 	page_counter=0
	 	Next_page=-1
	 	return_download_url = "inet|10,13|" & http_type & url_parent & "/index.php?page=dapi&s=post&q=index" & tags & "|" & http_type & url_parent & "/"
	Else
		return_download_url = "inet|10,13|" & http_type & url_parent & "/index.php?page=post&s=list" & tags & "&pid=" & page & "|" & http_type & url_parent & "/"
	End If
End If
return_download_url=return_download_url & vbcrlf & "User-Agent: Mozilla/4.0 (compatible; MSIE 8.00)"
OX163_urlpage_Referer=http_type & url_parent & "/" & vbcrlf & "User-Agent: Mozilla/4.0 (compatible; MSIE 8.00)"

End Function

'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next

Dim split_str, sid, pic_type
If page="pool" Then
	'pool���ֽ������Ƶ���tags���ֵĴ���
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
		'http://gelbooru.com/thumbnails/97/b1/thumbnail_97b1d4c93ed8c45ba3e1415f985faea0.jpg?2808368
		'http://simg4.gelbooru.com//images/97/b1/97b1d4c93ed8c45ba3e1415f985faea0.png?280836
		'html_str��ȡ"http://gelbooru.com/thumbnails/97/b1"����
		'url_str��ȡ"/thumbnail_97b1d4c93ed8c45ba3e1415f985faea0.jpg?2808368"����
	  html_str = Mid(url_str,1,InStr(LCase(url_str), "/thumbnail_")-1)
	  url_str = Mid(url_str, InStr(LCase(url_str), "/thumbnail_")+len("/thumbnail_"))
	  'html_str��ȡ"/97/b1"����
	  If InStr(html_str,"://")>0 Then html_str=Mid(html_str, InStr(html_str, "://")+3)
	  Do While InStr(html_str,"//")>0
	  	html_str=replace(html_str,"//","/")
	  loop
	  Do While Right(html_str,1)="/"
	  	html_str=Left(html_str,len(html_str)-1)
	  loop
	  html_str = Mid(html_str,InStr(html_str, "/")+1)
	  html_str = Mid(html_str,InStr(html_str, "/")+1)
	  'pic url
		url_str=http_type & url_parent & "/images/" & html_str & "/" & url_str
		If InStr(url_str,"?")>1 Then url_str=mid(url_str,1,InStr(url_str,"?")-1)
		pic_type=Mid(url_str,instrrev(url_str,"."))
		
		'pic name
	  split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "title=""")+len("title="""))
	  split_str(split_i) = Trim(Mid(split_str(split_i), 1, InStr(split_str(split_i), chr(34))-1))
	  split_str(split_i) = Trim(Mid(split_str(split_i), 1, InStr(split_str(split_i), "  ")-1))
	  If Len(split_str(split_i))>180 Then split_str(split_i)=Left(split_str(split_i),179) & "~"
    split_str(split_i) = replace(split_str(split_i),"|","_")
	  split_str(split_i) = "(" & Specific_web_name(url_parent) & ") " & sid & " - " & split_str(split_i) & pic_type
		split_str(split_i) = "|" & url_str & "|" & split_str(split_i) & "|"
  Next  
  return_download_list=join(split_str,vbCrLf) & vbCrLf

ElseIf Next_page<0 Then
	
	If InStr(LCase(html_str), "<posts count=""") > 0 Then
		Next_page=-2
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
		  If left(url_str,2)="//" Then url_str=http_type & Mid(url_str,3)
			pic_type=Mid(url_str,instrrev(url_str,"."))
			
			'pic name
	  	key_word=""" tags="""
		  html_str = Mid(split_str(split_i), InStr(split_str(split_i), key_word) + len(key_word))
		  html_str = Mid(html_str,1,InStr(html_str, chr(34))-1)
		  
		  If Len(html_str)>180 Then html_str=Left(html_str,179) & "~"
		  html_str = replace(html_str,"|","_")
		  html_str = "(" & Specific_web_name(url_parent) & ") " & sid & " - " & html_str & pic_type
			split_str(split_i) = "|" & url_str & "|" & html_str & "|"
	  Next
	  
	  return_download_list=join(split_str,vbCrLf) & vbCrLf
	  If (page+1)*100<page_counter Then
	  	page=page+1
	  	return_download_list=return_download_list & "1|inet|10,13|" & http_type & url_parent & "/index.php?page=dapi&s=post&q=index" & tags & "&pid=" & page
		End If
		
	ElseIf Next_page<-1 Then
		If retry_counter<3 Then
			retry_counter=retry_counter+1
			return_download_list="1|inet|10,13|" & http_type & url_parent & "/index.php?page=dapi&s=post&q=index" & tags & "&pid=" & page
		Else
			return_download_list="0"
		End If
	Else
		Next_page=0
		return_download_list = "1|inet|10,13|" & http_type & url_parent & "/index.php?page=post&s=list" & tags & "&pid=" & page & "|" & http_type & url_parent & "/"
	End If
  
ElseIf InStr(LCase(html_str), "class=""thumb""><a id=""") > 0 Then
	
	Next_page=0
	retry_counter=0
	If InStr(LCase(html_str), "alt=""next"">") > 0 Then
		url_str=Mid(html_str,1,InStr(LCase(html_str), "alt=""next"">")-1)
		url_str=Mid(url_str, InStrrev(LCase(url_str), "pid=")+4)
		url_str=Mid(url_str,1,InStr(url_str, chr(34))-1)
		If IsNumeric(url_str) Then Next_page=Int(url_str)
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
		'http://thumbs2.booru.org/safe/638/thumbnail_d6679254289b8e22c2462b172f8347c66327e8c9.jpg?643485
		'http://safebooru.org//images/638/d6679254289b8e22c2462b172f8347c66327e8c9.jpg
		'http://lolibooru.com/thumbnails//106/thumbnail_7d027b47b1bfcb4cf39775332437dea8ae52a514.jpeg
		'http://lolibooru.com/images/106/7d027b47b1bfcb4cf39775332437dea8ae52a514.jpeg
		'http://gelbooru.com/thumbnails/97/b1/thumbnail_97b1d4c93ed8c45ba3e1415f985faea0.jpg?2808368
		'http://simg4.gelbooru.com//images/97/b1/97b1d4c93ed8c45ba3e1415f985faea0.png?280836
		'html_str��ȡ"http://gelbooru.com/thumbnails/97/b1"����
		'url_str��ȡ"/thumbnail_97b1d4c93ed8c45ba3e1415f985faea0.jpg?2808368"����
	  html_str = Mid(url_str,1,InStr(LCase(url_str), "/thumbnail_")-1)
	  url_str = Mid(url_str, InStr(LCase(url_str), "/thumbnail_")+len("/thumbnail_"))
	  'html_str��ȡ"/97/b1"����
	  If InStr(html_str,"://")>0 Then html_str=Mid(html_str, InStr(html_str, "://")+3)
	  Do While InStr(html_str,"//")>0
	  	html_str=replace(html_str,"//","/")
	  loop
	  Do While Right(html_str,1)="/"
	  	html_str=Left(html_str,len(html_str)-1)
	  loop
	  html_str = Mid(html_str,InStr(html_str, "/")+1)
	  html_str = Mid(html_str,InStr(html_str, "/")+1)
	  'pic url
		url_str=http_type & url_parent & "/images/" & html_str & "/" & url_str
		'http://animalcrossingpatterns.booru.org/images/thumbnails/1/499940af57af3241e6bdd67d038dd8c3a2d88782.jpg
		url_str=replace(url_str,"/thumbnails/","/")
		If InStr(url_str,"?")>1 Then url_str=mid(url_str,1,InStr(url_str,"?")-1)
		pic_type=Mid(url_str,instrrev(url_str,"."))
		
		'pic name
	  split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "title=""")+len("title="""))
	  split_str(split_i) = Trim(Mid(split_str(split_i), 1, InStr(split_str(split_i), chr(34))-1))
	  split_str(split_i) = Trim(Mid(split_str(split_i), 1, InStr(split_str(split_i), "  ")-1))
	  If Len(split_str(split_i))>180 Then split_str(split_i)=Left(split_str(split_i),179) & "~"
	  split_str(split_i) = replace(split_str(split_i),"|","_")
	  split_str(split_i) = "(" & Specific_web_name(url_parent) & ") " & sid & " - " & split_str(split_i) & pic_type
		split_str(split_i) = "|" & url_str & "|" & split_str(split_i) & "|"
  Next
  return_download_list=join(split_str,vbCrLf) & vbCrLf
  If Next_page>0 Then
  	return_download_list=return_download_list & "1|inet|10,13|" & http_type & url_parent & "/index.php?page=post&s=list" & tags & "&pid=" & Next_page
	End If
ElseIf retry_counter<3 Then
	retry_counter=retry_counter+1
  return_download_list="1|inet|10,13|" & http_type & url_parent & "/index.php?page=post&s=list" & tags & "&pid=" & Next_page
Else
return_download_list = "0"
End If
End Function


'--------------------------------------------------------
Function Specific_web_site(ByVal web_site_url)
Specific_web_site=0
If InStr(LCase(web_site_url),"www.allthefallen.org")=1 Then
	XML_TF=0
	Specific_web_site=1
End If
End Function

Function Specific_web_name(ByVal web_site_name)
Specific_web_name=web_site_name
If InStr(LCase(web_site_name),"www.allthefallen.org")=1 Then
	Specific_web_name="allthefallen.org"
End If
If InStr(LCase(web_site_name),"www.")=1 and len(web_site_name)>4 Then
	web_site_name=mid(web_site_name,4)
End If
End Function