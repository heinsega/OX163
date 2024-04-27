'2024-4-27 163.shanhaijing.net
Dim http_type, url_parent, tags, page, page_counter, Next_page, retry_counter

Function return_download_url(ByVal url_str)
'https://rule34.paheal.net/post/view/6272998#search=A_Certain_Scientific_Railgun
'https://rule34.paheal.net/post/list
'https://rule34.paheal.net/post/list/2
'https://rule34.paheal.net/post/list/Ei_Raiden/1
'https://rule34.paheal.net/api/danbooru/find_posts?tags=Ei_Raiden&limit=100
'https://rule34.paheal.net/api/danbooru/find_posts?tags=Ei_Raiden&limit=100&pid=(0ÊÇµÚÒ»Ò³)
On Error Resume Next

return_download_url=""
url_parent=url_str

If InStr(LCase(url_str), "/post/list") >1 Then
	'tags
	tags=mid(url_str,InStr(LCase(url_str), "/post/list")+10)
	If InStr(tags, "/") = 1 Then
		tags=mid(tags,2)
		If InStr(tags, "/") > 1 Then			
			tags=Mid(tags,1,InStr(tags, "/")-1)
		Else
			tags=""
		End If
	Else
		tags=""
	End If
	Trim(tags)=""
	page=0
	page_counter=0
	retry_counter=0
	url_parent=url_str
	Next_page="https://rule34.paheal.net/api/danbooru/find_posts?limit=100&tags=" & tags & "&pid="
	return_download_url = "inet|10,13|" & Next_page & page & "|" & url_str
End If

return_download_url=return_download_url & vbcrlf & "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko"
OX163_urlpage_Referer=url_str & vbcrlf & "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko"

End Function

'--------------------------------------------------------

Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next	
Dim split_str, sid, pic_type
'<posts count='8791' offset='0'>
'<tag id='6298741' md5='8fb817fa80583ef344f5b59b2f6e8b29' file_name='117875758_p3.jpg' file_url='https://r34i.paheal-cdn.net/8f/b8/8fb817fa80583ef344f5b59b2f6e8b29' height='1200' width='1546' preview_url='/_thumbs/8fb817fa80583ef344f5b59b2f6e8b29/thumb.jpg' preview_height='149' preview_width='192' rating='?' date='2024-04-26 20:10:39' tags='acheron crossover Ei_Raiden Genshin_Impact Honkai:_Star_Rail saya_pixiv(16679661)' source='https://www.pixiv.net/en/artworks/117875758' score='0' author='y39x'>
'</tag>
'...
If InStr(LCase(html_str), "<posts count='") > 0 and InStr(LCase(html_str), "<tag id='") > 0 Then
	retry_counter=0
	Dim key_word
	key_word="<posts count='"
	url_str=Mid(html_str, InStr(LCase(html_str), key_word) + len(key_word))
	url_str=Mid(url_str,1,InStr(url_str, "'")-1)
	If IsNumeric(url_str) Then page_counter=Int(url_str)

	key_word="<tag id='"
	html_str = Mid(html_str, InStr(LCase(html_str), key_word) + len(key_word))
	split_str = Split(html_str, key_word)
  For split_i = 0 To UBound(split_str)
		html_str=""
		url_str=""
		sid=""
		pic_type=""
				
  	'sid
	  sid = Mid(split_str(split_i),1,InStr(split_str(split_i), "'")-1)

	  'file_url
  	key_word="' file_url='"
	  url_str = Mid(split_str(split_i), InStr(split_str(split_i), key_word) + len(key_word))
	  url_str = Mid(url_str,1,InStr(url_str, "'")-1)	
		If left(url_str,2)="//" Then url_str="https://" & Mid(url_str,3)
	  
	  'pic_type
  	key_word="' file_name='"
	  pic_type = Mid(split_str(split_i), InStr(split_str(split_i), key_word) + len(key_word))
	  pic_type = Mid(pic_type,1,InStr(pic_type, "'")-1)	
		pic_type = Mid(pic_type,instrrev(pic_type,"."))
		
		
		'pic name
  	key_word="' tags='"
	  html_str = Mid(split_str(split_i), InStr(split_str(split_i), key_word) + len(key_word))
	  html_str = Mid(html_str,1,InStr(html_str, "'")-1)
	  
	  If Len(html_str)>180 Then html_str=Left(html_str,179) & "~"
	  html_str = replace(html_str,"|","_")
	  html_str = replace(html_str,":","_")
	  html_str = "(rule34.paheal) " & sid & " - " & html_str & pic_type
	  
		split_str(split_i) = "|" & url_str & "|" & html_str & "|"
  Next
  
  return_download_list=join(split_str,vbCrLf) & vbCrLf
  If (page+1)*100<page_counter Then
  	page=page+1
  	return_download_list=return_download_list & "1|inet|10,13|" & Next_page & page
	End If
ElseIf retry_counter<3 Then
	retry_counter=retry_counter+1
	return_download_list="1|inet|10,13|" & Next_page & page
Else
	return_download_list = "0"
End If

End Function