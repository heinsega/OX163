'2012-11-8 163.shanhaijing.net
Dim page,tags,url_instr,page_retry,delay_tf,start_time,Next_page

Function return_download_url(ByVal url_str)
'http://gelbooru.com/index.php?page=post&s=list&pid=336960
'http://gelbooru.com/index.php?page=post&s=list&tags=all
'http://gelbooru.com/index.php?page=post&s=list
'http://gelbooru.com/index.php?page=post&s=list&tags=nakoruru
'http://gelbooru.com/index.php?page=post&s=list&tags=nakoruru&pid=60
'http://www.gelbooru.com/index.php?page=post&s=list&tags=canaan
'youhate.us www.youhate.us
On Error Resume Next
tags=""
Dim page_tmp
page_tmp=url_str
page_retry=0
page=0
delay_tf=0

'-----------------pool--------------------
If InStr(LCase(url_str), "page=pool") >1 Then
	Next_page="pool"
	return_download_url = "inet|10,13|" & url_str & "|http://gelbooru.com/"
	url_instr=url_str
End If

'-----------------tags--------------------
If InStr(LCase(url_str), "tags=") > 0 Then
    tags = Mid(url_str, InStr(url_str, "tags=") + 5)
    If InStr(tags, "&") > 0 Then tags = Mid(tags, 1, InStr(tags, "&") - 1)
    If InStr(tags, " ") > 0 Then tags = Mid(tags, 1, InStr(tags, " ") - 1)
    If LCase(tags)="all" Then tags=""
End If
If tags <> "" Then url_str = "http://gelbooru.com/index.php?page=post&s=list&tags=" & tags

If InStr(LCase(page_tmp),"&pid=")>len("gelbooru.com/") Or InStr(LCase(page_tmp),"?pid=")>len("gelbooru.com/") Then
	If InStr(LCase(page_tmp),"&pid=")>len("gelbooru.com/") Then page_tmp=Mid(page_tmp,InStr(LCase(page_tmp),"&pid=")+5)
	If InStr(LCase(page_tmp),"?pid=")>len("gelbooru.com/") Then page_tmp=Mid(page_tmp,InStr(LCase(page_tmp),"?pid=")+5)
	If InStr(LCase(page_tmp),"&")>0 Then page_tmp=Mid(page_tmp,1,InStr(page_tmp,"&")-1)
	If IsNumeric(page_tmp) Then
		If Int(page_tmp)>1 Then
			If MsgBox("本页为第" & Int(page_tmp/28)+1 & "页" & vbcrlf & "是否从第1页开始？", vbYesNo, "问题")=vbyes Then
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
url_instr=url_str
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

Dim split_str,url_temp,split_i,sid,pic_type
return_download_list = ""

If delay_tf=1 and DateDiff("s", start_time, Now()) < 12 Then
	return_download_list="1|inet|10,13|http://www.163.com/?Delay_15s-利用163页面延迟15秒"
	Exit Function
ElseIf delay_tf=1 Then
	delay_tf=2
	return_download_list = "1|inet|10,13|http://gelbooru.com/intermission.php"'skip ads
	Exit Function
ElseIf delay_tf>0 Then
	delay_tf=0
	return_download_list = "1|inet|10,13|" & url_instr
	Exit Function
End If

url_str=html_str

If Next_page="pool" Then
	'该部分代码复制b-sys.vbs，仅调整了文件命名方式和url格式
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
	  MsgBox url_str
    url_str = replace(replace(url_str,"/thumbnails/","/images/"),"thumbnail_","")
    url_str = replace(url_str,"/thumbs/","/images/")
    url_str = replace(url_str,"http://youhate.us/","http://gelbooru.com/")
    url_str = replace(url_str,"http://www.youhate.us/","http://gelbooru.com/")
	  MsgBox url_str
		If InStr(url_str,"?")>1 Then url_str=mid(url_str,1,InStr(url_str,"?")-1)
		pic_type=Mid(url_str,instrrev(url_str,"."))
		
		'pic name
	  split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "title=""")+len("title="""))
	  split_str(split_i) = Trim(Mid(split_str(split_i), 1, InStr(split_str(split_i), chr(34))-1))
	  split_str(split_i) = Trim(Mid(split_str(split_i), 1, InStr(split_str(split_i), "  ")-1))
	  If Len(split_str(split_i))>180 Then split_str(split_i)=Left(split_str(split_i),179) & "~"
    split_str(split_i) = replace(split_str(split_i),"|"," ")
	  split_str(split_i) = sid & "_" & split_str(split_i) & pic_type
		split_str(split_i) = "|" & url_str & "|" & split_str(split_i) & "|"
  Next  
  return_download_list=join(split_str,vbCrLf) & vbCrLf
ElseIf InStr(html_str, " class=""thumb"">") > 0 Then
	
	html_str = Mid(html_str, InStr(html_str, " class=""thumb"">") + 15)
	html_str = Mid(html_str, InStr(html_str, "<a id=""") + 7)

	split_str = Split(html_str, " class=""thumb""><a id=""")

    For split_i = 0 To UBound(split_str)
    html_str=Mid(split_str(split_i),1, InStr(split_str(split_i), Chr(34)) -1) & "_" 'p371667_
    split_str(split_i) = Mid(split_str(split_i), InStr(split_str(split_i), "<img src=""") +10)
    
    'url
    url_temp = Mid(split_str(split_i), 1,InStr(split_str(split_i), "?") -1)
    url_temp = replace(replace(url_temp,"/thumbnails/","/images/"),"thumbnail_","")
    url_temp = replace(url_temp,"/thumbs/","/images/")
    url_temp = replace(url_temp,"http://youhate.us/","http://gelbooru.com/")
    url_temp = replace(url_temp,"http://www.youhate.us/","http://gelbooru.com/")
    url_temp = replace(url_temp,"http://cdn2.","http://cdn1.")
    
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

	Next_page=0
	delay_tf=0
	
	If InStr(LCase(url_str), "alt=""next"">") > 0 Then
		url_str=Mid(url_str,1,InStr(LCase(url_str), "alt=""next"">")-1)
		url_str=Mid(url_str, InStrrev(LCase(url_str), "pid=")+4)
		url_str=Mid(url_str,1,InStr(url_str, chr(34))-1)
		If IsNumeric(url_str) Then Next_page=Int(url_str)
	End If
		
	If Next_page>0 Then
		page_retry=0
		url_instr="http://gelbooru.com/index.php?page=post&s=list&tags=" & tags & "&pid=" & Next_page
		return_download_list = return_download_list & "1|inet|10,13|" & url_instr
	Else
		return_download_list = return_download_list & "0"
	End If
ElseIf page_retry<5 Then
	page_retry=page_retry+1
	delay_tf=1
	start_time=Now()
	return_download_list="1|inet|10,13|http://www.163.com/?Delay_15s-利用163页面延迟15秒"
Else
	return_download_list = return_download_list & "0"
End If
End Function