'2014-2-22 http://www.shanhaijing.net/163
Dim sohu_ID,page_num
Function return_download_url(ByVal url_str)
On Error Resume Next
Dim Timestr
return_download_url = ""
'sohuÏà²áalbum
'http://pp.sohu.com/u/793927
'http://pp.sohu.com/u/793927/w6AGNSYzfVD
If instr(lcase(url_str),"http://pp.sohu.com/u/")=1 Then
	url_str=mid(url_str,len("http://pp.sohu.com/u/")+1)
	If instr(url_str,"/w")>0 Then
		url_str="http://pp.sohu.com/u/" & url_str
	Else
		sohu_ID=url_str
		page_num=0
		Timestr=Timer()
		Timestr=replace(Timestr,".","")
		If instr(sohu_ID,"?")>0 Then sohu_ID=mid(sohu_ID,1,instr(sohu_ID,"?")-1)
		If instr(sohu_ID,"#")>0 Then sohu_ID=mid(sohu_ID,1,instr(sohu_ID,"#")-1)
		If isnumeric(sohu_ID) Then
			url_str="http://pp.sohu.com/u/" & sohu_ID & "?" & Timestr
		Else
			url_str=""
			Exit Function
		End If
	End If
	return_download_url = "inet|10,13|" & url_str & "|http://pp.sohu.com/"
	OX163_urlpage_Referer = "X-Requested-With: XMLHttpRequest" & vbCrLf & _
	"Accept: application/json, text/javascript, */*; q=0.01" & vbCrLf & _
	"Referer: http://pp.sohu.com/" & vbCrLf & _
	"User-Agent: Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E)"
End If

End Function

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim key_word,split_str,split_str_i
Dim pp_showId,pp_description,pp_Name,pp_photoNum

return_albums_list=""
key_word="{""showId"":"""
If instr(html_str,key_word)>0 Then
	html_str=mid(html_str,instr(html_str,key_word)+len(key_word))
	split_str=split(html_str,key_word)
	For split_str_i=0 to ubound(split_str)
		pp_showId=""
		pp_Name=""
		pp_description=""
		pp_photoNum=""
		
		pp_showId=Mid(split_str(split_str_i),1,instr(split_str(split_str_i),chr(34))-1)
		
		key_word="""name"":"""
		pp_Name=Mid(split_str(split_str_i),instr(split_str(split_str_i),key_word)+len(key_word))
		pp_Name=Mid(pp_Name,1,instr(pp_Name,chr(34))-1)
		
		key_word="""description"":"""
		pp_description=Mid(split_str(split_str_i),instr(split_str(split_str_i),key_word)+len(key_word))
		pp_description=Mid(pp_description,1,instr(pp_description,chr(34))-1)
		
		key_word="""photoNum"":"
		pp_photoNum=Mid(split_str(split_str_i),instr(split_str(split_str_i),key_word)+len(key_word))
		pp_photoNum=Mid(pp_photoNum,1,instr(pp_photoNum,",")-1)
		If IsNumeric(pp_photoNum)=False Then pp_photoNum=""
		
		'http://pp.sohu.com/u/793927/w6AGNSYzfVD
		split_str(split_str_i)="0|" & pp_photoNum & "|http://pp.sohu.com/u/" & sohu_ID & "/w" & pp_showId & "|"& replace(pp_Name,"|","_") & "(AID_" & pp_showId & ")|" & pp_description
	Next
	html_str=join(split_str,vbcrlf)
	page_num=page_num+10
	url_str="1|inet|10,13|http://pp.sohu.com/u/" & sohu_ID & "?offset=" & page_num
	return_albums_list=html_str & vbcrlf & url_str
End If
End Function

'----------------------------------------------------------------------------------

Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim key_word,split_str,split_str_i
Dim pp_originalFilename,pp_width,pp_height,pp_originUrl

return_download_list=""
key_word="{""creatorId"":"
If instr(html_str,key_word)>0 Then
	html_str=mid(html_str,instr(html_str,key_word)+len(key_word))
	split_str=split(html_str,key_word)
	For split_str_i=0 to ubound(split_str)
		pp_originalFilename=""
		pp_width=""
		pp_height=""
		pp_originUrl=""
		
		key_word="""originalFilename"":"""
		pp_originalFilename=Mid(split_str(split_str_i),instr(split_str(split_str_i),key_word)+len(key_word))
		pp_originalFilename=Mid(pp_originalFilename,1,instr(pp_originalFilename,chr(34))-1)
		pp_originalFilename=replace(pp_originalFilename,"|","_")
		
		key_word="""originUrl"":"""
		pp_originUrl=Mid(split_str(split_str_i),instr(split_str(split_str_i),key_word)+len(key_word))
		pp_originUrl=Mid(pp_originUrl,1,instr(pp_originUrl,chr(34))-1)
		If Left(lcase(pp_originUrl),4)<>"http" Then pp_originUrl=""
		
		key_word="""width"":"
		pp_width=Mid(split_str(split_str_i),instr(split_str(split_str_i),key_word)+len(key_word))
		pp_width=Mid(pp_width,1,instr(pp_width,",")-1)
		
		key_word="""height"":"
		pp_height=Mid(split_str(split_str_i),instr(split_str(split_str_i),key_word)+len(key_word))
		pp_height=Mid(pp_height,1,instr(pp_height,",")-1)
		If IsNumeric(pp_width)=False or IsNumeric(pp_height)=False Then
			pp_width=""
		Else
			pp_width=pp_width & " x " & pp_height
		End If
		If pp_originUrl<> "" Then
			split_str(split_str_i)="|" & pp_originUrl & "|" & pp_originalFilename & "|" & pp_width & vbcrlf
		Else
			split_str(split_str_i)=""
		End If
	Next
	return_download_list=join(split_str,"") & "0"
End If
End Function