'2009-8-15 http://163.shanhaijing.net/
Dim album_ID,page,split_str
Function return_download_url(ByVal url_str)
On Error Resume Next
return_download_url = ""

'http://www.yupoo.com/photos/view?id=ff80808122babdae0122efb0c793301d&album=ff80808122babeff0122efb08c6e13a3
'http://www.yupoo.com/albums/view?id=ff8080811579ca7b011589fd691c27e0
'http://www.yupoo.com/photos/view?id=ff8080811579ca7b01158a00b11e29be&album=ff8080811579ca7b011589fd691c27e0
'http://www.yupoo.com/photos/slideshow/?album_id=ff80808111f9f35b0112003a7a035dd0
'http://moonpie.yupoo.com/albums/
'http://sanhuadidas.yupoo.com/

if instr(lcase(url_str),"http://www.yupoo.com/")>0 Then
	if instr(lcase(url_str),"http://www.yupoo.com/photos/view?")>0 Then
		url_str=Mid(url_str,instr(lcase(url_str),"album=")+6)
		if instr(url_str,"&")>1 Then url_str=Mid(url_str,1,instr(url_str,"&")-1)
		if instr(url_str,"#")>1 Then url_str=Mid(url_str,1,instr(url_str,"#")-1)
		if instr(url_str,"?")>1 Then url_str=Mid(url_str,1,instr(url_str,"?")-1)
	ElseIf instr(lcase(url_str),"http://www.yupoo.com/photos/slideshow/")>0 Then
		url_str=Mid(url_str,instr(lcase(url_str),"album_id=")+9)
		if instr(url_str,"&")>1 Then url_str=Mid(url_str,1,instr(url_str,"&")-1)
		if instr(url_str,"#")>1 Then url_str=Mid(url_str,1,instr(url_str,"#")-1)
		if instr(url_str,"?")>1 Then url_str=Mid(url_str,1,instr(url_str,"?")-1)
	ElseIf instr(lcase(url_str),"http://www.yupoo.com/albums/view?")>0 Then
		url_str=Mid(url_str,instr(lcase(url_str),"id=")+3)
		if instr(url_str,"&")>1 Then url_str=Mid(url_str,1,instr(url_str,"&")-1)
		if instr(url_str,"#")>1 Then url_str=Mid(url_str,1,instr(url_str,"#")-1)
		if instr(url_str,"?")>1 Then url_str=Mid(url_str,1,instr(url_str,"?")-1)
	Else
		url_str=""
	End If
	album_ID=url_str
	page=0
	If url_str<>"" Then return_download_url = "inet|10,13|http://www.yupoo.com/albums/view?id=" & url_str & "|http://www.yupoo.com/"

Else
url_str=Mid(url_str,instr(lcase(url_str),"http://")+7)
url_str=Mid(url_str,1,instr(lcase(url_str),".yupoo.com")-1)
return_download_url = "inet|10,13|http://" & url_str & ".yupoo.com/albums/|http://www.yupoo.com/"
End If

OX163_urlpage_Referer="Referer: http://www.yupoo.com/"
End Function
'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list = ""
If InStr(lcase(html_str), "<a class=""seta"" href=""") > 0 Then

	html_str = Mid(html_str, InStr(lcase(html_str), "<a class=""seta"" href=""")+22)

	Dim str_split
	str_split=split(html_str,"<a class=""Seta"" href=""")

	For i=0 to UBound(str_split)
		'url
		url_str=Mid(str_split(i),1,InStr(str_split(i),Chr(34))-1)
		
		'name
		str_split(i)=Mid(str_split(i),InStr(lcase(str_split(i)),"title=""")+7)
		html_str=str_split(i)
		str_split(i)=replace(rename_utf8(Mid(str_split(i),1,InStr(str_split(i),Chr(34))-1)),"|","_")
		If str_split(i)="" Then str_split(i)="No_Name_Album"
		
		'number
		html_str=Mid(html_str,InStr(lcase(html_str),"<b>")+3)
		html_str=Mid(html_str,1,InStr(lcase(html_str),"</b>")-1)
		If IsNumeric(html_str)=false Then html_str=""
		
	        return_albums_list = return_albums_list & "0|" & html_str & "|" & url_str & "|" & str_split(i) & "|" & vbcrlf
	Next
	return_albums_list = return_albums_list & "0"
End If
End Function
'----------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next

return_download_list=""

If page=0 Then
page=1
	'<script type="text/javascript">
	'//<![CDATA[
	'	var app_domain = 'yupoo.com';
	'	var media_root = 'http://www.yupoo.com';
	'	var auth_hash = '3EA4D24E-966F-DCA2-DF33-E4156EBAB633';
	'	var api_key = 'a55e3b68ea39f949d487193e4b4610b6';
	'	var api_secret = 'prz6fypcj9xftakf';
	'	var api_endpoint = '/api/rest/';
	'	var api_response_format = 'json';
	'	var current_user = null;
	'//]]>
	'</script>
'http://www.yupoo.com/api/rest/;jsessionid=EC0DE89D-141C-7713-C327-AF59D7A0E066?ypp=1&owner_detail=1&api_key=e8d6919bfe826b56622e6ace69501a35&method=yupoo.photos.search&album_id=ff80808122babeff0122efb08c6e13a3&page=1&per_page=108
'http://www.yupoo.com/api/rest/;jsessionid=[auth_hash]?ypp=1&owner_detail=1&api_key=[api_key]&method=yupoo.photos.search&album_id=[album_ID]&page=1&per_page=108
	If InStr(lcase(html_str), "var auth_hash = '") > 0 Then
		url_str=html_str
		'auth_hash
		html_str=Mid(html_str,InStr(lcase(html_str), "var auth_hash = '")+17)
		html_str=Mid(html_str,1,InStr(lcase(html_str), "'")-1)
		'api_key
		url_str=Mid(url_str,InStr(lcase(url_str), "var api_key = '")+15)
		url_str=Mid(url_str,1,InStr(lcase(url_str), "'")-1)		
		return_download_list ="1|inet|10,13|http://www.yupoo.com/api/rest/;jsessionid=" & html_str & "?ypp=1&owner_detail=1&api_key=" & url_str & "&method=yupoo.photos.search&album_id=" & album_ID & "&page=1&per_page=9999"
	End If
ElseIf page=1 Then
	'<photo id="ff80808122babdae0122efb0c793301d" owner="ff8080811dbe7500011dc24acaa63df8" ownername="sanhuadidas" ownericon="http://ico.yupoo.com/sanhuadidas/952367117f7f/" title="SMS °×·Ûºì" status="5" host="0" dir="sanhuadidas" filename="494037de42f3"/>
	If InStr(html_str, "<photo id=""") > 0 Then
		page=2
		html_str=Mid(html_str,InStr(html_str, "<photo id=""")+11)
		split_str=split(html_str,"<photo id=""")
		For i=0 to UBound(split_str)
			'id
			url_str=Mid(split_str(i),1,InStr(split_str(i), Chr(34))-1)
			'name
			html_str=Mid(split_str(i),InStr(LCase(split_str(i)), "title=""")+7)
			html_str=Mid(html_str,1,InStr(html_str, Chr(34))-1)
			'host
			split_str(i)=Mid(split_str(i),InStr(LCase(split_str(i)), "host=""")+6)
			split_str(i)=Mid(split_str(i),1,InStr(split_str(i), Chr(34))-1)
			split_str(i)=url_str & vbcrlf & split_str(i) & vbcrlf & html_str
		Next
		url_str=Mid(split_str(0),1,InStr(split_str(0), vbcrlf)-1)
		return_download_list ="1|inet|10,13|http://www.yupoo.com/photos/view?id=" & url_str & "&album=" & album_ID
	End If
ElseIf page>=2 Then
	Randomize
	'http://www.yupoo.com/photos/view?id=ff8080811579ca7b01158a00b11e29be&album=ff8080811579ca7b011589fd691c27e0
	If InStr(LCase(html_str), "var photo = {id: '") > 0 Then
		html_str=Mid(html_str,InStr(LCase(html_str), "var photo = {id: '"))
		html_str=Mid(html_str,InStr(LCase(html_str), "source:["))
		html_str=Mid(html_str,1,InStr(html_str, "]"))
		url_str=Mid(html_str,InStrrev(LCase(html_str), "{title:'"))
		'url
		url_str=Mid(url_str,InStr(LCase(url_str), "src:'")+5)
		url_str=Mid(url_str,1,InStr(url_str, "'")-1) & "?" & CDbl(Rnd)
		url_str=Mid(url_str,InStr(LCase(url_str), "."))
		'name
		split_str(page-2)=Mid(split_str(page-2),InStr(split_str(page-2), vbcrlf)+Len(vbcrlf))
		html_str=Mid(split_str(page-2),InStr(split_str(page-2), vbcrlf)+Len(vbcrlf))		
		html_str=rename_utf8(Mid(html_str,1,InStr(html_str, Chr(34))-1))
		If html_str="" Then html_str="noname_pic"
		'host
		split_str(page-2)=Mid(split_str(page-2),1,InStr(split_str(page-2), vbcrlf)-1)
		
		url_str="http://photo" & split_str(page-2) & url_str
		
		return_download_list ="jpg|" & url_str & "|" & html_str & "|" & vbcrlf
	End If
	
	If page-2<UBound(split_str) Then
		page=page+1
		url_str=Mid(split_str(page-2),1,InStr(split_str(page-2), vbcrlf)-1)
		return_download_list =return_download_list & "1|inet|10,13|http://www.yupoo.com/photos/view?id=" & url_str & "&album=" & album_ID
	End If

End If

End Function
'------------------------------------------------------
Function rename_utf8(byval utf8_Str)
If Len(utf8_Str)=0 Then Exit Function
For i=1 to Len(utf8_Str)
	If  Asc(Mid(utf8_Str,i,1))=63 Then utf8_Str=replace(utf8_Str,Mid(utf8_Str,i,1),"_")
Next
rename_utf8=utf8_Str
End Function