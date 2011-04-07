'2011-1-6 http://www.shanhaijing.net/163
Function return_download_url(ByVal url_str)
On Error Resume Next
return_download_url = ""

'_user.albums =[{id:'
'.com->.ca,.it,.fr,.de,.hk¡­¡­¡­¡­¡­¡­
'http://picasaweb.google.com/takabe.eri
'http://picasaweb.google.com/takabe.eri/
'http://picasaweb.google.com/takabe.eri/HxfMqJ
'http://picasaweb.google.com/takabe.eri/HxfMqJ#
'http://picasaweb.google.com/takabe.eri/HxfMqJ#5293250776970480754
'https://
Dim url_str_split,photo_tf

url_str_split=split(mid(url_str,instr(lcase(url_str),"picasaweb.google.")+Len("picasaweb.google.")),"/")

photo_tf=0

If InStr(url_str_split(1),"?")>6 Then url_str_split(1)=Mid(url_str_split(1),1,InStr(url_str_split(1),"?")-1)
	
If Is_username(url_str_split(1))=0 and url_str_split(1)<>"home" then Exit Function

If ubound(url_str_split) > 1 Then
	if instr(url_str_split(2),"#")>0 then url_str_split(2)=mid(url_str_split(2),1,instr(url_str_split(2),"#")-1)
	if len(url_str_split(2))>0  then photo_tf=1
end if

If photo_tf=1 Then
	return_download_url = "inet|10,13|https://picasaweb.google." & url_str_split(0) & "/" & url_str_split(1) & "/" & url_str_split(2)
Else
	return_download_url = "inet|10,13|https://picasaweb.google." & url_str_split(0) & "/" & url_str_split(1) & "/?showall=true"
End If

End Function
'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list = ""

If InStr(lcase(html_str), "_user.albums =[{id:'") > 0 Then

	html_str = Mid(html_str, InStr(lcase(html_str), "_user.albums =[{id:'")+20)
	html_str = Mid(html_str, 1, InStr(lcase(html_str), "</script>")-1)

	Dim str_split
	str_split=split(html_str,"},{id:'")

	For i=0 to UBound(str_split)
		str_split(i)=Mid(str_split(i),InStr(lcase(str_split(i)),"'},title:'")+Len("'},title:'"))
		'name
		html_str=rename_utf8(Trim(Mid(str_split(i),1,InStr(str_split(i),"'")-1)))
		html_str=vbsUnEscape(html_str)
		
		If html_str="" Then html_str="No_Name_Album"
		
		str_split(i)=Mid(str_split(i),InStr(lcase(str_split(i)),"',url:'")+7)
		'url
		url_str=replace(Mid(str_split(i),1,InStr(str_split(i),"'")-1),"\x2F","/")
		
		str_split(i)=Mid(str_split(i),InStr(lcase(str_split(i)),"',count:")+8)
		'pic number
		str_split(i)=Mid(str_split(i),1,InStr(str_split(i),",")-1)
		If IsNumeric(str_split(i))=false Then str_split(i)=""
		
	        return_albums_list = return_albums_list & "0|" & str_split(i) & "|" & url_str & "|" & html_str & vbcrlf
	Next
	return_albums_list = return_albums_list & "0"
End If
End Function
'----------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list=""

If InStr(lcase(html_str), "entry"":[{""gd$kind"":""photos#photo"",""id"":""") > 0 Then
	html_str = Mid(html_str, InStr(lcase(html_str), "entry"":[{""gd$kind"":""photos#photo"",""id"":""")+Len("entry"":[{""gd$kind"":""photos#photo"",""id"":"""))
	html_str = Mid(html_str, 1, InStr(lcase(html_str), "</script>")-1)

	Dim str_split,pic_type
	str_split=split(html_str,"},{""gd$kind"":""photos#photo"",""id"":""")

	For i=0 to UBound(str_split)
		str_split(i)=Mid(str_split(i),InStr(lcase(str_split(i)),",""title"":""")+10)
		'name
		html_str=rename_utf8(Trim(Mid(str_split(i),1,InStr(str_split(i),chr(34))-1)))

		url_str=Mid(str_split(i),InStr(lcase(str_split(i)),",""description"":""")+16)
		'summary
		url_str=rename_utf8(Trim(Mid(url_str,1,InStr(url_str,chr(34))-1)))
		url_str=vbsUnEscape(url_str)
				
		str_split(i)=Mid(str_split(i),InStr(lcase(str_split(i)),"{""content"":[{""url"":""")+20)
		'url
		str_split(i)=Mid(str_split(i),1,InStr(str_split(i),chr(34))-1)

		If html_str="" Then html_str=Mid(str_split(i),InStrrev(str_split(i),"/")+1)
		pic_type=lcase(Mid(str_split(i),InStrrev(str_split(i),".")))
		If lcase(Mid(html_str,InStrrev(html_str,".")))<>pic_type Then html_str=html_str & pic_type		
		
		return_download_list =return_download_list & "|" & str_split(i) & "|" & html_str & "|" & url_str & vbcrlf
	Next

end if
End Function
'------------------------------------------------------
Function Is_username(byval user_name)
Is_username=0
If Len(user_name)<6 Then Exit Function
For i=1 to Len(user_name)
	If  instr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.0123456789",mid(user_name,i,1))<1 Then Exit Function
Next
Is_username=1
End Function

Function rename_utf8(byval utf8_Str)
If Len(utf8_Str)=0 Then Exit Function
For i=1 to Len(utf8_Str)
	If  Asc(Mid(utf8_Str,i,1))=63 Then utf8_Str=replace(utf8_Str,Mid(utf8_Str,i,1),"_")
Next
rename_utf8=utf8_Str
End Function

Function vbsUnEscape(str) 
On Error Resume Next
    dim i,s,c
    s="" 
    For i=1 to Len(str) 
    DoEvents
        c=Mid(str,i,1) 
        If Mid(str,i,2)="\x" and i<=Len(str)-5 Then 
            If IsNumeric("&H" & Mid(str,i+2,2)) Then 
                s = s & CHRW(CInt("&H" & Mid(str,i+2,2))) 
                i = i+5 
            Else 
                s = s & c 
            End If
        Else 
            s = s & c 
        End If 
    Next 
    vbsUnEscape = replace(s,"\/","/") 
End Function