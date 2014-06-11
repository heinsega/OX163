Dim uid,aid,pid

Function return_download_url(ByVal url_str)
'http://bbs.3dmgame.com/home.php?mod=space&uid=2499190&do=album&view=me&from=space&page=3
'http://bbs.3dmgame.com/space-uid-2499190.html
'http://bbs.3dmgame.com/home.php?mod=space&uid=2499190&do=album&id=1243
return_download_url=""
uid=""
If instr(LCase(url_str),"&uid=")>0 Then
	uid=Mid(url_str,InStr(LCase(url_str),"&uid=")+5)
	If instr(uid,"&")>0 Then uid=Mid(uid,1,InStr(uid,"&")-1)
ElseIf InStr(LCase(url_str),"-uid-")>0 Then
	uid=Mid(url_str,InStr(LCase(url_str),"-uid-")+5)
	If instr(uid,"&")>0 Then uid=Mid(uid,1,InStr(uid,".")-1)
End If

If uid<>"" Then
	If instr(LCase(url_str),"&id=")>0 Then
			'photo
			aid=Mid(url_str,InStr(LCase(url_str),"&id=")+4)
			If instr(aid,"&")>0 Then aid=Mid(aid,1,InStr(aid,"&")-1)
			pid=1
			return_download_url="inet|10,13|http://bbs.3dmgame.com/home.php?mod=space&uid="&uid&"&do=album&id="&aid
		Else
			'album
			pid=1
			return_download_url="inet|10,13|http://bbs.3dmgame.com/home.php?mod=space&uid="&uid&"&do=album&view=me&from=space"
	End If		
End If
End Function
'--------------------------------------------------------------

Function return_albums_list(ByVal html_str, ByVal url_str)
return_albums_list=""
Dim key_word, next_page, split_str, A_link, A_name, A_count

'判断下一页
next_page=0
If InStr(html_str,">下一页<")>0 Then
	pid=pid+1
	next_page="1|inet|10,13|http://bbs.3dmgame.com/home.php?mod=space&uid="&uid&"&do=album&view=me&from=space&page="&pid
End If

key_word="<p class=""ptn"">"
If InStr(html_str,key_word)>0 Then
	html_str=Mid(html_str,InStr(html_str,key_word)+len(key_word))
	split_str=split(html_str,key_word)
	For i=o to ubound(split_str)
		'<a href="home.php?mod=space&amp;uid=2499190&amp;do=album&amp;id=23221" target="_blank"  class="xi2">亚莉克希亚新扮相</a> (13) </p>
		A_link=""
		A_name=""
		A_count=""
		
		'====>home.php?mod=space&amp;uid=2499190&amp;do=album&amp;id=23221
		A_link=mid(split_str(i),InStr(split_str(i),chr(34))+1)
		A_link=mid(A_link,1,InStr(A_link,chr(34))-1)
		A_link=replace(A_link,"&amp;","&")
		
		'====>亚莉克希亚新扮相
		A_name=mid(split_str(i),InStr(split_str(i),">")+1)
		A_count=A_name
		A_name=mid(A_name,1,InStr(A_name,"</a>")-1)
		If instr(A_name,"|")>0 Then A_name=replace(A_name,"|","_")
			
		'====>13
		A_count=mid(A_count,InStr(A_count,"(")+1)
		A_count=mid(A_count,1,InStr(A_count,")")-1)
		If not IsNumeric(A_count) Then A_count=""
		split_str(i)="0|"&A_count&"|http://bbs.3dmgame.com/"&A_link&"|"&A_name&"|"&vbCrLf
	Next
	return_albums_list=join(split_str,"")
End If

return_albums_list=return_albums_list & next_page

End Function
'--------------------------------------------------------------

Function return_download_list(ByVal html_str, ByVal url_str)
return_download_list = ""
Dim key_word, next_page, split_str, A_link, A_name, A_count

'判断下一页
next_page = 0
If InStr(html_str, ">下一页<") > 0 Then
    pid = pid + 1
    next_page = "1|inet|10,13|http://bbs.3dmgame.com/home.php?mod=space&uid=" & uid & "&do=album&id=" & aid & "&page=" & pid
End If

key_word = "<ul class=""ptw ml mlp cl"">"
If InStr(html_str, key_word) > 0 Then
    html_str = Mid(html_str, InStr(html_str, key_word) + Len(key_word))
    html_str = Mid(html_str, 1, InStr(html_str, "</ul>"))
    key_word = "<img src="""
    html_str = Mid(html_str, InStr(html_str, key_word) + Len(key_word))
    split_str = Split(html_str, key_word)
    key_word = ".thumb."
    For i = 0 To UBound(split_str)
        split_str(i) = Mid(split_str(i), 1, InStr(split_str(i),chr(34))-1)
        If InStr(split_str(i), key_word) > 0 Then
            split_str(i) = Mid(split_str(i), 1, InStr(split_str(i), key_word) - 1)
        End If
        split_str(i) = "|" & split_str(i) & "||" & vbCrLf
    Next
    return_download_list = Join(split_str, "")
End If

return_download_list = return_download_list & next_page

End Function