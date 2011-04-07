'2010-1-2 http://www.shanhaijing.net/163

Function return_download_url(byVal url_str)
return_download_url="inet|10,13,34|" & url_str
End Function

Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next

return_download_list = ""

If InStr(html_str,"mediaUrl':'")>0 Then
	Dim split_str,url
	html_str=Mid(html_str,InStr(html_str,"mediaUrl':'")+11)
	split_str=split(html_str,"mediaUrl':'")
	html_str=""
	For i=0 to UBound(split_str)
		url_str=Mid(split_str(i),1,InStr(split_str(i),"'")-1)
		If html_str<>url_str Then
			html_str=url_str
			return_download_list=return_download_list & "|" & url_str & "|" & rename_utf8(url_str) & "|" & vbCrLf
		End If
	Next
	
	return_download_list=return_download_list & "0"
	
Else
	return_download_list = "0"
End If
End Function

Function rename_utf8(byval utf8_Str)
If Len(utf8_Str)=0 Then Exit Function
For i=1 to Len(utf8_Str)
	If  Asc(Mid(utf8_Str,i,1))=63 Then utf8_Str=replace(utf8_Str,Mid(utf8_Str,i,1),"_")
Next
rename_utf8=utf8_Str
End Function