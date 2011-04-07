'2009-9-4 by visceroid

Dim board_id, root_str, dir_array(3), filename_array(3)
root_str = "http://2cat.twbbs.org/~kirur/touhou/"
dir_array(0) = ""
dir_array(1) = ""
dir_array(2) = "_2/"
dir_array(3) = "th12/"
filename_array(0) = ""
filename_array(1) = "th.htm"
filename_array(2) = "index.htm"
filename_array(3) = "index.htm"

Function return_download_url(ByVal url_str)
'http://2cat.twbbs.org/~kirur/touhou/th.htm
'http://2cat.twbbs.org/~kirur/touhou/1.htm
'http://2cat.twbbs.org/~kirur/touhou/_2/index.htm
'http://2cat.twbbs.org/~kirur/touhou/th12/index.htm
'http://2cat.twbbs.org/~kirur/touhou/th12/futaba.php?res=1
'http://2cat.twbbs.org/~kirur/touhou/th12/Stag.php?key=A
'http://2cat.twbbs.org/~kirur/touhou/_2/src/
On Error Resume Next
	return_download_url = ""
	Dim rglr_str
	get_board_id(url_str)
	If InStr(url_str, "res") > 0 Or InStr(url_str, "key") > 0 Then
		return_download_url = "inet|10,13|" & url_str & "|" & root_str
	ElseIf InStr(url_str, "src") > 0 Then
		return_download_url = "inet|10,13|" & root_str & dir_array(board_id) & "src/"
	Else
		rglr_str = root_str & dir_array(board_id) & filename_array(board_id)
		return_download_url = "inet|10,13|" & rglr_str & "|" & root_str
	End If
End Function

'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
	return_albums_list = ""
	Dim regex, matches, title_str, addr_str, desc_str
	Set regex = New RegExp
	regex.Global = True
	regex.Pattern = "(No.\d+)[^\[]*\[<a href=(\w+\.php\?res=\d+).*?>.*?</a>\].*?<blockquote>(.*?)</blockquote>"
	Set matches = regex.Execute(html_str)
	For Each match in matches
		title_str = match.SubMatches(0)
		addr_str = root_str & dir_array(board_id) & match.SubMatches(1)
		desc_str = post_process(match.SubMatches(2))
		return_albums_list = return_albums_list & "0|0|" & addr_str & "|" & title_str & "|" & desc_str & vbCrLf
	Next
	regex.Pattern = "<form action=""(\d+.htm)"" method=get><td><input type=submit value=""ÏÂÒ»í“""></td></form>"
	Set matches = regex.Execute(html_str)
	For Each match in matches
		addr_str = root_str & dir_array(board_id) & match.SubMatches(0)
		return_albums_list = return_albums_list & "1|inet|10,13|" & addr_str
		Exit Function
	Next
	return_albums_list = return_albums_list & "0"
End Function

'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
	return_download_list = ""
	Dim regex, matches, addr_str, name_str
	Set regex = New RegExp
	regex.Global = True
	If InStr(url_str, "res") > 0 Then
		regex.Pattern = "(?:FileName: |™nÃû£º)<a href=[""']?(src/(\d+\.\w+))[""']?.*?>\2</a>"
	Else
		regex.Pattern = "<a href=[""']?((?:src/)?(\d+\.\w+))[""']?.*?><.*?></a>"
	End If
	Set matches = regex.Execute(html_str)
	For Each match in matches
		If InStr(url_str, "src") > 0 Then
			addr_str = root_str & dir_array(board_id) & "src/" & match.SubMatches(0)
		Else
			addr_str = root_str & dir_array(board_id) & match.SubMatches(0)
		End If
		name_str = match.SubMatches(1)
		return_download_list = return_download_list & "|" & addr_str & "|" & name_str & "|" & vbCrLf
	Next
	regex.Pattern = "</b>\s*<a href='(index\.php\?page.*?)'>\[\d\]</a>"
	Set matches = regex.Execute(html_str)
	For Each match in matches
		addr_str = root_str & dir_array(board_id) & "src/" & match.SubMatches(0)
		return_download_list = return_download_list & "1|inet|10,13|" & addr_str
		Exit Function
	Next
	return_download_list = return_download_list & "0"
End Function

Function get_board_id(ByVal url_str)
	If InStr(url_str, "_2") > 0 Then
		board_id = 2
	ElseIf InStr(url_str, "th12") > 0 Then
		board_id = 3
	Else
		board_id = 1
	End If
End Function

Function post_process(ByVal raw_str)
	post_process = ""
	Dim regex
	Set regex = New RegExp
	regex.Global = True
	regex.Pattern = "<br\s*/>"
	raw_str = regex.Replace(raw_str, " ")
	regex.Pattern = "<.*?>"
	raw_str = regex.Replace(raw_str, "")
	post_process = raw_str
End Function