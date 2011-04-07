'2010-3-2 visceroid
Dim started, multi_page, brief_mode, brief_mode_rf, retries_count, cache_index, root_str, next_page_str, parent_next_page_str, matches_cache
started = False
multi_page = True
retries_count = 0
cache_index = 0
root_str = "http://www.pixiv.net"

Function return_download_url(ByVal url_str)
On Error Resume Next
	return_download_url = ""
	Dim sub_url_str, regex, matches
	Set regex = New RegExp
	regex.Global = True

	brief_mode_rf=""
	If Right(url_str,Len("&brief_mode=t"))="&brief_mode=t" Then
		brief_mode_rf="&brief_mode=t"
	ElseIf Right(url_str,Len("&brief_mode=f"))="&brief_mode=f" Then
		brief_mode_rf="&brief_mode=f"
	End If
			
			
	'http://www.pixiv.net/member.php?id=230836
	regex.Pattern = root_str & "/(\w+)\.php(?:\?(?:(?:((?:id|illust_id)=\d+)|((?:tag|word)=[%\w]+)|(type=(?:illust|user|reg_user))|(mode=(?:medium|all)|rest=(?:show|hide)|s_mode=(?:s_tc|s_tag))|[^&]+)(?:&|$))*)?"
	Set matches = regex.Execute(url_str)
	For Each match In matches
		Select Case match.SubMatches(0)
			Case "member", "member_illust"
				sub_url_str = "/member_illust.php?" & match.SubMatches(1) & "&" & match.SubMatches(4)
				multi_page = (match.SubMatches(4) = "")
			Case "tags"
				sub_url_str = "/tags.php?" & match.SubMatches(2)
			Case "search"
				sub_url_str = "/search.php?" & match.SubMatches(2) & "&" & match.SubMatches(4)
			Case "bookmark"
				sub_url_str = "/bookmark.php?" & match.SubMatches(1) & "&" & match.SubMatches(3) & "&" & match.SubMatches(4)
				multi_page = (InStr(match.SubMatches(3), "user") = 0)
			Case "response"
				sub_url_str = "/response.php?" & match.SubMatches(1) & "&" & match.SubMatches(4)
			Case "new_illust", "index"
				sub_url_str = "/new_illust.php"
			Case "bookmark_new_illust", "mypage"
				sub_url_str = "/bookmark_new_illust.php"
			Case Else
				Exit Function
		End Select
		regex.Pattern = "(?:(?:\?|&)+(?=$)|(\?|&)&+)"
		next_page_str = "1|inet|10,13|" & root_str & regex.Replace(sub_url_str, "$1")
		
		If brief_mode_rf="&brief_mode=t" Then
			brief_mode=-1
		ElseIf brief_mode_rf="&brief_mode=f" Then
			brief_mode=0
		Else
			brief_mode = (MsgBox("是否忽略漫画（采用简略分析方式）？", vbYesNo, "问题") = vbYes)
		End If
		Exit For
	Next
	
	return_download_url = "inet|10,13|" & root_str & "|" & root_str
End Function

Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
	return_albums_list = ""
	Dim name_filter_str, link_str, rename_str, description_str, regex, matches
	name_filter_str = ""
	Set regex = New RegExp
	regex.Global = True
	
	If started Then
		'<div style="width:140px;height:130px;float:left;text-align:center;">
		'<a href="member.php?id=395882"></a>
		'<div style="padding-top:5px;">ぽんねつ</div>
		'<a href="jump.php?http://www.e-bunny.net/" target="_blank">
		'</div>
		regex.Pattern = "<div[^>]*>\s*<a[^>]*href=""member\.php\?id=(\d+)[^""]*""[^>]*>(?:(?!</a>).)*</a>\s*<div[^>]*>([^<" & name_filter_str & "]+)[^<]*</div>(?:\s*<a[^>]*href=""jump.php\?([^""\s]+)[^""]*""[^>]*>|(?:(?!</div>).)+)+</div>"
		Set matches = regex.Execute(html_str)
		If matches.Count = 0 Then
			process_retry
		Else
			retries_count = 0
			For Each match In matches
				If brief_mode Then
					link_str = root_str & "/member_illust.php?id=" & match.SubMatches(0) & "&brief_mode=t"
				Else 
					link_str = root_str & "/member_illust.php?id=" & match.SubMatches(0) & "&brief_mode=f"
				End If
				rename_str = rename_utf8(match.SubMatches(1))
				description_str = match.SubMatches(2)
				return_albums_list = return_albums_list & "0||" & link_str & "|[" & match.SubMatches(0) & "]" & rename_str & "|" & description_str & vbCrLf
			Next
			next_page_str = get_next_page(html_str)
		End If
	Else
		check_login
	End If
	
	return_albums_list = return_albums_list & next_page_str
End Function

Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
	return_download_list = ""
	Dim page_count, regex, matches
	Set regex = new RegExp
	regex.Global = True
	
	If started Then	
		'<a href="member_illust.php?mode=medium&illust_id=8645263"><img src="http://img15.pixiv.net/img/hounori/8645263_s.jpg"alt="咲夜に背中流され隊" border="0" /></a>
		regex.Pattern = "<a[^>]*href=""(member_illust\.php\?mode=(\w+)&illust_id=(\d+))""[^>]*>\s*<img[^>]*src=""([^""]+)\3_(?:s|m)\.(\w+)[^""]*""[^>]*alt=""([^""]+)""[^>]*>\s*</a>"
		Set matches = regex.Execute(html_str)
		If matches.Count = 0 Then
			process_retry
		Else
			retries_count = 0
			For Each match In matches
				Select Case match.SubMatches(1)
					Case "medium"
						If InStr(next_page_str, match.SubMatches(2)) = 0 Then
							If brief_mode Then
								add_download_list_entry match, return_download_list, 0
							Else
								Set matches_cache = matches
								Exit For
							End If
						End If
					Case "big"
						If InStr(next_page_str, match.SubMatches(2)) > 0 Then
							add_download_list_entry match, return_download_list, 0
							Exit For
						End If
					Case "manga"
						If InStr(next_page_str, match.SubMatches(2)) > 0 Then
							'<span style="color:#666666;float:left;">全3ページ</span>
							regex.Pattern = "<span[^>]*>全(\d+)ページ"
							page_count = regex.Execute(html_str).Item(0).SubMatches(0)
							For page_index = 0 To page_count - 1
								add_download_list_entry match, return_download_list, page_index
							Next
							Exit For
						End If
					Case Else
						Exit Function
				End Select
			Next
			
			next_page_str = "0"
			If multi_page Then
				If cache_index = 0 Then
					next_page_str = get_next_page(html_str)
					parent_next_page_str = next_page_str
				End If
				If Not brief_mode Then
					If cache_index < matches_cache.Count Then
						next_page_str = "1|inet|10,13|" & root_str & "/" & matches_cache.Item(cache_index).SubMatches(0)
						cache_index = cache_index + 1
					Else
						next_page_str = parent_next_page_str
						cache_index = 0
					End If
					If next_page_str = "0" and brief_mode_rf="" Then
						MsgBox "分析已完成。", vbOKOnly, "提醒"
					End If
				End If
			End If
		End If
	Else
		check_login
	End If
	
	return_download_list = return_download_list & next_page_str
End Function

Function process_retry()
	retries_count = retries_count + 1
	If retries_count > 3 Then
		MsgBox "分析由于网络原因中断。", vbOKOnly + vbExclamation, "提醒"
		next_page_str = "0"
	End If
End Function

Function add_download_list_entry(ByRef match, ByRef download_list, ByVal page_index)
	Dim format_str, link_str, rename_str, description_str
	format_str = match.SubMatches(4)
	link_str = match.SubMatches(3) & match.SubMatches(2)
	rename_str = rename_utf8(match.SubMatches(5)) & "_" & match.SubMatches(2)
	description_str = rename_utf8(match.SubMatches(5))
	
	Select Case match.SubMatches(1)
		Case "medium", "big"
			link_str = link_str & "." & format_str
		Case "manga"
			link_str = link_str & "_p" & page_index & "." & format_str
			rename_str = rename_str & "_p" & page_index
			description_str = description_str & " - " & (page_index + 1)
		Case Else
			Exit Function
	End Select
	download_list = download_list & format_str & "|" & link_str & "|" & rename_str & "|" & description_str & vbCrLf
End Function

Function check_login()
	Dim regex, matches
	Set regex = new RegExp
	regex.Global = True
	
	'<input type="hidden" name="mode" value="login">
	regex.Pattern = "<input[^>]*value=""login""[^>]*>"
	If regex.Execute(html_str).Count > 0 Then
		MsgBox "您还没有登陆PIXIV。" & vbCrLf & "请使用内置浏览器登陆或使用IE类浏览器登陆" & vbCrLf & "并勾选“次回から自動的にログイン”保存cookies。", vbOKOnly + vbExclamation, "提醒"
		next_page_str = "0"
	Else
		started = True
	End If
End Function

Function get_next_page(ByVal html_str)
	get_next_page = "0"
	Dim regex, matches
	Set regex = New RegExp
	regex.Global = True
	
	'<a href=member_illust.php?id=230836&p=2>次の20件&gt;&gt;</a>
	regex.Pattern = "<a[^>]*href=([^>\s]+)[^>]*>次の\d+件&gt;&gt;</a>"
	Set matches = regex.Execute(html_str)
	For Each match In matches
		get_next_page = "1|inet|10,13|" & root_str & "/" & match.SubMatches(0)
		Exit For
	Next
End Function

Function rename_utf8(ByVal utf8_str)
	rename_utf8 = ""
	
	If Len(utf8_str) = 0 Then
		Exit Function
	End If
	For i = 1 to Len(utf8_str)
		If Asc(Mid(utf8_str, i, 1)) = 63 Then
			utf8_str = replace(utf8_str, Mid(utf8_str, i, 1), "_")
		End If
	Next
	rename_utf8 = replace(utf8_str, "|", "｜")
End Function