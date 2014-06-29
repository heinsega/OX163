'2014-6-29 visceroid & hein@shanghaijing.net
Dim started, multi_page, brief_mode, brief_mode_rf, retries_count, cache_index, root_str, next_page_str, parent_next_page_str, matches_cache, member_type, php_name
started = False
multi_page = True
retries_count = 0
cache_index = 0
ranking_page=0
ranking_url=""
root_str = "http://www.pixiv.net"

Function return_download_url(ByVal url_str)
On Error Resume Next
	member_type=0
	return_download_url = ""
	Dim sub_url_str, regex, matches, page_number, page_url
	Set regex = New RegExp
	regex.Global = True
	page_number=1
	brief_mode_rf=""
	If Right(url_str,Len("&brief_mode=t"))="&brief_mode=t" Then
		brief_mode_rf="&brief_mode=t"
	ElseIf Right(url_str,Len("&brief_mode=f"))="&brief_mode=f" Then
		brief_mode_rf="&brief_mode=f"
	End If
	url_str = replace(url_str,"mode=manga","mode=medium")
	url_str = replace(url_str,"mode=big","mode=medium")
	url_str = replace(url_str,"mode=manga_big","mode=medium")
	regex.Pattern = root_str & "/(\w+)\.php(?:\?(?:(?:((?:id|illust_id)=\d+)|((?:tag|word)=(?:[%\w\-]+\+?)+)|(type=(?:illust|user|reg_user))|(mode=(?:medium|all)|rest=(?:show|hide)|s_mode=(?:s_tc|s_tag))|(p=\d+)|[^&]+)(?:&|$))*)?"
	Set matches = regex.Execute(url_str)
	For Each match In matches
		Select Case LCase(match.SubMatches(0))
			Case "member", "member_illust"
				member_type=1
				php_name="member_illust.php"
				sub_url_str = "/member_illust.php?" & match.SubMatches(1) & "&" & match.SubMatches(2) & "&" & match.SubMatches(4)
				multi_page = (match.SubMatches(4) = "")
			Case "tags"
				php_name="search.php"
				'http://www.pixiv.net/tags.php?tag=%E3%80%90%E9%AD%94%E5%A5%B3%E3%81%AE%E5%A5%91%E7%B4%84%E3%80%91&tset=2
				'http://www.pixiv.net/search.php?s_mode=s_tag_full&word=%E3%80%90%E9%AD%94%E5%A5%B3%E3%81%AE%E5%A5%91%E7%B4%84%E3%80%91
				sub_url_str = replace(lcase(match.SubMatches(2)),"tag=","s_mode=s_tag_full&word=")
				sub_url_str = "/search.php?" & sub_url_str
			Case "search"
				php_name="search.php"
				sub_url_str=Mid(url_str,instr(url_str,"/search.php?")+12)
				sub_url_str=replace(sub_url_str,"&brief_mode=t","")
				sub_url_str=replace(sub_url_str,"&brief_mode=f","")
				If match.SubMatches(5)<> "" Then sub_url_str=replace(sub_url_str,"&" & match.SubMatches(5),"")
				sub_url_str = "/search.php?" & sub_url_str
			Case "bookmark"
				php_name="bookmark.php"
				sub_url_str = "/bookmark.php?" & match.SubMatches(1) & "&" & match.SubMatches(3) & "&" & match.SubMatches(4)
				multi_page = (InStr(match.SubMatches(3), "user") = 0)
			Case "response"
				php_name="response.php"
				sub_url_str = "/response.php?" & match.SubMatches(1) & "&" & match.SubMatches(4)
			Case "new_illust", "index"
				php_name="new_illust.php"
				sub_url_str = "/new_illust.php"
			Case "bookmark_new_illust", "mypage"
				php_name="bookmark_new_illust.php"
				sub_url_str = "/bookmark_new_illust.php"
			Case "bookmark_new_illust_r18"
				php_name="bookmark_new_illust_r18.php"
				sub_url_str = "/bookmark_new_illust_r18.php"
			Case "ranking"
				'http://www.pixiv.net/ranking.php?format=json&mode=daily&p=1
				php_name=LCase(match.SubMatches(0)) & ".php"
				sub_url_str=Mid(url_str,instr(url_str,".php?")+5)
				If match.SubMatches(5)<> "" Then sub_url_str=replace(sub_url_str,"&" & match.SubMatches(5),"")
				sub_url_str = "/" & php_name & "?" & sub_url_str & "&format=json"
				ranking_url = sub_url_str
				ranking_page=1
			Case "ranking_area"
				php_name=LCase(match.SubMatches(0)) & ".php"
				sub_url_str=Mid(url_str,instr(url_str,".php?")+5)
				If match.SubMatches(5)<> "" Then sub_url_str=replace(sub_url_str,"&" & match.SubMatches(5),"")
				sub_url_str = "/" & php_name & "?" & sub_url_str
			Case Else
				Exit Function
		End Select
		If match.SubMatches(5) <> "" Then
			If MsgBox("是否从第1页开始分析？", vbYesNo, "问题") = vbno Then
				sub_url_str = sub_url_str & "&" & match.SubMatches(5)
			End If
		End If
		regex.Pattern = "(?:(?:\?|&)+(?=$)|(\?|&)&+)"
		next_page_str = "1|inet|10,13|" & root_str & regex.Replace(sub_url_str, "$1")
		
		If instr(LCase(next_page_str),"illust_id=")>0 Then
			php_name="illust_id"
			brief_mode=vbYes
		End If
		
		If brief_mode_rf="&brief_mode=t" Then
			brief_mode=-1
		ElseIf brief_mode_rf="&brief_mode=f" Then
			brief_mode=0
		ElseIf php_name<>"illust_id" Then
			brief_mode = (MsgBox("是否忽略漫画（采用简略分析方式）？" & vbcrlf & vbcrlf & "2013年4月之后的作品必须选“否”才能正确分析", vbYesNo, "问题") = vbYes)
		End If
		Exit For
	Next
	
	return_download_url = "inet|10,13|" & root_str & "|" & root_str & vbcrlf & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.00; Windows 98)"
End Function

'---------------------------------------------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
	return_albums_list = ""
	Dim name_filter_str, link_str, rename_str, description_str, regex, matches
	name_filter_str = ""
	Set regex = New RegExp
	regex.Global = True
	
	If started Then
		'<li><input name="id[]" value="1593522" type="checkbox" /><div class="usericon"><a href="member.php?id=1593522"><img src="http://img46.pixiv.net/profile/kasetsu_03/mobile/3399441_80.jpg" alt="霞雪"/></a></div><div class="userdata"><a href="member.php?id=1593522">霞雪</a>はじめまして、“カセツ”と読みます。<br><span>&nbsp;</span></div></li>
		regex.Pattern = "<div[^>]*class=""userdata""[^>]*>\s*<a[^>]*href=""member\.php\?id=(\d+)[^""]*""[^>]*>\s*([^<" & name_filter_str & "]+)[^<]*</a>([\s\S^]*?)<br>"
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
		check_login html_str
	End If
	
	return_albums_list = return_albums_list & next_page_str
End Function

'---------------------------------------------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
	return_download_list = ""
	Dim page_count, regex, matches, ugoira_zip
	Set regex = new RegExp
	regex.Global = True
	ugoira_zip=0
	
	If started Then
		
		'清楚搜索页showcase内容
		If InStr(LCase(html_str), "<section class=""showcase"">")>0 Then
			Dim html_str_temp
			html_str_temp=mid(html_str,1,InStr(LCase(html_str), "<section class=""showcase"">")-1)
			html_str=mid(html_str,InStr(LCase(html_str), "<section class=""showcase"">")+len("<section class=""showcase"">"))
			html_str=html_str_temp & mid(html_str,InStr(LCase(html_str), "</section>")+len("</section>"))
		End If
		
		'转换ranking.php页面json数据
		If php_name="ranking.php" and cache_index = 0 Then
			html_str=format_ranking_html(html_str)
		End If
			
		'清除付费会员特殊格式
		If InStr(LCase(html_str), "data-src=")>1 Then
				html_str=format_transparent_html(html_str)
		End If
		
		'格式化非画师页面格式为画师页面格式
    '2013最新search.php
		'<li class="image-item"><a href="/member_illust.php?mode=medium&amp;illust_id=33318229" class="work"><img src="http://i1.pixiv.net/img05/img/kwcmm466/33318229_s.jpg" class="_thumbnail"><h1 class="title" title="レッド">レッド</h1></a><a href="/member_illust.php?id=41461" class="user" title="山城田うなぎ">山城田うなぎ</a><ul class="count-list"><li><a href="/bookmark_detail.php?illust_id=33318229" class="bookmark-count ui-tooltip" data-tooltip="2件のブックマーク"><span class="count-icon">&nbsp;</span>2</a></li></ul></li>
    '2013最新search.php <ul class="images autopagerize_page_element">重复部分
    '<li class="image"><a href="/member_illust.php?mode=medium&amp;illust_id=33318229"><p><img src="http://i1.pixiv.net/img05/img/kwcmm466/33318229_s.jpg"></p><h2>レッド</h2></a><p class="user"><a href="/member.php?id=41461">山城田うなぎ</a></p><ul class="count-list"><li><a href="/bookmark_detail.php?illust_id=33318229" class="bookmark-count ui-tooltip" data-tooltip="2件のブックマーク"><span class="count-icon">&nbsp;</span>2</a></li></ul></li>
    '2013最新bookmark_new_illust.php
		'<li class="image"><a href="/member_illust.php?mode=medium&amp;illust_id=33514129"><p><img src="http://i2.pixiv.net/img02/img/caelestis/33514129_s.jpg"></p><h2>冬</h2></a><p class="user"><a href="/member.php?id=14753">霜</a></p></li>
    '2013最新member_illust.php
		'<li class="image-item"><a href="/member_illust.php?mode=medium&amp;illust_id=33514129" class="work"><img src="http://i2.pixiv.net/img02/img/caelestis/33514129_s.jpg" class="_thumbnail"><h1 class="title" title="冬">冬</h1></a></li>
    '2013最新new_illust.php
		'<li class="image"><a href="/member_illust.php?mode=medium&amp;illust_id=33516395"><p><img src="http://i2.pixiv.net/img36/img/mid0nightcom3/33516395_s.jpg"></p><h2>【腐向け】一足早いキスをした【緑高】</h2></a><p class="user"><a href="/member.php?id=1050672">鳩丸</a></p></li>
    '2013最新bookmark.php
		'<a href="member_illust.php?mode=medium&amp;illust_id=33347566"><img src="http://i1.pixiv.net/img119/img/ms_sacory/33347566_s.jpg">ジブリがいっぱいコレクション</a>
		
		'old'<a href="member_illust.php?mode=medium&illust_id=17872081"><img src="http://img21.pixiv.net/img/youri19/17872081_s.png" alt="犬姫/悠" title="犬姫/悠" />犬姫</a></li>

		html_str=replace(html_str,"a href=""/member_illust.php","a href=""member_illust.php")
		
		If php_name="bookmark_new_illust.php" or php_name="bookmark_new_illust_r18.php" Then
			html_str=replace(html_str,"<h2>","<h1>")
			html_str=replace(html_str,"</h2>","</h1>")
			html_str=replace(html_str,"<p>","")
			html_str=replace(html_str,"</p>","")
		End If
		If cache_index=0 and php_name="ranking_area.php" Then
			regex.Pattern = "<a[^>]*href=""(member_illust\.php\?mode=(\w+)&(?:amp;)?illust_id=(\d+))[^""]*""[^>]*>\s*<img[^>]*(?:\s*(?:data-src=""([^""]+)\3_(?:s|m)\.(\w+)[^""]*""|alt=""([^""]+)""))+\s*>[\s\S]*?<h2><a[^>]*href=[^>]*>((?:(?!</a>).)*)</a></h2>"
		ElseIf brief_mode and php_name<>"bookmark.php" and php_name<>"illust_id" Then
			regex.Pattern = "<a[^>]*href=""(member_illust\.php\?mode=(\w+)&(?:amp;)?illust_id=(\d+))[^""]*""[^>]*>\s*<img(?:\s*(?:src=""([^""]+)\3_(?:s|m)\.(\w+)[^""]*""|alt=""([^""]+)""|\w+=""[^""]*""|))+\s*/?><h1[^>]*>((?:(?!</h1>).)*)</h1></a>"
		ElseIf instr(LCase(html_str),LCase("_ugoira1920x1080.zip"))>0 Then
			'{"src":"http:\/\/i2.pixiv.net\/img-zip-ugoira\/img\/2014\/06\/29\/14\/08\/25\/44387029_ugoira1920x1080.zip"
			regex.Pattern = "\{""src"":""(http)[^""]*(_ugoira1920x1080)\.zip"""'+\s*/?  --->  [^>]*
		ugoira_zip=1
		Else
			regex.Pattern = "<a[^>]*href=""(member_illust\.php\?mode=(\w+)&(?:amp;)?illust_id=(\d+))[^""]*""[^>]*>\s*<img(?:\s*(?:src=""([^""]+)\3_(?:s|m)\.(\w+)[^""]*""|alt=""([^""]+)""|\w+=""[^""]*""|))+\s*/?>((?:(?!</a>).)*)</a>"'+\s*/?  --->  [^>]*
		End If
		Set matches = regex.Execute(html_str)
		If matches.Count = 0 and php_name<>"illust_id" and instr(html_str,"<i class=""_icon sprites-mypixiv-badge""></i>")<1 and ugoira_zip=0 Then
			process_retry
		Else
			If ugoira_zip>0 Then
				add_download_list_ugoira html_str, return_download_list
			Else
				For Each match In matches
					Select Case match.SubMatches(1)
						Case "medium"
							If InStr(next_page_str, match.SubMatches(2)) = 0 Then
								If brief_mode Then
									add_download_list_entry match, return_download_list, 0
								ElseIf cache_index=0 Then
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
								'regex.Pattern = "<div[^>]*class=""works_data""[^>]*>\s*<p[^>]*>(?:(?!</p>).)*(?:漫画|漫畫|Manga) (\d+)P(?:(?!</p>).)*</p>"
								regex.Pattern = "<li>(?:漫画|漫畫|Manga) (\d+)P</li>"
								page_count = regex.Execute(html_str).Item(0).SubMatches(0)
								For page_index = 0 To page_count - 1
									add_download_list_entry match, return_download_list, page_index
								Next
								Exit For
							End If
						Case Else
							process_retry()
							return_download_list = return_download_list & next_page_str
							Exit Function
					End Select
				Next
			End If
			
			retries_count = 0
			next_page_str = "0"
			If multi_page Then
				If cache_index = 0 Then
					next_page_str = get_next_page(html_str)
					parent_next_page_str = next_page_str
				End If
				If Not brief_mode Then
					If cache_index < matches_cache.Count Then
						'http://source.pixiv.net/source/images/limit_mypixiv_s.png?20110520
						'http://source.pixiv.net/source/images/limit_unknown_s.png?20110520
						Do While instr(lcase(matches_cache.Item(cache_index).SubMatches(3) & matches_cache.Item(cache_index).SubMatches(2)),".pixiv.net/img")<1
							If cache_index < matches_cache.Count Then
							cache_index = cache_index + 1
							Else
							Exit Do
							End If
						loop
						If cache_index < matches_cache.Count Then
							next_page_str = "1|inet|10,13|" & root_str & "/" & matches_cache.Item(cache_index).SubMatches(0)
							next_page_str = replace(next_page_str,"&amp;","&")
							cache_index = cache_index + 1
						Else
						next_page_str = parent_next_page_str
						cache_index = 0						
						End If
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
		check_login html_str
	End If
	return_download_list = return_download_list & next_page_str
End Function
'----------------------------------------------------------------------------------------------------
Function process_retry()
	retries_count = retries_count + 1
	If retries_count > 3 Then
		retries_count=0
		If brief_mode Then
			If Len(parent_next_page_str)>2 Then
				next_page_str = parent_next_page_str
			Else
				next_page_str = "0"
			End If
		Else
		 If cache_index < matches_cache.Count Then
		 	next_page_str = "1|inet|10,13|" & root_str & "/" & matches_cache.Item(cache_index).SubMatches(0)
		 	next_page_str = replace(next_page_str,"&amp;","&")
		 	cache_index = cache_index + 1
		 	ElseIf Len(parent_next_page_str)>2 Then
				next_page_str = parent_next_page_str
				Else
				next_page_str = "0"
			End If
		End If
	End If
End Function
'----------------------------------------------------------------------------------------------------
Function add_download_list_entry(ByRef match, ByRef download_list, ByVal page_index)
	Dim format_str, link_str, rename_str, description_str, manga_big_str
	If match.SubMatches(3)="" Then Exit Function
	format_str = match.SubMatches(4)
	link_str = match.SubMatches(3) & match.SubMatches(2)
	If match.SubMatches(6) <> "" Then
		rename_str = "(pid-" & match.SubMatches(2) & ")" & rename_utf8(match.SubMatches(6))
	Else
		rename_str = "(pid-" & match.SubMatches(2) & ")" & rename_utf8(match.SubMatches(5))
	End If
	description_str = rename_utf8(match.SubMatches(5))
	
	Select Case match.SubMatches(1)
		Case "medium", "big"
			link_str = link_str & "." & format_str
		Case "manga" '11319936_big_p0.jpg 06/16/2010 20:43--------------11319930_p0.jpg 06/16/2010 20:43
			manga_big_str="_p"
			If match.SubMatches(2)<11319931 Then
				manga_big_str="_p"
			Else
				manga_big_str="_big_p"
			End If
			link_str = link_str & manga_big_str & page_index & "." & format_str
			rename_str = rename_str & manga_big_str & page_index
			description_str = description_str & " - " & (page_index + 1)
		Case Else
			Exit Function
	End Select
	download_list = download_list & format_str & "|" & link_str & "?" & (CDbl(Now()) * 10000000000) & "|" & rename_str & "|" & description_str & vbCrLf
End Function
'---------------------------------------------------------------------------------------------

Function add_download_list_ugoira(byval ugoira_str, ByRef download_list)
		'pixiv.context.illustId         = '44387029';
		'pixiv.context.illustTitle      = 'Hello ミク';pixiv.context.userId           = '395595';
		'pixiv.context.userName         = 'KD'
		'{"src":"http:\/\/i2.pixiv.net\/img-zip-ugoira\/img\/2014\/06\/29\/14\/08\/25\/44387029_ugoira1920x1080.zip"
		Dim file_ID,file_name,file_Url,file_description
		
		ugoira_str=mid(ugoira_str,1,instr(LCase(ugoira_str),LCase("_ugoira1920x1080.zip"))) & "ugoira1920x1080.zip"
		
		file_Url=Mid(ugoira_str,InStrrev(ugoira_str,chr(34))+1)
		file_Url=replace(file_Url,"\/","/")
		
		file_ID=mid(ugoira_str,InStr(LCase(ugoira_str),LCase("pixiv.context.illustId")))
		file_ID=mid(file_ID,InStr(file_ID,"'")+1)
		file_ID=mid(file_ID,1,InStr(file_ID,"'")-1)

		file_description=mid(ugoira_str,InStr(LCase(ugoira_str),LCase("pixiv.context.illustTitle")))
		file_description=mid(file_description,InStr(file_description,"'")+1)
		file_description=mid(file_description,1,InStr(file_description,"'")-1)
		file_name=file_description
		If Len(file_name)>200 Then file_name=left(file_name,200)
		file_name="(pid-" & file_ID & ")" & rename_utf8(file_name) & "_ugoira1920x1080.zip"
		
		download_list = "zip|" & file_Url & "?" & (CDbl(Now()) * 10000000000) & "|" & file_name & "|" & file_description & vbCrLf
End Function
'---------------------------------------------------------------------------------------------

Function check_login(ByVal html_str)
	Dim regex, matches
	Set regex = new RegExp
	regex.Global = True
	
	regex.Pattern = "<input[^>]*value=""login""[^>]*>"
	If regex.Execute(html_str).Count > 0 Then
		MsgBox "您还没有登陆PIXIV。" & vbCrLf & "请使用内置浏览器登陆或使用IE类浏览器登陆" & vbCrLf & "并勾选“次回から自動的にログイン”保存cookies。", vbOKOnly + vbExclamation, "提醒"
		next_page_str = "0"
	Else
		started = True
	End If
End Function
'---------------------------------------------------------------------------------------------
Function get_next_page(ByVal html_str)
	get_next_page = "0"
	Dim regex, matches
	Set regex = New RegExp
	regex.Global = True
	regex.Pattern = "<a[^>]*href=""([^>\s]+)""[^>]*rel=""next""[^>]*>.*?</a>\s*</li>"
								 '新search.php页面
								 '<li><a href="?word=sega&amp;order=date_d&amp;p=2" rel="next" title="下一面"><span class="_button-lite"><i class="_icon sprites-next"></i></span></a></li>
								 '作者页面与老search.php页面
								 '<li><a href="member_illust.php?id=517481&p=2" class="button" rel="next">下一面 ?</a></li>
								 '<li class="next"><a href="?word=sega&amp;order=date_d&amp;p=2" class="ui-button-light" rel="next" title="下一面">&gt;</a></li>
								 
								 '新ranking.php页面
								 '<span class="next"><a href="?mode=daily&amp;p=2&amp;ref=rn-h-next" rel="next" class="_button" title="下一面"><i class="_icon sprites-next-linked"></i></a></span>
	If php_name="ranking.php" and ranking_page>0 and ranking_page<10 Then
		ranking_page=ranking_page+1
		get_next_page = "1|inet|10,13|" & root_str & "/" & ranking_url & "&p=" & ranking_page
		Exit Function
	End If

	Set matches = regex.Execute(html_str)
	'InputBox next_page_str,next_page_str,next_page_str
	For Each match In matches
		get_next_page = replace(match.SubMatches(0),"&amp;","&")
		If Left(get_next_page,1)="?" Then
			get_next_page=php_name & get_next_page
		End If
		get_next_page = "1|inet|10,13|" & root_str & "/" & get_next_page
		Exit For
	Next
End Function
'---------------------------------------------------------------------------------------------
Function format_transparent_html(ByVal html_str)
    format_transparent_html = html_str
    '2013最新search.php
    '<LI class="image-item"><A class="work" href="http://www.pixiv.net/member_illust.php?mode=medium&amp;illust_id=33484357"><DIV class="layout-thumbnail"><IMG class="_thumbnail ui-scroll-view" alt="" src="http://source.pixiv.net/source/images/common/transparent.gif" data-user-id="2876335" data-tags="オリジナル 女の子 落書き" data-src="http://i2.pixiv.net/img72/img/ttt0106/33484357_s.jpg" data-filter="thumbnail-filter lazy-image"></DIV><H1 class="title" title="一人旅">一人旅</H1></A><A class="user" title="たいそす" href="http://www.pixiv.net/member_illust.php?id=2876335">たいそす</A></LI>
		'转换为
		'<LI class="image-item"><A class="work" href="http://www.pixiv.net/member_illust.php?mode=medium&amp;illust_id=33484357"><img src="http://i2.pixiv.net/img72/img/ttt0106/33484357_s.jpg"><H1 class="title" title="一人旅">一人旅</H1></A><A class="user" title="たいそす" href="http://www.pixiv.net/member_illust.php?id=2876335">たいそす</A></LI>
    '2013最新search.php <ul class="images autopagerize_page_element">重复部分
		'<li class="image"><a href="/member_illust.php?mode=medium&amp;illust_id=33484357"><p><div class="layout-thumbnail"><img alt="" class="ui-scroll-view" src="http://source.pixiv.net/source/images/common/transparent.gif" data-filter="thumbnail-filter lazy-image" data-src="http://i2.pixiv.net/img72/img/ttt0106/33484357_s.jpg" data-tags="オリジナル 女の子 落書き"></div></p><h2>一人旅</h2></a><p class="user"><a href="/member.php?id=2876335">たいそす</a></p></li>
		'转换为
		'<li class="image"><a href="/member_illust.php?mode=medium&amp;illust_id=33484357"><img src="http://i2.pixiv.net/img72/img/ttt0106/33484357_s.jpg"><h1>一人旅</h1></a><p class="user"><a href="/member.php?id=2876335">たいそす</a></li>
    Dim split_str, matches(4),temp(2)
    temp(0) = Mid(html_str,1,InStr(LCase(html_str), "<div class=""layout-thumbnail"">")-1)
    html_str = Mid(html_str,InStr(LCase(html_str), "<div class=""layout-thumbnail"">")+len("<div class=""layout-thumbnail"">"))
    'temp(1) = Mid(html_str,InStr(LCase(html_str), "<div class=""clear""></div>"))
    split_str = Split(html_str, "<div class=""layout-thumbnail"">")
    For i = 0 To UBound(split_str)
    		matches(0) = ""
    		matches(1) = ""
    		matches(2) = ""
    		matches(3) = ""
    		matches(4) = ""
    		'del transparent.gif
    		matches(0)="<img src="""
    		matches(1)=Mid(split_str(i), InStr(LCase(split_str(i)), """ data-src=""") + Len(""" data-src="""))
    		matches(2) = Mid(matches(1), InStr(matches(1), ">"))
    		matches(1) = Mid(matches(1),1,InStr(matches(1), """"))
    		'<img src="http://i2.pixiv.net/img72/img/ttt0106/33484357_s.jpg">....
    		split_str(i)=matches(0) & matches(1) & matches(2)
    		'del data-tags
    		'del </div>
        matches(3) = Mid(split_str(i),1,InStr(LCase(split_str(i)), "</div>")-1)
        matches(4) = Mid(split_str(i),InStr(LCase(split_str(i)), "</div>")+6)
        split_str(i)=matches(3) & matches(4)
    Next
    format_transparent_html =temp(0) & Join(split_str, "")
    format_transparent_html=replace(format_transparent_html,"<h2>","<h1>")
		format_transparent_html=replace(format_transparent_html,"</h2>","</h1>")
		format_transparent_html=replace(format_transparent_html,"<p>","")
		format_transparent_html=replace(format_transparent_html,"</p>","")
End Function

Function format_ranking_html(ByVal html_str)
    format_ranking_html = html_str
    '{"illust_id":40019442,
    '"title":"\u305d\u3063\u3068 \u308f\u305f\u3057\u306f \u5927\u4eba\u306b\u306a\u3063\u305f","width":1000,"height":1000,"date":"2013\u5e7411\u670830\u65e5 15:35","tags":["\u30aa\u30ea\u30b8\u30ca\u30eb","\u3075\u3064\u304f\u3057\u3044",    "\u306a\u306b\u3053\u308c\u7d20\u6575","\u9ed2\u30bb\u30fc\u30e9\u30fc","\u30aa\u30ea\u30b8\u30ca\u30eb500users\u5165\u308a"],
    '"url":"http:\/\/i1.pixiv.net\/img77\/img\/fff365\/mobile\/40019442_240mw.jpg",
    '"user_id":3118206,
    '"user_name":"\u53e4\u753a","profile_img":"http:\/\/i1.pixiv.net\/img77\/profile\/fff365\/6437486_s.png","rank":53,"yes_rank":0,"total_score":3959,"view_count":3560},
		'转换为
		'<li class="image"><a href="/member_illust.php?mode=medium&amp;illust_id=33484357">
		'<img src="http://i2.pixiv.net/img72/img/ttt0106/33484357_s.jpg"><h1>一人旅</h1></a>
		'<p class="user"><a href="/member.php?id=2876335">たいそす</a></li>

		Dim split_str, matches(4)
    If InStr(html_str, "{""illust_id"":") > 0 Then
        html_str = Mid(html_str, InStr(LCase(html_str), "{""illust_id"":") + Len("{""illust_id"":"))
    Else
        format_ranking_html = ""
        Exit Function
    End If
		html_str=replace(html_str,"\/","/")
    split_str = Split(html_str, "{""illust_id"":")
    For i = 0 To UBound(split_str)
    		matches(0) = ""
    		matches(1) = ""
    		matches(2) = ""
    		matches(3) = ""
    		matches(4) = ""
    		'illust_id
    		matches(0)=Mid(split_str(i),1,InStr(split_str(i),",")-1)
    		'url
    		'http://i1.pixiv.net/img35/img/snika5800/mobile/44331825_240mw.jpg
    		'--->http://i1.pixiv.net/img35/img/snika5800/44331825_s.jpg
    		'http://i1.pixiv.net/c/240x480/img-master/img/2014/06/27/14/45/03/44340318_master1200.jpg
    		'--->http://i1.pixiv.net/img-inf/img/2014/06/27/14/45/03/44340318_s.jpg
    		matches(1)=Mid(split_str(i), InStr(LCase(split_str(i)), """url"":""") + Len("""url"":"""))
    		matches(1)=Mid(matches(1),1,InStr(matches(1),"""")-1)
    		If InStr(matches(1),"/mobile/")>0 Then
    			matches(1)=replace(matches(1),"/mobile/","/")
    		ElseIf InStr(matches(1),"/img-master/")>0 Then
    			matches(1)=Mid(matches(1),1,instr(matches(1),".pixiv.net/")+9) & "/img-inf/" & Mid(matches(1),InStr(matches(1),"/img-master/")+12)
    		End if
    		matches(1)=Mid(matches(1),1,InStrrev(matches(1),"_")-1) & "_s" & Mid(matches(1),InStrrev(matches(1),"."))
    		'title
    		matches(2)=Mid(split_str(i), InStr(LCase(split_str(i)), """title"":""") + Len("""title"":"""))
    		matches(2)=Mid(matches(2),1,InStr(matches(2),"""")-1)
    		matches(2)=fix_Unicode_Name(matches(2))
    		matches(2)=replace(matches(2),">","&gt;")
    		matches(2)=replace(matches(2),"<","&lt;")
    		'user_id
    		matches(3)=Mid(split_str(i), InStr(LCase(split_str(i)), """user_id"":") + Len("""user_id"":"))
    		matches(3)=Mid(matches(3),1,InStr(matches(3),",")-1)
    		'user_name
    		matches(4)=Mid(split_str(i), InStr(LCase(split_str(i)), """user_name"":""") + Len("""user_name"":"""))
    		matches(4)=Mid(matches(4),1,InStr(matches(4),"""")-1)
    		matches(4)=fix_Unicode_Name(matches(4))
    		matches(4)=replace(matches(4),">","&gt;")
    		matches(4)=replace(matches(4),"<","&lt;")

    		split_str(i)="<li class=""image""><a href=""/member_illust.php?mode=medium&amp;illust_id=" & matches(0) & """><img src=""" & matches(1) & """><h1>" & matches(2) & "</h1></a><p class=""user""><a href=""/member.php?id=" & matches(3) & """>" & matches(4) & "</a></li>"
    Next
    format_ranking_html =Join(split_str, "")
End Function
'---------------------------------------------------------------------------------------------
Function fix_Unicode_Name(ByVal sLongFileName)
    Dim i,fixed_Unicode_tf,split_str,fix_Unicode    
    fix_Unicode_Name = sLongFileName 
    split_str = Split(sLongFileName, "\u")
    If UBound(split_str) >= 1 Then
        For i = 1 To UBound(split_str)
            fixed_Unicode_tf = False
            If Len(split_str(i)) > 3 Then
                fix_Unicode = Mid(split_str(i), 1, 4)
                If Len(split_str(i)) > 4 Then
                	split_str(i) = Mid(split_str(i), 5)
                Else
                	split_str(i) = ""
                End If
                
                If is_Hex_code(fix_Unicode) Then
                    fix_Unicode = ChrW(Int("&H" & fix_Unicode))
                    fixed_Unicode_tf = True
                End If
                
                If fixed_Unicode_tf = False Then
                    split_str(i) = "\u" & fix_Unicode & split_str(i)
                Else
                    split_str(i) = fix_Unicode & split_str(i)
                End If
            End If
        Next
        fix_Unicode_Name = Join(split_str, "")
    End If
End Function

Function is_Hex_code(ByVal Hex_code)
    Dim i
    is_Hex_code = True
    If Len(Hex_code)>0 And Len(Hex_code)<7 Then
        For i=1 To Len(Hex_code)
            If InStr("ABCDEFabcdef0123456789", Mid(Hex_code, i, 1)) < 1 Then is_Hex_code = False: Exit Function
        Next
    Else
        is_Hex_code = False
    End If
End Function

'----------------------------------------------------------------------
' 保存文本文件
Function SaveEncodedTextFile(sFilePath, sCharset, s)
    Dim oStream
    Set oStream = CreateObject("ADODB.Stream")
    ' 以文本模式
    oStream.Type = 2
    oStream.Mode = 3
    If Len(sCharset) > 0 Then
        On Error Resume Next
        oStream.Charset = sCharset
        If Err.number <> 0 Then
            oStream.Charset = "_autodetect_all"
        End If
        On Error Goto 0
    End If
    oStream.Open
    oStream.WriteText s
    ' 2 - adSaveCreateOverwrite
    On Error Resume Next
    oStream.SaveToFile sFilePath, 2
    If Err.number <> 0 Then
        SaveEncodedTextFile = False
    Else
        SaveEncodedTextFile = True
    End If
    On Error Goto 0
    Set oStream = Nothing
End Function
    
Function rename_utf8(ByVal utf8_str)
	rename_utf8 = ""	
	If Len(utf8_str) = 0 Then
		Exit Function
	End If	
	utf8_str=Hex_unicode_str(utf8_str)
	
	For i = 1 to Len(utf8_str)
		If Asc(Mid(utf8_str, i, 1)) = 63 Then
			utf8_str = replace(utf8_str, Mid(utf8_str, i, 1), "_")
		End If
	Next
	rename_utf8 = replace(utf8_str, "|", "｜")
End Function

Function Hex_unicode_str(ByVal old_String)
    Dim i, UnAnsi_Str, Hex_UnAnsi_Str
    For i = 1 To Len(old_String)
        If Asc(Mid(old_String, i, 1)) = 63 Then UnAnsi_Str = UnAnsi_Str & Mid(old_String, i, 1)
    Next
        
    For i = 1 To Len(UnAnsi_Str)
        Hex_UnAnsi_Str = Mid(UnAnsi_Str, i, 1)
        Hex_UnAnsi_Str = "&H" & Hex(AscW(Hex_UnAnsi_Str))
        old_String = Replace(old_String, Mid(UnAnsi_Str, i, 1), "&#" & Int(Hex_UnAnsi_Str) & ";")
    Next
    Hex_unicode_str = old_String
End Function
