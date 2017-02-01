'2017-2-1 visceroid & hein@shanghaijing.net
Dim started, multi_page, brief_mode, reg_bigmode, brief_mode_rf, retries_count, cache_index, root_str, next_page_str, parent_next_page_str, matches_cache, php_name
Dim manga_count, ids_count, ids_max, ids_split, limit_ids_max, ids_string
started = False
multi_page = True
retries_count = 0
cache_index = 0
ranking_page=0
manga_count=0
ids_count=0
ids_max=0
limit_ids_max=200
ranking_url=""
root_str = "http://www.pixiv.net"

Function return_download_url(ByVal url_str)
On Error Resume Next
	Dim sub_url_str, regex, matches, page_number, page_url
	return_download_url = ""
	Set regex = New RegExp
	regex.Global = True
	regex.IgnoreCase = True
	
	page_number=1
	brief_mode_rf=""
	reg_bigmode=""
	If Right(url_str,Len("&brief_mode=t"))="&brief_mode=t" Then
		brief_mode_rf="&brief_mode=t"
	ElseIf Right(url_str,Len("&brief_mode=f"))="&brief_mode=f" Then
		brief_mode_rf="&brief_mode=f"
	End If
	If instr(url_str,"#")>0 Then url_str=mid(url_str,1,instr(url_str,"#")-1)
	regex.Pattern = root_str & "/(\w+)\.php(?:\?(?:(?:((?:id|illust_id)=\d+)|((?:tag|word)=(?:[%\w\-]+\+?)+)|(type=(?:illust|user|reg_user))|(mode=(?:medium|big|manga|manga_big|all)|rest=(?:show|hide)|s_mode=(?:s_tc|s_tag))|(p=\d+)|[^&]+)(?:&|$))*)?"
	Set matches = regex.Execute(url_str)
	For Each match In matches
		Select Case LCase(match.SubMatches(0))
			Case "member", "member_illust"
				php_name="member_illust.php"
				multi_page = (match.SubMatches(4) = "")
				If not multi_page Then
					'match.SubMatches(1)=illust_id=48061189
					reg_bigmode="json"
					sub_url_str = "/rpc/illust_list.php?illust_ids=" & replace(match.SubMatches(1),"illust_id=","") & "&verbosity="
				Else
					sub_url_str = "/member_illust.php?" & match.SubMatches(1) & "&" & match.SubMatches(2) & "&" & match.SubMatches(4)
				End If
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
				sub_url_str = "/bookmark.php?" & match.SubMatches(1) & "&" & match.SubMatches(2) & "&" & match.SubMatches(3) & "&" & match.SubMatches(4)
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
				reg_bigmode="json"
				ranking_page=2
			  parent_next_page_str = "1|inet|10,13|" & root_str & "/" & ranking_url & "&p=2"
			Case "ranking_area"
				php_name=LCase(match.SubMatches(0)) & ".php"
				sub_url_str=Mid(url_str,instr(url_str,".php?")+5)
				If match.SubMatches(5)<> "" Then sub_url_str=replace(sub_url_str,"&" & match.SubMatches(5),"")
				sub_url_str = "/" & php_name & "?" & sub_url_str
			Case "bookmark_detail"
				'http://www.pixiv.net/bookmark_detail.php?illust_id=49734016
				'http://www.pixiv.net/rpc/recommender.php?type=illust&sample_illusts=49734016&num_recommendations=1000&tt=e75f2fba47c534cb303d889d383cacb1
				sub_url_str = "/rpc/recommender.php?type=illust&sample_illusts=" & replace(match.SubMatches(1),"illust_id=","") & "&num_recommendations=1000"
			Case Else
				Exit Function
		End Select
		If match.SubMatches(5) <> "" Then
			If MsgBox("是否从第1页开始分析？", vbYesNo, "问题") = vbno Then
				If instr(sub_url_str,"?")<1 Then
					sub_url_str = sub_url_str & "?" & match.SubMatches(5)
				Else
					sub_url_str = sub_url_str & "&" & match.SubMatches(5)
				End If
			End If
		End If
		regex.Pattern = "(?:(?:\?|&)+(?=$)|(\?|&)&+)"
		next_page_str = "1|inet|10,13|" & root_str & regex.Replace(sub_url_str, "$1")
		
		If instr(LCase(next_page_str),"illust_id=")>0 Then
			php_name="illust_id"
			brief_mode=vbYes
		End If
		
		brief_mode=0
		'	brief_mode = (MsgBox("是否忽略漫画（采用简略分析方式）？" & vbcrlf & vbcrlf & "2013年4月之后的作品必须选“否”才能正确分析", vbYesNo, "问题") = vbYes)
		
		Exit For
	Next
	
	return_download_url = "inet|10,13|" & root_str & "|" & root_str & vbcrlf & "User-Agent: Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/7.0)"
	OX163_urlpage_Referer = "http://www.pixiv.net/member_illust.php?mode=medium&illust_id=12345" & vbCrLf & "User-Agent: Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko"

End Function

'---------------------------------------------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
	return_albums_list = ""
	Dim name_filter_str, link_str, rename_str, description_str, regex, matches
	name_filter_str = ""
	Set regex = New RegExp
	regex.Global = True
	regex.IgnoreCase = True
	
	If started Then
		'<li><input name="id[]" value="1593522" type="checkbox" /><div class="usericon"><a href="member.php?id=1593522"><img src="http://img46.pixiv.net/profile/kasetsu_03/mobile/3399441_80.jpg" alt="霞雪"/></a></div><div class="userdata"><a href="member.php?id=1593522">霞雪</a>はじめまして、“カセツ”と読みます。<br><span>&nbsp;</span></div></li>
		regex.Pattern = "<div[^>]*class=""userdata""[^>]*>\s*<a[^>]*href=""member\.php\?id=(\d+)[^""]*""[^>]*>\s*([^<" & name_filter_str & "]+)[^<]*</a>([\s\S^]*?)<br>"
		Set matches = regex.Execute(html_str)
		If matches.Count = 0 Then
			If process_retry=ture Then	return_albums_list = next_page_str
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
	Dim page_count, regex, matches, ids
	Set regex = new RegExp
	regex.Global = True
	regex.IgnoreCase = True

	'检测是否登陆
	If not started Then
		check_login html_str 
		return_download_list = next_page_str
		Exit Function
	End If
	
		'清除搜索页showcase内容
		If InStr(LCase(html_str), "<section class=""showcase"">")>0 Then
			Dim html_str_temp
			html_str_temp=mid(html_str,1,InStr(LCase(html_str), "<section class=""showcase"">")-1)
			html_str=mid(html_str,InStr(LCase(html_str), "<section class=""showcase"">")+len("<section class=""showcase"">"))
			html_str=html_str_temp & mid(html_str,InStr(LCase(html_str), "</section>")+len("</section>"))
		End If
			
		'清除付费会员特殊格式
		If InStr(LCase(html_str), "data-src=")>1 Then
				html_str=format_transparent_html(html_str)
		End If

		'判断页面类型使用对应正则
		'setp3,获取动图
		
		'部分单图页面其实是manga，需要判断
		If manga_count=0 and reg_bigmode="big" Then
			regex.Pattern = "<li>(?:複数枚投稿|一次性投稿多張作品|一次性投稿多张作品) (\d+)P</li>"
			Set matches = regex.Execute(html_str)
			If matches.count>0 Then
				reg_bigmode="manga"
			End If
		End If
		
		If reg_bigmode="ugoira" Then
			If InStr(LCase(html_str),LCase("_ugoira1920x1080.zip"))>0 Then
				add_download_list_ugoira html_str, return_download_list
				Call Set_next_json_url
			Else
				If process_retry=false Then	Call Set_next_json_url
			End If
			
		ElseIf reg_bigmode="big" Then
			If InStr(LCase(html_str),LCase("class=""original-image"""))>0 Then
				add_download_list_big html_str, return_download_list
				Call Set_next_json_url
			Else
				If process_retry=false Then	Call Set_next_json_url
			End If
					
		ElseIf reg_bigmode="manga" Then
			If manga_count=0 Then
				'<span class="total">5</span>
				regex.Pattern = "<li>(?:複数枚投稿|一次性投稿多張作品|一次性投稿多张作品) (\d+)P</li>"				
				page_count = regex.Execute(html_str).Item(0).SubMatches(0)
				If IsNumeric(page_count) Then manga_count=int(page_count)
				If manga_count>0 Then
					next_page_str="1|inet|10,13|" & root_str & "/member_illust.php?mode=manga_big&illust_id=" & matches_cache.Item(cache_index-1).SubMatches(0) & "&page=0"
				Else
					manga_count=1
					next_page_str="1|inet|10,13|" & root_str & "/member_illust.php?mode=big&illust_id=" & matches_cache.Item(cache_index-1).SubMatches(0)
				End If
			ElseIf manga_count>0 and InStr(LCase(html_str),LCase("<img src="""))>0 Then
				add_download_list_manga html_str, return_download_list
				Call Set_next_json_url
			Else
				If process_retry=false Then	Call Set_next_json_url
			End If
			
		'setp2,分析json数据
		ElseIf reg_bigmode="json" Then
			'转换ranking.php页面json数据
			If php_name="ranking.php" and cache_index = 0 Then html_str=format_ranking_html(html_str)
			regex.Pattern = """illust_id"":""([0-9]{1,})"",[\s\S]*?""(?:illust_title)"":""([^""]*)"",[\s\S]*?""illust_type"":""([012])"""
			Set matches = regex.Execute(html_str)
			
			If matches.Count = 0 Then
				If process_retry=false Then reg_bigmode="":cache_index=0:next_page_str=parent_next_page_str:parent_next_page_str=""
			Else
				Set matches_cache = matches
				Call Set_next_json_url				
			End If
			
		'setp1图片列表获取ids
		ElseIf cache_index=0 and reg_bigmode="" Then
			regex.Pattern ="<a(?:\s*href=""/?(member_illust\.php\?mode=(\w+)&(?:amp;)?illust_id=(\d+))(?:(?!ref=)[^""])*""\s*class=""(work[^""]*)"")[^>]*?>[\s\S]*?""_layout-thumbnail""[\s\S]*?<img(?:\s*(?:src=""([^""]+)(?:_(?:s|m|(?:master\d+))\.)(\w+)[^""]*""|\w+=""[^""]*""))[^>]*?>" 
			Set matches = regex.Execute(html_str)
			'{"recommendations":[49730141,49844639,49762616,50011024]}
			If matches.Count = 0 and Left(LCase(html_str),len("{""recommendations"":["))<>"{""recommendations"":[" Then
				If process_retry=false Then next_page_str=parent_next_page_str:parent_next_page_str=""
			Else
				ids=""
				If Left(LCase(html_str),len("{""recommendations"":["))="{""recommendations"":[" Then
					ids=Mid(html_str,21)
					ids=Left(ids,len(ids)-2)
				Else
					For Each match In matches
						Select Case match.SubMatches(1)
							Case "medium"
								If IsNumeric(match.SubMatches(2)) Then ids=ids & match.SubMatches(2) & ","
						End Select
					Next
				End If
				
				If Right(ids,1)="," Then ids=Left(ids,Len(ids)-1)
				retries_count = 0
				next_page_str = "0"
				reg_bigmode="json"
				'获取下一版面地址
				If cache_index = 0 Then
				
				'bookmark_detail 超过200分段进行
					If ids_count>0 Then
						ids_count=ids_count+1
						ids=limit_ids
						If ids_count=ids_max Then ids_count=0
					Else
						ids_string=ids
						ids_max=ubound(split(ids,","))
						ids_max=-1 * Int(-1*(ids_max+1)/limit_ids_max)
						ids_count=1
						ids=limit_ids()
						If ids_count=ids_max Then ids_count=0:ids=ids_string
						next_page_str = get_next_page(html_str)
						parent_next_page_str = next_page_str
					End If
				End If
				'获取json
				next_page_str = "1|inet|10,13|" & root_str & "/rpc/illust_list.php?illust_ids=" & ids & "&verbosity=" 'Replace(ids,",","%2C")
			End If
		End If
	If (next_page_str = "0" or next_page_str="") and brief_mode_rf="" Then
		'bookmark_detail 超过200分段进行
		If ids_count>0 Then
			next_page_str="1|inet|10,13|" & root_str
		Else
			MsgBox "分析已完成。", vbOKOnly, "提醒"
		End If
	End If
		
	return_download_list = return_download_list & next_page_str

End Function
'----------------------------------------------------------------------------------------------------
Function limit_ids()
Dim ids_split
ids_split=split(ids_string,",")
For i=0 to ubound(ids_split) 
	If i>=((ids_count-1)*limit_ids_max) and i<(ids_count*limit_ids_max) Then
		ids_split(i)= ids_split(i) & ","
	Else
		ids_split(i)= ""
	End If
Next
limit_ids=join(ids_split,"")
If Right(limit_ids,1)="," Then limit_ids=Left(limit_ids,Len(limit_ids)-1)
End Function
'----------------------------------------------------------------------------------------------------
Function process_retry()
On Error Resume Next
	process_retry=True
	retries_count = retries_count + 1
	If retries_count > 3 Then
		retries_count=0
		process_retry=False
	End If
End Function
'----------------------------------------------------------------------------------------------------
'step2
Function Set_next_json_url()
	Set_next_json_url=""
	retries_count=0
	manga_count=0
	If cache_index<matches_cache.count Then
		Select Case matches_cache.Item(cache_index).SubMatches(2)
		Case "1"
			reg_bigmode="manga"
			next_page_str="1|inet|10,13|" & root_str & "/member_illust.php?mode=medium&illust_id=" & matches_cache.Item(cache_index).SubMatches(0)
		Case "2"
			reg_bigmode="ugoira"
			next_page_str="1|inet|10,13|" & root_str & "/member_illust.php?mode=medium&illust_id=" & matches_cache.Item(cache_index).SubMatches(0)
		Case Else
			reg_bigmode="big"
			next_page_str="1|inet|10,13|" & root_str & "/member_illust.php?mode=medium&illust_id=" & matches_cache.Item(cache_index).SubMatches(0)
		End Select
		cache_index=cache_index+1
	Else
		reg_bigmode=""
		cache_index=0
		next_page_str = parent_next_page_str
		parent_next_page_str=""
		If php_name="ranking.php" and ranking_page>0 and ranking_page<11 Then
			reg_bigmode="json"
			ranking_page=ranking_page+1
			If ranking_page<11 then parent_next_page_str = "1|inet|10,13|" & root_str & "/" & ranking_url & "&p=" & ranking_page	
		End If
	End If
End Function
'---------------------------------------------
Function format_ranking_html(ByVal html_str)
On Error Resume Next
    format_ranking_html = html_str
		'{"contents":[{"title":"\u9759\u8b10\u3061\u3083\u3093",
		'"date":"2017\u5e7401\u670830\u65e5 00:10",
		'"tags":["Fate\/GrandOrder"],
		'"url":"http:\/\/i3.pixiv.net\/c\/240x480\/img-master\/img\/2017\/01\/30\/00\/10\/10\/61183086_p0_master1200.jpg",
		'"illust_type":"0",
		'"illust_book_style":"0",
		'"illust_page_count":"1",
		'"user_name":"\u3057\u3089\u3073",
		'"profile_img":"http:\/\/i2.pixiv.net\/user-profile\/img\/2016\/11\/04\/06\/07\/50\/11706125_fcc9cf69109f56fe4dd6faaaafc8b9c7_50.jpg",
		'"illust_content_type":{"sexual":0,"lo":false,"grotesque":false,"violent":false,"homosexual":false,"drug":false,"thoughts":false,"antisocial":false,"religion":false,"original":false,"furry":false,"bl":false,"yuri":false},
		'"illust_id":61183086,
		'"width":855,
		'"height":960,
		'"user_id":216403,
		'"rank":1,
		'"yes_rank":3,
		'"total_score":41373,
		'"view_count":89096,
		'"illust_upload_timestamp":1485702610,
		'"attr":""
		'},{
    '转换为
		'{
		'"tags":[],
		'"url":"http:\/\/i1.pixiv.net\/img-inf\/img\/2015\/01\/08\/00\/31\/35\/48053836_s.jpg",
		'"user_name":"\u3042\u3065\u306a",
		'"illust_id":"48053836",
		'"illust_title":"Happy new year\u3010\u30ea\u30f4\u30a1\u30a4\u73ed\u3011",
		'"illust_user_id":"8500710",
		'"illust_restrict":"0",
		'"illust_x_restrict":"0",
		'"illust_type":"2"
		'}

		Dim split_str, matches(5)
    If InStr(html_str, ":[{") > 0 Then
        html_str = Mid(html_str, InStr(LCase(html_str), ":[{") + Len(":[{"))
    Else
        format_ranking_html = ""
        Exit Function
    End If
    split_str = Split(html_str, "},{")
    For i = 0 To UBound(split_str)
    		matches(0) = ""
    		matches(1) = ""
    		matches(2) = ""
    		matches(3) = ""
    		matches(4) = ""
    		matches(5) = ""
    		'illust_id
    		matches(0)=Mid(split_str(i), InStr(LCase(split_str(i)), """illust_id"":") + Len("""illust_id"":"))
    		matches(0)=Mid(matches(0),1,InStr(matches(0),",")-1)
    		'url
    		matches(1)=Mid(split_str(i), InStr(LCase(split_str(i)), """url"":""") + Len("""url"":"""))
    		matches(1)=Mid(matches(1),1,InStr(matches(1),"""")-1)
    		'title
    		matches(2)=Mid(split_str(i), InStr(LCase(split_str(i)), """title"":""") + Len("""title"":"""))
    		matches(2)=Mid(matches(2),1,InStr(matches(2),"""")-1)
    		'user_id
    		matches(3)=Mid(split_str(i), InStr(LCase(split_str(i)), """user_id"":") + Len("""user_id"":"))
    		matches(3)=Mid(matches(3),1,InStr(matches(3),",")-1)
    		'user_name
    		matches(4)=Mid(split_str(i), InStr(LCase(split_str(i)), """user_name"":""") + Len("""user_name"":"""))
    		matches(4)=Mid(matches(4),1,InStr(matches(4),"""")-1)
    		'illust_type
    		matches(5)=Mid(split_str(i), InStr(LCase(split_str(i)), """illust_type"":""") + Len("""illust_type"":"""))
    		matches(5)=Mid(matches(5),1,InStr(matches(5),"""")-1)
    		split_str(i)="{""tags"":[],""url"":""" & matches(1) & """,""user_name"":""" & matches(4) & """,""illust_id"":""" & matches(0) & """,""illust_title"":""" & matches(2) & """,""illust_user_id"":""" & matches(3) & """,""illust_type"":""" & matches(5) & """}"
    Next
    format_ranking_html ="[" & Join(split_str, ",") & "]"
End Function
'----------------------------------------------------------------------------------------------------
Function add_download_list_big(byval big_str, ByRef download_list)
On Error Resume Next
	Dim file_ID,file_name,file_Url
	'data-src="http://i2.pixiv.net/img-original/img/2015/01/09/00/04/21/48069269_p0.jpg" class="original-image">
	big_str=mid(big_str,1,InStr(LCase(big_str),"class=""original-image"""))
	big_str=mid(big_str,InStrrev(LCase(big_str),"data-src=""")+10)
	file_Url=mid(big_str,1,instr(big_str,"""")-1)
	file_Url=Cls_Chr63(file_Url)
	big_str=mid(file_Url,InStrrev(file_Url,"."))
	
	file_ID=matches_cache.Item(cache_index-1).SubMatches(0)
	file_name=fix_Unicode_Name(matches_cache.Item(cache_index-1).SubMatches(1))
	file_name="(pid-" & file_ID & ")" & rename_utf8(file_name) & big_str
	download_list = "|" & file_Url & "?" & (CDbl(Now()) * 10000000000) & "|" & file_name & "|" & vbCrLf
End Function
'----------------------------------------------------------------
Function add_download_list_manga(byval manga_str, ByRef download_list)
On Error Resume Next
	Dim file_ID,file_name,file_Url,manga_i
	
	'<img src="http://i1.pixiv.net/img-original/img/2015/01/08/02/13/17/48055260_p0.jpg" onclick="(window.open('', '_self')).close()">
	manga_str=mid(manga_str,InStr(LCase(manga_str),"<img src=""")+10)
	file_Url=mid(manga_str,1,instr(manga_str,"""")-1)
	file_Url=Cls_Chr63(file_Url)
	manga_str=mid(file_Url,InStrrev(file_Url,"."))
	file_Url=mid(file_Url,1,InStrrev(file_Url,"_p")+1)
	
	'原漫画改版前后分界线为ID=11319931（最新改版带img-original无big）
	'11319936_big_p0.jpg 06/16/2010 20:43--------------11319930_p0.jpg 06/16/2010 20:43
	file_ID=matches_cache.Item(cache_index-1).SubMatches(0)
	file_name=fix_Unicode_Name(matches_cache.Item(cache_index-1).SubMatches(1))
	file_name="(pid-" & file_ID & ")" & rename_utf8(file_name)
	If right(LCase(file_Url),6)<>"_big_p" Then
		file_name=file_name & "_p"
	Else
		file_name=file_name & "_big_p"
	End If
	For manga_i=0 to manga_count-1
		download_list =download_list & "|" & file_Url & manga_i & manga_str & "?" & (CDbl(Now()) * 10000000000) & "|" & file_name & manga_i & manga_str & "|" & vbCrLf
	Next
End Function
'----------------------------------------------------------------
Function add_download_list_ugoira(byval ugoira_str, ByRef download_list)
On Error Resume Next
		'pixiv.context.illustId         = "44387029";
		'pixiv.context.illustTitle      = "Hello ミク";pixiv.context.userId           = "395595";
		'pixiv.context.userName         = "KD"
		'{"src":"http:\/\/i2.pixiv.net\/img-zip-ugoira\/img\/2014\/06\/29\/14\/08\/25\/44387029_ugoira1920x1080.zip"
		Dim file_ID,file_name,file_Url,file_description
		
		ugoira_str=mid(ugoira_str,1,instr(LCase(ugoira_str),LCase("_ugoira1920x1080.zip"))) & "ugoira1920x1080.zip"
		
		file_Url=Mid(ugoira_str,InStrrev(ugoira_str,chr(34))+1)
		file_Url=replace(file_Url,"\/","/")
		
		file_ID=mid(ugoira_str,InStr(LCase(ugoira_str),LCase("pixiv.context.illustId")))
		file_ID=mid(file_ID,InStr(file_ID,"""")+1)
		file_ID=mid(file_ID,1,InStr(file_ID,"""")-1)

		file_description=mid(ugoira_str,InStr(LCase(ugoira_str),LCase("pixiv.context.illustTitle")))
		file_description=mid(file_description,InStr(file_description,"""")+1)
		file_description=mid(file_description,1,InStr(file_description,"""")-1)
		file_description=fix_Unicode_Name(file_description)
		file_name=file_description
		
		If Len(file_name)>200 Then file_name=left(file_name,200)
		file_name="(pid-" & file_ID & ")" & rename_utf8(file_name) & "_ugoira1920x1080.zip"
		
		download_list = "zip|" & file_Url & "?" & (CDbl(Now()) * 10000000000) & "|" & file_name & "|" & file_description & vbCrLf
End Function

Function Cls_Chr63(ByVal file_url)
On Error Resume Next
Dim file_type
	file_type=""
	file_type=mid(file_Url,InStrrev(file_Url,"."))
	If InStrrev(file_type,"?")>2 Then file_url=mid(file_Url,1,InStrrev(file_Url,"?")-1)
	Cls_Chr63 = file_url
End Function
'---------------------------------------------------------------------------------------------
Function check_login(ByVal html_str)
On Error Resume Next
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
On Error Resume Next
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
On Error Resume Next
	Dim FTH_regex
	Set FTH_regex = new RegExp
	FTH_regex.Global = True
	FTH_regex.IgnoreCase = True
	'<img src="http://i2.pixiv.net/img-inf/img/2011/02/11/23/03/36/16592817_s.jpg" alt="" class="_thumbnail ui-scroll-view" data-filter="thumbnail-filter" data-src="http://i2.pixiv.net/img-inf/img/2011/02/11/23/03/36/16592817_s.jpg" data-tags="大股開き" data-user-id="465458">
	FTH_regex.Pattern = "<img[^>]*?src=""([^""]+)""[^>]*?data-src=""([^""]+)""[^>]*?>"
	html_str=FTH_regex.replace(html_str,"<img src=""$2"" alt="""">")
	'<h2><a href="member_illust.php?mode=medium&amp;illust_id=44501062">オオダマ</a></h2>
	'--><h1 class="title" title="オオダマ">オオダマ</h1>
	FTH_regex.Pattern = "<h2><a[^>]*?>((?:(?!</a>).)*)</a></h2>"
	html_str=FTH_regex.replace(html_str,"<h1 class=""title"" title=""$1"">$1</h1>")
	format_transparent_html=html_str
End Function

'---------------------------------------------------------------------------------------------
Function fix_Unicode_Name(ByVal sLongFileName)
On Error Resume Next
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
        fix_Unicode_Name = replace(fix_Unicode_Name,"\/","/") 
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
    
