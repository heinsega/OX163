'2012-12-22 163.shanhaijing.net
Dim vid
Function return_download_url(ByVal url_str)
On Error Resume Next
'http://www.moeyo.com/2012/04/review_beat_ikkitousen_chouun_1.html
return_download_url = "inet|10,13|" & url_str
End Function

'--------------------------------------------------------

Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim key_word, page_title, file_type
return_download_list = ""

key_word="<div class=""entry-body"">"
If InStr(LCase(html_str),LCase(key_word)) > 0 Then
	'page title
	key_word="<h2 class=""entry-header"">"
	page_title=mid(html_str,InStr(LCase(html_str),LCase(key_word))+len(key_word))
	page_title=mid(page_title,1,InStr(LCase(page_title),"</h2>")-1)
	page_title=Trim(replace(page_title,"|","_"))
	
	key_word="<div class=""entry-body"">"
	html_str=mid(html_str,InStr(LCase(html_str),LCase(key_word)))
	key_word="<div class=""entry-footer"">"
	html_str=mid(html_str,1,InStr(LCase(html_str),LCase(key_word)))

	Dim regex, matches
	Set regex = new RegExp
	regex.Global = True
	'<a href="http://www2.moeyo.com/img/12/04/07/review_beat_ikkitousen_chouun/001.html" target="_blank"><img src="http://www2.moeyo.com/img/12/04/07/review_beat_ikkitousen_chouun/s001.jpg" width="500" height="748" alt="¥µ¥à¥Í¥¤¥ë»­Ïñ£º¥¯¥ê¥Ã¥¯¤Ç’ˆ´ó±íÊ¾" class="pict" /></a>
	regex.Pattern = "<a[^>]*href=""http://[^""]*.html""[^>]*target=""_blank""[^>]*>\s*<img src=""(http://[^""]*)""[^>]*class=""pict""[^>]*>\s*</a>"
	Set matches = regex.Execute(html_str)
	If matches.Count >0 Then
		For Each match In matches
			url_str=""
			url_str=match.SubMatches(0)
			
			'http://www2.moeyo.com/img/12/04/07/review_beat_ikkitousen_chouun/stop.jpg
			'http://www2.moeyo.com/img/12/04/07/review_beat_ikkitousen_chouun/s001.jpg
			'http://www2.moeyo.com/img/12/04/07/review_beat_ikkitousen_chouun/s048.jpg
			'http://www2.moeyo.com/img/11/01/11/comike79_cosplay_tonacos_highlights/cosplay_s061.jpg
			key_word=""
			key_word=mid(url_str,instrrev(url_str,"/")+1)
			url_str=mid(url_str,1,instrrev(url_str,"/"))
			file_type="jpg"
			file_type=mid(key_word,instrrev(key_word,"."))
			key_word=mid(key_word,1,instrrev(key_word,".")-1)
						
			If LCase(Right(key_word,4))<>"stop" Then
				html_str=""
				html_str=Right(key_word,1)
				key_word=left(key_word,Len(key_word)-1)
				Do While isnumeric(html_str)
					html_str=Right(key_word,1) & html_str
					key_word=left(key_word,Len(key_word)-1)
				loop
				If lcase(Left(html_str,1))="s" Then html_str=mid(html_str,2)
				key_word=key_word & html_str
				return_download_list = return_download_list & "|" & url_str & key_word & file_type & "|" & page_title & key_word & file_type & "|" & page_title & key_word & vbcrlf
			End If
		Next
	End if
End If
End Function