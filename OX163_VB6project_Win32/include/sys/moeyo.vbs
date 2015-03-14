'2014-1-5 163.shanhaijing.net
Dim vid
Function return_download_url(ByVal url_str)
On Error Resume Next
'old
'http://www.moeyo.com/2012/04/review_beat_ikkitousen_chouun_1.html
'new
'http://www.moeyo.com/article/44648
return_download_url = "inet|10,13|" & url_str & "|http://www.moeyo.com/" & vbcrlf & "User-Agent: Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0)"
End Function

'--------------------------------------------------------

Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim key_word, page_title

return_download_list = ""

key_word="class=""colorbox"">"
If InStr(LCase(html_str),LCase(key_word)) > 0 Then
	'page title
	key_word="<title>"
	page_title=mid(html_str,InStr(LCase(html_str),LCase(key_word))+len(key_word))
	page_title=mid(page_title,1,InStr(LCase(page_title)," | moeyo.com</title>")-1)
	page_title=Trim(replace(page_title,"|","_"))
	
	Dim regex, matches
	Set regex = new RegExp
	regex.Global = True
	'old
	'<a href="http://www2.moeyo.com/img/12/04/07/review_beat_ikkitousen_chouun/001.html" target="_blank"><img src="http://www2.moeyo.com/img/12/04/07/review_beat_ikkitousen_chouun/s001.jpg" width="500" height="748" alt="¥µ¥à¥Í¥¤¥ë»­Ïñ£º¥¯¥ê¥Ã¥¯¤Ç’ˆ´ó±íÊ¾" class="pict" /></a>
	'regex.Pattern = "<a[^>]*href=""http://[^""]*.html""[^>]*target=""_blank""[^>]*>\s*<img src=""(http://[^""]*)""[^>]*class=""pict""[^>]*>\s*</a>"
	'http://www2.moeyo.com/img/12/04/07/review_beat_ikkitousen_chouun/stop.jpg
	'http://www2.moeyo.com/img/12/04/07/review_beat_ikkitousen_chouun/s001.jpg
	'http://www2.moeyo.com/img/11/01/11/comike79_cosplay_tonacos_highlights/cosplay_s061.jpg
	'new
	'<a href="http://cdn.moeyo.com/2013/0905/03/002.jpg"class="colorbox"><img src="http://cdn.moeyo.com/2013/0905/03/002s.jpg" width="180" height="120"></a>
	regex.Pattern = "<a[^>]*href=""(http://[^""]*.moeyo.com/[0-9]{4}/[0-9]{4}/[0-9]{1,}/[0-9]{3,}.[A-Za-z]{3,4})""[^>]*class=""colorbox""[^>]*>"
	Set matches = regex.Execute(html_str)
	If matches.Count >0 Then
		For Each match In matches
			url_str=""
			url_str=match.SubMatches(0)
			
			key_word=""
			key_word=mid(url_str,instrrev(url_str,"/")+1)
			
			return_download_list = return_download_list & "|" & url_str & "|" & page_title & key_word & "|" & page_title & key_word & vbcrlf
		Next
	End If
End If
End Function