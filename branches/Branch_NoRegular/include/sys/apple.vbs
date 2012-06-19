'2011-8-11 163.shanhaijing.net

Function return_download_url(ByVal url_str)
'http://www.apple.com/trailers/wb/wherethewildthingsare/
'http://trailers.apple.com/trailers/paramount/wtc/
On Error Resume Next
return_download_url = "inet|10,13|" & url_str & "/includes/playlists/web.inc|User-Agent: QuickTime/7.6.2 (qtver=7.6.2;os=Windows NT 5.1Service Pack 2)"
End Function

'--------------------------------------------------------

Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""

If InStr(LCase(html_str), LCase(".mov""")) > 0 Then

split_str = Split(html_str, ".mov""", -1, 1)
Dim end_i
		end_i=UBound(split_str)-1
    For split_i = 0 To end_i
    'url
    split_str(split_i) = Mid(split_str(split_i), InStrrev(LCase(split_str(split_i)), chr(34))+1)
    split_str(split_i) = split_str(split_i) & ".mov"
    split_str(split_i) = replace(split_str(split_i),"_480p.mov","_h480p.mov")
    split_str(split_i) = replace(split_str(split_i),"_720p.mov","_h720p.mov")
    split_str(split_i) = replace(split_str(split_i),"_1080p.mov","_h1080p.mov")
    'name
    html_str=Mid(split_str(split_i),InStrrev(split_str(split_i), "/")+1)
    
    return_download_list = return_download_list & "|" & split_str(split_i) & "|" & html_str & "|" & vbCrLf
    Next
End If

return_download_list = return_download_list & "0"

End Function