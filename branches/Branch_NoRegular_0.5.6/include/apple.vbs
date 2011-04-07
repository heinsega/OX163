'2010-8-26 163.shanhaijing.net

Function return_download_url(ByVal url_str)
'http://www.apple.com/trailers/wb/wherethewildthingsare/
'http://trailers.apple.com/trailers/paramount/wtc/
On Error Resume Next
return_download_url = "web|10,13|" & url_str & "|User-Agent: QuickTime/7.6.2 (qtver=7.6.2;os=Windows NT 5.1Service Pack 2)"
End Function

'--------------------------------------------------------

Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""

If InStr(LCase(html_str), LCase("<H4>HD</H4>")) > 0 Then

html_str = Mid(html_str, InStr(LCase(html_str), LCase("<H4>HD</H4>")))
html_str = Mid(html_str, InStr(LCase(html_str), LCase("class=hd"))+8)
split_str = Split(html_str, "class=hd", -1, 1)
'<A class=hd style="FILTER: progid:DXImageTransform.Microsoft.AlphaImageLoader(src='http://trailers.apple.com/trailers/images/hud_button_square.png',sizingMethod='crop'); BACKGROUND-IMAGE: none" href="http://trailers.apple.com/movies/paramount/world_trade_center/world_trade_center-tlr1_h480p.mov" s_oc="null">
    For split_i = 0 To UBound(split_str)
    'url
    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "href=""")+6)
    split_str(split_i) = Mid(split_str(split_i),1,InStr(LCase(split_str(split_i)), Chr(34))-1)
    
    'name
    html_str=Mid(split_str(split_i),InStrrev(split_str(split_i), "/")+1)
    
    return_download_list = return_download_list & "|" & split_str(split_i) & "|" & html_str & "|" & vbCrLf
    Next
End If

return_download_list = return_download_list & "0"

End Function