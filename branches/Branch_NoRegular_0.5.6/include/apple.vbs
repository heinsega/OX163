'2009-8-28 163.shanhaijing.net

Function return_download_url(ByVal url_str)
'http://www.apple.com/trailers/wb/wherethewildthingsare/
On Error Resume Next
return_download_url = "inet|10,13|" & url_str & "|User-Agent: QuickTime/7.6.2 (qtver=7.6.2;os=Windows NT 5.1Service Pack 2)"
End Function

'--------------------------------------------------------

Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""

If InStr(LCase(html_str), "<a class=""hd""") > 0 Then

html_str = Mid(html_str, InStr(LCase(html_str), "<a class=""hd""") + 13)

split_str = Split(html_str, "<a class=""hd""", -1, 1)
'<a class="hd" href="http://movies.apple.com/movies/wb/wherethewildthingsare/wtwta-tlr2_480p.mov" class="480p">
    For split_i = 0 To UBound(split_str)
    'url
    split_str(split_i) = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), "href=""")+6)
    split_str(split_i) = Mid(split_str(split_i),1,InStr(LCase(split_str(split_i)), Chr(34))-1)
    url_str=Mid(split_str(split_i),InStrrev(split_str(split_i), "_")+1)
    split_str(split_i) = Mid(split_str(split_i),1,InStrrev(split_str(split_i), "_"))
    split_str(split_i)=split_str(split_i) & "h" & url_str
    
    'name
    html_str=Mid(split_str(split_i),InStrrev(split_str(split_i), "/")+1)
    
    return_download_list = return_download_list & "|" & split_str(split_i) & "|" & html_str & "|" & vbCrLf
    Next
End If

return_download_list = return_download_list & "0"

End Function