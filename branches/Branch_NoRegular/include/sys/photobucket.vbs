'2011-1-9 http://www.shanhaijing.net/163

'http://s132.photobucket.com/albums/q35/ShadowManipulation/Infinity%20Additions%20Febuary%202010/
'http://photobucket.com/images/nicole%20richie/?page=15
'http://smg.photobucket.com/home/nococ/index#!cpZZ1QQtppZZ36
'http://smg.photobucket.com/albums/v243/nococ/

Dim next_page, thumbnail_ID, Fix_Url

Function return_download_url(ByVal url_str)
next_page = 1
thumbnail_ID = -1
If MsgBox("是否查找该页面以后的所有页面媒体？" & vbcrlf & vbcrlf & "[是]全部页面" & vbcrlf & "[否]仅当前页", vbYesNo, "询问") = vbNo Then next_page = 0
If InStr(url_str,"#")>0 Then url_str=Mid(url_str,1,InStr(url_str,"#")-1)
Fix_Url=url_str
return_download_url = "inet|10,13|" & url_str
End Function

Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""
If next_page = 1 Then
    url_str = html_str
    url_str = Mid(url_str, 1, InStr(url_str, "tr('pagination_next_click');") - 1)
    url_str = Mid(url_str, InStrRev(url_str, "<a href=""") + Len("<a href="""))
    If InStr(LCase(url_str), "style=""display:none;""") > 0 Or InStr(LCase(url_str), "javascript:void(0)") > 0 Then
        url_str = ""
    Else
        url_str = Mid(url_str, 1, InStr(url_str, """") - 1)
        If Left(url_str,7)<>"http://" and Left(url_str,8)<>"https://" Then url_str=""
    End If
Else
    url_str = ""
End If

If InStr(html_str, "<div id=""thumbnail_") > 0 Then
    Dim split_str, url, media_ID, temp_str
    html_str = Mid(html_str, InStr(html_str, "<div id=""thumbnail_") + Len("<div id=""thumbnail_"))
    split_str = Split(html_str, "<div id=""thumbnail_")
    html_str = ""
    For i = 0 To UBound(split_str)
        temp_str = Mid(split_str(i), 1, InStr(split_str(i), """") - 1)
        If IsNumeric(temp_str) Then
            media_ID = Int(temp_str)
            temp_str = ""
        Else
            media_ID = 0
        End If
        If thumbnail_ID < media_ID Then
            thumbnail_ID = media_ID
            split_str(i) = photobucket_UnEscape(split_str(i))
            split_str(i) = Mid(split_str(i), InStr(split_str(i), ",'mediaUrl':'") + Len(",'mediaUrl':'"))
            url = Mid(split_str(i), 1, InStr(split_str(i), "'") - 1)
            split_str(i) = ""
            split_str(i) = Mid(url, InStrRev(url, "/") + 1)
            If split_str(i) = "" Then split_str(i) = "noname_file"
            return_download_list = return_download_list & "|" & url & "|" & rename_utf8(split_str(i)) & "|" & vbCrLf
        End If
    Next

If url_str <> "" Then
    return_download_list = return_download_list & "1|inet|10,13|" & url_str
Else
    return_download_list = return_download_list & "0"
End If
    
Else
    return_download_list = "0"
End If
End Function

Function rename_utf8(ByVal utf8_Str)
If Len(utf8_Str) = 0 Then Exit Function
For i = 1 To Len(utf8_Str)
    If Asc(Mid(utf8_Str, i, 1)) = 63 Then utf8_Str = Replace(utf8_Str, Mid(utf8_Str, i, 1), "_")
Next
rename_utf8 = utf8_Str
End Function

Function photobucket_UnEscape(str)
    str = Replace(str, "&#039;", "'")
    photobucket_UnEscape = Replace(str, "\/", "/")
End Function