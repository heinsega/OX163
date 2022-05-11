'2012-10-10 http://www.shanhaijing.net/163
Dim page_num,user_id,Page_utl
Function return_download_url(ByVal url_str)
On Error Resume Next
return_download_url = ""
page_num = "1"
user_id = ""
Page_utl = ""
If InStr(LCase(url_str), "media/set/") > 0 Then

		return_download_url = "inet|10,13|" & url_str

	ElseIf InStr(LCase(url_str), "photos_albums") > 0 Then
    
		return_download_url = "inet|10,13|" & url_str
Else
    	return_download_url = "inet|10,13|" & url_str & "/photos_albums"

    End If
	
End Function

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list = ""
	
	If InStr(LCase(html_str), "<a class=""phototexttitle"" href=""") > 0 Then

	html_str = Mid(html_str,InStr(LCase(html_str),"<a class=""phototexttitle"" href=""")+32)

	Dim str_split
	str_split=split((LCase(html_str)),"<a class=""phototexttitle"" href=""")
	For i=0 to UBound(str_split)

		'url
		url_str=Mid(str_split(i),1,InStr(str_split(i),"""")-1)
		str_split(i)=Mid(str_split(i),InStr(str_split(i),"<strong>")+8)

		'name
		html_str=Trim(Mid(str_split(i),1,InStr(str_split(i),"</strong>")-1))
		If html_str="" Then html_str="No_Name_Album"
		str = html_str
		utf8_str = vbsUnEscape(str)
		rename_utf8=utf8_str
		html_str=rename_utf8(utf8_str)
		html_str=replace(html_str,"&#8231","")

		'pic number
		str_split(i)=Mid(str_split(i),InStr(str_split(i),""">")+2)
		str_split(i)=Mid(str_split(i),1,InStr(str_split(i)," ")-1)

		If IsNumeric(str_split(i))=false Then str_split(i)="0"

			return_albums_list = return_albums_list & "0|" & str_split(i) & "|" & url_str & "|" & html_str & vbcrlf

	Next

End If 

End Function
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim photo_split
return_download_list = ""

		If page_num < 2 Then
		user_id=Mid(html_str,InStr(html_str,"""user"":""")+8)
		user_id=Mid(user_id,1,InStr(user_id,"""")-1)
		End If 

		Page_utl=Mid(html_str,InStr(html_str,"TimelinePhotosAlbumPagelet")+30)
		Page_utl=Mid(Page_utl,1,InStr(Page_utl,"}"))

		Page_utl=replace(Page_utl,"\","")


If InStr(LCase(html_str),"amp;src=") > 0 Then

		html_str=Mid(html_str,InStr(html_str,"amp;src=")+8)
		html_str=Mid(html_str,1,InStr(html_str,"class=""fbTimelinePhotosScroller""")-1)
		photo_split=split(html_str,"amp;src=")
		For i=0 to UBound(photo_split)
			
			'url
			url_str=Mid(photo_split(i),1,InStr(photo_split(i),"&amp")-1)
			url_str=replace(url_str,"u00253A",":")
			url_str=replace(url_str,"u00252F","/")
			url_str=replace(url_str,"\","")
			url_str=replace(url_str,"%3A",":")
			url_str=replace(url_str,"%2F","/")

	If InStr(return_download_list,url_str) > 0 Then

			return_download_list=return_download_list
		Else
			return_download_list=return_download_list & "|" & url_str & "|" & "|" & vbcrlf

	End If 

	Next	

		If page_num > 1 Then
		If UBound(photo_split)<63 Then

		return_albums_list = return_albums_list & "0"
	Else
		page_num=page_num+1

		return_download_list= return_download_list & "1|inet|10,13|" & "http://www.facebook.com/ajax/pagelet/generic.php/TimelinePhotosAlbumPagelet?ajaxpipe=1&ajaxpipe_token=AXhfhYGef0zq6L_0&no_script_path=1&data=" & Page_utl & "&__user=" & user_id & "&__a=1&__adt=" &  page_num

		End If
		End If 

		If page_num < 2 Then
		If UBound(photo_split)<55 Then
		return_albums_list = return_albums_list & "0"
	Else
		page_num=page_num+1
		
		return_download_list= return_download_list & "1|inet|10,13|" & "http://www.facebook.com/ajax/pagelet/generic.php/TimelinePhotosAlbumPagelet?ajaxpipe=1&ajaxpipe_token=AXhfhYGef0zq6L_0&no_script_path=1&data=" & Page_utl & "&__user=" & user_id & "&__a=1&__adt=" &  page_num
		
		End If
		End If 

End If
End Function
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Function vbsUnEscape(str)
On Error Resume Next
    Dim i, s, c
    s = ""
    For i = 1 To Len(str)
    DoEvents
        c = Mid(str, i, 1)
        If Mid(str, i, 3) = "&#x" And i <= Len(str) - 7 And Mid(str, i + 7, 1) = ";" Then
            If IsNumeric("&H" & Mid(str, i + 3, 4)) Then
                s = s & ChrW(CInt("&H" & Mid(str, i + 3, 4)))
                i = i + 7
            Else
                s = s & c
            End If
        Else
            s = s & c
        End If
    Next
    vbsUnEscape = Replace(s, "\/", "/")
End Function

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
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
    rename_utf8 = replace(utf8_str, "|", "£ü")
End Function

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
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