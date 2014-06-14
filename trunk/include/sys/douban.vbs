'2014-6-15 163.shanhaijing.net
Dim retry_time, page, album_ID, html_type, nav_item, nav_str

Function return_download_url(ByVal url_str)
    On Error Resume Next
    return_download_url = ""
    retry_time = 0
    
    If InStr(LCase(url_str), "http://www.douban.com/photos/") = 1 Then
        'http://www.douban.com/photos/album/11028188/?start=18
        If InStr(LCase(url_str), "http://www.douban.com/photos/album/") = 1 Then
            url_str = Mid(url_str, InStr(LCase(url_str), "http://www.douban.com/photos/album/") + Len("http://www.douban.com/photos/album/"))
            If InStr(url_str, "/") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "/") - 1)
            If InStr(url_str, "?") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "?") - 1)
            If InStr(url_str, "#") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "#") - 1)
            If IsNumeric(url_str) And Len(url_str) > 0 Then
                album_ID = "http://www.douban.com/photos/album/" & url_str & "/"
                html_type = "album"
                page = 1
                return_download_url = "inet|10,13|" & album_ID
            End If
        ElseIf InStr(LCase(url_str), "http://www.douban.com/photos/photo/") = 1 Then
            'http://www.douban.com/photos/photo/125705371/
            'http://www.douban.com/photos/photo/125716028/#next_photo
            album_ID = url_str
            html_type = "photo"
            page = 1
            return_download_url = "inet|10,13|" & album_ID
        End If
        
    ElseIf InStr(LCase(url_str), "http://site.douban.com/") = 1 And (InStr(LCase(url_str), "/widget/photos/") > 1 Or InStr(LCase(url_str), "/widget/public_album/") > 1) Then
        'http://site.douban.com/151879/widget/photos/16469008/
        'http://site.douban.com/106689/widget/public_album/6112058/
        If InStr(LCase(url_str), "/widget/photos/") > 1 Then
            album_ID = Mid(url_str, 1, InStr(LCase(url_str), "/widget/photos/") + 15 - 1)
            url_str = Mid(url_str, InStr(LCase(url_str), "/widget/photos/") + 15)
        ElseIf InStr(LCase(url_str), "/widget/public_album/") > 1 Then
            album_ID = Mid(url_str, 1, InStr(LCase(url_str), "/widget/public_album/") + 21 - 1)
            url_str = Mid(url_str, InStr(LCase(url_str), "/widget/public_album/") + 21)
        End If
        If InStr(url_str, "/") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "/") - 1)
        If InStr(url_str, "?") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "?") - 1)
        If InStr(url_str, "#") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "#") - 1)
        If IsNumeric(url_str) And Len(url_str) > 0 Then
            album_ID = album_ID & url_str & "/"
            html_type = "site_album"
            page = 1
            return_download_url = "inet|10,13|" & album_ID
        End If
        
    ElseIf InStr(LCase(url_str), "http://site.douban.com/") = 1 Then
        '多分类
        'http://site.douban.com/151879/
        '多页room
        'http://site.douban.com/106689/room/545141/
        html_type = "site_room"
        page = 0
        If MsgBox("您是否要尝试下载该小站所有分类的相册？" & vbCrLf & "是:尝试下载所有分类相册" & vbCrLf & "否:仅下载当前分类页面", vbYesNo, "询问") = vbYes Then
        	page = -1
        	If InStr(LCase(url_str), "/room/") > 1 Then url_str=mid(url_str,1,InStr(LCase(url_str), "/room/"))
        End If
        If InStr(url_str, "?") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "?") - 1)
        If InStr(url_str, "#") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "#") - 1)
        album_ID = url_str
        return_download_url = "inet|10,13|" & album_ID
        
    ElseIf InStr(LCase(url_str), "http://www.douban.com/people/") = 1 Then
        'http://www.douban.com/people/royzhong/
        'http://www.douban.com/people/royzhong/notes
        'http://www.douban.com/people/royzhong/photos
        'http://www.douban.com/people/royzhong/photos?start=32
        url_str = Mid(url_str, InStr(LCase(url_str), "http://www.douban.com/people/") + Len("http://www.douban.com/people/"))
        If InStr(url_str, "/") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "/") - 1)
        If InStr(url_str, "?") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "?") - 1)
        If InStr(url_str, "#") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "#") - 1)
        album_ID = "http://www.douban.com/people/" & url_str & "/photos"
        page = 1
        html_type = "people"
        return_download_url = "inet|10,13|" & album_ID
    End If

OX163_urlpage_Referer = "http://www.douban.com/" & vbcrlf & "User-Agent: Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0)"

End Function

'--------------------------------------------------------
Function Get_album_ID(ByVal url_str)
    On Error Resume Next
    Get_album_ID = ""
    If InStr(LCase(url_str), "http://www.douban.com/photos/album/") = 1 Then
        url_str = Mid(url_str, InStr(LCase(url_str), "http://www.douban.com/photos/album/") + Len("http://www.douban.com/photos/album/"))
        If InStr(url_str, "/") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "/") - 1)
        If InStr(url_str, "?") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "?") - 1)
        If InStr(url_str, "#") > 1 Then url_str = Mid(url_str, 1, InStr(url_str, "#") - 1)
        If IsNumeric(url_str) And Len(url_str) > 0 Then Get_album_ID = url_str
    End If
End Function

'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
    On Error Resume Next
    return_albums_list = ""
    Dim key_word, split_str, album_title(3)
    
    If page=-1 Then
    	url_str=html_str
    	nav_item=0
    	key_word="<div class=""nav-items"">"
    	If InStr(LCase(url_str), key_word) > 0 Then
    		url_str=Mid(url_str, InStr(LCase(url_str), LCase(key_word)) + Len(key_word))
    		url_str=Mid(url_str, InStr(url_str, "<a href=""") + 9)
    		url_str=Mid(url_str,1, InStr(url_str, "</div>") - 1)
    		nav_str=split(url_str,"<a href=""")
    		For ni = 0 To UBound(nav_str)
    			nav_str(ni)=mid(nav_str(ni),1,instr(nav_str(ni),"""")-1)
    		Next
    		If ubound(nav_str)>0 Then nav_item=1
    	End If
    page=0
    End If
    
    If html_type = "site_room" And InStr(LCase(html_str), "<div class=""mod""") > 0 Then
        retry_time = 0
        url_str = html_str
        If page = 0 Then
            key_word = "<div class=""mod"" id="""
        Else
            key_word = "<div class=""mod"" archives=""1"" id="""
        End If
        html_str = Mid(html_str, InStr(LCase(html_str), LCase(key_word)) + Len(key_word))
        split_str = Split(html_str, key_word)
        
        For i = 0 To UBound(split_str)
            key_word = ""
            pic_title = ""
            album_title(0) = ""
            album_title(1) = ""
            album_title(2) = ""
            If InStr(split_str(i), "public_album-") = 1 Or InStr(split_str(i), "photos-") = 1 Then
                'album ID
                key_word = Mid(split_str(i), InStr(split_str(i), "-") + 1)
                key_word = Mid(key_word, 1, InStr(key_word, """") - 1)
                
                If IsNumeric(key_word) And Len(key_word) > 0 Then
                    'title
                    album_title(0) = Mid(split_str(i), InStr(split_str(i), "<span>") + 6)
                    If page>0 Then 
                    	album_title(0) = Mid(album_title(0), InStr(album_title(0), "<span>") + 6)
                    	album_title(0) = Mid(album_title(0), InStr(album_title(0), ">") + 1)
                    	album_title(0) = Mid(album_title(0), 1, InStr(album_title(0), "</a>") - 1)
                    Else
                    	album_title(0) = Mid(album_title(0), 1, InStr(album_title(0), "</span>") - 1)
                    End if
                    album_title(0) = Trim(Replace(album_title(0), "|", "｜"))
                    album_title(0) = "AID" & key_word & "-" & album_title(0)
                    
                    'url
                    album_title(1) = Mid(split_str(i), InStr(split_str(i), " href=""") + 7)
                    album_title(1) = Mid(album_title(1),1,InStr(album_title(1), """") - 1)
                    'pic number
                    album_title(2) = Mid(split_str(i), InStr(split_str(i), "<p class=""rec-num"">") + Len("<p class=""rec-num"">"))
                    album_title(2) = Mid(album_title(2), 1, InStr(album_title(2), "张照片") - 1)
                    album_title(2) = Trim(album_title(2))
                    If IsNumeric(album_title(2)) = False Then album_title(2) = ""
                    
                    return_albums_list = return_albums_list & "0|" & album_title(2) & "|" & album_title(1) & "|" & album_title(0) & "|" & vbCrLf
                End If
            End If
        Next
        
        If page = 0 Then
            key_word = "<div class=""mod"" id=""div_archives"""
        ElseIf page>0 Then
        		key_word = "<link rel=""next"" href="""
        End If
        
        If InStr(LCase(url_str), LCase(key_word)) > 0 Then
            url_str = Mid(url_str, InStr(LCase(url_str), LCase(key_word)) + Len(key_word))
            If page = 0 Then url_str = Mid(url_str, InStr(LCase(url_str), "<a href=""") + 9)
            url_str = Mid(url_str, 1, InStr(LCase(url_str), """") - 1)
            page = page + 1
            album_ID = url_str
            return_albums_list = return_albums_list & "1|inet|10,13|" & url_str
        ElseIf nav_item>0 Then
        		If nav_item<=ubound(nav_str) Then
        			return_albums_list = return_albums_list & "1|inet|10,13|" & nav_str(nav_item)
        			page = 0
        			nav_item=nav_item+1
        		End If
        End If
        
    ElseIf page > 0 And html_type = "people" And InStr(LCase(html_str), "<a class=""album_photo"" href=""") > 0 Then
        retry_time = 0
        url_str = html_str
        key_word = "<div class=""albumlst"">"
        html_str = Mid(html_str, InStr(LCase(html_str), LCase(key_word)) + Len(key_word))
        split_str = Split(html_str, key_word)
        
        For i = 0 To UBound(split_str)
            key_word = ""
            pic_title = ""
            album_title(0) = ""
            album_title(1) = ""
            album_title(2) = ""
            'album ID
            key_word = Mid(split_str(i), InStr(split_str(i), "<a class=""album_photo"" href=""") + Len("<a class=""album_photo"" href="""))
            key_word = Mid(key_word, 1, InStr(key_word, """") - 1)
            key_word = Get_album_ID(key_word)
            
            If IsNumeric(key_word) And Len(key_word) > 0 Then
                'title
                album_title(0) = Mid(split_str(i), InStr(split_str(i), "<div class=""pl2"">") + 25)
                album_title(0) = Mid(album_title(0), InStr(album_title(0), ">") + 1)
                album_title(0) = Mid(album_title(0), 1, InStr(album_title(0), "</a>") - 1)
                album_title(0) = Trim(Replace(album_title(0), "|", "｜"))
                album_title(0) = "AID" & key_word & "-" & album_title(0)
                
                'desc
                album_title(1) = Mid(split_str(i), InStr(split_str(i), "<div class=""albumlst_descri"">") + Len("<div class=""albumlst_descri"">"))
                album_title(1) = Mid(album_title(1), 1, InStr(album_title(1), "</div>") - 1)
                
                'pic number
                album_title(2) = Mid(split_str(i), InStr(split_str(i), "<span class=""pl"">") + Len("<span class=""pl"">"))
                album_title(2) = Mid(album_title(2), 1, InStr(album_title(2), "张照片") - 1)
                album_title(2) = Trim(album_title(2))
                If IsNumeric(album_title(2)) = False Then album_title(2) = ""
                
                return_albums_list = return_albums_list & "0|" & album_title(2) & "|http://www.douban.com/photos/album/" & key_word & "/|" & album_title(0) & "|" & album_title(1) & vbCrLf
            End If
        Next
        
        key_word = "<link rel=""next"" href="""
        If InStr(LCase(url_str), LCase(key_word)) > 0 Then
            url_str = Mid(url_str, InStr(LCase(url_str), LCase(key_word)) + Len(key_word))
            url_str = Mid(url_str, 1, InStr(LCase(url_str), """") - 1)
            page = page + 1
            album_ID = url_str
            return_albums_list = return_albums_list & "1|inet|10,13|" & url_str
        End If
        
        'ElseIf Then
        'retry_time=retry_time+1
        'return_download_list="1|inet|10,13|" & album_ID
    End If
    
    
End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
    On Error Resume Next
    'http://otho.douban.com/view/photo/thumb/xvvoA0chC9-RG7NT0eh9pw/x1159718333.jpg
    'http://otho.douban.com/view/photo/photo/afk-0HXr3gp8fXp3ehM_8g/x1159718333.jpg
    'http://img3.douban.com/view/photo/thumb/public/p1159718333.jpg
    'http://img3.douban.com/view/photo/photo/public/p1159718333.jpg
    return_download_list = ""
    Dim key_word, split_str, pic_title
    If page > 0 And html_type = "site_album" And InStr(LCase(html_str), "<div class=""photo-item"">") > 0 Then
        retry_time = 0
        url_str = html_str
        key_word = "<div class=""photo-item"">"
        html_str = Mid(html_str, InStr(LCase(html_str), LCase(key_word)) + Len(key_word))
        split_str = Split(html_str, key_word)
        For i = 0 To UBound(split_str)
            key_word = ""
            pic_title = ""
            'title
            pic_title = Mid(split_str(i), InStr(split_str(i), "title=""") + Len("title="""))
            pic_title = Mid(pic_title, 1, InStr(pic_title, """") - 1)
            'url
            key_word = Mid(split_str(i), InStr(split_str(i), "<img src=""") + Len("<img src="""))
            key_word = Mid(key_word, 1, InStr(key_word, """") - 1)
            'http://img5.douban.com/view/photo/thumb/public/p1244052979.jpg
            '=>http://img3.douban.com/view/photo/photo/public/p1211013910.jpg
            key_word = Replace(key_word, "/thumb/", "/photo/")
            If Left(key_word, Len("http://img1.")) <> "http://img1." Then key_word = "http://img1." & Mid(key_word, InStr(key_word, ".") + 1)
            If Len(key_word) > 0 Then
                'file name
                split_str(i) = Mid(key_word, InStrRev(key_word, "/") + 1)
                return_download_list = return_download_list & "|" & key_word & "|" & split_str(i) & "|" & pic_title & vbCrLf
            End If
        Next
        
        key_word = "<link rel=""next"" href="""
        If InStr(LCase(url_str), LCase(key_word)) > 0 Then
            url_str = Mid(url_str, InStr(LCase(url_str), LCase(key_word)) + Len(key_word))
            url_str = Mid(url_str, 1, InStr(LCase(url_str), """") - 1)
            If InStr(LCase(url_str), LCase("http://site.douban.com")) <> 1 Then url_str = "http://site.douban.com" & url_str
            page = page + 1
            album_ID = url_str
            return_download_list = return_download_list & "1|inet|10,13|" & url_str
        End If
        
    ElseIf page > 0 And html_type = "album" And InStr(LCase(html_str), "<div class=""photo_wrap"">") > 0 Then
        retry_time = 0
        url_str = html_str
        key_word = "<div class=""photo_wrap"">"
        html_str = Mid(html_str, InStr(LCase(html_str), LCase(key_word)) + Len(key_word))
        split_str = Split(html_str, key_word)
        For i = 0 To UBound(split_str)
            key_word = ""
            pic_title = ""
            'title
            pic_title = Mid(split_str(i), InStr(split_str(i), "title=""") + Len("title="""))
            pic_title = Mid(pic_title, 1, InStr(pic_title, """") - 1)
            'url
            key_word = Mid(split_str(i), InStr(split_str(i), "<img src=""") + Len("<img src="""))
            key_word = Mid(key_word, 1, InStr(key_word, """") - 1)
            key_word = Replace(key_word, "/thumb/", "/photo/")
            If Left(key_word, Len("http://img1.")) <> "http://img1." Then key_word = "http://img1." & Mid(key_word, InStr(key_word, ".") + 1)
            If Len(key_word) > 0 Then
                'file name
                split_str(i) = Mid(key_word, InStrRev(key_word, "/") + 1)
                return_download_list = return_download_list & "|" & key_word & "|" & split_str(i) & "|" & pic_title & vbCrLf
            End If
        Next
        
        key_word = "<link rel=""next"" href="""
        If InStr(LCase(url_str), LCase(key_word)) > 0 Then
            url_str = Mid(url_str, InStr(LCase(url_str), LCase(key_word)) + Len(key_word))
            url_str = Mid(url_str, 1, InStr(LCase(url_str), """") - 1)
            page = page + 1
            album_ID = url_str
            return_download_list = return_download_list & "1|inet|10,13|" & url_str
        End If
        
    ElseIf page = 0 And html_type = "photo" And InStr(LCase(html_str), "<span class=""rr"">") > 0 Then
        retry_time = 0
        page = 1
        html_type = "album"
        html_str = Mid(html_str, InStr(LCase(html_str), "<span class=""rr"">"))
        html_str = Mid(html_str, InStr(LCase(html_str), "<a href=""") + 9)
        html_str = Mid(html_str, 1, InStr(html_str, """") - 1)
        album_ID = html_str
        return_download_list = "1|inet|10,13|" & album_ID
        
    ElseIf retry_time < 4 Then
        retry_time = retry_time + 1
        return_download_list = "1|inet|10,13|" & album_ID
    End If
End Function