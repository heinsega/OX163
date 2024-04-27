'2022-5-6 163.shanhaijing.net
Dim deep_DL, split_str, split_c0, split_c1
Dim tags, page, page_counter, url_instr, pool, url_head, url_http
Dim retry_time, retry_url, delay_time, start_time, delay_url
Function return_download_url(ByVal url_str)
    'idol.sankakucomplex.com
    'http://chan.sankakucomplex.com/
    'http://chan.sankakucomplex.com/post/
    'http://chan.sankakucomplex.com/?page=3
    ''http://chan.sankakucomplex.com/post/index
    ''http://chan.sankakucomplex.com/post?page=3
    
    'http://chan.sankakucomplex.com/post/index?tags=tagme
    'http://chan.sankakucomplex.com/?page=2&tags=chibi
    ''http://chan.sankakucomplex.com/post?tags=emma
    ''http://chan.sankakucomplex.com/post?page=5&tags=miko
    
    'http://chan.sankakucomplex.com/pool/show/596?page=3
	
    'http://chan.sankakucomplex.com/wiki/show?title=park_sung-woo
    'http://chan.sankakucomplex.com/wiki/show?page=3&title=park_sung-woo
    
    'http://chan.sankakucomplex.com/post/show/9506/cg-d-o-_-publisher-eigoukaiki-eroge-ino-tagme
    On Error Resume Next
    tags = ""
    retry_url = ""
    retry_time = 0
    page_counter = 0
    page = 1
    delay_time=1
    
    'idol.sankakucomplex.com
    'chan.sankakucomplex.com
    If InStr(LCase(url_str), "http:") = 1 Then
        url_http = "http:"
    Else
        url_http = "https:"
    End If
     
    'idol.sankakucomplex.com
    'chan.sankakucomplex.com
    If InStr(LCase(url_str), "//idol.sankakucomplex.com") > 1 Then
        url_head = url_http & "//idol"
    Else
        url_head = url_http & "//chan"
    End If
    
'---------------------------����ҳ��-----------------------------------------
    If InStr(LCase(url_str), "/post/show/") >0 Then
        pool = "post"
        return_download_url = "inet|10,13|" & url_str
        retry_url = return_download_url
    		return_download_url = return_download_url & "|User-agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.147 Safari/537.36"
				OX163_urlpage_Referer = "User-agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.147 Safari/537.36"
        Exit Function
    End If
'----------------------------------------------------------------------------


    Dim page_str
    page_str = ""
    If InStr(LCase(url_str), "?page=") > 10 Then
        page_str = Mid(url_str, InStr(LCase(url_str), "?page=") + 6)
        url_str = Mid(url_str, 1, InStr(LCase(url_str), "?page=") - 1)
        If InStr(page_str, "&") > 0 Then
            url_str = url_str & "?" & Mid(page_str, InStr(page_str, "&") + 1)
            page_str = Mid(page_str, 1, InStr(page_str, "&") - 1)
        End If
    ElseIf InStr(LCase(url_str), "&page=") > 10 Then
        page_str = Mid(url_str, InStr(LCase(url_str), "&page=") + 6)
        url_str = Mid(url_str, 1, InStr(LCase(url_str), "&page=") - 1)
        If InStr(page_str, "&") > 0 Then
            url_str = url_str & "&" & Mid(page_str, InStr(page_str, "&") + 1)
            page_str = Mid(page_str, 1, InStr(page_str, "&") - 1)
        End If
    End If
    
    retry_url = ""
    url_instr = url_str
    return_download_url = "inet|10,13|" & url_instr
    
    If page_str <> "" And IsNumeric(page_str) = True Then
        If MsgBox("���������ҳ��ַ���Ǵӵ�һҳ��ʼ�ģ�" & vbCrLf & "�Ƿ�ӵ�һҳ��ʼ���أ�" & vbCrLf & vbCrLf & "[YES]�ӵ�һҳ��ʼ" & vbCrLf & "[NO]�ӵ�ǰҳ��ʼ", vbYesNo, "ѯ��") = vbNo Then
            If Int(page_str) > 1 Then
                page = Int(page_str)
                If InStr(LCase(url_instr), "?") > 10 Then
                    return_download_url = return_download_url & "&page=" & page
                Else
                    return_download_url = return_download_url & "?page=" & page
                End If
            End If
        End If
    End If
    
    retry_url = return_download_url
    
    return_download_url = return_download_url & "|" & url_head & ".sankakucomplex.com/post/show/1936676" & vbcrlf & "User-Agent: Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0)"
		OX163_urlpage_Referer = "User-Agent: Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0)"
    deep_DL = MsgBox("�Ƿ�ʹ�ÿ��ٷ�����" & vbCrLf & "(���ַ�JPGͼƬ��PNG/GIF�ȿ����޷�������ȡ)" & vbCrLf & vbCrLf & "[YES]���ٷ���" & vbCrLf & "[NO]�������", vbYesNo, "ѯ��")
End Function
'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
    On Error Resume Next
    If delay_time=0 Then
    	If DateDiff("s", start_time, Now()) < 30 Then
    		return_download_list="1|inet|10,13|http://www.163.com/?Delay_30s-����163ҳ���ӳ�30��"
    		Exit Function
    	Else
    		return_download_list=delay_url
    		delay_time=1
    	End If
    End If
    
    Dim key_str, add_temp, file_url, preview_url, md5_code, file_type

    return_download_list = ""
    
    If pool = "post" Or pool = "deep_DL" Then
        'http://chan.sankakucomplex.com/post/show/9506/cg-d-o-_-publisher-eigoukaiki-eroge-ino-tagme
        Dim pic_alt
        If InStr(LCase(html_str), "<li>original:") > 0 or InStr(LCase(html_str), LCase(">Save this file (right click and save as)</a>")) > 0 Then
            retry_time = 0
            'ID
            key_str="<meta content=""Post "
            url_str = Mid(html_str, InStr(html_str, key_str)+len(key_str))
            url_str = Mid(url_str, 1, InStr(url_str, """") - 1)
            url_str = "p" & url_str
            'alt
            pic_alt = get_tags(html_str)
						pic_alt = Replace(pic_alt, "|", "&#124;")
						pic_alt = Replace(pic_alt, "\\", "\")
            url_str = url_str & "_" & pic_alt
            If Len(url_str) > 180 Then url_str = Left(url_str, 179) & "~"
            
            'file_url
            If InStr(LCase(html_str), "<li>original:") > 0 Then
	            html_str = Mid(html_str, InStr(LCase(html_str), "<li>original:"))
	            html_str = Mid(html_str, InStr(LCase(html_str), "<a href=""") + 9)
	            html_str = Mid(html_str, 1, InStr(html_str, Chr(34)) - 1)
	          ElseIf InStr(LCase(html_str),LCase(">Save this file (right click and save as)</a>")) > 0 Then
            	html_str = Mid(html_str,1,InStr(LCase(html_str), LCase(">Save this file (right click and save as)</a>")))
            	html_str = Mid(html_str,1,InStrrev(html_str, Chr(34)) - 1)
            	html_str = Mid(html_str,InStrrev(html_str, Chr(34)) + 1)
          	End If
          	If left(html_str,2)="//" Then html_str=url_http & html_str
						'https://v.sankakucomplex.com/data/c2/f8/c2f8d27d9b679821204309f6939e59a0.gif?e=1651808910&amp;m=bvY8eIsbcxehslh0uEcKKA&amp;expires=1651808910&amp;token=qa0Qi3fLtFIFyiGAOkulgF73EcF0cf12YzyAp6PR07w
						'convert &
						'https://v.sankakucomplex.com/data/c2/f8/c2f8d27d9b679821204309f6939e59a0.gif?e=1651808910&m=bvY8eIsbcxehslh0uEcKKA&expires=1651808910&token=qa0Qi3fLtFIFyiGAOkulgF73EcF0cf12YzyAp6PR07w
						html_str=replace(html_str,"&amp;","&")
          	
          	'file_type
          	key_str=Mid(html_str, InStrRev(html_str, "."))
          	
            If instr(key_str,"?")>2 Then key_str=Mid(key_str,1,InStr(key_str, "?")-1)
            url_str = Replace(url_str, " ", "-") & key_str
            
            If pool = "deep_DL" Then
            		split_str(split_c0)=""
                Do While split_str(split_c0) = "" And split_c0 < split_c1
                    split_c0 = split_c0 + 1
                Loop
                If split_c0 <= split_c1 and split_str(split_c0)<>"" Then
                    key_str = "1|inet|10,13|" & url_head & ".sankakucomplex.com/post/show/" & split_str(split_c0)
                    retry_url = "inet|10,13|" & url_head & ".sankakucomplex.com/post/show/" & split_str(split_c0)
                Else
                		pool=""
                    key_str = check_nextpage()
                End If
                
                key_str=delay_time_tf(key_str)
                
                return_download_list = "|" & html_str & "|" & url_str & "|" & vbCrLf & key_str
            Else
                return_download_list = "|" & html_str & "|" & url_str & "|" & vbCrLf & "0"
            End If
        ElseIf retry_time < 5 Then
            retry_time = retry_time + 1
            return_download_list = "1|" & retry_url
        Else
            If pool = "deep_DL" Then
            		split_str(split_c0)=""
                Do While split_str(split_c0) = "" And split_c0 < split_c1
                    split_c0 = split_c0 + 1
                Loop
                If split_c0 <= split_c1 and split_str(split_c0)<>"" Then
                    key_str = "1|inet|10,13|" & url_head & ".sankakucomplex.com/post/show/" & split_str(split_c0)
                    retry_url = "inet|10,13|" & url_head & ".sankakucomplex.com/post/show/" & split_str(split_c0)
                Else
                		pool=""
                    key_str = check_nextpage()
                End If
                key_str=delay_time_tf(key_str)
                
                return_download_list = key_str
            Else
                return_download_list = "0"
            End If
        End If
        Exit Function
    End If
    
    url_str = html_str
    key_str = "Post.register({"
    If InStr(LCase(html_str), LCase(key_str)) > 0 Then
        retry_time = 0
        html_str = Mid(html_str, InStr(LCase(html_str), LCase(key_str)) + Len(key_str))
        split_str = Split(html_str, key_str, -1, 1)

        For split_i = 0 To UBound(split_str)
            If deep_DL = vbNo Then
                split_str(split_i) = Mid(split_str(split_i), InStr(split_str(split_i), ",""id"":") + len(",""id"":"))
                If InStr(split_str(split_i),",")>1 Then
                	split_str(split_i) = Mid(split_str(split_i), 1, InStr(split_str(split_i), ",") - 1)
                ElseIf InStr(split_str(split_i),"});")>1 Then
                	split_str(split_i) = Mid(split_str(split_i), 1, InStr(split_str(split_i), "});") - 1)
                End if
                If IsNumeric(split_str(split_i)) = False Then split_str(split_i) = ""
            Else
                'tags
                html_str = ""
                key_str = ",""tags"":"""
                html_str = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + Len(key_str))
                html_str = Mid(html_str, 1, InStr(html_str, Chr(34)) - 1)
                html_str = Replace(html_str, "|", "&#124;")
                html_str = Replace(html_str, "\\", "\")
                
                'file_url
                file_url = ""
                preview_url = ""
                'animated_gif,gif_artifacts,transparent_png,animated_png,swf->http://chan.sankakucomplex.com/download-preview.png
                file_type = ""
                'http://chan.sankakustatic.com/data/12/20/1220fc93e930e816b5da07b7b24b9379.swf
                '"md5":"1220fc93e930e816b5da07b7b24b9379"->12/20/1220fc93e930e816b5da07b7b24b9379
                md5_code = ""
                ''preview_url
                key_str = ",""preview_url"":"""
                preview_url = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + Len(key_str))
                preview_url = Mid(preview_url, 1, InStr(preview_url, Chr(34)) - 1)
                ''MD5
                key_str = ",""md5"":"""
                md5_code = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + Len(key_str))
                md5_code = Mid(md5_code, 1, InStr(md5_code, Chr(34)) - 1)
                If InStr(LCase(preview_url), "download-preview.png")>0 Then
                    file_type = ".swf"
                Else
                    file_type = check_gif_png(html_str)
                    If file_type = "" Then file_type = Mid(preview_url, InStrRev(preview_url, "."))
                End If
                If url_head = "http://idol" or url_head = "https://idol" Then
                		file_url = url_http & "//is.sankakucomplex.com/data/" & Left(md5_code, 2) & "/" & Mid(md5_code, 3, 2) & "/" & md5_code & file_type
    						ElseIf url_head = "http://chan" or url_head = "https://chan" Then
                		file_url = url_http & "//cs.sankakucomplex.com/data/" & Left(md5_code, 2) & "/" & Mid(md5_code, 3, 2) & "/" & md5_code & file_type
								End If
								
                'ID
                add_temp = ""
                key_str = ",""id"":"
                add_temp = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase(key_str)) + Len(key_str))
                If InStr(add_temp, "}") Then add_temp = Mid(add_temp, 1, InStr(add_temp, "}") - 1)
                If InStr(add_temp, ",") Then add_temp = Mid(add_temp, 1, InStr(add_temp, ",") - 1)
                If IsNumeric(add_temp) = False Then add_temp = ""
                
                'file name
                split_str(split_i) = "p" & add_temp & "_" & Trim(html_str)
                If Len(split_str(split_i)) > 180 Then split_str(split_i) = Left(split_str(split_i), 179) & "~"
                split_str(split_i) = Replace(split_str(split_i), " ", "-") & Mid(file_url, InStrRev(file_url, "."))
                
                return_download_list = return_download_list & "|" & file_url & "|" & split_str(split_i) & "|" & html_str & vbCrLf
            End If
        Next
        
    ElseIf retry_time < 5 Then
        retry_time = retry_time + 1
        return_download_list = "2|" & retry_url
        Exit Function
    End If
    
    key_str = ""
    
    If InStr(LCase(url_str), "next-page-url=""") > 0 Then
        page_counter=page+1
    Else
        page_counter = 1
    End If
    
    If deep_DL = vbNo Then
        key_str = Join(split_str, "")
        If key_str <> "" Then
            split_c1 = 0
            split_c1 = UBound(split_str)
            split_c0 = 0
            pool = "deep_DL"
            Do While split_str(split_c0) = "" And split_c0 < split_c1
                split_c0 = split_c0 + 1
            Loop
            If split_c0 <= split_c1 and split_str(split_c0)<>"" Then
                return_download_list = "1|inet|10,13|" & url_head & ".sankakucomplex.com/post/show/" & split_str(split_c0)
                retry_url = "inet|10,13|" & url_head & ".sankakucomplex.com/post/show/" & split_str(split_c0)
                Exit Function
            End If
        End If
    End If
    
    pool = ""
    return_download_list = return_download_list & delay_time_tf(check_nextpage())

End Function
'--------------------------------------------
Function delay_time_tf(byval nextpage)
    If delay_time<25 Then
    	delay_time=delay_time+1
    	delay_time_tf=nextpage
    Else
    	delay_time=0
    	delay_url=""
    	delay_url=nextpage
    	delay_time_tf="1|inet|10,13|http://www.163.com/?Delay_5s-����163ҳ���ӳ�30��"
    	start_time=Now()
    End If
End Function

Function check_nextpage()
    check_nextpage = 0
    If page < page_counter Then
        page = page + 1
        If InStr(LCase(url_instr), "?") > 10 Then
            check_nextpage = page_counter & "|inet|10,13|" & url_instr & "&page=" & page
            retry_url = "inet|10,13|" & url_instr & "&page=" & page
        Else
            check_nextpage = page_counter & "|inet|10,13|" & url_instr & "?page=" & page
            retry_url = "inet|10,13|" & url_instr & "?page=" & page
        End If
    Else
        check_nextpage = "0"
    End If
End Function

Function get_tags(ByVal html_str)
		On Error Resume Next
		get_tags=""
    html_str = Mid(html_str, InStr(html_str, "<ul id=tag-sidebar>"))
    html_str = Mid(html_str, InStr(html_str, "<li class=tag-type-")+len("<li class=tag-type-"))
    html_str = Mid(html_str, 1, InStr(html_str, "</ul>")-1)
    Dim split_str
    split_str=split(html_str,"<li class=tag-type-")
    For i=0 to ubound(split_str)
    	split_str(i)=Mid(split_str(i),InStr(split_str(i),"<a href=")+8)
    	split_str(i)=Mid(split_str(i),InStr(split_str(i),">")+1)
    	split_str(i)=Trim(Mid(split_str(i),1,InStr(split_str(i),"</a>")-1))
    	split_str(i)=replace(split_str(i)," ","_")
  	Next
  	html_str=join(split_str," ")
  	get_tags=html_str
End Function
'--------------------------------------------

Function check_gif_png(ByVal tags_str)
    On Error Resume Next
    check_gif_png = ""
    '1:animated_png,2:animated_gif,3:gif_artifacts,4:transparent_png
    If InStr(LCase(tags_str), "animated_png ") = 1 Then check_gif_png = ".png": Exit Function
    If InStr(LCase(tags_str), " animated_png ") > 1 Then check_gif_png = ".png": Exit Function
    If InStr(LCase(tags_str), " animated_png") = (Len(tags_str) - Len(" animated_png") + 1) Then check_gif_png = ".png": Exit Function
    
    If InStr(LCase(tags_str), "animated_gif ") = 1 Then check_gif_png = ".gif": Exit Function
    If InStr(LCase(tags_str), " animated_gif ") > 1 Then check_gif_png = ".gif": Exit Function
    If InStr(LCase(tags_str), " animated_gif") = (Len(tags_str) - Len(" animated_gif") + 1) Then check_gif_png = ".gif": Exit Function
    
    If InStr(LCase(tags_str), "gif_artifacts ") = 1 Then check_gif_png = ".gif": Exit Function
    If InStr(LCase(tags_str), " gif_artifacts ") > 1 Then check_gif_png = ".gif": Exit Function
    If InStr(LCase(tags_str), " gif_artifacts") = (Len(tags_str) - Len(" gif_artifacts") + 1) Then check_gif_png = ".gif": Exit Function
    
    If InStr(LCase(tags_str), "transparent_png ") = 1 Then check_gif_png = ".png": Exit Function
    If InStr(LCase(tags_str), " transparent_png ") > 1 Then check_gif_png = ".png": Exit Function
    If InStr(LCase(tags_str), " transparent_png") = (Len(tags_str) - Len(" transparent_png") + 1) Then check_gif_png = ".png": Exit Function
End Function