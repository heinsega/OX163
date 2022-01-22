'2020-5-5 163.shanhaijing.net
Dim url_parent, Next_page, retry_counter, url_host

Function return_download_url(ByVal url_str)
'http://www.zerochan.net/?o=1927151
'http://kpop.asiachan.com/?p=1
'http://www.zerochan.net/Izumi+no+Kami+Kanesada?p=2
'http://www.zerochan.net/Izumi+no+Kami+Kanesada?d=1&p=2
'http://www.zerochan.net/Kuroko+no+Basuke?d=2&o=1257633
'https://www.zerochan.net/Takada+Akemi?p=2
On Error Resume Next
return_download_url = ""
Next_page = ""
Page = 0
p_oid = 0
p_str=""
url_parent = Mid(url_str, InStr(url_str, "//") + 2)
url_host = Mid(url_parent,1, InStr(url_parent, "/") - 1)

If InStr(LCase(url_str), "?p=") > 1 Then
    url_parent = Mid(url_str, 1, InStr(LCase(url_str), "?p=") - 1)
    Page = Mid(url_str, InStr(LCase(url_str), "?p=") + 3)
    p_str = "?"
ElseIf InStr(LCase(url_str), "&p=") > 1 Then
    url_parent = Mid(url_str, 1, InStr(LCase(url_str), "&p=") - 1)
    Page = Mid(url_str, InStr(LCase(url_str), "&p=") + 3)
    p_str = "&"
ElseIf InStr(LCase(url_str), "?o=") > 1 Then
    url_parent = Mid(url_str, 1, InStr(LCase(url_str), "?o=") - 1)
    p_oid = Mid(url_str, InStr(LCase(url_str), "?o=") + 3)
    p_str = "?"
ElseIf InStr(LCase(url_str), "&o=") > 1 Then
    url_parent = Mid(url_str, 1, InStr(LCase(url_str), "&o=") - 1)
    p_oid = Mid(url_str, InStr(LCase(url_str), "&o=") + 3)
    p_str = "&"
Else
		p_str = "?"
End If

If Page > 1 Or p_oid > 0 Then
    If MsgBox("本页不是第1页" & vbCrLf & "是否从第1页开始？", vbYesNo, "问题") = vbYes Then
        url_str = url_parent
    End If
End If

retry_counter = 0
    
return_download_url = "inet|10,13|" & url_str & "|" & url_str
OX163_urlpage_Referer = url_str

'http://www.zerochan.net/Kuroko+no+Basuke?d=2&
'url_parent = url_parent & p_str

End Function

'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim split_str, sid, key_str, Next_oID

'<a href="?p=2" tabindex="1" rel="next">Next &raquo;</a>
key_str = "rel=""next"""
Next_oID = ""
If InStr(LCase(html_str), LCase(key_str)) > 0 Then
	Next_oID = Mid(html_str, 1, InStr(LCase(html_str), key_str))
	Next_oID = Mid(Next_oID, InStrRev(Next_oID, "?"))
	Next_oID = Mid(Next_oID, 1, InStr(Next_oID, """") - 1)
End If

'{"@type": ["ListItem","ImageObject"], "position": "18", "name": "Ayukawa Madoka", "url": "https://static.zerochan.net/Ayukawa.Madoka.full.457682.jpg", "thumbnailUrl": "https://s3.zerochan.net/Ayukawa.Madoka.240.457682.jpg"}
key_str = "{""@type"": [""ListItem"",""ImageObject""],"
If InStr(LCase(html_str), LCase(key_str)) > 0 Then
  Next_page = ""
  retry_counter = 0
    
  html_str = Mid(html_str, InStr(LCase(html_str), LCase(key_str)) + Len(key_str))
  'html_str = Mid(html_str, InStr(LCase(html_str), LCase("""name"": """)) + 9)
  html_str = Mid(html_str, 1, InStr(LCase(html_str), "</script>"))
    
  split_str = Split(html_str, key_str)
  
  For split_i = 0 To UBound(split_str)
  
		'http://host/name.full.id.type
		url_str = "" 'url
		html_str = "" 'name
		sid = "" 'id
		key_str = "" 'type
		
		'url
		url_str = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), LCase("""url"": """)) + 8)
		url_str = Mid(url_str, 1, InStr(url_str, Chr(34)) - 1)
		'https://static.zerochan.net/Ayukawa.Madoka.full.457682.jpg

		'name
		'Ayukawa.Madoka.full.457682.jpg
		html_str = Mid(url_str, InStr(url_str, ".zerochan.net/") + len(".zerochan.net/"))
		'Ayukawa.Madoka.full.457682.jpg
    
		'OK
		split_str(split_i) = ""
		split_str(split_i) = "|" & url_str & "|(" & url_host & ")" & UTF8DecodeURI(html_str) & "|"

  Next
  return_download_list = Join(split_str, vbCrLf) & vbCrLf
  If Next_oID <>"" Then  	
		Next_page = Next_oID
    return_download_list = return_download_list & "1|inet|10,13|" & url_parent & Next_page
  End If
  
ElseIf retry_counter < 3 Then
  retry_counter = retry_counter + 1
  return_download_list = "1|inet|10,13|" & url_parent & Next_page
  
Else
	return_download_list = "0"
	
End If
End Function

Function UTF8DecodeURI(ByVal strIn)
    UTF8DecodeURI = ""
    Dim sl: sl = 1
    Dim tl: tl = 1
    Dim key: key = "%"
    Dim kl: kl = Len(key)
    sl = InStr(sl, strIn, key)
    Do While sl > 0
        If (tl = 1 And sl <> 1) Or tl < sl Then
            UTF8DecodeURI = UTF8DecodeURI & Mid(strIn, tl, sl - tl)
        End If
        Dim hh, hi, hl
        Dim a
        Select Case UCase(Mid(strIn, sl + kl, 1))
        Case "U": 'Unicode URLEncode
            a = Mid(strIn, sl + kl + 1, 4)
            UTF8DecodeURI = UTF8DecodeURI & ChrW("&H" & a)
            sl = sl + 6
        Case "E": 'UTF-8 URLEncode
            hh = Mid(strIn, sl + kl, 2)
            a = Int("&H" & hh) 'ascii码
            If Abs(a) < 128 Then
                sl = sl + 3
                UTF8DecodeURI = UTF8DecodeURI & Chr(a)
            Else
                hi = Mid(strIn, sl + 3 + kl, 2)
                hl = Mid(strIn, sl + 6 + kl, 2)
                a = ("&H" & hh And &HF) * 2 ^ 12 Or ("&H" & hi And &H3F) * 2 ^ 6 Or ("&H" & hl And &H3F)
                If a < 0 Then a = a + 65536
                UTF8DecodeURI = UTF8DecodeURI & ChrW(a)
                sl = sl + 9
            End If
        Case Else: 'Asc URLEncode
            hh = Mid(strIn, sl + kl, 2) '高位
            a = Int("&H" & hh) 'ascii码
            If Abs(a) < 128 Then
                sl = sl + 3
            Else
                hi = Mid(strIn, sl + 3 + kl, 2) '低位
                a = Int("&H" & hh & hi) '非ascii码
                sl = sl + 6
            End If
            UTF8DecodeURI = UTF8DecodeURI & Chr(a)
        End Select
        tl = sl
        sl = InStr(sl, strIn, key)
    Loop
    UTF8DecodeURI = UTF8DecodeURI & Mid(strIn, tl)
End Function