'2012-11-2 163.shanhaijing.net
Dim page_num
Function return_download_url(ByVal url_str)
On Error Resume Next
'http://www.rosiyy.com/
page_num=0
url_str=UTF8EncodeURI(url_str)
return_download_url = "inet|10,13|" & url_str
'OX163_urlpage_Referer="http://" & url_parent & "/" & vbcrlf & "User-Agent: Mozilla/4.0 (compatible; MSIE 8.00; Windows XP)"

End Function

'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
Dim key_word,split_str,str_temp
return_albums_list = ""
str_temp=html_str

key_word="<ol class=""page-navigator""><li><a class=""prev"""
If page_num=0 and instr(LCase(html_str),key_word)>0 Then
	page_num=1
	html_str=mid(html_str,instr(LCase(html_str),key_word)+len(key_word))
	html_str=mid(html_str,instr(LCase(html_str),"<li><a href=""")+len("<li><a href="""))
	html_str=mid(html_str,1,instr(LCase(html_str),"""")-1)
	return_albums_list="1|inet|10,13|" & UTF8EncodeURI(html_str)
	Exit Function
ElseIf page_num=0 Then
	page_num=1
End If

html_str=str_temp
key_word="<div id=""content"" class=""list"">"
If instr(LCase(html_str),key_word)>0 Then
	html_str=mid(html_str,instr(LCase(html_str),key_word))
	html_str=mid(html_str,1,instr(LCase(html_str),"<ol class=""page-navigator"">")-1)
	
	key_word="<div class=""photo""><a href="""
	html_str=mid(html_str,instr(LCase(html_str),key_word)+len(key_word))	
	split_str=split(html_str,key_word)
	For i=0 to ubound(split_str)
		html_str=""
		url_str=""
		'url
		url_str=mid(split_str(i),1,instr(split_str(i),"""")-1)
		'alt
		html_str=mid(split_str(i),InStr(split_str(i),"alt=""")+5)
		html_str=mid(html_str,1,InStr(html_str,"""")-1)
		split_str(i)="0||" & url_str & "|" & html_str
	Next
	return_albums_list=join(split_str,vbcrlf) & vbcrlf
End If

html_str=str_temp
key_word="<a class=""next"" href="""
If InStr(LCase(html_str),key_word)>0 Then
	html_str=mid(html_str,instr(LCase(html_str),key_word)+len(key_word))
	html_str=mid(html_str,1,instr(LCase(html_str),"""")-1)
	return_albums_list= return_albums_list & "1|inet|10,13|" & UTF8EncodeURI(html_str)
Else
	return_albums_list= return_albums_list & "0"
End If

End Function

'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
If InStr(LCase(html_str), "<div class=""imglist"">") > 0 Then
    Dim split_str
    html_str = Mid(html_str, InStr(LCase(html_str), "<div class=""imglist"">"))
    html_str = Mid(html_str, 1, InStr(LCase(html_str), "</div>") - 1)
    html_str = Mid(html_str, InStr(LCase(html_str), "<a class=""fancybox"" rel=""group"" href=""")+len("<a class=""fancybox"" rel=""group"" href="""))
    split_str = Split(html_str, "<a class=""fancybox"" rel=""group"" href=""")
    For i = 0 To UBound(split_str)
        html_str = ""
        url_str = ""
        html_str = split_str(i)
        'title="Rosi模特写真_2012光棍节特别篇-小莉高清写真套图104P-第1张"
        html_str = Mid(html_str, InStr(LCase(html_str), "alt=""") + 5)
        html_str = Mid(html_str, 1, InStr(LCase(html_str), """") - 1)
        html_str=replace(html_str,"|","_")
        If len(html_str)>100 Then html_str=left(html_str,99) & "~"
        
        '/photo/2012ggj/rosiyy-ggj-001.jpg" class="highslide"
        'http://www.rosiyy.com/photo/2012ggj/rosiyy-ggj-001.jpg
        url_str = "http://www.rosiyy.com" & Mid(split_str(i), 1, InStr(split_str(i), """") - 1)
        split_str(i)=Mid(url_str,InStrRev(url_str, "/") + 1)        
        split_str(i) = "|" & url_str & "|" & Mid(split_str(i),1,InStrRev(split_str(i), ".") -1) & "(" & html_str & ")" & Mid(url_str,InStrRev(url_str, ".")) & "|" & html_str
    Next
    return_download_list = Join(split_str, vbCrLf) & vbCrLf
End If
    return_download_list = return_download_list & "0"
End Function

'------------------------------------------------------------
Function UTF8EncodeURI(ByVal szInput)
On Error Resume Next
    Dim wch, uch, szRet
    Dim x
    Dim nAsc, nAsc2, nAsc3

    If szInput = "" Then
        UTF8EncodeURI = szInput
        Exit Function
    End If

    For x = 1 To Len(szInput)
        wch = Mid(szInput, x, 1)
        nAsc = AscW(wch)

        If nAsc < 0 Then nAsc = nAsc + 65536

        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next

    UTF8EncodeURI = szRet
End Function