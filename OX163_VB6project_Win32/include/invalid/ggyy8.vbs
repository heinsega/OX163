'2011-10-1 163.shanhaijing.net
Function return_download_url(ByVal url_str)
return_download_url="inet|10,13|" & url_str & "|Referer: " & url_str & vbcrlf & "Host: img.ggyy8.cc"
End Function
'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list = ""
If InStr(html_str, "<!--章节列表开始-->") > 0 Then

Dim album_list,url_temp
url_temp="http://www.ggyy8.cc/"

html_str = Mid(html_str, InStr(html_str, "<!--漫画信息开始-->"))
url_str = Mid(html_str,1, InStr(html_str, "</h1>")-1)
url_str = Mid(url_str,InStr(url_str, "<h1>")+4)
html_str = Mid(html_str, InStr(html_str, "<!--章节列表开始-->")+len("<!--章节列表开始-->"))
html_str = Mid(html_str,1, InStr(html_str, "<!--章节列表结束-->")-1)
html_str = Mid(html_str,InStr(html_str, "<li><a href=""")+len("<li><a href="""))
    
album_list = Split(html_str, "<li><a href=""")

For i = 0 To UBound(album_list)
				html_str=""
        If InStr(LCase(album_list(i)), ".html") > 0 Then
        'url
        html_str = Mid(album_list(i),1,InStr(album_list(i),Chr(34))-1)
        html_str = url_temp & html_str
        album_list(i)=Mid(album_list(i),InStr(album_list(i), ">")+1)
        album_list(i)=replace(Mid(album_list(i),1,InStr(album_list(i), "<span>")-1),"&nbsp;"," ")
        album_list(i)=url_str & "_" & album_list(i)
        return_albums_list = return_albums_list & "0||" & html_str & "|" & album_list(i) & "|" & album_list(i) & vbcrlf
	End If
Next
return_albums_list = return_albums_list & "0"

Else	
return_albums_list = return_albums_list & "0"
End If
End Function
'----------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list=""
'var __arr = ['0_86.jpg','1_5c.jpg','17_bc.jpg'];var __p='/comic/海贼王/102241/'
'http://img.ggyy8.cc/comic/海贼王/102241/0_86.jpg
If InStr(LCase(html_str), "var __arr =")>0 and InStr(LCase(html_str), "var __p=") Then
	
	url_str=Mid(html_str,InStr(LCase(html_str), "var __p="))
	url_str=Mid(url_str,InStr(url_str, "'")+1)
	url_str=Mid(url_str,1,InStr(url_str, "'")-1)
	url_str=UTF8EncodeURI(url_str)
	url_str="http://img.ggyy8.cc" & url_str
	html_str=Mid(html_str,InStr(LCase(html_str), "var __arr ="))
	html_str=Mid(html_str,InStr(html_str, "'")+1)
	html_str=Mid(html_str,1,InStr(html_str, "']")-1)
	
	Dim split_str
	split_str=Split(html_str, "','")
	
	For i = 0 To UBound(split_str)
    split_str(i)=UTF8EncodeURI(split_str(i))
		return_download_list=return_download_list & "|" & url_str & split_str(i) & "|" & split_str(i) & "|" & vbcrlf
	next
	return_download_list = return_download_list & "0"
Else
    return_download_list = "0"
End If
End Function
'--------------------------------------------------------
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