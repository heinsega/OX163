'2010-1-11 163.shanhaijing.net

Function OX163_Web_Browser_ctrl(ByVal URL,ByVal Flags,ByVal TargetFrameName,ByVal PostData,ByVal Headers)
On Error Resume Next
OX163_Web_Browser_ctrl="" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & ""
'If InStr(LCase(URL),"http://95.211.21.16/s/")=1 Then'95.211.21.16 www.hentaiverse.net
'OX163_Web_Browser_ctrl=replace(LCase(URL),"http://95.211.21.16/s/","http://g.e-hentai.org/s/") & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "Host: 95.211.21.16"

If InStr(LCase(URL),"http://g.e-hentai.org/s/")=1 Then
	OX163_Web_Browser_ctrl=URL & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "Host: r.e-hentai.org"

ElseIf InStr(LCase(URL),"http://g.e-hentai.org/")=1 and Right(LCase(URL),8)="/1-m-y/0" Then
	URL=Left(URL,Len(URL)-7)
	OX163_Web_Browser_ctrl=URL & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "Host: 95.211.21.16"

ElseIf InStr(LCase(URL),"http://g.e-hentai.org")=1 Then
	If Right(LCase(URL),8)="/1-m-y/0" Then URL=Left(URL,Len(URL)-7)
	OX163_Web_Browser_ctrl=URL & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "Host: 95.211.21.16"

ElseIf InStr(LCase(URL),"getchu.com")>8 and InStr(LCase(URL),"getchu.com")<15 Then
	If InStr(LCase(URL),"http://www.getchu.com")=1 Then
		OX163_Web_Browser_ctrl=replace(LCase(URL),"http://www.getchu.com","http://210.155.150.152") & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & ""
	ElseIf InStr(LCase(URL),"http://getchu.com")=1 Then
		OX163_Web_Browser_ctrl=replace(LCase(URL),"http://getchu.com","http://210.155.150.152") & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & ""
	End If
End If
End Function

Function OX163_Web_Browser_url(ByVal URL)
OX163_Web_Browser_url=URL
If InStr(LCase(URL),"http://95.211.21.16")=1 Then
	OX163_Web_Browser_url=replace(LCase(URL),"http://95.211.21.16","http://g.e-hentai.org")
End If
End Function