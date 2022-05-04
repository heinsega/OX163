'2022-5-4 163.shanhaijing.net

Function OX163_Web_Browser_ctrl(ByVal URL,ByVal Flags,ByVal TargetFrameName,ByVal PostData,ByVal Headers)
On Error Resume Next
OX163_Web_Browser_ctrl="" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & ""
'If InStr(LCase(URL),"http://95.211.21.16/s/")=1 Then'95.211.21.16 www.hentaiverse.net
'If InStr(LCase(URL),"http://g.e-hentai.org")=1 Then
'	If Right(LCase(URL),8)="/1-m-y/0" Then URL=Left(URL,Len(URL)-7)
'	OX163_Web_Browser_ctrl=replace(URL,"http://g.e-hentai.org","http://r.e-hentai.org") & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & ""
If InStr(LCase(URL),"b http://")=1 or InStr(LCase(URL),"b%20http://")=1 Then
  URL=Mid(URL,InStrrev(LCase(URL),"http://"))
	OX163_Web_Browser_ctrl=URL & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & ""
ElseIf InStr(LCase(URL),"b https://")=1 or InStr(LCase(URL),"b%20https://")=1 Then
  URL=Mid(URL,InStrrev(LCase(URL),"https://"))
	OX163_Web_Browser_ctrl=URL & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & ""
ElseIf InStr(LCase(URL),"http://b%20http//")=1 or InStr(LCase(URL),"http://b%20https//")=1 Then
  If InStr(LCase(URL),"http://b%20http//")=1 Then URL="http://" & Mid(URL,18)
  If InStr(LCase(URL),"http://b%20https//")=1 Then URL="https://" & Mid(URL,19)
	OX163_Web_Browser_ctrl=URL & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & ""

ElseIf InStr(LCase(URL),"http://picasaweb.google.")=1 or InStr(LCase(URL),"picasaweb.google.")=1 Then
	URL="https://picasaweb.google." & Mid(URL,InStr(LCase(URL),"picasaweb.google.")+Len("picasaweb.google."))
	OX163_Web_Browser_ctrl=URL & vbCrLf & vbCrLf & Flags & vbCrLf & vbCrLf & TargetFrameName & vbCrLf & vbCrLf & PostData & vbCrLf & vbCrLf & Headers
ElseIf InStr(LCase(URL),"http://behoimi.org")=1 or InStr(LCase(URL),"http://www.behoimi.org")=1 Then
	Headers="User-Agent: QuickTime/7.6.2 (qtver=7.6.2;os=Windows NT 5.1Service Pack 2)"
	OX163_Web_Browser_ctrl=URL & vbCrLf & vbCrLf & Flags & vbCrLf & vbCrLf & TargetFrameName & vbCrLf & vbCrLf & PostData & vbCrLf & vbCrLf & Headers
ElseIf InStr(LCase(URL),"http://exhentai.org")=1 or InStr(LCase(URL),"http://g.e-hentai.org")=1 Then
	URL="https://" & Mid(URL,8)
	OX163_Web_Browser_ctrl=URL & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & "" & vbCrLf & vbCrLf & ""
End If
End Function

Function OX163_Web_Browser_url(ByVal URL)
OX163_Web_Browser_url=URL
'If InStr(LCase(URL),"http://r.e-hentai.org")=1 Then
'	OX163_Web_Browser_url=replace(LCase(URL),"http://r.e-hentai.org","http://g.e-hentai.org")
If InStr(LCase(URL),"b http://")=1 or InStr(LCase(URL),"b%20http://")=1 or InStr(LCase(URL),"http://b%20http//")=1 Then
  URL=Mid(URL,InStr(LCase(URL),"b"))
  URL="b http:" & Mid(URL,InStr(LCase(URL),"//"))
  OX163_Web_Browser_url=URL
End If

End Function