'2013-5-25 163.shanhaijing.net
Dim new_trailer_page, new_trailer_large, new_trailer_extralarge, includes_html, trailer_url_split, trailer_url_i, trailer_url_ubound
Function return_download_url(ByVal url_str)
'http://www.apple.com/trailers/wb/wherethewildthingsare/
'http://trailers.apple.com/trailers/paramount/wtc/
On Error Resume Next
new_trailer_page=0
new_trailer_large=0
new_trailer_extralarge=0
includes_html=""
return_download_url = "inet|10,13|" & url_str & "/includes/playlists/web.inc|User-Agent: QuickTime/7.6.2 (qtver=7.6.2;os=Windows NT 5.1Service Pack 2)"
OX163_urlpage_Referer = "User-Agent: QuickTime/7.6.2 (qtver=7.6.2;os=Windows NT 5.1Service Pack 2)"
End Function

'--------------------------------------------------------

Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next

Dim split_str, end_i, split_i
return_download_list = ""

If InStr(LCase(html_str), LCase(".mov""")) > 0 and new_trailer_page=0 Then
	
	'old_trailer_list
	split_str = Split(html_str, ".mov""", -1, 1)
			end_i=UBound(split_str)-1
	    For split_i = 0 To end_i
	    'url
	    split_str(split_i) = Mid(split_str(split_i), InStrrev(LCase(split_str(split_i)), chr(34))+1)
	    split_str(split_i) = split_str(split_i) & ".mov"
	    split_str(split_i) = replace(split_str(split_i),"_480p.mov","_h480p.mov")
	    split_str(split_i) = replace(split_str(split_i),"_720p.mov","_h720p.mov")
	    split_str(split_i) = replace(split_str(split_i),"_1080p.mov","_h1080p.mov")
	    'name
	    html_str=Mid(split_str(split_i),InStrrev(split_str(split_i), "/")+1)
	    
	    return_download_list = return_download_list & "|" & split_str(split_i) & "|" & html_str & "|" & vbCrLf
	    Next
	return_download_list = return_download_list & "0"
	Exit Function

Else
	'new_trailer_list
	'step1
	If new_trailer_page=0 Then
		If InStr(LCase(html_str), LCase("includes/large.html")) > 0 Then
			new_trailer_large=1
			new_trailer_page=1
		End If
		If InStr(LCase(html_str), LCase("includes/extralarge.html")) > 0 Then
			new_trailer_extralarge=1
			new_trailer_page=1
		End If
		
		If new_trailer_large=1 Then
			new_trailer_large=0
			return_download_list = "1|inet|10,13|" & url_str & "/includes/large.html"
			Exit Function
		ElseIf new_trailer_extralarge=1 Then
			new_trailer_extralarge=0
			return_download_list = "1|inet|10,13|" & url_str & "/includes/extralarge.html"
			Exit Function
		Else
			return_download_list = 0
			Exit Function
		End If
		
	'ElseIf new_trailer_page=1 and new_trailer_extralarge=1 Then
	'		includes_html=return_ncludes_list(html_str)
	'		new_trailer_extralarge=0
	'		return_download_list = "1|inet|10,13|" & url_str & "/includes/extralarge.html"
	'		Exit Function

	'step2
	ElseIf new_trailer_page=1 Then
		'If includes_html="" Then
			includes_html=return_ncludes_list(html_str)
		'Else
		'	includes_html=includes_html & vbcrlf & return_ncludes_list(html_str)
		'End If
		
		If includes_html<>"" Then
			new_trailer_page=2
			trailer_url_split=split(includes_html,vbcrlf)
			'teaser/large.html
			'trailer4/large.html
			'trailer3/extralarge.html
			trailer_url_i=0
			trailer_url_ubound=UBound(trailer_url_split)
			return_download_list = "1|inet|10,13|" & url_str & "/includes/" & trailer_url_split(trailer_url_i)
			Exit Function
			
		Else
			return_download_list = 0
			Exit Function
		End If
		
	'step3 list trailer url
	ElseIf new_trailer_page=2 Then
		If InStr(LCase(html_str), LCase(".mov")) > 0 Then
			split_str = Split(html_str, ".mov", -1, 1)
			end_i=UBound(split_str)-1
	    For split_i = 0 To end_i
		    'url
		    split_str(split_i) = Mid(split_str(split_i), InStrrev(LCase(split_str(split_i)), chr(34))+1)
		    split_str(split_i) = split_str(split_i) & ".mov"
		    'split_str(split_i) = replace(split_str(split_i),"_480p.mov","_h480p.mov")
		    'split_str(split_i) = replace(split_str(split_i),"_720p.mov","_h720p.mov")
		    'split_str(split_i) = replace(split_str(split_i),"_1080p.mov","_h1080p.mov")
		    split_str(split_i) = Mid(split_str(split_i),1,InStrrev(split_str(split_i), "_"))
		    'name
		    html_str=Mid(split_str(split_i),InStrrev(split_str(split_i), "/")+1)
		    
		    return_download_list = return_download_list & "|" & split_str(split_i) & "h480p.mov|" & html_str & "h480p.mov|" & vbCrLf
		    return_download_list = return_download_list & "|" & split_str(split_i) & "h720p.mov|" & html_str & "h720p.mov|" & vbCrLf
		    return_download_list = return_download_list & "|" & split_str(split_i) & "h1080p.mov|" & html_str & "h1080p.mov|" & vbCrLf
	    Next
		End If
	    
	  If trailer_url_i<trailer_url_ubound Then
	  	trailer_url_i=trailer_url_i+1
			return_download_list = return_download_list & "1|inet|10,13|" & url_str & "/includes/" & trailer_url_split(trailer_url_i)
			Exit Function
	  Else
	   	return_download_list = return_download_list & "0"
			Exit Function
	  End If
	
	Else
		return_download_list = 0
		Exit Function
	End If
End If

End Function

'----------------------------------------------------------------
Function return_ncludes_list(ByVal html_str)
return_ncludes_list=""
If InStr(LCase(html_str), LCase("<a href=""includes/")) > 0 Then
	Dim split_str,split_i
	html_str=mid(LCase(html_str),InStr(LCase(html_str), LCase("<a href=""includes/"))+len("<a href=""includes/"))
	split_str=split(html_str,"<a href=""includes/")
	For split_i=0 to UBound(split_str)
		split_str(split_i)=mid(split_str(split_i),1,InStr(split_str(split_i), "#")-1)
	Next
	return_ncludes_list=join(split_str,vbcrlf)
End If
End Function