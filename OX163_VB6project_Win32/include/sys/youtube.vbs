'2014-7-19 163.shanhaijing.net
Dim vid, use_Function
Function return_download_url(ByVal url_str)
On Error Resume Next
return_download_url =""
'http://www.youtube.com/watch?v=rjZds93XHeQ&list=PL932AB0958975441F&index=32&feature=plpp_video
'http://www.youtube.com/watch?v=vY8PBYQTtUM&feature=plcp
'http://www.youtube.com/get_video_info?video_id=vY8PBYQTtUM
If instr(lcase(url_str),"://www.youtube.com/watch?v=") Then
	url_str=Mid(url_str,instr(lcase(url_str),"://www.youtube.com/watch?v=")+len("://www.youtube.com/watch?v="))
	If instr(url_str,"&")>0 Then url_str=Mid(url_str,1,instr(url_str,"&")-1)
	vid=url_str
	return_download_url = "inet|10,13|http://www.youtube.com/get_video_info?video_id=" & url_str
	
	If MsgBox("是否使用新方式查找?" & vbcrlf & "是:新方式可列出独立音频格式" & vbcrlf & "否:老方式高清视频可能为独立视频无外挂音频",vbYesNo,"询问")=vbYes Then
		use_Function=1
	Else
		use_Function=0
	End If
End If
End Function

'--------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
	If use_Function=1 Then
		return_download_list=return_download_list_new(html_str, url_str)
	Else
		return_download_list=return_download_list_old(html_str, url_str)
	End If
End Function

Function return_download_list_new(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list_new = ""
html_str = UTF8DecodeURI(html_str)
If InStr(LCase(html_str), "&url=") > 0 Or InStr(LCase(html_str), ",url=") > 0 Then
    Dim split_str, split_i, split_title, key_word, split_size, file_type
    'clen=59622600
    'url=https%3A%2F%2Fr3---sn-q4f7sn7z.googlevideo.com%2Fvideoplayback%3Fupn%3DCTNNFpaND3g%26requiressl%3Dyes%26gir%3Dyes%26key%3Dyt5%26ip%3D107.178.200.34%26mws%3Dyes%26dur%3D175.475%26ipbits%3D0%26lmt%3D1405555397933550%26sparams%3Dclen%252Cdur%252Cgir%252Cid%252Cip%252Cipbits%252Citag%252Clmt%252Crequiressl%252Csource%252Cupn%252Cexpire%26source%3Dyoutube%26ms%3Dau%26clen%3D59622600%26id%3Do-ADML4g_zTo8ByJ3eolfHcjrmAUN0abL3ffTPaiVQIW0h%26itag%3D137%26mt%3D1405759349%26sver%3D3%26expire%3D1405782000%26signature%3DB6B47B0D22B6E7B8AC55017B46304EC1D1CB986C.418A7FF3A02EBF589AD832DC78F84B22A858BAC6%26fexp%3D901802%252C902408%252C914074%252C924222%252C926400%252C927622%252C930008%252C934024%252C934030%252C934804%252C938628%252C945307%252C945533%252C946013%26mv%3Du
    'itag=137
    'type=video%2Fmp4%3B+codecs%3D%22avc1.640028%22
    'bitrate=4262216
    'size=1920x1080
    'lmt=1405555397933550
    'index=711-1162,init=0-234
    key_word = "&clen="
    url_str = ""
    url_str = Mid(html_str, InStr(LCase(html_str), key_word))
    
    '&title=イラストレーター+KEI+Live+Painting&
    key_word = "&title="
    split_title = ""
    split_title = Mid(html_str, InStr(LCase(html_str), key_word) + Len(key_word))
    split_title = Mid(split_title, 1, InStr(LCase(split_title), "&") - 1)
    'title
    split_title = Replace(Replace(split_title, "+", " "), "|", "_")
    
    If InStr(LCase(url_str), "&url=") > 0 Then
        key_word = "&url="
    Else
        key_word = ",url="
    End If
    url_str = Mid(url_str, InStr(LCase(url_str), key_word) + Len(key_word))
    split_str = Split(url_str, key_word)
    
    For split_i = 0 To UBound(split_str)
        url_str = ""
        html_str = ""
        split_size = ""
        url_str = Mid(split_str(split_i), 1, InStr(split_str(split_i), "&") - 1)
        url_str = UTF8DecodeURI(url_str)
        
        '&sig=581456E233472AC93ABAF398F360D22F8C1860B4.7D12462F28E2BFE7F5974C5EBC03AF0A24A253F9&
        '->&signature=581456E233472AC93ABAF398F360D22F8C1860B4.7D12462F28E2BFE7F5974C5EBC03AF0A24A253F9
        key_word = "&sig="
        html_str = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), key_word) + Len(key_word))
        html_str = Mid(html_str, 1, InStr(html_str, "&") - 1)
        'url
        url_str = url_str & "&signature=" & html_str
        'If instr(LCase(url_str),"http://")=1 Then url_str="https://" & mid(url_str,8)
        
        'size
        key_word = "&size="
        split_size = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), key_word) + Len(key_word))
        split_size = Mid(split_size, 1, InStr(split_size, "&") - 1)
        '&type=video/webm;+codecs="vp8.0,+vorbis"&
        '&type=video/x-flv&
        '&type=video/mp4;+codecs="avc1.42001E,+mp4a.40.2"&
        '&type=video/x-flv&
        '&type=video/3gpp;+codecs="mp4v.20.3,+mp4a.40.2"&
        '&type=video/3gpp;+codecs="mp4v.20.3,+mp4a.40.2"&
        'Watch online   (WebM(VP8), 640 x 360, Stereo 44KHz Vorbis)
        'Download   (FLV, 640 x 360, Stereo 44KHz AAC)
        'Download   (MP4(H.264), 640 x 360, Stereo 44KHz AAC)
        'Download   (FLV, 400 x 240, Mono 44KHz MP3)
        'Download   (3GP, 320 x 240, Stereo 44KHz AAC)
        'Download   (3GP, 176 x 144, Stereo 44KHz AAC)
        key_word = "&type="
        file_type = ""
        html_str = Mid(split_str(split_i), InStr(LCase(split_str(split_i)), key_word) + Len(key_word))
        html_str = Mid(html_str, 1, InStr(html_str, "&") - 1)
        html_str = UTF8DecodeURI(html_str)
        If InStr(LCase(html_str), "video/webm") = 1 Then
            file_type = "webm"
        ElseIf InStr(LCase(html_str), "video/3gpp") = 1 Then
            file_type = "3gpp"
        ElseIf InStr(LCase(html_str), "video/x-flv") = 1 Then
            file_type = "flv"
        ElseIf InStr(LCase(html_str), "video/mp4") = 1 Then
            file_type = "mp4"
        ElseIf InStr(LCase(html_str), "audio/webm") = 1 Then
            file_type = "audio.webm"
            split_size = "audio"
        ElseIf InStr(LCase(html_str), "audio/mp4") = 1 Then
            file_type = "audio.mp4"
            split_size = "audio"
        ElseIf InStr(LCase(html_str), "video/") = 1 Then
            file_type = "flv"
        ElseIf InStr(LCase(html_str), "audio/") = 1 Then
            file_type = "mp3"
            split_size = "audio"
        Else
            file_type = ""
        End If

        If InStr(split_size, "x") < 2 Then
            split_size = ""
        Else
            key_word = Mid(split_size, 1, InStr(split_size, "x") - 1)
            If IsNumeric(key_word) = False Then
                split_size = ""
            End If
        End If
        
        If file_type <> "" Then
            return_download_list_new = return_download_list_new & file_type & "|" & url_str & "|" & split_title & "(Vid_" & vid & "_" & split_size & ")." & file_type & "|" & html_str & vbCrLf
        End If
    Next
    return_download_list_new = return_download_list_new & "0"
End If
End Function

Function return_download_list_old(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list_old = ""
html_str=UTF8DecodeURI(html_str)
If InStr(LCase(html_str),"&url=") > 0 Then
	Dim split_str,split_i,split_title,key_word,split_size,file_type
	'&fmt_list=43/640x360/99/0/0,34/640x360/9/0/115,18/640x360/9/0/115,5/320x240/7/0/0,36/320x240/99/0/0,17/176x144/99/0/0&
	key_word="&fmt_list="
	url_str=""
	url_str=Mid(html_str,InStr(LCase(html_str), key_word)+len(key_word))
	url_str=Mid(url_str,1,InStr(LCase(url_str), "&")-1)
	split_size=split(url_str,",")
	
	'&title=イラストレーター+KEI+Live+Painting&
	key_word="&title="
	url_str=""
	url_str=Mid(html_str,InStr(LCase(html_str), key_word)+len(key_word))
	url_str=Mid(url_str,1,InStr(LCase(url_str), "&")-1)
	'title
	split_title=replace(replace(url_str,"+"," "),"|","_")
	
	key_word="&url="
	html_str=Mid(html_str,InStr(LCase(html_str), key_word)+len(key_word))
	split_str = Split(html_str, "&url=")
	
	For split_i=0 to UBound(split_str)
		url_str=""
		html_str=""
		url_str=Mid(split_str(split_i),1,InStr(split_str(split_i), "&")-1)
		url_str=UTF8DecodeURI(url_str)
		
		'&sig=581456E233472AC93ABAF398F360D22F8C1860B4.7D12462F28E2BFE7F5974C5EBC03AF0A24A253F9&
		'->&signature=581456E233472AC93ABAF398F360D22F8C1860B4.7D12462F28E2BFE7F5974C5EBC03AF0A24A253F9
		key_word="&sig="
		html_str=Mid(split_str(split_i),InStr(LCase(split_str(split_i)), key_word)+len(key_word))
		html_str=Mid(html_str,1,InStr(html_str, "&")-1)
		'url
		url_str=url_str & "&signature=" & html_str
		'If instr(LCase(url_str),"http://")=1 Then url_str="https://" & mid(url_str,8)
		
		'size
		If instr(split_size(split_i),"/")>0 Then split_size(split_i)=mid(split_size(split_i),instr(split_size(split_i),"/")+1)
		If instr(split_size(split_i),"/")>0 Then split_size(split_i)=mid(split_size(split_i),1,instr(split_size(split_i),"/")-1)
	
		'&type=video/webm;+codecs="vp8.0,+vorbis"&
		'&type=video/x-flv&
		'&type=video/mp4;+codecs="avc1.42001E,+mp4a.40.2"&
		'&type=video/x-flv&
		'&type=video/3gpp;+codecs="mp4v.20.3,+mp4a.40.2"&
		'&type=video/3gpp;+codecs="mp4v.20.3,+mp4a.40.2"&
		'Watch online   (WebM(VP8), 640 x 360, Stereo 44KHz Vorbis)
		'Download   (FLV, 640 x 360, Stereo 44KHz AAC)
		'Download   (MP4(H.264), 640 x 360, Stereo 44KHz AAC)
		'Download   (FLV, 400 x 240, Mono 44KHz MP3)
		'Download   (3GP, 320 x 240, Stereo 44KHz AAC)
		'Download   (3GP, 176 x 144, Stereo 44KHz AAC)
		key_word="&type="
		file_type=""
		html_str=Mid(split_str(split_i),InStr(LCase(split_str(split_i)), key_word)+len(key_word))
		html_str=Mid(html_str,1,InStr(html_str, "&")-1)
		html_str=UTF8DecodeURI(html_str)
		If InStr(LCase(html_str), "video/webm")=1 Then
			file_type="webm"
		ElseIf InStr(LCase(html_str), "video/3gpp")=1 Then
			file_type="3gpp"
		ElseIf InStr(LCase(html_str), "video/x-flv")=1 Then
			file_type="flv"			
		ElseIf InStr(LCase(html_str), "video/mp4")=1 Then
			file_type="mp4"		
		Else
			file_type="flv"		
		End If
		
		If 	file_type<>"" Then
			return_download_list_old = return_download_list_old & file_type & "|" & url_str & "|" & split_title & "(Vid_" & vid & "_" & split_size(split_i) & ")." & file_type & "|" & html_str & vbcrlf
		End If
	Next
	return_download_list_old=return_download_list_old & "0"
End If
End Function

'--------------------------------------------------------
Function UTF8DecodeURI(ByVal strIn)
UTF8DecodeURI = ""
Dim sl: sl = 1
Dim tl: tl = 1
Dim key: key = "%"
Dim kl: kl = Len(key)
sl = InStr(sl, strIn, key, 1)
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
sl = InStr(sl, strIn, key, 1)
Loop
UTF8DecodeURI = UTF8DecodeURI & Mid(strIn, tl)
End Function