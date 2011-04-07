'2011-02-10 http://www.shanhaijing.net/163
Dim user_id,album_id,page_id,retry_times,get_info

Function return_download_url(ByVal url_str)
On Error Resume Next

'http://?*.blog.163.com/album/?albumId=?*
'http://?*.blog.163.com/*
'http://blog.163.com/?*
'http://?*.blog.163.com/album/?albumId=?*#start_page=1
'http://blog.163.com/wehi/album/?albumId=fks_085067080085086071087085074065082086084068086
'http://hyde3331201.blog.163.com/prevAlbumsInUser.do?albumId=fks_087065080081080065082083084095087085082071086080085069
'http://hyde3331201.blog.163.com/album/
'http://blog.163.com/hyde3331201/album/

'http://blog.163.com/wehi/album/?albumId=63181763
'http://wehi.blog.163.com/album/#m=1&aid=63181763&p=1

get_info=""
If InStr(LCase(url_str), "http://s")=1 and LCase(Right(url_str, 3))=".js" and Len(url_str)>25 Then
	get_info="js"
	return_download_url="inet|10,13|" & url_str
	
ElseIf InStr(1, url_str, "?albumId=", 1) > 0 or InStr(1, url_str, "&aid=", 1) > 0 Then

	If InStr(1, url_str, ".blog.163.com/", 1)>0 Then
		url_str=Mid(url_str,InStr(1, url_str, "http://", 1)+7)
		user_id=Mid(url_str,1,InStr(1, url_str, ".blog.163.com/", 1)-1)
	Else
		url_str=Mid(url_str,InStr(1, url_str, "http://blog.163.com/", 1)+20)
		user_id=Mid(url_str,1,InStr(url_str, "/")-1)	
	End If

	
	If InStr(1, url_str, "&aid=", 1) > 0 Then
		url_str=Mid(url_str,InStr(1, url_str, "&aid=", 1)+5)
		url_str=Mid(url_str,1,InStr( url_str, "&")-1)
	Else
		url_str=Mid(url_str,InStr(1, url_str, "albumId=", 1)+8)
	End If

	If InStr(url_str, "#")>0 Then		
		album_id=Mid(url_str,1,InStr(url_str, "#")-1)		
		url_str=Mid(url_str,InStrRev(url_str, "#")+1)
		If InStr(url_str, "start_page=")=1 Then
			url_str=Mid(url_str,12)
			If IsNumeric(url_str) Then
				page_id=CLng(url_str)
				If page_id<2 Then page_id=1
			Else
				page_id=1
			End If
		Else
			page_id=1
		End If
	Else
		page_id=1
		album_id=url_str
	End If
	
	If is_username(user_id) = True and IsNumeric(album_id) Then
		get_info="newblog"
	End If
	
	return_download_url="inet|10,13|http://" & user_id & ".blog.163.com/#m=1&aid=" & album_id
	
Else
	If InStr(1, url_str, "http://blog.163.com/", 1) > 0 Then
		url_str=Mid(url_str,InStr(1, url_str, "http://blog.163.com/", 1)+20)
		If InStr(url_str,"/")>0 Then url_str=Mid(url_str,1,InStr(url_str,"/")-1)
		user_id=url_str			
	Else
		url_str=Mid(url_str,InStr(1, url_str, "http://", 1)+7)
		user_id=Mid(url_str,1,InStr(1, url_str, ".blog.163.com/", 1)-1)	
	End If
	return_download_url=""
	'http://driger1024.blog.163.com/album/all/
	'http://blog.163.com/lihang./album/
	If is_username(user_id) =True Then return_download_url="inet|10,13|http://blog.163.com/" & user_id & "/album/"
	
End If
retry_times=0

End Function
'--------------------------------------------------------
Function return_albums_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_albums_list=""
If retry_times=0 Then
	retry_times=retry_times+1
	If InStr(LCase(html_str), ",photo163hostname:'") > 0 Then
		Dim photo163hostname
		photo163hostname=Mid(html_str,InStr(LCase(html_str), ",photo163hostname:'")+Len(",photo163hostname:'"))
		photo163hostname=Mid(photo163hostname,1,InStr(photo163hostname, "'")-1)
		If is_username(photo163hostname) =True Then user_id=photo163hostname
	End If	
End If

If InStr(LCase(html_str), ",cachefileurl:""") > 0 Then

	html_str=Mid(html_str,InStr(LCase(html_str), ",cachefileurl:""")+15)
	html_str="http://" & Trim(Mid(html_str,1,InStr(LCase(html_str),Chr(34))-1))
	return_albums_list="1|inet|10,13|" & html_str

ElseIf InStr(LCase(html_str), "g_albumlist.push({id:") > 0 Then

html_str=Mid(html_str,InStr(LCase(html_str), "g_albumlist.push({id:")+21)
html_str=Mid(html_str,InStr(LCase(html_str), "'")+1)
html_str=Mid(html_str,1,InStr(LCase(html_str), "</script>")-1)

'html_str=rename_utf8(html_str)

Dim album_list,psw_tf,page_add

album_list=split(html_str,"g_albumList.push({id: '")
	
For i = 0 To UBound(album_list)
DoEvents
	If InStr(LCase(album_list(i)), "albumname:  '") > 0 Then
		'id
		album_id=Mid(album_list(i),1,InStr(album_list(i), "'")-1)
		
		album_list(i)=Mid(album_list(i),InStr(LCase(album_list(i)), "albumname:  '"))
		album_list(i)=Mid(album_list(i),InStr( album_list(i), "'")+1)
		'name
		html_str=replace(vbsUnEscape(Trim(Mid(album_list(i),1,InStr(album_list(i), "'")-1))),"|","_")
		
		album_list(i)=Mid(album_list(i),InStr(LCase(album_list(i)), "albumdescription: '"))
		album_list(i)=Mid(album_list(i),InStr( album_list(i), "'")+1)
		
		'albumDescription
		url_str= replace(vbsUnEscape(Trim(Mid(album_list(i),1,InStr(album_list(i), "'")-1))),"|","_")
		
		album_list(i)=Mid(album_list(i),InStr(album_list(i), "'")+1)
		'locked?
		If InStr(LCase(album_list(i)), "coverpassword: 1") > 0 Then
		psw_tf="1|"
		Else
		psw_tf="0|"
		End If
		
		album_list(i)=Mid(album_list(i),InStr(LCase(album_list(i)), "photocount:")+11)
		'pic_num
		album_list(i)=Trim(Mid(album_list(i),1,InStr(album_list(i), ",")-1))
			
		If IsNumeric(album_list(i))= False Then
			'http://user_id.blog.163.com/album/?albumId=album_id
			album_list(i)=""
			album_list(i)=psw_tf & album_list(i) & "|http://blog.163.com/" & user_id & "/album/?albumId=" & album_id & "|" & html_str & "|" & url_str & vbcrlf
		Else
			page_id=CLng(album_list(i))
			If page_id>9999 Then
				page_add=1
				album_list(i)=""
				Do While page_id>0
				If page_id>9999 Then
				'http://user_id.blog.163.com/album/?albumId=album_id#start_page=n
				album_list(i)=album_list(i) & psw_tf & "9999|http://blog.163.com/" & user_id & "/album/?albumId=" & album_id & "#start_page=" & page_add & "|" & html_str & "|" & url_str & vbcrlf
				Else
				'http://user_id.blog.163.com/album/?albumId=album_id#start_page=n
				album_list(i)=album_list(i) & psw_tf & page_id & "|http://blog.163.com/" & user_id & "/album/?albumId=" & album_id & "#start_page=" & page_add & "|" & html_str & "|" & url_str & vbcrlf
				End If
				page_id=page_id-9999
				page_add=page_add+1				
				loop
			Else
				'http://user_id.blog.163.com/album/?albumId=album_id
				album_list(i)=psw_tf & page_id & "|http://blog.163.com/" & user_id & "/album/?albumId=" & album_id & "|" & html_str & "|" & url_str & vbcrlf
			End If
		End If
	End If
Next
    return_albums_list=join(album_list,"") & "0"
    
ElseIf InStr(html_str, "=[{id:") > 0 Then

        'var g_a$514028s='1187485;1187484;1187472;1187470;1187468;1187464;1187460;1187457;1187456;1187453;1530930;';
        'var g_a$514028d=[{id:
        '1187468,name:'虫袄 虫师二十景 漆原友纪画集 ',s:3,desc:'蟲師二十景 漆原友紀画集',st:1,au:0,count:14,t:1220710254100,ut:0,curl:'396/HjWuimtpsp-486EMHXLQ3A==/3070610520936616491.jpg',surl:'396/OO0u-aWixlqZ2iVH5rT2vg==/3070610520936616515.jpg',dmt:1220924333238,alc:true,comm:'',comdmt:0,kw:'',purl:'s1.photo.163.com/2vNO5QX8iwqKXVr2xX2Oiw==/72620543991354232.js'
        '},{id:
        '1530930,name:'password_text',s:0,desc:'password_text',st:1,au:1,count:0,t:1221048756165,ut:0,curl:'',surl:'',dmt:1221583000801,alc:true,comm:'',comdmt:0,kw:'',purl:''}];
        '63181790,name:'Yours 堀部秀郎',s:1,desc:'Yours - 堀部秀郎 ART WORKS\r(source BMP: 3500px height, 300dpi; convert to JPG: 3000px height, 72dpi)',st:0,au:0,count:126,t:1173461254669,ut:0,curl:'49/photo/NFYlplbThnFFe7rqi0CFRQ==/2314287258515572734.jpg',surl:'9/photo/J7isfp1SXK5Yl7JNuVDErQ==/2549037389092630546.jpg',dmt:1250102957187,alc:true,comm:'',comdmt:1250102957187,comnum:0,kw:'',purl:'s5.ph.126.net/F8zW7A3OzurlBDM5rqKMMQ==/121034239989244051.js'
	'},{id:63181775,name:'hah&aha',s:1,desc:'',st:1,au:1,count:13,
	't:1173460408684,ut:0,curl:'48/photo/zTq_fcPUQ4sepeF6yB36Qg==/4278982595956160994.jpg',
	'surl:'31/photo/jhIalo7fRvWq5IbBaoqjdg==/4241827899029719017.jpg',
	'dmt:1250145227443,alc:true,comm:'',comdmt:1250145227443,comnum:0,kw:'',purl:''},
	      
        html_str = Mid(html_str, InStr(html_str, "=[{id:") + 6) '定位到第一个相册的ID头
        html_str = Mid(html_str, 1, InStr(html_str, "'}];") - 1) '定位最后一个相册
        
        Dim albumsINFO,temp(3),iCount,albumsID
        
        albumsINFO = Split(html_str, "'},{id:")        
        html_str = ""        
        iCount = UBound(albumsINFO)
                
        For cout_num = 0 To iCount
            
            temp(0) = Mid(albumsINFO(cout_num), InStr(albumsINFO(cout_num), ",name:'") + 7)
            temp(3) = temp(0)
            
            temp(0) = Trim(Mid(temp(0), 1, InStr(temp(0), "'") - 1))
            If temp(0) = "" Then temp(0) = user_id & "[Noname_Albums]"            
            
            temp(3) = Mid(temp(3), InStr(temp(3), "'") + 1)
            temp(3) = Mid(temp(3), InStr(temp(3), ",desc:'") + 7)
            temp(2) = temp(3)
            temp(1) = temp(3)
            
            temp(3) = Trim(Mid(temp(3), 1, InStr(temp(3), "'") - 1))
            
            temp(1) = Mid(temp(1), InStr(temp(1), "'") + 1)
            temp(1) = Mid(temp(1), InStr(temp(1), "au:") + 3)
            temp(1) = Trim(Mid(temp(1), 1, InStr(temp(1), ",") - 1))
            
            temp(2) = Mid(temp(2), InStr(temp(2), "'") + 1)
            temp(2) = Mid(temp(2), InStr(temp(2), "count:") + 6)
            temp(2) = Trim(Mid(temp(2), 1, InStr(temp(2), ",") - 1))
            If IsNumeric(temp(2))=flase Then temp(2) = ""
            
            albumsID = ""
            
            albumsID = Trim(Mid(albumsINFO(cout_num), InStrRev(albumsINFO(cout_num), "'") + 1))
            
            If albumsID = "" Then
            albumsID = "http://blog.163.com/" & user_id & "/album/?albumId=" & Mid(albumsINFO(cout_num), 1, InStr(albumsINFO(cout_num), ",") - 1)
            Else
            albumsID = "http://" & albumsID
            End If
            
            If temp(1) <> "1" Then temp(1) = "0"
            
            'book_name temp(0))
            'book_psw temp(1)
            'book_ID
            'book_number temp(2)
            'book_disc temp(3)
            albumsINFO(cout_num)=temp(1) & "|" & temp(2) & "|" & albumsID & "|" & temp(0) & "|" & temp(3) & vbcrlf
                  
        Next
        
	return_albums_list=join(albumsINFO,"") & "0"
	
ElseIf retry_times<3 Then
	
	retry_times=retry_times+1
	return_albums_list="1|inet|10,13|http://photo.163.com/photo/dwrcross/" & user_id & "/u/" & user_id & "/dwr/call/plaincall/UserSpaceBean.getUserSpace.dwr?callCount=1&scriptSessionId=%24%7BscriptSessionId%7D957&c0-scriptName=UserSpaceBean&c0-methodName=getUserSpace&c0-id=0&c0-param0=string%3A" & user_id & "&batchId=" & Int(Time() * 1000000) & "|http://photo.163.com/photo/"
	
ElseIf retry_times<5 Then
	retry_times=retry_times+1
	'http://photo.163.com/photo/hera_er/dwr/call/plaincall/UserSpaceBean.getUserSpace.dwr
	'photo.163.com/photo/hera_er/dwr/call/plaincall/UserSpaceBean.getUserSpace.dwr
	return_albums_list="callCount=1" & "&for_ox163_replace_vbcrlf&"
	return_albums_list=return_albums_list & "scriptSessionId=${scriptSessionId}187" & "&for_ox163_replace_vbcrlf&"
	return_albums_list=return_albums_list & "c0-scriptName=UserSpaceBean" & "&for_ox163_replace_vbcrlf&"
	return_albums_list=return_albums_list & "c0-methodName=getUserSpace" & "&for_ox163_replace_vbcrlf&"
	return_albums_list=return_albums_list & "c0-id=0" & "&for_ox163_replace_vbcrlf&"
	return_albums_list=return_albums_list & "c0-param0=string:" & user_id & "&for_ox163_replace_vbcrlf&"
	return_albums_list=return_albums_list & "batchId=" & Int(Time() * 1000000)
	return_albums_list="1|inet|10,13|http://photo.163.com/photo/" & user_id & "/dwr/call/plaincall/UserSpaceBean.getUserSpace.dwr|http://photo.163.com|" & return_albums_list
	MsgBox user_id
Else
    return_albums_list = "0"
End If
End Function
'----------------------------------------------------------
Function return_download_list(ByVal html_str, ByVal url_str)
On Error Resume Next
return_download_list = ""
If retry_times=0 and get_info<>"js" Then
	retry_times=retry_times+1
	If InStr(LCase(html_str), ",photo163hostname:'") > 0 Then
		Dim photo163hostname
		photo163hostname=Mid(html_str,InStr(LCase(html_str), ",photo163hostname:'")+Len(",photo163hostname:'"))
		photo163hostname=Mid(photo163hostname,1,InStr(photo163hostname, "'")-1)
		If is_username(photo163hostname) =True Then user_id=photo163hostname
	End If
	
	If is_username(user_id) = True and IsNumeric(album_id) Then
		'return_download_list="inet|10,13|http://photo.163.com/photo/dwrcross/" & user_id & "/dwr/call/plaincall/AlbumBean.getAlbumData.dwr?callCount=1&scriptSessionId=%24%7BscriptSessionId%7D822&c0-scriptName=AlbumBean&c0-methodName=getAlbumData&c0-id=0&c0-param0=number%3A" & album_id & "&c0-param1=string%3A&c0-param2=string%3Afromblog&c0-param3=number%3A&c0-param4=boolean%3Afalse&batchId=7&ntime=" & CDbl(Now())
		'return_download_list="inet|10,13|http://photo.163.com/photo/dwrcross/" & user_id & "/u/" & user_id & "/dwr/call/plaincall/AlbumBean.getAlbumData.dwr?callCount=1&scriptSessionId=%24%7BscriptSessionId%7D822&c0-scriptName=AlbumBean&c0-methodName=getAlbumData&c0-id=0&c0-param0=number%3A" & album_id & "&c0-param1=string%3A&c0-param2=string%3Afromblog&c0-param3=number%3A&c0-param4=boolean%3Afalse&batchId=4&ntime=" & CDbl(Now())
		return_download_list="1|inet|10,13|http://photo.163.com/photo/" & user_id & "/dwr/call/plaincall/AlbumBean.getAlbumData.dwr?callCount=1&scriptSessionId=%24%7BscriptSessionId%7D187&c0-scriptName=AlbumBean&c0-methodName=getAlbumData&c0-id=0&c0-param0=number%3A" & album_id & "&c0-param1=string%3A&c0-param2=string%3Afromblog&c0-param3=string%3A32350899&c0-param4=boolean%3Afalse&batchId=" & Int(Time() * 1000000)
	ElseIf is_username(user_id) = True Then
		return_download_list="1|inet|10,13|http://blog.163.com/" & user_id & "/album/dwr/call/plaincall/Photo.getPhotosInAlbum.dwr||" & replace(post_str(page_id),vbcrlf,"&for_ox163_replace_vbcrlf&")
	End If
	Exit Function
End If




If get_info="newblog" and InStr(html_str, ".js"")") > 10 Then
	html_str=Mid(html_str,1,InStr(html_str, ".js"")")+2)
	html_str=Mid(html_str,InStr(html_str, Chr(34))+1)
	If LCase(Left(html_str,7))<>"http://" Then html_str="http://" & html_str
	get_info="js"
	return_download_list="1|inet|10,13|" & html_str
	
ElseIf get_info="js" and InStr(html_str, "=[{id:") > 0 Then
	
	'定位到第一张图片的文本头
	html_str = Mid(html_str, InStr(html_str, "=[{id:") + 6)
	'定位到最后一张图片
	html_str = Mid(html_str, 1, InStr(html_str, "}];") - 3)
	
	Dim a, b,new163pic_str_split,ourl
	Dim cout_num
	
	a=""
	b=""
	cout_num = 0
	
	new163pic_str_split = Split(html_str, "},{id:")
    
    For i = 0 To UBound(new163pic_str_split)
	ourl = ""

	'blog
	'{id:2665422496,s:1,
	'ourl:'3/photo/bveEQxqzGf3-iLP4ihV4yQ==/855402454224501762.jpg',
	'ow:7449,oh:3000,
	'murl:'3/photo/V1BxMjQ9vNeTZiwKlmBfZA==/855402454224501764.jpg',
	'surl:'3/photo/yX5FI7wVmU0bOFdwz2a5qg==/855402454224501766.jpg',
	'turl:'47/photo/3Gy7l6-IIgSEXdgW2it6Fw==/844706405109833346.jpg',
	'qurl:'3/photo/OGfb2qN6Az7V5rd0K89R_w==/855402454224501767.jpg',
	'desc:'colors000-1',t:1224488234491,comm:'',comdmt:0,comnum:0,exif:'',kw:',e^unknow,e^unknow'
	'},{id:

	If InStr(LCase(new163pic_str_split(i)), ",ourl:'") > 1 Then
	ourl = Mid(new163pic_str_split(i), InStr(LCase(new163pic_str_split(i)), ",ourl:'") + 7)
	ourl = Trim(Mid(ourl, 1, InStr(ourl, "'") - 1))
	End If

	new163pic_str_split(i) = Mid(new163pic_str_split(i), InStr(LCase(new163pic_str_split(i)), ",murl:'") + 7)
    
    
    	If ourl = "" Then
        	a = Mid(new163pic_str_split(i), 1, InStr(LCase(new163pic_str_split(i)), "'") - 1)
    	Else
        	a = ourl
    	End If    
    
	'第一种
	'616/bq4wr0XiQkbDUgWICDBoTg==/1026539240063803524.jpg
	'http://img616.photo.163.com/bq4wr0XiQkbDUgWICDBoTg==/1026539240063803524.jpg
	'第二种
	'/photo/nzovvldOrJcsKJ2iLjW8rA==/2845149064591786998.jpg
	'http://img.bimg.126.net/photo/nzovvldOrJcsKJ2iLjW8rA==/2845149064591786998.jpg
	b = Mid(a, 1, InStr(a, "/") - 1)
	a = Mid(a, InStr(a, "/"))
    
    	'M pic url or Ourl
     	If Left(LCase(a), 7) = "/photo/" Then
		a = "http://img" & b & ".bimg.126.net" & a
	Else
		a = "http://img" & b & ".photo.163.com" & a
	End If
    
	new163pic_str_split(i) = Mid(new163pic_str_split(i), InStr(LCase(new163pic_str_split(i)), "',desc:'") + 8)
    
	'描述
	b = Trim(Mid(new163pic_str_split(i), 1, InStr(new163pic_str_split(i), "'") - 1))
    
	If b = "" Then b = Mid(a, InStrRev(a, "/") + 1)
	new163pic_str_split(i) = ""
	new163pic_str_split(i) = LCase(Mid(b, InStrRev(b, ".")))
    
	If new163pic_str_split(i) <> LCase(Mid(a, InStrRev(a, "."))) Then b = b & Mid(a, InStrRev(a, "."))
    
	new163pic_str_split(i) = "|" & a & "|" & b & "|"
        
    Next
	return_download_list = join(new163pic_str_split,vbcrlf) & vbcrlf & "0"	
	
ElseIf InStr(html_str,".photoName=""")>20 Then

	html_str=Mid(html_str,InStr(html_str,".photoName=""")+12)
	Dim photo_list
	photo_list=split(html_str,".photoName=""")

For i = 0 To UBound(photo_list)
DoEvents
	'name
	html_str=Trim(Mid(photo_list(i),1,InStr(photo_list(i),Chr(34))-1))
	html_str=replace(vbsUnEscape(html_str),"|","_")
	If Len(html_str)>30 Then html_str=Left(html_str,29) & "~"
	If html_str="" Then html_str="no_name"
	'url
	photo_list(i)=Mid(photo_list(i),InStr(photo_list(i),Chr(34))+1)
	photo_list(i)=Mid(photo_list(i),InStr(photo_list(i),".url=""")+6)
	url_str=Mid(photo_list(i),1,InStr(photo_list(i),Chr(34))-1)
	'name.pic_type
	If Mid(LCase(url_str),instrrev(url_str,".")) <> Mid(LCase(html_str),instrrev(html_str,".")) Then html_str=html_str & Mid(url_str,instrrev(url_str,"."))
	photo_list(i) = "|" & url_str & "|" & html_str & "|"
Next
	return_download_list = join(photo_list,vbcrlf) & vbcrlf & "0"	
Else
	return_download_list = ""
End If

End Function
'------------------------------------------------------------------
Function return_password_rules(ByVal html_str, ByVal pass_word)
On Error Resume Next
If get_info="newblog" Then
	pass_word=urlencode(UTF8EncodeURI(urldecode(pass_word)))
	return_password_rules = "return_ad_password_rules|inet|10,13|http://photo.163.com/photo/dwrcross/" & user_id & "/u/" & user_id & "/dwr/call/plaincall/AlbumBean.getAlbumData.dwr?callCount=1&scriptSessionId=%24%7BscriptSessionId%7D822&c0-scriptName=AlbumBean&c0-methodName=getAlbumData&c0-id=0&c0-param0=number%3A" & album_id & "&c0-param1=string%3A" & pass_word & "&c0-param2=string%3Afromblog&c0-param3=number%3A&c0-param4=boolean%3Afalse&batchId=4&ntime=" & CDbl(Now())
Else
	Dim post_pw
	post_pw = ""
	post_pw=post_pw & "callCount=1" & vbcrlf
	post_pw=post_pw & "scriptSessionId=${scriptSessionId}175" & vbcrlf
	post_pw=post_pw & "c0-scriptName=Album" & vbcrlf
	post_pw=post_pw & "c0-methodName=checkAlbumPassword" & vbcrlf
	post_pw=post_pw & "c0-id=0" & vbcrlf
	post_pw=post_pw & "c0-param0=string:" & album_id & vbCrLf
	post_pw=post_pw & "c0-param1=string:" & pass_word & vbcrlf
	post_pw=post_pw & "c0-param2=boolean:false" & vbcrlf
	post_pw=post_pw & "batchId=7"
	return_password_rules = "http://blog.163.com/" & user_id & "/album/dwr/call/plaincall/Album.checkAlbumPassword.dwr|" & post_pw & "||1|var s0=[];"
End If
MsgBox return_password_rules
End Function

Function return_ad_password_rules(ByVal html_str, ByVal url_str, ByVal pass_word)
On Error Resume Next
	return_ad_password_rules=""
	If InStr(html_str, ".js"")") > 10 Then
		return_ad_password_rules="password_correct"
	End If
MsgBox return_ad_password_rules
End Function
'------------------------------------------------------------------
Function is_username(ByVal username)
On Error Resume Next
is_username =True
If Len(username) > 2 And Len(username) < 19 Then
For i = 1 To Len(username)
DoEvents
If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.-_@", Mid(username, i, 1)) < 1 Then
	is_username = False
	Exit Function
End if
Next
Else
is_username = False
End If
End Function

Function vbsUnEscape(str) 
On Error Resume Next
    dim i,s,c
    s="" 
    For i=1 to Len(str) 
    DoEvents
        c=Mid(str,i,1) 
        If Mid(str,i,2)="\u" and i<=Len(str)-5 Then 
            If IsNumeric("&H" & Mid(str,i+2,4)) Then 
                s = s & CHRW(CInt("&H" & Mid(str,i+2,4))) 
                i = i+5 
            Else 
                s = s & c 
            End If
        Else 
            s = s & c 
        End If 
    Next 
    vbsUnEscape = replace(s,"\/","/") 
End Function 

Function post_str(byval start_page) 
post_str = ""
start_page=(start_page-1)*9999
post_str = post_str & "callCount=1" & vbCrLf
post_str = post_str & "scriptSessionId=${scriptSessionId}375" & vbCrLf
post_str = post_str & "c0-scriptName=Photo" & vbCrLf
post_str = post_str & "c0-methodName=getPhotosInAlbum" & vbCrLf
post_str = post_str & "c0-id=0" & vbCrLf
post_str = post_str & "c0-param0=string:" & album_id & vbCrLf
post_str = post_str & "c0-param1=number:1" & vbCrLf '排序方式
post_str = post_str & "c0-param2=number:" & start_page & vbCrLf '起始图片为0开始
post_str = post_str & "c0-param3=number:9999" & vbCrLf '单次下载图片数量 100
post_str = post_str & "c0-param4=boolean:false" & vbCrLf
post_str = post_str & "batchId=0"
End Function


'------------------------------------------------------------------
Function is_username(ByVal username)
On Error Resume Next
is_username =True
If Len(username) > 2 And Len(username) < 19 Then
For i = 1 To Len(username)
DoEvents
If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.-_@", Mid(username, i, 1)) < 1 Then
	is_username = False
	Exit Function
End if
Next
Else
is_username = False
End If
End Function

Function vbsUnEscape(str) 
On Error Resume Next
    dim i,s,c
    s="" 
    For i=1 to Len(str) 
    DoEvents
        c=Mid(str,i,1) 
        If Mid(str,i,2)="\u" and i<=Len(str)-5 Then 
            If IsNumeric("&H" & Mid(str,i+2,4)) Then 
                s = s & CHRW(CInt("&H" & Mid(str,i+2,4))) 
                i = i+5 
            Else 
                s = s & c 
            End If
        Else 
            s = s & c 
        End If 
    Next 
    vbsUnEscape = replace(s,"\/","/") 
End Function

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

Function URLDecode(strURL)
On Error Resume Next
    Dim I

    If InStr(strURL, "%") = 0 Then
        URLDecode = strURL
        Exit Function
    End If

    For I = 1 To Len(strURL)
        If Mid(strURL, I, 1) = "%" Then
            If Eval("&H" & Mid(strURL, I + 1, 2)) > 127 Then
                URLDecode = URLDecode & Chr(Eval("&H" & Mid(strURL, I + 1, 2) & Mid(strURL, I + 4, 2)))
                I = I + 5
            Else
                URLDecode = URLDecode & Chr(Eval("&H" & Mid(strURL, I + 1, 2)))
                I = I + 2
            End If
        Else
            URLDecode = URLDecode & Mid(strURL, I, 1)
        End If
    Next
End Function

Function URLEncode(ByVal vstrIn)
On Error Resume Next
strReturn = ""
vstrIn = Replace(vstrIn, "%", "%25")
    Dim i
For i = 1 To Len(vstrIn)
ThisChr = Mid(vstrIn, i, 1)
If Abs(Asc(ThisChr)) < &HFF Then
strReturn = strReturn & ThisChr
Else
innerCode = Asc(ThisChr)
If innerCode < 0 Then
innerCode = innerCode + &H10000
End If
Hight8 = (innerCode And &HFF00) \ &HFF
Low8 = innerCode And &HFF
strReturn = strReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
End If
Next
strReturn = Replace(strReturn, "!", "%21")
strReturn = Replace(strReturn, Chr(34), "%22")
strReturn = Replace(strReturn, "#", "%20")
strReturn = Replace(strReturn, "$", "%24")
strReturn = Replace(strReturn, "&", "%26")
strReturn = Replace(strReturn, "'", "%27")
strReturn = Replace(strReturn, "(", "%28")
strReturn = Replace(strReturn, ")", "%29")
strReturn = Replace(strReturn, "*", "%2A")
strReturn = Replace(strReturn, "+", "%2B")
strReturn = Replace(strReturn, ",", "%2C")
strReturn = Replace(strReturn, ".", "%2E")
strReturn = Replace(strReturn, "/", "%2F")
strReturn = Replace(strReturn, ":", "%3A")
strReturn = Replace(strReturn, ";", "%3B")
strReturn = Replace(strReturn, "<", "%3C")
strReturn = Replace(strReturn, "=", "%3D")
strReturn = Replace(strReturn, ">", "%3E")
strReturn = Replace(strReturn, "?", "%3F")
strReturn = Replace(strReturn, "@", "%40")
strReturn = Replace(strReturn, "[", "%5B")
strReturn = Replace(strReturn, "\", "%5C")
strReturn = Replace(strReturn, "]", "%5D")
strReturn = Replace(strReturn, "^", "%5E")
strReturn = Replace(strReturn, "`", "%60")
strReturn = Replace(strReturn, "{", "%7B")
strReturn = Replace(strReturn, "|", "%7C")
strReturn = Replace(strReturn, "}", "%7D")
strReturn = Replace(strReturn, "~", "%7E")
strReturn = Replace(strReturn, Chr(32), "%20")
URLEncode = strReturn
End Function