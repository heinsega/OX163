'2009-11-24 163.shanhaijing.net
dim main_url

Function return_download_url(ByVal url_str)
'http://www.wallcoo.com/cartoon/The_fiery_English_alphabets_and_fiery_numbers_1920_1600/html/wallpaper3.html
'http://www.wallcoo.com/cartoon/The_fiery_English_alphabets_and_fiery_numbers_1920_1600/index.html
'http://www.wallcoo.com/cartoon/Abstract_Blue_backgrounds_1920x1200/index.html
'http://www.wallcoo.com/film/2009_11_2012/index.html
'http://www.wallcoo.com/cartoon/The_fiery_English_alphabets_and_fiery_numbers_1920_1600/html/wallpaper16.html

On Error Resume Next
dim split_str

return_download_url=""

url_str=mid(url_str,24)
split_str=split(url_str,"/")
if ubound(split_str)>0 then
	main_url="http://www.wallcoo.com/" & split_str(0) & "/" & split_str(1) & "/"
	'http://www.wallcoo.com/cartoon/The_fiery_English_alphabets_and_fiery_numbers_1920_1600/
	return_download_url = "inet|10,13|http://www.wallcoo.com/" & split_str(0) & "/" & split_str(1) & "/html/wallpaper1.html"
end if
'OX163_urlpage_Referer="Referer: http://www.wallcoo.com/" & vbCrLf & "User-Agent: Mozilla/4.0 (compatible: MSIE 8.0)"
End Function

'--------------------------------------------------------

Function return_download_list(ByVal html_str, ByVal url_str)
'http://www.wallcoo.com/cartoon/The_fiery_English_alphabets_and_fiery_numbers_1920_1600/wallpapers/
'1024x768/The_fiery_numbers_picture_4108983.jpg
'1280x800/The_fiery_numbers_picture_4108983.jpg
'1600x1200/The_fiery_numbers_picture_4108983.jpg

On Error Resume Next
dim screen_split,url_split

return_download_list = ""
url_str=html_str


If instr(html_str,"<select name=""imagelist""")>0 and instr(html_str,"<dt>¿ÉÏÂÔØ³ß´ç£º</dt>")>0 Then
	html_str=mid(html_str,instr(html_str,"<select name=""imagelist"""))
	html_str=mid(html_str,1,instr(html_str,"</select>"))
	html_str=mid(html_str,instr(html_str,"<option ")+8)
	url_split=split(html_str,"<option ")
	
	url_str=mid(url_str,instr(url_str,"<dt>¿ÉÏÂÔØ³ß´ç£º</dt>"))
	url_str=mid(url_str,instr(url_str,"<dd class=""size"">")+17)
	url_str=mid(url_str,1,instr(url_str,"</dd>")-1)
	if right(url_str,1)="|" then url_str=mid(url_str,1,len(url_str)-1)
	screen_split=split(url_str,"|")
	for i=0 to ubound(screen_split)
		screen_split(i)=replace(trim(screen_split(i)),"*","x")
	next
	html_str=""

	for i=0 to ubound(url_split)
		url_split(i)=mid(url_split(i),instr(url_split(i),">")+2)
		url_split(i)=mid(url_split(i),instr(url_split(i),". ")+2)
		url_split(i)=trim(mid(url_split(i),1,instr(url_split(i),"×ÀÃæ±ÚÖ½")-1))
		html_str=""
		for j=0 to ubound(screen_split)
			If screen_split(j)<>"" Then html_str=html_str & "|" & main_url & "wallpapers/" & screen_split(j) & "/" & url_split(i) & ".jpg|" & url_split(i) & "(" & screen_split(j) & ").jpg|" & screen_split(j) & vbCrLf
		next
		url_split(i)=html_str
	next

	return_download_list=join(url_split,"")

end if
return_download_list = return_download_list & "0"
End Function