Attribute VB_Name = "Parsing"
Option Explicit

Public Type ScriptInfo
    'tom.vbs|vbscript|GB2312|album|http://photo.tom.com/pim.php?*
    FileName As String
    Language As String
    Encoding As String
    HandleType As String
    Criteria As String
End Type

Public Type downloadInfo
    'inet|10,13|url|url_Referer|POST method
    isFinal As Boolean
    mode As DownloadMode
    excludeChar As String
    downloadURL As String
    refererINFO As String
    POSTmethod As String
End Type

Public Type AlbumInfo
    hasPassword As Boolean
    picCount As String
    URL As String
    AlbumName As String
    description As String
End Type

Public Enum DownloadMode
    OX_INET
    OX_WEB
End Enum



Public Function ParseInclude(ByVal sourceString As String) As ScriptInfo
    On Error Resume Next
    'include中的内容
    'tom.vbs|vbscript|GB2312|album|http://photo.tom.com/pim.php?*
    '判断完include之后发送到list_album_script或者list_photo_script中的内容
    'eht.vbs|vbscript|UTF-8|album|http://g.e-hentai.org/g/132226/e6f6d72be1/
    Dim script_ParseStr() As String
    ParseInclude.FileName = ""
    ParseInclude.Language = ""
    ParseInclude.Encoding = ""
    ParseInclude.HandleType = ""
    ParseInclude.Criteria = ""
    script_ParseStr = Split(sourceString, "|")
    If UBound(script_ParseStr) >= 0 Then ParseInclude.FileName = script_ParseStr(0)
    If UBound(script_ParseStr) >= 1 Then ParseInclude.Language = LCase$(script_ParseStr(1))
    If UBound(script_ParseStr) >= 2 Then ParseInclude.Encoding = script_ParseStr(2)
    If UBound(script_ParseStr) >= 3 Then ParseInclude.HandleType = LCase$(script_ParseStr(3))
    If UBound(script_ParseStr) >= 4 Then ParseInclude.Criteria = script_ParseStr(4)
End Function

Public Function ParseDownloadURL(ByVal sourceString As String) As downloadInfo
    On Error Resume Next
    'inet|10,13|url|url_Referer|POST method
    Dim script_ParseStr() As String
    Dim script_ParseStr_counts As Byte
    
    script_ParseStr_counts = 0
    ParseDownloadURL.isFinal = True
    ParseDownloadURL.mode = OX_INET
    ParseDownloadURL.excludeChar = ""
    ParseDownloadURL.downloadURL = ""
    ParseDownloadURL.refererINFO = ""
    ParseDownloadURL.POSTmethod = ""
    
    script_ParseStr = Split(sourceString, "|")
    If IsNumeric(script_ParseStr(0)) Then
        '0|inet|10,13|url|url_Referer|POST method
        script_ParseStr_counts = 1
        If Int(script_ParseStr(0)) > 0 Then ParseDownloadURL.isFinal = False
    End If
    If UBound(script_ParseStr) < (2 + script_ParseStr_counts) Then Exit Function
    
    If LCase$(script_ParseStr(0 + script_ParseStr_counts)) = "web" Then ParseDownloadURL.mode = OX_WEB
    ParseDownloadURL.excludeChar = script_ParseStr(1 + script_ParseStr_counts)
    ParseDownloadURL.downloadURL = Trim$(script_ParseStr(2 + script_ParseStr_counts))
    If UBound(script_ParseStr) > (2 + script_ParseStr_counts) Then ParseDownloadURL.refererINFO = Trim$(Replace(Replace(script_ParseStr(3 + script_ParseStr_counts), "&for_ox163_replace_vbcrlf&", vbCrLf), "&for_ox163_replace_vline&", "|"))
    If UBound(script_ParseStr) > (3 + script_ParseStr_counts) Then ParseDownloadURL.POSTmethod = Trim$(Replace(Replace(script_ParseStr(4 + script_ParseStr_counts), "&for_ox163_replace_vbcrlf&", vbCrLf), "&for_ox163_replace_vline&", "|"))
End Function

Public Function ParseAlbum(ByVal sourceString As String) As AlbumInfo()
    On Error Resume Next
    '0|23|http://comic.92wy.com/go/comicshow.aspx?id=1389&nameid=57|BLAME_第1卷|BLAME_第1卷
    Dim ParseAlbumStrSplit, ParseAlbumInfoSplit
    Dim i As Long, j As Long, ParseAlbumInfo() As AlbumInfo
    ParseAlbumStrSplit = Split(sourceString, vbCrLf)
    '判断数组
    If LBound(ParseAlbumStrSplit) <> 0 Then Exit Function
    ReDim ParseAlbumInfo(UBound(ParseAlbumStrSplit)) As AlbumInfo
    
    For i = 0 To UBound(ParseAlbumStrSplit)
        DoEvents
        If Form1.form_quit = True Then Exit Function
        ParseAlbumInfoSplit = Split(ParseAlbumStrSplit(i), "|")
        '判断部位最后一行
        If i < UBound(ParseAlbumStrSplit) Then
            'list_album_name
            ParseAlbumInfo(i).AlbumName = fix_code(ParseAlbumInfoSplit(3))
            'list_album_password
            ParseAlbumInfo(i).hasPassword = IIf(ParseAlbumInfoSplit(0) <> "0", True, False)
            'list_album_url
            ParseAlbumInfo(i).URL = ParseAlbumInfoSplit(2)
            'list_album_photo_numbers
            If IsNumeric(ParseAlbumInfoSplit(1)) Then
                ParseAlbumInfo(i).picCount = IIf(ParseAlbumInfoSplit(1) > 0, Format$(ParseAlbumInfoSplit(1), "00000") & "张", "")
            Else
                ParseAlbumInfo(i).picCount = ""
            End If
            'list_album_disc
            For j = 4 To UBound(ParseAlbumInfoSplit)
                ParseAlbumInfo(i).description = ParseAlbumInfo(i).description & ParseAlbumInfoSplit(j)
            Next j
            ParseAlbumInfo(i).description = fix_code(Trim$(ParseAlbumInfo(i).description))
        '最后一行为下一页下载信息
        Else
            ParseAlbumInfo(i).URL = Trim$(ParseAlbumStrSplit(i))
        End If
    Next i
    ParseAlbum = ParseAlbumInfo
End Function
