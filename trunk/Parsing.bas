Attribute VB_Name = "Parsing"
Option Explicit

Public Type ScriptInfo
    filename As String
    language As String
    encoding As String
    handleType As String
    criteria As String
End Type

Public Type downloadInfo
    isFinal As Boolean
    mode As DownloadMode
    excludeChar() As String
    regularURL As String
    refererURL As String
    method As String
End Type

Public Type AlbumInfo
    hasPassword As Boolean
    picCount As Integer
    URL As String
    dirName As String
    description As String
End Type
Public Function ParseInclude(ByVal sourceString As String, escFormat As EscapeFormat) As ScriptInfo
    Dim expression As New RegExp, results As MatchCollection, result As Match
    expression.Global = True
    expression.IgnoreCase = True
    expression.MultiLine = True
    'tom.vbs|vbscript|GB2312|album|http://photo.tom.com/pim.php?*
    expression.Pattern = "(" & OX_ESCAPED & "\.(?:vbs|js))" & OX_SEPARATOR & "(vbscript|javascript)" & OX_SEPARATOR & _
    "(" & OX_ESCAPED & ")" & OX_SEPARATOR & "(album|photo)" & OX_SEPARATOR & "(" & OX_ESCAPED & ")$"
    Set results = expression.Execute(sourceString)
    Debug.Assert results.count > 0
    
    Set result = results.Item(0)
    ParseInclude.filename = result.SubMatches(0)
    ParseInclude.language = LCase$(result.SubMatches(1))
    ParseInclude.encoding = result.SubMatches(2)
    ParseInclude.handleType = LCase$(result.SubMatches(3))
    ParseInclude.criteria = DeEscape(result.SubMatches(4), escFormat)
End Function

Public Function ParseAlbum(ByVal sourceString As String, escFormat As EscapeFormat) As AlbumInfo()
    Dim expression As New RegExp, results As MatchCollection, result As Match, infos() As AlbumInfo
    expression.Global = True
    expression.IgnoreCase = True
    expression.MultiLine = True
    '0|23|http://comic.92wy.com/go/comicshow.aspx?id=1389&nameid=57|BLAME_µÚ1¾í|BLAME_µÚ1¾í
    expression.Pattern = "(-?\d+)" & OX_SEPARATOR & "(\d+)?" & OX_SEPARATOR & "(" & OX_ESCAPED & ")" & OX_SEPARATOR & _
    "(" & OX_ESCAPED & ")" & OX_SEPARATOR & "(" & OX_ESCAPED & ")$"
    Set results = expression.Execute(sourceString)
    Dim index As Integer
    index = IIf(results.count > 0, 0, -1)
    ReDim infos(index To results.count - 1) As AlbumInfo
    For Each result In results
        infos(index).hasPassword = (CInt(result.SubMatches(0)) <> 0)
        infos(index).picCount = CInt(result.SubMatches(1))
        infos(index).URL = DeEscape(result.SubMatches(2), escFormat)
        infos(index).dirName = DeEscape(result.SubMatches(3), escFormat)
        infos(index).description = DeEscape(result.SubMatches(4), escFormat)
        index = index + 1
    Next
    ParseAlbum = infos
End Function
Public Function ParseDownloadURL(ByVal sourceString As String, escFormat As EscapeFormat) As downloadInfo
    Dim expression As New RegExp, results As MatchCollection, result As Match
    expression.Global = True
    expression.IgnoreCase = True
    expression.MultiLine = False
    'inet|10,13|url|url_Referer|POST method
    expression.Pattern = "(-?\d+)?" & OX_SEPARATOR & "?(inet|web)?" & OX_SEPARATOR & "?((?:-?\d+,)*-?\d+)?" & _
    OX_SEPARATOR & "?(" & OX_ESCAPED & ")?" & OX_SEPARATOR & "?(" & OX_ESCAPED & ")?" & OX_SEPARATOR & "?(POST)?$"
    Set results = expression.Execute(sourceString)
    Debug.Assert results.count > 0
    
    Set result = results.Item(0)
    ParseDownloadURL.isFinal = IIf(result.SubMatches(0) = "", False, CInt(result.SubMatches(0)) = 0)
    If ParseDownloadURL.isFinal Then Exit Function
    ParseDownloadURL.mode = IIf(LCase$(result.SubMatches(1)) = "inet", OX_INET, OX_WEB)
    ParseDownloadURL.excludeChar = Split(result.SubMatches(2), ",")
    ParseDownloadURL.regularURL = Trim$(DeEscape(result.SubMatches(3), escFormat))
    ParseDownloadURL.refererURL = DeEscape(result.SubMatches(4), escFormat)
    ParseDownloadURL.method = result.SubMatches(5)
End Function
