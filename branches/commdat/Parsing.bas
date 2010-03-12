Attribute VB_Name = "Parsing"
Public Function ParseInclude(ByVal sourceString As String, escFormat As EscapeFormat) As ScriptData
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
    ParseInclude.FileName = result.SubMatches(0)
    ParseInclude.Language = LCase$(result.SubMatches(1))
    ParseInclude.encoding = result.SubMatches(2)
    ParseInclude.handleType = LCase$(result.SubMatches(3))
    ParseInclude.criteria = DeEscape(result.SubMatches(4), escFormat)
End Function

Public Function ParseAlbum(ByVal sourceString As String, escFormat As EscapeFormat) As AlbumData()
    Dim expression As New RegExp, results As MatchCollection, result As Match, infos() As AlbumData
    expression.Global = True
    expression.IgnoreCase = True
    expression.MultiLine = True
    '0|23|http://comic.92wy.com/go/comicshow.aspx?id=1389&nameid=57|BLAME_µÚ1¾í|BLAME_µÚ1¾í
    expression.Pattern = "(-?\d+)" & OX_SEPARATOR & "(\d+)?" & OX_SEPARATOR & "(" & OX_ESCAPED & ")" & OX_SEPARATOR & _
    "(" & OX_ESCAPED & ")" & OX_SEPARATOR & "(" & OX_ESCAPED & ")$"
    Set results = expression.Execute(sourceString)
    
    Dim Index As Integer
    Index = IIf(results.count > 0, 0, -1)
    ReDim infos(Index To results.count - 1) As AlbumData
    For Each result In results
     Debug.Print result.Value
        infos(Index).hasPassword = (CInt(result.SubMatches(0)) <> 0)
        infos(Index).picCount = CInt(result.SubMatches(1))
        infos(Index).URL = DeEscape(result.SubMatches(2), escFormat)
        infos(Index).dirName = DeEscape(result.SubMatches(3), escFormat)
        infos(Index).Description = DeEscape(result.SubMatches(4), escFormat)
        Index = Index + 1
    Next
    ParseAlbum = infos
End Function
Public Function ParseDownloadURL(ByVal sourceString As String, escFormat As EscapeFormat) As URLData
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
    'ParseDownloadURL.excludeChar = Split(result.SubMatches(2), ",")
    ParseDownloadURL.regularURL = Trim$(DeEscape(result.SubMatches(3), escFormat))
    ParseDownloadURL.refererURL = DeEscape(result.SubMatches(4), escFormat)
    ParseDownloadURL.method = result.SubMatches(5)
End Function
