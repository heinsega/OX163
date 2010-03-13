Private Const OX_ENTRY_URL = 0, OX_ENTRY_ALBUM = 1, OX_ENTRY_PICT = 2, OX_ENTRY_SCRIPT = -1

Private datBundle
Set datBundle = CreateObject("OX163TypeLib.DataBundle")

Private Function Script(ByVal filename, ByVal language, ByVal encoding, ByVal handleType, ByVal criteria)
	Set Script = CreateObject("OX163TypeLib.ScriptData")
	Script.filename = filename
	Script.language = language
	Script.encoding = encoding
	Script.handleType = handleType
	Script.criteria = criteria
End Function

Private Function Page(ByVal isFinal, ByVal mode, ByVal excludeChars, ByVal regularURL, ByVal refererURL, ByVal method)
	Set Page = CreateObject("OX163TypeLib.URLData")
	Page.isFinal = isFinal
	Page.mode = mode
	Page.excludeChars = excludeChars
	Page.regularURL = regularURL
	Page.refererURL = refererURL
	Page.method = method
End Function

Private Function Album(ByVal hasPassword, ByVal picCount, ByVal URL, ByVal dirName, ByVal description)
	Set Album = CreateObject("OX163TypeLib.AlbumData")
	Album.hasPassword = hasPassword
	Album.picCount = picCount
	Album.URL = URL
	Album.dirName = dirName
	Album.description = description
End Function

Private Sub ResetBundle()
	Set datBundle = Nothing
	Set datBundle = CreateObject("OX163TypeLib.DataBundle")
End Sub

Private Sub Entry(ByVal entryVal, ByVal entryType)
	Select Case entryType
		Case OX_ENTRY_URL
			Call datBundle.urlEntries.Add(entryVal)
		Case OX_ENTRY_ALBUM
			Call datBundle.albumEntries.Add(entryVal)
		Case OX_ENTRY_PICT
			Call datBundle.pictEntries.Add(entryVal)
		Case Else
			Debug.Assert False
	End Select
End Sub

Function GetBundle()
	Set GetBundle = datBundle
	Call ResetBundle
End Function