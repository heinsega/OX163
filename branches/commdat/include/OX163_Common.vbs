Private Const REGEXP_FILENAME = "[^/\\:\*\?""<>\|]+"

Dim OX163_urlpage_Referer, OX163_urlpage_Cookies

Function set_cookies(ByVal set_str)
	OX163_urlpage_Cookies = set_str
End Function

Function return_script_data(ByVal script_str)
	Dim expression, results, result
	Set expression = New RegExp
	expression.Global = True
	expression.IgnoreCase = True
	expression.MultiLine = True
	'tom.vbs|vbscript|GB2312|album|http://photo.tom.com/pim.php?*
	expression.Pattern = "(" & REGEXP_FILENAME & "\.(?:vbs|js))\|(vbscript|javascript)\|([^\|]+)\|(album|photo)\|(\S+)\s*$"
	Set results = expression.Execute(script_str)
	
	Set result = results.Item(0)
	Dim filename_str, language_str, encoding_str, type_str, criteria_str
	filename_str = result.SubMatches(0)
	language_str = LCase(result.SubMatches(1))
	encoding_str = result.SubMatches(2)
	type_str = LCase(result.SubMatches(3))
	criteria_str = result.SubMatches(4)
	Set return_script_data = Script(filename_str, language_str, encoding_str, type_str, criteria_str)
End Function