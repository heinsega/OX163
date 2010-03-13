Attribute VB_Name = "Scripting"
Public Function ReadTextFile(file_name) As String
    On Error Resume Next
    
    Dim fso As Object, textStream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set textStream = fso.OpenTextFile(file_name, 1, False, 0)
    ReadTextFile = textStream.ReadAll
    textStream.Close
    Set fso = Nothing
End Function

Public Function LoadScript(ByVal flags As Byte, ParamArray scriptArgs() As Variant) As ScriptControl
    Set LoadScript = New ScriptControl
    Dim scriptCode As String, extension As String
    scriptCode = ""
    
    If UBound(scriptArgs) >= 0 Then
        LoadScript.Language = scriptArgs(0).Language
    Else
        LoadScript.Language = "vbscript"
    End If
    Select Case LoadScript.Language
    Case "vbscript"
        extension = ".vbs"
    Case "javascript"
        extension = ".js"
    Case Else
        Debug.Assert False
    End Select
    
    If flags And OX_INCL_COMMON <> 0 Then
        scriptCode = scriptCode & ReadTextFile(App.Path & "\include\" & OX_PATH_COMMON_INCL & extension) & vbCrLf
    End If
    If flags And OX_INCL_TYPELIB <> 0 Then
        scriptCode = scriptCode & ReadTextFile(App.Path & "\include\" & OX_PATH_TYPELIB_INCL & extension) & vbCrLf
    End If
    If UBound(scriptArgs) >= 0 Then
        Dim index As Integer
        For index = 0 To UBound(scriptArgs)
            scriptCode = scriptCode & ReadTextFile(App.Path & "\include\" & scriptArgs(index).FileName) & vbCrLf
        Next index
    End If
    Call LoadScript.AddCode(scriptCode)
End Function
