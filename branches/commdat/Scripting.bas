Attribute VB_Name = "Scripting"
Public Function ReadTextFile(file_name) As String
    On Error Resume Next
    
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TristateFalse = 0, TristateTrue = -1, TristateUseDefault = -2
    Dim FSO As Object, textStream As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set textStream = FSO.OpenTextFile(file_name, ForReading, False, TristateFalse)
    ReadTextFile = textStream.ReadAll
    textStream.Close
    Set textStream = Nothing
    Set FSO = Nothing
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
