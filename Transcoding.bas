Attribute VB_Name = "Transcoding"
Option Explicit

Public Enum CompareMode
    OX_BINARY
    OX_TEXT
End Enum

Public Enum TextFormat
    OX_PLAIN
    OX_INTERNAL
    OX_VBSCRIPT
    OX_JAVASCRIPT
End Enum

'返回内部字符串
Public Function CInternal(ByVal sourceString As String, Optional ByVal sourceType As TextFormat) As String
    Select Case sourceType
    Case OX_PLAIN
        sourceString = Replace$(sourceString, """", """""")
    Case OX_INTERNAL
    Case OX_VBSCRIPT
        sourceString = Replace$(sourceString, """", """""")
        sourceString = Replace$(sourceString, vbCr, """ & vbCr & """)
        sourceString = Replace$(sourceString, vbLf, """ & vbLf & """)
    Case OX_JAVASCRIPT
        sourceString = Replace$(sourceString, """", "\""")
        sourceString = Replace$(sourceString, vbLf, """ + String.fromCharCode(10) + """)
        sourceString = Replace$(sourceString, vbCr, """ + String.fromCharCode(13) + """)
    Case Else
        Debug.Assert False
    End Select
    CInternal = """" & sourceString & """"
End Function

'过滤指定关键字集
Public Function Filter(ByVal sourceString As String, ByRef keywords() As String, Optional ByVal mode As CompareMode = OX_BINARY) As String
    Dim keyword As Variant
    For Each keyword In keywords
        Select Case mode
        Case OX_BINARY
            sourceString = Replace$(sourceString, Chr(CLng(keyword)), "")
        Case OX_TEXT
            sourceString = Replace$(sourceString, keyword, "")
        Case Else
            Debug.Assert False
        End Select
    Next
    Filter = sourceString
End Function

