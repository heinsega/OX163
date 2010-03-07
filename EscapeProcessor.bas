Attribute VB_Name = "EscapeProcessor"
Public Type EscapeFormat
    escChar As String
    escBase As String
End Type

Public Function Escape(ByVal literal As String, ByVal escChar As String, ByVal escBase As String) As String
    Dim expression As New regExp
    expression.Global = True
    expression.Pattern = "([" & escBase & escChar & "])"
    Escape = expression.Replace(literal, escChar & "$1")
End Function

Public Function DeEscape(ByVal escaped As String, ByVal escChar As String, ByVal escBase As String) As String
    Dim expression As New regExp
    expression.Global = True
    expression.Pattern = escChar & "([" & escBase & escChar & "])"
    DeEscape = expression.Replace(escaped, "$1")
End Function
