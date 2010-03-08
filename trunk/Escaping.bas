Attribute VB_Name = "Escaping"
Public Type EscapeFormat
    escChar As String
    escBase As String
End Type

Public Function Escape(ByVal literal As String, ByRef escFormat As EscapeFormat) As String
    Escape = Translate(literal, "(" & escFormat.escBase & ")", escFormat.escChar & "$1")
End Function

Public Function DeEscape(ByVal escaped As String, ByRef escFormat As EscapeFormat) As String
    DeEscape = Translate(escaped, escFormat.escChar & "(" & escFormat.escBase & ")", "$1")
End Function

Private Function Translate(ByVal literal As String, ByVal source As String, ByVal target As String) As String
    Dim expression As New RegExp
    expression.Global = True
    expression.Pattern = source
    Translate = expression.Replace(literal, target)
End Function
