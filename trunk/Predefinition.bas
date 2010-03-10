Attribute VB_Name = "Predefinition"
Option Explicit

Public Enum DownloadMode
     OX_INET
     OX_WEB
End Enum

'Reserved Character RegExp Predefs
Public Const OX_SEPARATOR As String = "\|"
'Escape Sequence RegExp Predefs
Public Const OX_ESC_CHAR As String = "\\"
Private Const OX_RESERVED_SET As String = OX_SEPARATOR & OX_ESC_CHAR
Public Const OX_RESERVED As String = "[" & OX_RESERVED_SET & "]", OX_PLAIN As String = "[^" & OX_RESERVED_SET & "]"
Public Const OX_ESC_SEQUENCE As String = OX_ESC_CHAR & OX_RESERVED
Public Const OX_ESCAPED As String = "(?:" & OX_PLAIN & "|" & OX_ESC_SEQUENCE & ")*"
