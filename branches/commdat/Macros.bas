Attribute VB_Name = "Macros"
Option Explicit

'Include Flags
Public Const OX_INCL_NONE As Byte = &H0
Public Const OX_INCL_COMMON As Byte = &H1
Public Const OX_INCL_TYPELIB As Byte = &H2
Public Const OX_INCL_ALL As Byte = OX_INCL_COMMON Or OX_INCL_TYPELIB

'External Paths
Public Const OX_PATH_TYPELIB As String = "OX163_TypeLib"
Public Const OX_PATH_COMMON_INCL As String = "OX163_Common", OX_PATH_TYPELIB_INCL As String = "OX163_TypeLibInterface"

'RegExp Macros
Public Const OX_SEPARATOR As String = "\|"

Public Const OX_ESC_CHAR As String = "\\"
Private Const OX_RESERVED_SET As String = OX_SEPARATOR & OX_ESC_CHAR & "\f\n\r\t\v"
Public Const OX_RESERVED As String = "[" & OX_RESERVED_SET & "]", OX_PLAIN As String = "[^" & OX_RESERVED_SET & "]"
Public Const OX_ESC_SEQUENCE As String = OX_ESC_CHAR & OX_RESERVED
Public Const OX_ESCAPED As String = "(?:" & OX_PLAIN & "|" & OX_ESC_SEQUENCE & ")*"
