VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "URLData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum DownloadMode
    OX_INET
    OX_WEB
End Enum

Public isFinal As Boolean
Public dlmMode As DownloadMode
Private arrExcludeChars() As String
Public regularURL As String
Public refererURL As String
Public method As String

Public Property Let mode(ByVal modeString As String)
    Select Case LCase$(modeString)
    Case "inet"
        dlmMode = OX_INET
    Case "web"
        dlmMode = OX_WEB
    Case ""
        dlmMode = Empty
    Case Else
        Debug.Assert False
    End Select
End Property

Public Property Let excludeChars(ByVal delimitedString As String)
    arrExcludeChars = Split(delimitedString, ",")
End Property

Public Property Get excludeCharsArray() As String()
    excludeCharsArray = arrExcludeChars
End Property
