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
Public mode As DownloadMode
Private m_excludeChars() As String
Public regularURL As String
Public refererURL As String
Public method As String

Public Property Get excludeChar(ByVal index As Integer) As String
    excludeChar = m_excludeChars(index)
End Property

Public Property Let excludeChar(ByVal index As Integer, ByVal excludeChar As String)
    m_excludeChars(index) = excludeChar
End Property

Public Property Get excludeChars() As String()
    excludeChars = m_excludeChars
End Property
