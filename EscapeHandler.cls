VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EscapeHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mEscFormat As EscapeFormat

Public Sub SetFormat(ByVal escChar As String, ByVal escBase As String)
    mEscFormat.escChar = escChar
    mEscFormat.escBase = escBase
End Sub

Public Function Escape(ByVal literal As String) As String
    Escape = Escaping.Escape(literal, mEscFormat)
End Function

Public Function DeEscape(ByVal escaped As String) As String
    DeEscape = Escaping.DeEscape(escaped, mEscFormat)
End Function


