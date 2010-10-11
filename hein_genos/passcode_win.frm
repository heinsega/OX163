VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form passcode_win 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "填写验证码"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3405
   Icon            =   "passcode_win.frx":0000
   LinkTopic       =   "passcode_win"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3405
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   855
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   810
      Left            =   120
      ScaleHeight     =   810
      ScaleWidth      =   2175
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   2175
      Begin SHDocVwCtl.WebBrowser WebBrowser 
         Height          =   2775
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3855
         ExtentX         =   6800
         ExtentY         =   4895
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
End
Attribute VB_Name = "passcode_win"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isDown As Integer
Public html_str As String

Private Sub Command1_Click()
On Error Resume Next
Text1.Text = Replace(Replace(Text1.Text, Chr(13), ""), Chr(10), "")
If Len(Text1.Text) > 0 Then
    If Len(Text1.Text) > 6 Then alt_msg = MsgBox("验证码不正确，仍然发送？", vbYesNo + vbExclamation + vbDefaultButton2, "警告")
    If alt_msg = vbNo Then Exit Sub
    
    Form1.pass_code = "&encrypted_code=" & URLEncode(html_str) & "&code=" & Trim(Text1.Text)

    Unload Me
Else
    MsgBox "验证码不能为空！", vbOKOnly + vbExclamation, "警告"
End If
End Sub

Private Sub Form_Load()
Form1.always_on_top False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Form1.WindowState = 0 Then Form1.always_on_top sysSet.always_top
If isDown = 0 Then alt_msg = MsgBox("是否测试验证码的正确性？", vbYesNo + vbExclamation, "询问")
If alt_msg = vbYes Then
Call Form1.check_pass_code(True, isDown)
End If
Form1.Enabled = True
End Sub

Public Sub show_pass_code()
On Error Resume Next
If InStr(html_str, "/captcha.php?code=") > 0 Then
Dim url_links As String
url_links = Mid$(html_str, InStr(html_str, "/captcha.php?code="))
url_links = "http://photo.163.com" & Mid$(url_links, 1, InStr(url_links, Chr(34)) - 1)
WebBrowser.Navigate url_links
html_str = Mid$(html_str, InStr(html_str, "input name=" & Chr(34) & "encrypted_code" & Chr(34) & " type="))
html_str = Mid$(html_str, InStr(html_str, "value=") + 7)
html_str = Mid$(html_str, 1, InStr(html_str, Chr(34)) - 1)
Else

MsgBox "可能已经通过验证", vbOKOnly, "提醒"
Form1.pass_code = "&encrypted_code=&code="
Unload Me

End If
End Sub

Private Sub Text1_DblClick()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
ElseIf KeyAscii = 27 Then
Unload Me
End If
End Sub
