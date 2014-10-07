VERSION 5.00
Begin VB.Form passcode_win 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login 163"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3345
   Icon            =   "passcode_win.frx":0000
   LinkTopic       =   "passcode_win"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3345
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   570
      Left            =   2040
      ScaleHeight     =   570
      ScaleWidth      =   1335
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1335
      Begin VB.CommandButton Command1 
         Caption         =   "确定"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label passcode_Name_password 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请填写163密码:"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1260
   End
   Begin VB.Label passcode_Name 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请填写163帐号:"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1260
   End
End
Attribute VB_Name = "passcode_win"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public passport163name As String

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
        'Call Form1.check_pass_code(True, isDown)
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
