VERSION 5.00
Begin VB.Form password_win 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PassWord Input"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3285
   Icon            =   "password_win.frx":0000
   LinkTopic       =   "password_win"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3285
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "password_win.frx":406A
      Left            =   120
      List            =   "password_win.frx":407D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   800
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   120
      MaxLength       =   30
      TabIndex        =   1
      Top             =   360
      Width           =   3060
   End
   Begin VB.Label password_win_title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PassWord:"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "password_win"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isDown As Integer '是否下载
Dim C_quit As Boolean

Private Sub Command1_Click()
    On Error Resume Next
    Text1.Text = Replace(Replace(Text1.Text, Chr(13), ""), Chr(10), "")
    If Len(Text1.Text) > 0 Then
        If Len(Text1.Text) > 30 Then alt_msg = MsgBox("密码过长，超过30位，是否继续？", vbYesNo + vbExclamation + vbDefaultButton2, "警告")
        If alt_msg = vbNo Then Exit Sub
        
        If isDown = 0 Then
            Form1.edit_psw Combo1.ListIndex, Text1.Text
        ElseIf isDown > 0 Then
            Form1.user_list.ListItems(isDown).ListSubItems(2).Text = Text1.Text
            Combo1.Visible = True
        Else
            '登陆相册用户
            Form1.url_input.Text = Text1.Text
            Combo1.Visible = True
        End If
        C_quit = True
        Unload Me
    Else
        MsgBox "密码不能为空！", vbOKOnly + vbExclamation, "警告"
    End If
End Sub

Private Sub Form_Load()
    Form1.always_on_top False
    Combo1.ListIndex = 0
    C_quit = False
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'UnloadMode = 0 click ESC button
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    password_win_title.caption = "PassWord:"
    If isDown < 0 And C_quit = False Then Form1.url_input.Text = ""
    If Form1.WindowState = 0 Then Form1.always_on_top sysSet.always_top
    Form1.Enabled = True
End Sub

Private Sub Text1_DblClick()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 65 And Shift = vbCtrlMask Then
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1_Click
    ElseIf KeyAscii = 27 Then
        Unload Me
    End If
End Sub
