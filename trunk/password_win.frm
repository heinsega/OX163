VERSION 5.00
Begin VB.Form password_win 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PassWord Input"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3525
   Icon            =   "password_win.frx":0000
   LinkTopic       =   "password_win"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3525
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "password_win.frx":406A
      Left            =   45
      List            =   "password_win.frx":407D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   50
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   840
      TabIndex        =   2
      Top             =   45
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前密码:"
      Height          =   180
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   810
   End
End
Attribute VB_Name = "password_win"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isDown As Integer
Dim C_quit As Boolean
Dim ESC_quit As Boolean

Private Sub Command1_Click()
On Error Resume Next
Text1.Text = Replace(Replace(Text1.Text, Chr(13), ""), Chr(10), "")
If Len(Text1.Text) > 0 Then
    If Len(Text1.Text) > 30 Then alt_msg = MsgBox("密码过长，超过30位，是否继续？", vbYesNo + vbExclamation + vbDefaultButton2, "警告")
    If alt_msg = vbNo Then Exit Sub
    
    If isDown = 0 Then
    Form1.edit_psw Combo1.ListIndex, Text1.Text
    ElseIf isDown > 0 Then
    Form1.user_list.ListItems(isDown).ListSubItems(1).Text = Text1.Text
    Combo1.Visible = True
    Else
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
ESC_quit = False
Form1.always_on_top False
Combo1.ListIndex = 0
C_quit = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then ESC_quit = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If isDown < 0 And C_quit = False Then Form1.url_input.Text = ""
If Form1.WindowState = 0 Then Form1.always_on_top sysSet.always_top

If Form1.pass_code = "163" And isDown = 0 And ESC_quit = False Then
Call Form1.check_pass_code(False, isDown)
ElseIf Form1.pass_code = "new163_pass" And isDown = 0 And ESC_quit = False Then
Call Form1.new163_check_passcode(False, isDown)
Else
Form1.Enabled = True
End If

End Sub

Private Sub Text1_Change()
Label2.Caption = Text1.Text
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
ESC_quit = True
Unload Me
End If
End Sub
