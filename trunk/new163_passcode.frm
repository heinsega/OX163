VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form new163_passcode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "填写验证码"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   Icon            =   "new163_passcode.frx":0000
   LinkTopic       =   "new163_passcode"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3960
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2640
      Top             =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "刷新"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   2775
      TabIndex        =   2
      Top             =   480
      Width           =   2775
      Begin SHDocVwCtl.WebBrowser WebBrowser 
         Height          =   2055
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3015
         ExtentX         =   5318
         ExtentY         =   3625
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
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
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "new163_passcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isDown As Integer

Private Sub Command1_Click()
    On Error Resume Next
    Text1.Text = Replace(Replace(Text1.Text, Chr(13), ""), Chr(10), "")
    If Text1.Text = "163" Then MsgBox "验证码不正确！", vbOKOnly, "警告": Exit Sub
    If Len(Text1.Text) > 0 Then
        If Len(Text1.Text) > 5 Then alt_msg = MsgBox("验证码不正确，仍然发送？", vbYesNo + vbExclamation + vbDefaultButton2, "警告")
        If alt_msg = vbNo Then Exit Sub
        
        Form1.pass_code = Trim(Text1.Text)
        
        Unload Me
    Else
        MsgBox "验证码不能为空！", vbOKOnly + vbExclamation, "警告"
    End If
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Me.Enabled = False
    WebBrowser.Refresh
    Do While WebBrowser.Busy = True
        DoEvents
    Loop
    Me.Enabled = True
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    Me.Enabled = False
    WebBrowser.Silent = True
    Form1.always_on_top False
    show_pass_code
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Form1.WindowState = 0 Then Form1.always_on_top sysSet.always_top
    If isDown = 0 Then alt_msg = MsgBox("是否测试验证码的正确性？", vbYesNo + vbExclamation, "询问")
    If alt_msg = vbYes Then
        Call Form1.new163_check_passcode(True, isDown)
    End If
    Form1.Enabled = True
End Sub

Public Sub show_pass_code()
    On Error Resume Next
    Dim url_links As String
    url_links = "http://photo.163.com/photo/cap/captcha.jpgx?parentId=" & Int(Time() * 100000000) & "&t=" & Int(Time() * 10000000000#)
    WebBrowser.Navigate url_links
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

Private Sub Timer1_Timer()
    On Error Resume Next
    Timer1.Interval = 0
    Do While WebBrowser.Busy = True
        DoEvents
    Loop
    Command2_Click
End Sub
