VERSION 5.00
Begin VB.Form OX_Finish_Download 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OX163下载完成"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
   Icon            =   "OX_Finish_Download.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6255
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3360
      Top             =   120
   End
   Begin VB.PictureBox bg_Picture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1785
      Left            =   120
      ScaleHeight     =   1785
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   120
      Width           =   4500
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   2760
         Top             =   0
      End
      Begin VB.CommandButton Command2 
         Caption         =   "取消"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "打开"
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   525
         Left            =   0
         Picture         =   "OX_Finish_Download.frx":406A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "下载完成,是否打开文件夹?"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   600
         TabIndex        =   1
         Top             =   0
         Width           =   2160
      End
      Begin VB.Image Image2 
         Height          =   525
         Left            =   0
         Picture         =   "OX_Finish_Download.frx":4108
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
   End
End
Attribute VB_Name = "OX_Finish_Download"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Floders As String
Dim time_step As Byte

Private Sub Command1_Click()
    Shell "explorer.exe """ & Floders & """", vbNormalFocus
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = windows_destop_Height - 1140 - 900
    Me.Left = windows_destop_Width - 3090 - 75
    Me.Width = 1
    Me.Height = 1
    Finish_Download_on_top
    time_step = 0
End Sub

Private Sub Timer1_Timer()
    '3090 1140
    time_step = time_step + 1
    Me.Width = time_step * 309
    Me.Height = time_step * 114
    If time_step >= 10 Then Timer1.Enabled = False
    Timer2.Enabled = True
End Sub

Private Sub Finish_Download_on_top()
    Dim flags As Integer
    flags = SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags
End Sub

Private Sub Timer2_Timer()
    Image1.Visible = Not Image1.Visible
End Sub
