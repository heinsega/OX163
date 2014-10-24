VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form sys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OX163程序设置"
   ClientHeight    =   14565
   ClientLeft      =   270
   ClientTop       =   720
   ClientWidth     =   22080
   Icon            =   "sys.frx":0000
   LinkTopic       =   "sys"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14565
   ScaleWidth      =   22080
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton frame_rec 
      Caption         =   "调用INI恢复本栏设置"
      Height          =   495
      Left            =   4800
      TabIndex        =   70
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton frame_def 
      Caption         =   "恢复本栏默认设置"
      Height          =   495
      Left            =   6960
      TabIndex        =   71
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Frame FrameL 
      Caption         =   "网易相册设置"
      ForeColor       =   &H00C00000&
      Height          =   5295
      Index           =   7
      Left            =   2400
      TabIndex        =   11
      Top             =   10200
      Visible         =   0   'False
      Width           =   6375
      Begin VB.PictureBox Picture23 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   480
         ScaleHeight     =   615
         ScaleWidth      =   5595
         TabIndex        =   64
         Top             =   600
         Width           =   5595
         Begin VB.OptionButton new163passrule 
            Caption         =   "否(我有老相册用到中文密码)"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   66
            Top             =   240
            Width           =   3135
         End
         Begin VB.OptionButton new163passrule 
            Caption         =   "是(使用博客相册合并后的新密码规则)"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   65
            Top             =   0
            Width           =   3615
         End
      End
      Begin VB.TextBox Text1 
         Height          =   1695
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   21
         Text            =   "sys.frx":406A
         Top             =   2760
         Width           =   5175
      End
      Begin VB.TextBox passcode_text 
         Height          =   270
         Index           =   2
         Left            =   960
         TabIndex        =   15
         Text            =   "asd"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox passcode_text 
         Height          =   270
         Index           =   1
         Left            =   960
         TabIndex        =   14
         Text            =   "1530930"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox passcode_text 
         Height          =   270
         Index           =   0
         Left            =   960
         TabIndex        =   13
         Text            =   "wehi"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.PictureBox Picture17 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   3720
         ScaleHeight     =   1095
         ScaleWidth      =   2535
         TabIndex        =   19
         Top             =   1440
         Width           =   2535
         Begin VB.CommandButton Auto_Password_com 
            Caption         =   "自动填写"
            Height          =   975
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.Label FrameL7_1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "是否修正163相册中文密码问题:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   67
         Top             =   360
         Width           =   2520
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "以上内容用于测试验证码(新版相册可省略)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   3
         Left            =   960
         TabIndex        =   18
         Top             =   2520
         Width           =   3420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密  码:"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   2205
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "相册ID:"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   1845
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户名:"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1485
         Width           =   630
      End
   End
   Begin VB.Frame FrameL 
      Caption         =   "内置浏览器设置"
      ForeColor       =   &H00C00000&
      Height          =   5295
      Index           =   8
      Left            =   9000
      TabIndex        =   103
      Top             =   9000
      Width           =   6375
      Begin VB.PictureBox Picture29 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   360
         ScaleHeight     =   375
         ScaleWidth      =   1500
         TabIndex        =   137
         Top             =   2400
         Width           =   1500
         Begin VB.CommandButton Comm_edit_black 
            Caption         =   "编辑黑名单"
            Height          =   300
            Left            =   0
            TabIndex        =   138
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture28 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3120
         ScaleHeight     =   375
         ScaleWidth      =   1500
         TabIndex        =   135
         Top             =   2400
         Width           =   1500
         Begin VB.CommandButton Comm_edit_white 
            Caption         =   "编辑白名单"
            Height          =   300
            Left            =   0
            TabIndex        =   136
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture27 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   1395
         TabIndex        =   132
         Top             =   2040
         Width           =   1395
         Begin VB.OptionButton ie_black_list 
            Caption         =   "是"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   134
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton ie_black_list 
            Caption         =   "否"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   133
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture26 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3120
         ScaleHeight     =   255
         ScaleWidth      =   1395
         TabIndex        =   129
         Top             =   2040
         Width           =   1395
         Begin VB.OptionButton ie_white_list 
            Caption         =   "否"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   131
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton ie_white_list 
            Caption         =   "是"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   130
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture25 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3120
         ScaleHeight     =   255
         ScaleWidth      =   1395
         TabIndex        =   124
         Top             =   1320
         Width           =   1395
         Begin VB.OptionButton ie_local_window 
            Caption         =   "否"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   126
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton ie_local_window 
            Caption         =   "是"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   125
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture22 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   1395
         TabIndex        =   107
         Top             =   600
         Width           =   1395
         Begin VB.OptionButton ox163_window 
            Caption         =   "否"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   109
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton ox163_window 
            Caption         =   "是"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   108
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture10 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   1395
         TabIndex        =   104
         Top             =   1320
         Width           =   1395
         Begin VB.OptionButton ie_window 
            Caption         =   "是"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   106
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton ie_window 
            Caption         =   "否"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   105
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "启用黑名单？"
         Height          =   180
         Index           =   19
         Left            =   360
         TabIndex        =   128
         ToolTipText     =   "非特定需求建议选择(是)"
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "启用白名单？"
         Height          =   180
         Index           =   18
         Left            =   3120
         TabIndex        =   127
         ToolTipText     =   "非特定需求建议选择(是)"
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "不阻止相同域新开窗口？"
         Height          =   180
         Index           =   17
         Left            =   3120
         TabIndex        =   123
         ToolTipText     =   "非特定需求建议选择(是)"
         Top             =   1080
         Width           =   1980
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "是否用OX163打开新窗口？"
         Height          =   180
         Index           =   14
         Left            =   360
         TabIndex        =   111
         ToolTipText     =   "浏览特定网站请建议选择(是)"
         Top             =   360
         Width           =   2070
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "是否阻止浏览器弹出新开窗口？"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   110
         ToolTipText     =   "非特定需求建议选择(是)"
         Top             =   1080
         Width           =   2520
      End
   End
   Begin VB.CommandButton sys_apply 
      Caption         =   "应用"
      Height          =   465
      Left            =   7680
      TabIndex        =   78
      Top             =   5520
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":415F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":41CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":424B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":42CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":434A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":43C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":443E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":44BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":453C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":45B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":4624
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":46A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":4722
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":479F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":4816
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":4893
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":4913
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sys.frx":4991
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView SysTreeView 
      Height          =   5295
      Left            =   120
      TabIndex        =   77
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   9340
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      Style           =   7
      Scroll          =   0   'False
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FrameL 
      Caption         =   "系统检测结果"
      ForeColor       =   &H00C00000&
      Height          =   5295
      Index           =   9
      Left            =   15480
      TabIndex        =   63
      Top             =   9000
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox OX_Start_log_Text 
         Height          =   3375
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   150
         Text            =   "sys.frx":4A09
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.Frame FrameL 
      Caption         =   "热键与警告框"
      ForeColor       =   &H00C00000&
      Height          =   5295
      Index           =   6
      Left            =   15480
      TabIndex        =   42
      Top             =   4680
      Visible         =   0   'False
      Width           =   6375
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3600
         ScaleHeight     =   255
         ScaleWidth      =   1395
         TabIndex        =   112
         Top             =   2040
         Width           =   1395
         Begin VB.OptionButton quitOp 
            Caption         =   "是"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   114
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton quitOp 
            Caption         =   "否"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   113
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2475
         TabIndex        =   57
         Top             =   2040
         Width           =   2475
         Begin VB.OptionButton saveOp 
            Caption         =   "询问"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton saveOp 
            Caption         =   "直接保存"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   58
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1395
         TabIndex        =   54
         Top             =   2760
         Width           =   1395
         Begin VB.OptionButton changepsw 
            Caption         =   "否"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   56
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton changepsw 
            Caption         =   "是"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture9 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3600
         ScaleHeight     =   255
         ScaleWidth      =   1395
         TabIndex        =   51
         Top             =   2760
         Width           =   1395
         Begin VB.OptionButton askfloder 
            Caption         =   "是"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton askfloder 
            Caption         =   "否"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   52
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture14 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2475
         TabIndex        =   46
         Top             =   1200
         Width           =   2475
         Begin VB.OptionButton ubb_copy 
            Caption         =   "Alt+C"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   48
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton ubb_copy 
            Caption         =   "Ctrl+C"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture13 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2475
         TabIndex        =   43
         Top             =   480
         Width           =   2475
         Begin VB.OptionButton list_copy 
            Caption         =   "Ctrl+C"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton list_copy 
            Caption         =   "Alt+C"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   44
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "程序执行时，打开退出询问？"
         Height          =   180
         Left            =   3600
         TabIndex        =   115
         ToolTipText     =   "建议选择(是)防止误操作"
         Top             =   1800
         Width           =   2340
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "目录错误时，保存到download目录？"
         Height          =   180
         Left            =   240
         TabIndex        =   62
         ToolTipText     =   "建议选择(是)防止无意保存"
         Top             =   1800
         Width           =   2880
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密码错误时，询问是否重填密码？"
         Height          =   180
         Left            =   240
         TabIndex        =   61
         ToolTipText     =   "非特定需求建议选择(是)"
         Top             =   2520
         Width           =   2700
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下载完成后，是否出现提示框？"
         Height          =   180
         Left            =   3600
         TabIndex        =   60
         ToolTipText     =   "非特定需求建议选择(是)"
         Top             =   2520
         Width           =   2520
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "复制UBB标签:[url], 到剪贴板"
         Height          =   180
         Index           =   9
         Left            =   240
         TabIndex        =   50
         ToolTipText     =   "如果出现错误或者程序假死请选择(否)"
         Top             =   960
         Width           =   2430
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "复制链接(url && Lst)到剪贴板"
         Height          =   180
         Index           =   8
         Left            =   240
         TabIndex        =   49
         ToolTipText     =   "如果出现错误或者程序假死请选择(否)"
         Top             =   240
         Width           =   2430
      End
   End
   Begin VB.Frame FrameL 
      Caption         =   "代理服务器设置"
      ForeColor       =   &H00C00000&
      Height          =   5295
      Index           =   4
      Left            =   2400
      TabIndex        =   22
      Top             =   6000
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox web_proxy_box 
         Caption         =   "对内置浏览器启用代理A (还原IE代理模式后可能需要重启软件)"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   151
         Top             =   3480
         Value           =   1  'Checked
         Width           =   6015
      End
      Begin VB.ComboBox ProxyComb 
         Height          =   300
         Index           =   1
         ItemData        =   "sys.frx":4AE3
         Left            =   120
         List            =   "sys.frx":4AF0
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1440
         Width           =   3135
      End
      Begin VB.ComboBox ProxyComb 
         Height          =   300
         Index           =   0
         ItemData        =   "sys.frx":4B2F
         Left            =   120
         List            =   "sys.frx":4B3C
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox proxy_txt2 
         Height          =   270
         Index           =   2
         Left            =   4080
         TabIndex        =   34
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox proxy_txt2 
         Height          =   270
         Index           =   1
         Left            =   4080
         TabIndex        =   33
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox proxy_txt1 
         Height          =   270
         Index           =   2
         Left            =   4080
         TabIndex        =   30
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox proxy_txt1 
         Height          =   270
         Index           =   1
         Left            =   4080
         TabIndex        =   29
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox proxy_txt1 
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   3135
      End
      Begin VB.PictureBox proxy_pic 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3000
         ScaleHeight     =   375
         ScaleWidth      =   3135
         TabIndex        =   24
         Top             =   2880
         Width           =   3135
         Begin VB.CommandButton proxy_com3 
            Caption         =   "复制B到A"
            Height          =   375
            Left            =   2040
            TabIndex        =   39
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton proxy_com2 
            Caption         =   "复制A到B"
            Height          =   375
            Left            =   840
            TabIndex        =   38
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton proxy_com1 
            Caption         =   "清空"
            Height          =   375
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.TextBox proxy_txt2 
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "密  码:"
         Height          =   180
         Index           =   7
         Left            =   3360
         TabIndex        =   36
         Top             =   1860
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "用户名:"
         Height          =   180
         Index           =   6
         Left            =   3360
         TabIndex        =   35
         Top             =   1485
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "密  码:"
         Height          =   180
         Index           =   4
         Left            =   3360
         TabIndex        =   32
         Top             =   900
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "用户名:"
         Height          =   180
         Index           =   3
         Left            =   3360
         TabIndex        =   31
         Top             =   525
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   $"sys.frx":4B7B
         ForeColor       =   &H000000FF&
         Height          =   540
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   4590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "代理设置B: 下载图片内容"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   2070
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "代理设置A: 下载页面内容"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2070
      End
   End
   Begin VB.Frame FrameL 
      Caption         =   "网络下载设置"
      ForeColor       =   &H00C00000&
      Height          =   5295
      Index           =   1
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   6375
      Begin VB.PictureBox FrameL1_bg 
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   120
         ScaleHeight     =   4335
         ScaleWidth      =   6135
         TabIndex        =   172
         Top             =   240
         Width           =   6135
         Begin VB.PictureBox FrameL1_bgs 
            BorderStyle     =   0  'None
            Height          =   6735
            Left            =   0
            ScaleHeight     =   6735
            ScaleWidth      =   5865
            TabIndex        =   174
            Top             =   0
            Width           =   5865
            Begin VB.HScrollBar VS_retry 
               Height          =   225
               Left            =   840
               Max             =   20
               TabIndex        =   206
               Top             =   1650
               Value           =   20
               Width           =   2055
            End
            Begin VB.HScrollBar VS_timeout 
               Height          =   225
               Left            =   840
               Max             =   120
               Min             =   10
               TabIndex        =   205
               Top             =   1170
               Value           =   120
               Width           =   2055
            End
            Begin VB.CommandButton LST_Help 
               Caption         =   "下载列表文件使用说明"
               Height          =   300
               Left            =   2520
               TabIndex        =   204
               Top             =   2520
               Width           =   2295
            End
            Begin VB.ComboBox Combo_lst 
               ForeColor       =   &H000000FF&
               Height          =   300
               ItemData        =   "sys.frx":4C14
               Left            =   0
               List            =   "sys.frx":4C21
               Style           =   2  'Dropdown List
               TabIndex        =   196
               Top             =   2520
               Width           =   2415
            End
            Begin VB.PictureBox FrameL1_Picture 
               BorderStyle     =   0  'None
               Height          =   735
               Index           =   1
               Left            =   0
               ScaleHeight     =   735
               ScaleWidth      =   4875
               TabIndex        =   189
               Top             =   240
               Width           =   4875
               Begin VB.HScrollBar downHS 
                  Height          =   225
                  Left            =   840
                  Max             =   1000
                  Min             =   1
                  TabIndex        =   195
                  Top             =   375
                  Value           =   4
                  Width           =   2055
               End
               Begin VB.TextBox downText 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   225
                  Left            =   3000
                  TabIndex        =   194
                  Text            =   "1KB"
                  Top             =   405
                  Width           =   735
               End
               Begin VB.OptionButton downOp 
                  Caption         =   "自定义"
                  Height          =   255
                  Index           =   3
                  Left            =   0
                  TabIndex        =   193
                  Top             =   360
                  Width           =   855
               End
               Begin VB.OptionButton downOp 
                  Caption         =   "100 KB"
                  Height          =   255
                  Index           =   2
                  Left            =   2160
                  TabIndex        =   192
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton downOp 
                  Caption         =   "50 KB"
                  Height          =   255
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   191
                  Top             =   0
                  Width           =   975
               End
               Begin VB.OptionButton downOp 
                  Caption         =   "10 KB"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   190
                  Top             =   0
                  Width           =   975
               End
            End
            Begin VB.Frame FrameL1_Frame1 
               Height          =   2925
               Left            =   0
               TabIndex        =   175
               Top             =   3720
               Width           =   5775
               Begin VB.PictureBox FrameL1_Picture 
                  BorderStyle     =   0  'None
                  Height          =   675
                  Index           =   2
                  Left            =   120
                  ScaleHeight     =   675
                  ScaleWidth      =   5535
                  TabIndex        =   184
                  Top             =   2160
                  Width           =   5535
                  Begin VB.TextBox Customize_UA_txt 
                     BeginProperty Font 
                        Name            =   "宋体"
                        Size            =   10.5
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   0
                     TabIndex        =   186
                     Top             =   360
                     Width           =   5535
                  End
                  Begin VB.ComboBox Combo_UA_List 
                     Height          =   300
                     ItemData        =   "sys.frx":4C62
                     Left            =   1920
                     List            =   "sys.frx":4C84
                     Style           =   2  'Dropdown List
                     TabIndex        =   185
                     Top             =   0
                     Width           =   3615
                  End
                  Begin VB.Label FrameL1_lab 
                     AutoSize        =   -1  'True
                     Caption         =   "用户代理(User-Agent):"
                     ForeColor       =   &H00C00000&
                     Height          =   180
                     Index           =   7
                     Left            =   0
                     MousePointer    =   14  'Arrow and Question
                     TabIndex        =   187
                     ToolTipText     =   $"sys.frx":4DB8
                     Top             =   45
                     Width           =   1890
                  End
               End
               Begin VB.PictureBox FrameL1_Picture 
                  BorderStyle     =   0  'None
                  Height          =   1695
                  Index           =   0
                  Left            =   120
                  ScaleHeight     =   1695
                  ScaleWidth      =   5295
                  TabIndex        =   176
                  Top             =   240
                  Width           =   5295
                  Begin VB.CheckBox Check_no_store 
                     Caption         =   "强制 Cache-Control: no-store"
                     Height          =   255
                     Left            =   1200
                     MousePointer    =   14  'Arrow and Question
                     TabIndex        =   180
                     ToolTipText     =   "向 服务器/客户端 强制发送 Cache-Control: no-store 头信息：""不进行缓存"""
                     Top             =   1320
                     Width           =   3135
                  End
                  Begin VB.CheckBox Check_no_cache 
                     Caption         =   "强制 Pragma: no-cache"
                     Height          =   255
                     Left            =   1200
                     MousePointer    =   14  'Arrow and Question
                     TabIndex        =   179
                     ToolTipText     =   "向 服务器/客户端 强制发送 Pragma: no-cache 头信息：""本地没有缓存"""
                     Top             =   960
                     Width           =   2415
                  End
                  Begin VB.ComboBox Combo_DelCache_BefDL 
                     Height          =   300
                     ItemData        =   "sys.frx":4E30
                     Left            =   1800
                     List            =   "sys.frx":4E40
                     Style           =   2  'Dropdown List
                     TabIndex        =   178
                     Top             =   0
                     Width           =   2775
                  End
                  Begin VB.ComboBox Combo_DelCache_AftDL 
                     Height          =   300
                     ItemData        =   "sys.frx":4E8E
                     Left            =   1800
                     List            =   "sys.frx":4EA1
                     Style           =   2  'Dropdown List
                     TabIndex        =   177
                     Top             =   480
                     Width           =   2775
                  End
                  Begin VB.Label FrameL1_lab 
                     AutoSize        =   -1  'True
                     Caption         =   "下载前"
                     Height          =   180
                     Index           =   8
                     Left            =   1200
                     MousePointer    =   14  'Arrow and Question
                     TabIndex        =   183
                     ToolTipText     =   $"sys.frx":4F09
                     Top             =   45
                     Width           =   540
                  End
                  Begin VB.Label FrameL1_lab 
                     AutoSize        =   -1  'True
                     Caption         =   "下载后"
                     Height          =   180
                     Index           =   9
                     Left            =   1200
                     MousePointer    =   14  'Arrow and Question
                     TabIndex        =   182
                     ToolTipText     =   $"sys.frx":4F7B
                     Top             =   525
                     Width           =   540
                  End
                  Begin VB.Label FrameL1_lab 
                     AutoSize        =   -1  'True
                     Caption         =   "历史缓存设置:"
                     ForeColor       =   &H00C00000&
                     Height          =   180
                     Index           =   5
                     Left            =   0
                     MousePointer    =   14  'Arrow and Question
                     TabIndex        =   181
                     ToolTipText     =   $"sys.frx":4FED
                     Top             =   45
                     Width           =   1170
                  End
               End
               Begin VB.Label FrameL1_move_lab 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "上下拖拽移动面板"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   7.5
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000010&
                  Height          =   150
                  Index           =   0
                  Left            =   4560
                  TabIndex        =   207
                  Top             =   0
                  Width           =   1200
               End
               Begin VB.Label FrameL1_lab 
                  AutoSize        =   -1  'True
                  Caption         =   "高级设置"
                  ForeColor       =   &H00C00000&
                  Height          =   180
                  Index           =   6
                  Left            =   120
                  TabIndex        =   188
                  Top             =   0
                  Width           =   720
               End
            End
            Begin VB.Label Combo_lst1 
               AutoSize        =   -1  'True
               Caption         =   $"sys.frx":5037
               ForeColor       =   &H000000FF&
               Height          =   360
               Left            =   0
               TabIndex        =   203
               Top             =   2880
               Width           =   2610
            End
            Begin VB.Label FrameL1_lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "默认下载列表格式:"
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   3
               Left            =   0
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   202
               ToolTipText     =   $"sys.frx":5074
               Top             =   2280
               Width           =   1530
            End
            Begin VB.Label FrameL1_lab 
               AutoSize        =   -1  'True
               Caption         =   "下载区块(写入缓存):"
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   0
               Left            =   0
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   201
               ToolTipText     =   $"sys.frx":50C1
               Top             =   0
               Width           =   1710
            End
            Begin VB.Label FrameL1_lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "连接超时:"
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   1
               Left            =   0
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   200
               ToolTipText     =   $"sys.frx":511F
               Top             =   1200
               Width           =   810
            End
            Begin VB.Label FrameL1_lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "失败重连:"
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   2
               Left            =   0
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   199
               ToolTipText     =   "当连接超时后，对超时地址重新连接的次数"
               Top             =   1680
               Width           =   810
            End
            Begin VB.Label LB_timeout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "30秒"
               Height          =   180
               Left            =   3000
               TabIndex        =   198
               ToolTipText     =   "非特定需求建议选择(是)"
               Top             =   1200
               Width           =   360
            End
            Begin VB.Label LB_retry 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "无限重试"
               Height          =   180
               Left            =   3000
               TabIndex        =   197
               ToolTipText     =   "非特定需求建议选择(是)"
               Top             =   1680
               Width           =   720
            End
         End
         Begin VB.VScrollBar FrameL1_bgvs 
            Height          =   4335
            Left            =   5880
            TabIndex        =   173
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.Frame FrameL 
      Caption         =   "常规参数设置"
      ForeColor       =   &H00C00000&
      Height          =   5295
      Index           =   5
      Left            =   9000
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame FrameL5_Frame1 
         Height          =   1575
         Left            =   120
         TabIndex        =   152
         Top             =   1680
         Width           =   6015
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   1635
            TabIndex        =   163
            Top             =   1200
            Width           =   1635
            Begin VB.OptionButton listOp 
               Caption         =   "否"
               Height          =   255
               Index           =   0
               Left            =   840
               TabIndex        =   165
               Top             =   0
               Width           =   495
            End
            Begin VB.OptionButton listOp 
               Caption         =   "是"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   164
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture18 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3240
            ScaleHeight     =   255
            ScaleWidth      =   1755
            TabIndex        =   160
            Top             =   1200
            Width           =   1755
            Begin VB.OptionButton set_sbar 
               Caption         =   "否"
               Height          =   255
               Index           =   0
               Left            =   840
               TabIndex        =   162
               Top             =   0
               Width           =   495
            End
            Begin VB.OptionButton set_sbar 
               Caption         =   "是"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   161
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture19 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3240
            ScaleHeight     =   255
            ScaleWidth      =   1755
            TabIndex        =   157
            Top             =   480
            Width           =   1755
            Begin VB.OptionButton set_checkall 
               Caption         =   "是"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   159
               Top             =   0
               Width           =   495
            End
            Begin VB.OptionButton set_checkall 
               Caption         =   "否"
               Height          =   255
               Index           =   0
               Left            =   840
               TabIndex        =   158
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   1755
            TabIndex        =   154
            Top             =   480
            Width           =   1755
            Begin VB.OptionButton set_tray 
               Caption         =   "是"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   156
               Top             =   0
               Width           =   495
            End
            Begin VB.OptionButton set_tray 
               Caption         =   "否"
               Height          =   255
               Index           =   0
               Left            =   840
               TabIndex        =   155
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "分析页面时，显示列表清单？"
            Height          =   180
            Left            =   240
            TabIndex        =   169
            ToolTipText     =   "建议选择(否)加快刷新速度"
            Top             =   960
            Width           =   2340
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "是否显示信息栏？"
            Height          =   180
            Index           =   11
            Left            =   3240
            TabIndex        =   168
            ToolTipText     =   "信息栏用于提示OX163的最新信息"
            Top             =   960
            Width           =   1440
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "是否自动全部标记多选框？"
            Height          =   180
            Index           =   12
            Left            =   3240
            TabIndex        =   167
            ToolTipText     =   "列表后自动全选功能"
            Top             =   240
            Width           =   2160
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "是否最小化到系统托盘？"
            Height          =   180
            Index           =   4
            Left            =   240
            TabIndex        =   166
            ToolTipText     =   "如果出现错误或者程序假死请选择(否)"
            Top             =   240
            Width           =   1980
         End
         Begin VB.Label sys_lab5 
            AutoSize        =   -1  'True
            Caption         =   "常规设置"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   153
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.TextBox update_host_Text 
         Height          =   270
         Left            =   1440
         TabIndex        =   141
         Text            =   "http://www.shanhaijing.net"
         Top             =   920
         Width           =   4335
      End
      Begin VB.ComboBox update_host_Combo 
         Height          =   300
         ItemData        =   "sys.frx":518B
         Left            =   1440
         List            =   "sys.frx":518D
         TabIndex        =   139
         Text            =   "update_host_Combo"
         ToolTipText     =   "以""http://""开头, 默认""http://www.shanhaijing.net/163/"""
         Top             =   1200
         Width           =   4335
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   5655
         TabIndex        =   6
         Top             =   480
         Width           =   5655
         Begin VB.CommandButton Update_now_Command 
            Caption         =   "立即检查更新"
            Height          =   300
            Left            =   1560
            TabIndex        =   10
            ToolTipText     =   "手动检查更新"
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton autoOp 
            Caption         =   "否"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   8
            Top             =   50
            Width           =   495
         End
         Begin VB.OptionButton autoOp 
            Caption         =   "是"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   50
            Width           =   495
         End
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "设置更新服务:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   120
         TabIndex        =   140
         ToolTipText     =   "以""http://""开头, 默认""http://www.shanhaijing.net/163/"""
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "自动检查更新:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "建议选择(是)自动进行新版本检查"
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.CommandButton sys_def_com 
      Caption         =   "恢复全部默认设置"
      Height          =   465
      Index           =   0
      Left            =   2160
      TabIndex        =   9
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton sys_no 
      Caption         =   "取消(&C)"
      Height          =   465
      Left            =   6360
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton sys_rec 
      Caption         =   "调用INI恢复全部设置"
      Height          =   465
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton sys_yes 
      Caption         =   "确定(&A)"
      Height          =   465
      Left            =   4920
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Frame FrameL 
      Caption         =   "文件目录操作"
      ForeColor       =   &H00C00000&
      Height          =   5295
      Index           =   2
      Left            =   9000
      TabIndex        =   68
      Top             =   120
      Width           =   6375
      Begin VB.ComboBox Combo_unicode_ctrl 
         Height          =   300
         Index           =   1
         ItemData        =   "sys.frx":518F
         Left            =   240
         List            =   "sys.frx":519C
         Style           =   2  'Dropdown List
         TabIndex        =   122
         Top             =   1080
         Width           =   5895
      End
      Begin VB.ComboBox Combo_unicode_ctrl 
         Height          =   300
         Index           =   0
         ItemData        =   "sys.frx":51FB
         Left            =   240
         List            =   "sys.frx":5208
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   480
         Width           =   5895
      End
      Begin VB.Frame Frame2 
         ForeColor       =   &H00C00000&
         Height          =   1575
         Left            =   240
         TabIndex        =   80
         Top             =   2880
         Width           =   5895
         Begin VB.ComboBox Combo_rar 
            Height          =   300
            ItemData        =   "sys.frx":5274
            Left            =   3720
            List            =   "sys.frx":5281
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox Combo_rar_name 
            Height          =   300
            ItemData        =   "sys.frx":52A3
            Left            =   2880
            List            =   "sys.frx":52A5
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   900
            Width           =   1335
         End
         Begin VB.TextBox fix_name_Text 
            Height          =   270
            Left            =   2880
            MaxLength       =   15
            TabIndex        =   99
            Top             =   570
            Width           =   2655
         End
         Begin VB.PictureBox Picture15 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   4200
            ScaleHeight     =   495
            ScaleWidth      =   1500
            TabIndex        =   97
            Top             =   840
            Width           =   1500
            Begin VB.CommandButton Command1 
               Caption         =   "添加后缀"
               Height          =   300
               Left            =   120
               TabIndex        =   98
               Top             =   60
               Width           =   1215
            End
         End
         Begin VB.PictureBox Picture16 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   480
            ScaleHeight     =   975
            ScaleWidth      =   2115
            TabIndex        =   92
            Top             =   480
            Width           =   2115
            Begin VB.OptionButton file_compare 
               Caption         =   "跳过同名文件"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   95
               ToolTipText     =   "skip same name files"
               Top             =   360
               Width           =   2055
            End
            Begin VB.OptionButton file_compare 
               Caption         =   "跳过同名同尺寸文件"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   94
               ToolTipText     =   "skip same files"
               Top             =   0
               Width           =   2055
            End
            Begin VB.OptionButton file_compare 
               Caption         =   "不比较 直接改名"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   93
               ToolTipText     =   "rename as new files"
               Top             =   720
               Width           =   2055
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "文件判断"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   116
            Top             =   0
            Width           =   720
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "伪图检查："
            Height          =   180
            Index           =   3
            Left            =   2880
            TabIndex        =   102
            ToolTipText     =   "建议使用自动改名"
            Top             =   300
            Width           =   900
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "下载文件名已存在："
            Height          =   180
            Index           =   10
            Left            =   240
            TabIndex        =   96
            ToolTipText     =   "请按用户需要设定"
            Top             =   240
            Width           =   1620
         End
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H00C00000&
         Height          =   1335
         Left            =   240
         TabIndex        =   79
         Top             =   1440
         Width           =   5895
         Begin VB.PictureBox Picture20 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3360
            ScaleHeight     =   255
            ScaleWidth      =   1635
            TabIndex        =   88
            Top             =   910
            Width           =   1635
            Begin VB.OptionButton set_url_folder 
               Caption         =   "是"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   90
               Top             =   0
               Width           =   495
            End
            Begin VB.OptionButton set_url_folder 
               Caption         =   "否"
               Height          =   255
               Index           =   0
               Left            =   840
               TabIndex        =   89
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.TextBox def_path_txt 
            Enabled         =   0   'False
            Height          =   270
            Left            =   150
            TabIndex        =   86
            Top             =   510
            Width           =   4935
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1440
            ScaleHeight     =   255
            ScaleWidth      =   1515
            TabIndex        =   83
            Top             =   240
            Width           =   1515
            Begin VB.OptionButton def_path 
               Caption         =   "启用"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   85
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton def_path 
               Caption         =   "关闭"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   84
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.PictureBox Picture11 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   5160
            ScaleHeight     =   255
            ScaleWidth      =   555
            TabIndex        =   81
            Top             =   510
            Width           =   555
            Begin VB.CommandButton def_path_com 
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   255
               Left            =   0
               TabIndex        =   82
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "目录设定"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   117
            Top             =   0
            Width           =   720
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "下载时，是否以网页地址作为目录？"
            Height          =   180
            Index           =   13
            Left            =   120
            TabIndex        =   91
            ToolTipText     =   "（如：C:\163blog.vbs_vbscript_GB2312\http：／／blog.163.com／aaa／\）"
            Top             =   960
            Width           =   2880
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "下载默认路径："
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   87
            ToolTipText     =   "建议自定义设置"
            Top             =   270
            Width           =   1260
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "其他Unicode文本字符："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   16
         Left            =   240
         TabIndex        =   121
         Top             =   840
         Width           =   1890
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unicode文件夹\文件名："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   15
         Left            =   240
         TabIndex        =   119
         Top             =   240
         Width           =   1980
      End
   End
   Begin VB.Frame FrameL 
      Caption         =   "脚本控制"
      ForeColor       =   &H00C00000&
      Height          =   5295
      Index           =   3
      Left            =   15480
      TabIndex        =   69
      Top             =   120
      Width           =   6375
      Begin VB.Frame Frame3 
         Height          =   3615
         Left            =   240
         TabIndex        =   142
         Top             =   960
         Width           =   5895
         Begin VB.FileListBox scriptFile 
            Height          =   2970
            Hidden          =   -1  'True
            Left            =   3600
            Pattern         =   "*.txt"
            TabIndex        =   148
            Top             =   480
            Width           =   2175
         End
         Begin VB.CommandButton IncLstCtrl_Com1 
            Height          =   255
            Index           =   1
            Left            =   3120
            Picture         =   "sys.frx":52A7
            Style           =   1  'Graphical
            TabIndex        =   147
            ToolTipText     =   "Remove Include File"
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton IncLstCtrl_Com1 
            Height          =   255
            Index           =   0
            Left            =   3120
            Picture         =   "sys.frx":5303
            Style           =   1  'Graphical
            TabIndex        =   146
            ToolTipText     =   "Add Incule File"
            Top             =   480
            Width           =   375
         End
         Begin MSComctlLib.ListView scriptList 
            DragIcon        =   "sys.frx":535F
            Height          =   2820
            Left            =   120
            TabIndex        =   149
            ToolTipText     =   "拖拽排列顺序"
            Top             =   480
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   4974
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "include_name"
               Object.Tag             =   "name_include"
               Text            =   "include列表"
               Object.Width           =   3881
            EndProperty
         End
         Begin VB.Label custom_sLabel2 
            AutoSize        =   -1  'True
            Caption         =   "Include Flie(优先级由高至低)"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   9
            Left            =   120
            TabIndex        =   145
            Top             =   240
            Width           =   2520
         End
         Begin VB.Label custom_sLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "custom文件夹(点击打开)"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   8
            Left            =   3600
            MouseIcon       =   "sys.frx":93C9
            MousePointer    =   99  'Custom
            TabIndex        =   144
            Top             =   240
            Width           =   1980
         End
         Begin VB.Label custom_sLabel1 
            AutoSize        =   -1  'True
            Caption         =   "include列表"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   7
            Left            =   120
            TabIndex        =   143
            Top             =   0
            Width           =   990
         End
      End
      Begin VB.PictureBox Picture12 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   2895
         TabIndex        =   72
         Top             =   600
         Width           =   2895
         Begin VB.OptionButton scriptOP 
            Caption         =   "延后"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   75
            ToolTipText     =   "在程序分析完是否为163相册后执行"
            Top             =   20
            Width           =   735
         End
         Begin VB.OptionButton scriptOP 
            Caption         =   "优先"
            Height          =   255
            Index           =   1
            Left            =   915
            TabIndex        =   74
            ToolTipText     =   "优先执行外部脚本"
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton scriptOP 
            Caption         =   "关闭"
            Height          =   255
            Index           =   2
            Left            =   1785
            TabIndex        =   73
            ToolTipText     =   "关闭外部脚本执行"
            Top             =   20
            Width           =   735
         End
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "脚本信息报告"
         Height          =   180
         Left            =   4080
         TabIndex        =   118
         Top             =   360
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "脚本调用设置:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   76
         ToolTipText     =   "请按用户需要设定"
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame ToolTip_Frame 
      Caption         =   "功能说明"
      Height          =   615
      Left            =   120
      TabIndex        =   170
      Top             =   6120
      Width           =   8775
      Begin VB.Label ToolTip_Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   240
         TabIndex        =   171
         Top             =   240
         Width           =   90
      End
   End
End
Attribute VB_Name = "sys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_dragItem As ListItem
Dim m_dragNode As Node
Dim move_tf As Byte
Dim pos_y As Single
Private Const VS_int = 600

Private Sub Combo_lst_Click()
    'lst (for flashget)
    'htm(for All Tools)
    'txt & bat(for All)
    If Combo_lst.ListIndex = 0 Then
        Combo_lst1.caption = "导出带有自动更名的LST下载列表" & vbCrLf & "适用于flashget1.96等经典版本"
    ElseIf Combo_lst.ListIndex = 1 Then
        Combo_lst1.caption = "导出带有全部下载信息的htm页面" & vbCrLf & "可以直接调用迅雷等下载软件"
    ElseIf Combo_lst.ListIndex = 2 Then
        Combo_lst1.caption = "导出一个仅有下载地址的txt文档" & vbCrLf & "同时生成一个bat文档用于重命名"
    End If
End Sub

Private Sub Combo_lst_KeyPress(KeyAscii As Integer)
    'lst (for flashget)
    'htm(for All Tools)
    'txt & bat(for All)
    If Combo_lst.ListIndex = 0 Then
        Combo_lst1.caption = "导出带有自动更名的LST下载列表" & vbCrLf & "适用于flashget1.96等经典版本"
    ElseIf Combo_lst.ListIndex = 1 Then
        Combo_lst1.caption = "导出带有全部下载信息的htm页面" & vbCrLf & "可以直接调用迅雷等下载软件"
    ElseIf Combo_lst.ListIndex = 2 Then
        Combo_lst1.caption = "导出一个仅有下载地址的txt文档" & vbCrLf & "同时生成一个bat文档用于重命名"
    End If
End Sub



Private Sub Combo_rar_name_Click()
    If Combo_rar_name.ListIndex > 0 Then
        fix_name_Text.Text = Combo_rar_name.List(Combo_rar_name.ListIndex)
        Command1.caption = "修改后缀"
    Else
        fix_name_Text.Text = ""
        Command1.caption = "添加后缀"
    End If
End Sub


Private Sub Auto_Password_com_Click()
    On Error Resume Next
    sys.Enabled = False
    Auto_Password_com.caption = "正在查找,请等待..."
    Dim html, split_html
    html = Form1.update.OpenURL(sysSet.update_host & "passcode_inf.txt?ntime=" & CDbl(Now()))
    split_html = Split(html, vbCrLf)
    Randomize
    a_count = Int(Rnd * (UBound(split_html) + 1))
    html = Split(split_html(a_count), "|")
    If UBound(html) > 1 Then
        If html(0) <> "" And html(1) <> "" And html(2) <> "" Then
            passcode_text(0) = html(0)
            passcode_text(1) = html(1)
            passcode_text(2) = html(2)
        End If
    End If
    Auto_Password_com.caption = "自动填写"
    sys.Enabled = True
End Sub


Private Sub Combo_UA_List_Click()
    Customize_UA_txt = OX_UA_Const(Combo_UA_List.ListIndex)
End Sub

Private Sub Combo_UA_List_KeyPress(KeyAscii As Integer)
    Customize_UA_txt = OX_UA_Const(Combo_UA_List.ListIndex)
End Sub

Private Sub Customize_UA_txt_DblClick()
    Customize_UA.SelStart = 0
    Customize_UA.SelLength = Len(Customize_UA.Text)
End Sub

Private Sub Customize_UA_txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 65 And Shift = vbCtrlMask Then
        Customize_UA_txt_DblClick
    End If
End Sub

Private Sub show_ToolTip()
If ToolTip_Lab.caption <> "" Then ToolTip_Frame.Height = 195 + ToolTip_Lab.Top + ToolTip_Lab.Height: sys.Height = 6510 + ToolTip_Frame.Height + 150
End Sub

Private Sub hide_ToolTip()
sys.Height = 6510
ToolTip_Lab.caption = ""
End Sub

Private Sub FrameL1_bgs_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
move_tf = 1
FrameL1_bgs.MousePointer = 7
pos_y = Y
End Sub

Private Sub FrameL1_bgs_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Static mov_y As Single
mov_y = FrameL1_bgs.Top
If move_tf = 1 Then
mov_y = mov_y - (pos_y - Y)
If mov_y > 0 Then mov_y = 0
If mov_y < FrameL1_bg.Height - FrameL1_bgs.Height Then mov_y = FrameL1_bg.Height - FrameL1_bgs.Height
FrameL1_bgs.Top = mov_y
FrameL1_bgvs.Value = Int(-FrameL1_bgs.Top / VS_int)
End If
End Sub

Private Sub FrameL1_bgs_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
move_tf = 0
pos_y = 0
FrameL1_bgs.MousePointer = 0
End Sub

Private Sub FrameL1_bgvs_Change()
If move_tf <> 1 Then FrameL1_bgs.Top = 0 - FrameL1_bgvs.Value * VS_int
End Sub

Private Sub FrameL1_bgvs_Scroll()
FrameL1_bgs.Top = 0 - FrameL1_bgvs.Value * VS_int
End Sub

Private Sub FrameL1_Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call FrameL1_bgs_MouseDown(Button, Shift, x, Y)
End Sub

Private Sub FrameL1_Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call FrameL1_bgs_MouseMove(Button, Shift, x, Y)
End Sub

Private Sub FrameL1_Frame1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call FrameL1_bgs_MouseUp(Button, Shift, x, Y)
End Sub

Private Sub FrameL1_move_lab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Call FrameL1_bgs_MouseDown(Button, Shift, x, Y)
End Sub

Private Sub FrameL1_move_lab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Call FrameL1_bgs_MouseMove(Button, Shift, x, Y)
End Sub

Private Sub FrameL1_move_lab_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Call FrameL1_bgs_MouseUp(Button, Shift, x, Y)
End Sub

Private Sub FrameL1_Picture_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Call FrameL1_bgs_MouseDown(Button, Shift, x, Y)
End Sub

Private Sub FrameL1_Picture_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Call FrameL1_bgs_MouseMove(Button, Shift, x, Y)
End Sub

Private Sub FrameL1_Picture_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Call FrameL1_bgs_MouseUp(Button, Shift, x, Y)
End Sub

Private Sub FrameL1_lab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
ToolTip_Lab.caption = FrameL1_lab(Index).ToolTipText
show_ToolTip
End Sub

Private Sub FrameL1_lab_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
hide_ToolTip
End Sub
Private Sub Check_no_cache_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
ToolTip_Lab.caption = Check_no_cache.ToolTipText
show_ToolTip
End Sub

Private Sub Check_no_cache_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
hide_ToolTip
End Sub

Private Sub Check_no_store_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
ToolTip_Lab.caption = Check_no_store.ToolTipText
show_ToolTip
End Sub

Private Sub Check_no_store_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
hide_ToolTip
End Sub


Private Sub IncLstCtrl_Com1_Click(Index As Integer)
    Select Case Index
    Case 0
        If Len(scriptFile.fileName) > 0 And sys_CheckIncLst_NoThisfile(scriptFile.fileName) = True Then
            scriptList.ListItems.Add , , scriptFile.fileName
            scriptList.ListItems(scriptList.ListItems.count).Checked = True
        End If
        
    Case 1
        If MsgBox("是否要移除选中的自定义Inc文件?" & vbCrLf & "(sys_163与sys_include不会被移除)", vbYesNo, "询问") = vbYes Then
            For i = scriptList.ListItems.count To 1 Step -1
                DoEvents
                If scriptList.ListItems(i).Selected = True And scriptList.ListItems(i).Text <> "sys_163" And scriptList.ListItems(i).Text <> "sys_include" Then
                    scriptList.ListItems.Remove i
                End If
            Next
        End If
        
    End Select
End Sub

Private Function sys_CheckIncLst_NoThisfile(scriptFileName As String) As Boolean
    sys_CheckIncLst_NoThisfile = True
    For i = 1 To scriptList.ListItems.count
        If scriptFileName = scriptList.ListItems(i).Text Then
            sys_CheckIncLst_NoThisfile = False
            MsgBox "文件已存在", vbOKOnly, "警告"
        End If
    Next
End Function



Private Sub scriptFile_DblClick()
    Call IncLstCtrl_Com1_Click(0)
End Sub

Private Sub scriptList_DblClick()
    scriptList.SelectedItem.Checked = Not scriptList.SelectedItem.Checked
End Sub


Private Sub set_sbar_Click(Index As Integer)
If Index = 0 Then MsgBox "建议开启状态栏, 否则你无法看到下载提示信息以及快速功能!", vbOKOnly + vbExclamation, "警告"
End Sub

Private Sub sys_lab1_Click(Index As Integer)

End Sub

Private Sub Update_now_Command_Click()
    Form1.Timer3.Enabled = True
    Update_now_Command.caption = "再次检查更新"
End Sub

Private Sub def_path_Click(Index As Integer)
    If def_path(1).Value = True Then
        def_path_com.Enabled = True
        def_path_txt.Enabled = True
    Else
        def_path_com.Enabled = False
        def_path_txt.Enabled = False
    End If
End Sub

Private Sub def_path_com_Click()
    On Error Resume Next
retry:
    Folder_path = ""
    If Right(def_path_txt, 1) <> "\" Then def_path_txt = def_path_txt & "\"
    Folder_path = GetFolder("请选择默认文件夹", def_path_txt, True)
    
    If Mid$(Folder_path, 2, 2) = ":\" Then
        If (GetFileAttributes(Folder_path) = -1) Then MsgBox "该路径不能保存文件", vbOKOnly + vbExclamation, "警告": GoTo retry
        def_path_txt = Folder_path
    End If
End Sub

Private Sub downHS_Change()
    downText.Text = downHS.Value & "KB"
    downOp(3).Value = True
End Sub
Private Sub downHS_Scroll()
    downText.Text = downHS.Value & "KB"
    downOp(3).Value = True
End Sub
Private Sub fix_name_Text_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1_Click
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    sys.Width = 9105
    sys.Height = 6510
    Dim i As Byte
    For i = 1 To 9
        FrameL(i).Top = 120
        FrameL(i).Left = 2520
    Next i
    FrameL1_bgs.Top = 0
    FrameL1_bgs.Left = 0
    FrameL1_bgvs.Max = Int((FrameL1_bgs.Height - FrameL1_bg.Height) / VS_int)
    Call Build_TVW_Menu
    Call SysTreeView_NodeClick(SysTreeView.Nodes(1))
    Form1.always_on_top False
    'Dim flags As Integer
    'flags = SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
    'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags
    
    
    Form1.Enabled = False
    scriptFile.Path = App_path & "\include\custom"
    scriptList.Height = scriptFile.Height
    OX_Start_log_Text = OX_Start_log_Text & vbCrLf & vbCrLf & OX_Start_log
    Call sys_def(0)
    Call load_ini(0)
End Sub

Private Sub Build_TVW_Menu()
    Call SysTreeView.Nodes.Add(, 4, "TVW1", "网络下载设置", 1)
    Call SysTreeView.Nodes.Add("TVW1", 4, "TVW2", "文件目录操作", 2)
    Call SysTreeView.Nodes.Add("TVW1", 4, "TVW3", "脚本控制", 3)
    Call SysTreeView.Nodes.Add("TVW1", 4, "TVW4", "代理服务器", 4)
    Call SysTreeView.Nodes.Add(, 4, "TVW5", "常规参数设置", 5)
    Call SysTreeView.Nodes.Add("TVW5", 4, "TVW6", "热键与警告框", 6)
    Call SysTreeView.Nodes.Add("TVW5", 4, "TVW7", "网易相册设置", 7)
    Call SysTreeView.Nodes.Add(, 4, "TVW8", "内置浏览器", 8)
    Call SysTreeView.Nodes.Add(, 4, "TVW9", "维护与工具", 9)
    Dim nodx As Node
    For Each nodx In SysTreeView.Nodes
        nodx.Expanded = True
    Next
    SysTreeView.Nodes(1).Selected = True
    'Set myNod = 控件名.Nodes.Add(a, b, key, text, image)
    '参数说明:
    'a: 参照物,在谁的基础上建节点,a就是谁的key值,如果是跟节点,可省略.
    'b: 参照物和本身的关系 , 如果是父子关系, 值为tvwchile, 如果是兄弟关系, 值为tvwnext
    'tvwlast--1；该节点置于任何其他的在relative中被命名的同一级别的节点的后面
    'tvwNext--2；该节点置于在relative中被命名节点的后面
    'tvwPrevius--3；该节点置于在relative中被命名的节点的前面
    'tvwChild--4；该节点成为在relative中被命名的节点的的子节点
    'key: 关键字,唯一的.
    'text: 节点上显示的文字
    'image: 节点前的小图标 , 需要配合图标控件用, 可省略
    '    With Me.TreeView1.Nodes
    '        .Add , 4, "K1", "分类（一）", Form1.user_list_save.Picture
    '        .Add "K1", 4, , "小分类1"
    '        .Add "K1", 4, , "小分类2"
    '        .Add , 4, "K2", "分类（二）"
    '        .Add "K2", 4, , "小分类1"
    '        .Add "K2", 4, , "小分类2"
    '    End With
    
End Sub

Private Sub custom_sLabel_Click(Index As Integer)
    Shell "explorer.exe " & App_path & "\include\custom", vbNormalFocus
End Sub

Private Sub SysTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim i As Byte
    For i = 1 To 9
        If SysTreeView.Nodes(i).Image <> i Then SysTreeView.Nodes(i).Image = i
        FrameL(i).Visible = False
    Next i
    Node.Image = 9 + Node.Index
    FrameL(Node.Index).Visible = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Form1.WindowState = 0 Then Form1.always_on_top sysSet.always_top
    Form1.Enabled = True
End Sub

Private Sub frame_def_Click()
    Call sys_def(SysTreeView.SelectedItem.Index)
End Sub

Private Sub frame_rec_Click()
    Call sys_def(SysTreeView.SelectedItem.Index)
    Call load_ini(SysTreeView.SelectedItem.Index)
End Sub

Private Sub list_copy_Click(Index As Integer)
    If list_copy(1).Value = True Then
        ubb_copy(0).Value = True
    Else
        ubb_copy(1).Value = True
    End If
End Sub

Private Sub LST_Help_Click()
    Dim help_str As String
    help_str = "HTML下载列表(默认):适用于各种下载工具" & vbCrLf & _
    "操作方法: 直接用浏览器打开操作" & vbCrLf & vbCrLf & _
    "LST下载列表:适用于flashget1.96等经典版本" & vbCrLf & _
    "操作方法:文件菜单->导入LST下载列表" & vbCrLf & vbCrLf & _
    "TXT& BAT下载列表: 适用于适用于各种下载工具" & vbCrLf & _
    "操作方法:文件菜单->导入TXT下载列表" & vbCrLf & _
    "          下载完成后将BAT文件放入下载目录运行" & vbCrLf & _
    "          完成重命名工作"
    Call MsgBox(help_str, vbOKOnly, "下载列表文件帮助说明")
End Sub

Private Sub proxy_com1_Click()
    proxy_txt1(0).Text = ""
    proxy_txt1(1).Text = ""
    proxy_txt1(2).Text = ""
    proxy_txt2(0).Text = ""
    proxy_txt2(1).Text = ""
    proxy_txt2(2).Text = ""
    ProxyComb(0).ListIndex = 0
    ProxyComb(1).ListIndex = 0
End Sub

Private Sub proxy_com2_Click()
    proxy_txt2(0).Text = proxy_txt1(0).Text
    proxy_txt2(1).Text = proxy_txt1(1).Text
    proxy_txt2(2).Text = proxy_txt1(2).Text
    ProxyComb(1).ListIndex = ProxyComb(0).ListIndex
End Sub

Private Sub proxy_com3_Click()
    proxy_txt1(0).Text = proxy_txt2(0).Text
    proxy_txt1(1).Text = proxy_txt2(1).Text
    proxy_txt1(2).Text = proxy_txt2(2).Text
    ProxyComb(0).ListIndex = ProxyComb(1).ListIndex
End Sub

Private Sub sys_apply_Click()
    On Error Resume Next
    sys.Enabled = False
    '网络下载设置------------------------------------
    '下载区块
    For i = 0 To 3
        If downOp(i).Value = True Then
            Select Case i
            Case 0
                WriteIniStr "maincenter", "downloadblock", "10240"
            Case 1
                WriteIniStr "maincenter", "downloadblock", "51200"
            Case 2
                WriteIniStr "maincenter", "downloadblock", "102400"
            Case Else
                WriteIniStr "maincenter", "downloadblock", CLng(downHS.Value) * 1024
            End Select
            Exit For
        End If
    Next i
    '超时
    WriteIniStr "maincenter", "time_out", VS_timeout.Value
    '重试
    WriteIniStr "maincenter", "retry_times", VS_retry.Value
    '导出下载列表格式
    WriteIniStr "maincenter", "list_type", Combo_lst.ListIndex
    '历史缓存设置
    WriteIniStr "maincenter", "DelCache_BefDL", Combo_DelCache_BefDL.ListIndex
    WriteIniStr "maincenter", "DelCache_AftDL", Combo_DelCache_AftDL.ListIndex
        'http头强制发送no-cache
      WriteIniStr "maincenter", "Cache_no_cache", Check_no_cache.Value
        'http头强制发送no-store
      WriteIniStr "maincenter", "Cache_no_store", Check_no_store.Value
    '用户代理(User Agent)store
    Customize_UA_txt = Trim(Customize_UA_txt)
    Customize_UA_txt = Replace(Customize_UA_txt, Chr(10), "")
    Customize_UA_txt = Replace(Customize_UA_txt, Chr(13), "")
    Customize_UA_txt = Replace(Customize_UA_txt, vbNullChar, "")
    If Customize_UA_txt = "" Then Combo_UA_List.ListIndex = 0
      WriteIniStr "maincenter", "Customize_UA", Customize_UA_txt
    
    '文件目录操作------------------------------------
    'Unicode文件夹\文件名
    WriteIniStr "maincenter", "Unicode_File", Combo_unicode_ctrl(0).ListIndex
    '其他Unicode文本字符
    WriteIniStr "maincenter", "Unicode_Str", Combo_unicode_ctrl(1).ListIndex
    '下载默认路径
    If def_path(1).Value = True Then
        WriteIniTF "maincenter", "def_path_tf", True
        WriteIniStr "maincenter", "def_path", def_path_txt.Text
    Else
        WriteIniTF "maincenter", "def_path_tf", False
        WriteIniStr "maincenter", "def_path", ""
    End If
    '下载时，是否以网页地址作为目录
        WriteIniTF "maincenter", "url_folder", set_url_folder(1).Value
    '下载文件名已存在
    If file_compare(1).Value = True Then
        WriteIniStr "maincenter", "file_compare", "1"
    ElseIf file_compare(2).Value = True Then
        WriteIniStr "maincenter", "file_compare", "2"
    Else
        WriteIniStr "maincenter", "file_compare", "0"
    End If
    '伪图检查
    WriteIniStr "maincenter", "fix_rar", Combo_rar.ListIndex
    '伪图类型
    fix_rar_name = ""
    If Combo_rar_name.ListCount > 1 Then
        For i = 1 To Combo_rar_name.ListCount - 1
            fix_rar_name = fix_rar_name & Combo_rar_name.List(i) & "|"
        Next i
    End If
    If Right$(fix_rar_name, 1) = "|" Then fix_rar_name = Left$(fix_rar_name, Len(fix_rar_name) - 1)
    WriteIniStr "maincenter", "fix_rar_name", fix_rar_name
    
    '脚本控制----------------------------------------
    '脚本调用设置
    If scriptOP(0).Value = True Then
        WriteIniStr "maincenter", "include_script", "delay"
    ElseIf scriptOP(1).Value = True Then
        WriteIniStr "maincenter", "include_script", "first"
    Else
        WriteIniStr "maincenter", "include_script", "close"
    End If
    'include列表
    Dim sys_scriptlist_str As String
    sys_scriptlist_str = ""
    For i = 1 To scriptList.ListItems.count
        sys_scriptlist_str = sys_scriptlist_str & "|" & scriptList.ListItems(i).Text
        If scriptList.ListItems(i).Checked = True Then
            sys_scriptlist_str = sys_scriptlist_str & ",1"
        Else
            sys_scriptlist_str = sys_scriptlist_str & ",0"
        End If
    Next
    sys_scriptlist_str = OX_Check_include_scriptlist(Mid(sys_scriptlist_str, 2), False) 'Mid去除第一个"|"
    WriteIniStr "maincenter", "include_scriptList", sys_scriptlist_str
    
    '代理服务器设置----------------------------------
    Select Case ProxyComb(0).ListIndex
    Case 1
        WriteIniStr "proxyset", "proxy_A_type", "icDirect"
    Case 2
        WriteIniStr "proxyset", "proxy_A_type", "icNamedProxy"
    Case Else
        WriteIniStr "proxyset", "proxy_A_type", "icUseDefault"
    End Select
    
    Select Case ProxyComb(1).ListIndex
    Case 1
        WriteIniStr "proxyset", "proxy_B_type", "icDirect"
    Case 2
        WriteIniStr "proxyset", "proxy_B_type", "icNamedProxy"
    Case Else
        WriteIniStr "proxyset", "proxy_B_type", "icUseDefault"
    End Select
    
    proxy_txt1(0) = Trim(Replace(Replace(proxy_txt1(0), Chr(10), ""), Chr(13), ""))
    proxy_txt1(1) = Trim(Replace(Replace(proxy_txt1(1), Chr(10), ""), Chr(13), ""))
    proxy_txt1(2) = Trim(Replace(Replace(proxy_txt1(2), Chr(10), ""), Chr(13), ""))
    proxy_txt2(0) = Trim(Replace(Replace(proxy_txt2(0), Chr(10), ""), Chr(13), ""))
    proxy_txt2(1) = Trim(Replace(Replace(proxy_txt2(1), Chr(10), ""), Chr(13), ""))
    proxy_txt2(2) = Trim(Replace(Replace(proxy_txt2(2), Chr(10), ""), Chr(13), ""))
    'proxy_A
    WriteIniStr "proxyset", "proxy_A", proxy_txt1(0)
    WriteIniStr "proxyset", "proxy_A_user", proxy_txt1(1)
    WriteIniStr "proxyset", "proxy_A_pw", proxy_txt1(2)
    'proxy_B
    WriteIniStr "proxyset", "proxy_B", proxy_txt2(0)
    WriteIniStr "proxyset", "proxy_B_user", proxy_txt2(1)
    WriteIniStr "proxyset", "proxy_B_pw", proxy_txt2(2)
    '对内置浏览器启用代理A
    WriteIniStr "proxyset", "web_proxy", web_proxy_box.Value
    '常规参数设置------------------------------------
    '自动更新
        WriteIniTF "maincenter", "autocheck", autoOp(1).Value
    '更新服务器设置
    If Left(LCase(update_host_Text), 7) = "http://" Or Left(LCase(update_host_Text), 8) = "https://" Then
        WriteIniStr "maincenter", "update_host", update_host_Text
    Else
        WriteIniStr "maincenter", "update_host", "http://www.shanhaijing.net/163/"
    End If
    '列表时显示列表
        WriteIniTF "maincenter", "listshow", listOp(1).Value = True
    '是否最小化到系统托盘
        WriteIniTF "maincenter", "sysTray", set_tray(1).Value
    '是否显示信息栏
        WriteIniTF "maincenter", "bottom_StatusBar", set_sbar(1).Value
    '是否自动全部标记多选框
        WriteIniTF "maincenter", "check_all", set_checkall(1).Value
    '热键与警告框------------------------------------
    '复制链接(url && Lst)到剪贴板
        WriteIniTF "maincenter", "list_copy", list_copy(1).Value
    '目录错误时，保存到download目录
        WriteIniTF "maincenter", "savedef", saveOp(1).Value
    '程序执行时，打开退出询问
        WriteIniTF "maincenter", "askquit", quitOp(1).Value
    '密码错误时，询问是否重填密码？
        WriteIniTF "maincenter", "change_psw", changepsw(1).Value
    '下载完成后，是否出现提示框？
        WriteIniTF "maincenter", "openfloder", askfloder(1).Value
    '网易相册设置------------------------------------
    '163相册验证码
    passcode_text(0) = Replace(Replace(passcode_text(0), Chr(10), ""), Chr(13), "")
    passcode_text(1) = Replace(Replace(passcode_text(1), Chr(10), ""), Chr(13), "")
    passcode_text(2) = Replace(Replace(passcode_text(2), Chr(10), ""), Chr(13), "")
    If passcode_text(0) = "" Or passcode_text(1) = "" Or passcode_text(2) = "" Then
            passcode_text(0) = "wehi"
            passcode_text(1) = "1530930"
            passcode_text(2) = "asd"
    End If
    WriteIniStr "maincenter", "new163passcode_user", passcode_text(0)
    WriteIniStr "maincenter", "new163passcode_album", passcode_text(1)
    WriteIniStr "maincenter", "new163passcode_pw", passcode_text(2)
    '是否修正163相册中文密码问题
        WriteIniTF "maincenter", "new163pass_rules", new163passrule(1).Value
    '内置浏览器设置------------------------------------
    '是否阻止浏览器弹出新开窗口
        WriteIniTF "maincenter", "new_ie_win", ie_window(1).Value
    '是否用OX163打开新窗口
        WriteIniTF "maincenter", "ox163_ie_win", ox163_window(1).Value
    '---------------------------------------------------
    
    '重新载入设定
    OX_GetIni_Setting sysSet
    
    If sysSet.bottom_StatusBar = True Then
        Form1.show_StatusBar = 255
        Form1.StatusBar.Visible = True
        If Form1.form_height < 3000 Then Form1.form_height = 1470 + Form1.show_StatusBar
        If Form1.Height < 1470 + Form1.show_StatusBar Then Form1.Height = 1470 + Form1.show_StatusBar
        Form1.frame_resize
    Else
        Form1.show_StatusBar = 0
        Form1.StatusBar.Visible = False
        If Form1.form_height < 3000 Then Form1.form_height = 1470
        Form1.frame_resize
    End If
    
    If sysSet.def_path_tf = True Then
        sysSet.def_path = GetIniStr("maincenter", "def_path")
        If Mid$(sysSet.def_path, 2, 2) <> ":\" And Len(sysSet.def_path) > 2 Then GoTo reset_path
        If Right(sysSet.def_path, 1) = "\" Then sysSet.def_path = Mid$(sysSet.def_path, 1, Len(sysSet.def_path) - 1): WriteIniStr "maincenter", "def_path", sysSet.def_path
        If (GetFileAttributes(sysSet.def_path) = -1) Then GoTo reset_path
    Else
reset_path:
        sysSet.def_path_tf = False
        def_path(0).Value = True
        If sysSet.def_path <> "" Then sysSet.def_path = "": WriteIniStr "maincenter", "def_path", ""
    End If
        
    Call Proxy_set
    
    If sysSet.list_type >= 0 And sysSet.list_type <= 2 Then
        Form1.list_output.Picture = Form1.ImageLibrary_Normal.ListImages(10 + sysSet.list_type).Picture
        Form1.out_all.Picture = Form1.ImageLibrary_Normal.ListImages(10 + sysSet.list_type).Picture
        Form1.user_list_output.Picture = Form1.ImageLibrary_Normal.ListImages(10 + sysSet.list_type).Picture
    End If
    
    sys.Enabled = True
End Sub

Private Sub ubb_copy_Click(Index As Integer)
    If ubb_copy(1).Value = True Then
        list_copy(0).Value = True
    Else
        list_copy(1).Value = True
    End If
End Sub

Private Sub sys_def_com_Click(Index As Integer)
    Call sys_def(0)
End Sub

Private Sub sys_def(ByVal frameID As Byte)
    On Error Resume Next
    '网络下载设置------------------------------------
    If frameID = 0 Or frameID = 1 Then
        '下载区块
        downOp(1).Value = True
        '超时
        VS_timeout.Value = 30
        '重试
        VS_retry.Value = 5
        '导出下载列表格式
        Combo_lst.ListIndex = 1
        '历史缓存设置
        Combo_DelCache_BefDL.ListIndex = 0
        Combo_DelCache_AftDL.ListIndex = 0
        'http头强制发送no-cache
        Check_no_cache.Value = 0
        'http头强制发送no-store
        Check_no_store.Value = 0
        '用户代理(User Agent)
        Combo_UA_List.ListIndex = 1
        Combo_UA_List.ListIndex = 0
    End If
    
    '文件目录操作------------------------------------
    If frameID = 0 Or frameID = 2 Then
        'Unicode文件夹\文件名
        Combo_unicode_ctrl(0).ListIndex = 0
        '其他Unicode文本字符：
        Combo_unicode_ctrl(1).ListIndex = 0
        '下载默认路径
        def_path(0).Value = True
        def_path_txt = ""
        '下载时，是否以网页地址作为目录
        set_url_folder(0).Value = True
        '下载文件名已存在
        file_compare(1).Value = True
        '伪图检查
        Combo_rar.ListIndex = 1
        fix_name_list "RAR|ZIP|7Z|PNG|BMP"
        Combo_rar_name.ListIndex = 0
    End If
    
    '脚本控制---------------------------------------
    If frameID = 0 Or frameID = 3 Then
        scriptOP(0).Value = True
        scriptList.ListItems.Clear
        scriptList.ListItems.Add , , "sys_163"
        scriptList.ListItems(scriptList.ListItems.count).Checked = True
        scriptList.ListItems.Add , , "sys_include"
        scriptList.ListItems(scriptList.ListItems.count).Checked = True
    End If
    
    '代理服务器设置------------------------------------
    If frameID = 0 Or frameID = 4 Then
        proxy_txt1(0).Text = ""
        proxy_txt1(1).Text = ""
        proxy_txt1(2).Text = ""
        proxy_txt2(0).Text = ""
        proxy_txt2(1).Text = ""
        proxy_txt2(2).Text = ""
        ProxyComb(0).ListIndex = 0
        ProxyComb(1).ListIndex = 0
        web_proxy_box.Value = 1
    End If
    
    '常规参数设置------------------------------------
    If frameID = 0 Or frameID = 5 Then
        '自动更新
        autoOp(1).Value = True
        '更新服务器设置
        update_host_Combo.List(0) = "默认设置|" & "http://www.shanhaijing.net/163/"
        update_host_1 = Split(update_host_info1, "|")
        update_host_2 = Split(update_host_info2, "|")
        For i = 0 To UBound(update_host_1)
            update_host_Combo.List(i + 1) = update_host_2(i) & "|" & update_host_1(i)
        Next i
        update_host_Combo.ListIndex = 0
        '列表时显示列表
        listOp(0).Value = True
        '是否最小化到系统托盘
        set_tray(1).Value = True
        '是否显示信息栏
        set_sbar(1).Value = True
        '是否自动全部标记多选框
        set_checkall(1).Value = True
        
    End If
    
    '热键与警告框------------------------------------
    If frameID = 0 Or frameID = 6 Then
        '复制链接(url && Lst)到剪贴板
        list_copy(1).Value = True
        '目录错误时，保存到download目录
        saveOp(1).Value = True
        '程序执行时，打开退出询问
        quitOp(1).Value = True
        '密码错误时，询问是否重填密码？
        changepsw(1).Value = True
        '下载完成后，是否出现提示框？
        askfloder(1).Value = True
    End If
    
    '网易相册设置------------------------------------
    If frameID = 0 Or frameID = 7 Then
        '163相册验证码
        passcode_text(0) = "wehi"
        passcode_text(1) = "1530930"
        passcode_text(2) = "asd"
        '是否修正163相册中文密码问题
        new163passrule(1).Value = True
    End If
    
    '内置浏览器设置------------------------------------
    If frameID = 0 Or frameID = 8 Then
        '是否阻止浏览器弹出新开窗口
        ie_window(1).Value = True
        '是否用OX163打开新窗口
        ox163_window(1).Value = True
    End If
End Sub

Private Sub sys_no_Click()
    Unload sys
End Sub

Private Sub load_ini(ByVal frameID As Byte)
    On Error Resume Next
    Dim load_ini_sysSetting As sysSetting
    
    load_ini_sysSetting = OX_Default_Setting
    OX_GetIni_Setting load_ini_sysSetting
    
    '网络下载设置------------------------------------
    If frameID = 0 Or frameID = 1 Then
        '下载区块
        Select Case load_ini_sysSetting.downloadblock
        Case 10240
            downOp(0).Value = True
        Case 51200
            downOp(1).Value = True
        Case 102400
            downOp(2).Value = True
        Case Else
            downOp(3).Value = True
            If load_ini_sysSetting.downloadblock <= 1024000 Then
                downHS.Value = Int(load_ini_sysSetting.downloadblock / 1024)
            Else
                downHS.Value = 1000
            End If
        End Select
        '超时
        If load_ini_sysSetting.time_out <= 200 And load_ini_sysSetting.time_out >= 10 Then
            VS_timeout.Value = load_ini_sysSetting.time_out
        Else
            VS_timeout.Value = 30
        End If
        '重试
        If load_ini_sysSetting.retry_times <= 255 And load_ini_sysSetting.retry_times >= 0 Then
            VS_retry.Value = load_ini_sysSetting.retry_times
        Else
            VS_retry.Value = 5
        End If
        '导出下载列表格式
        If load_ini_sysSetting.list_type >= 0 And load_ini_sysSetting.list_type <= 2 Then
            Combo_lst.ListIndex = load_ini_sysSetting.list_type
        End If
        '历史缓存设置
        If load_ini_sysSetting.DelCache_BefDL >= 0 And load_ini_sysSetting.DelCache_BefDL <= 3 Then Combo_DelCache_BefDL.ListIndex = load_ini_sysSetting.DelCache_BefDL
        If load_ini_sysSetting.DelCache_AftDL >= 0 And load_ini_sysSetting.DelCache_AftDL <= 4 Then Combo_DelCache_AftDL.ListIndex = load_ini_sysSetting.DelCache_AftDL
        'http头强制发送no-cache
        Check_no_cache.Value = load_ini_sysSetting.Cache_no_cache
        'http头强制发送no-store
        Check_no_store.Value = load_ini_sysSetting.Cache_no_store
        '用户代理(User-Agent)
        Customize_UA_txt = load_ini_sysSetting.Customize_UA
    End If
    
    '文件目录操作------------------------------------
    If frameID = 0 Or frameID = 2 Then
        'Unicode文件夹\文件名
        If load_ini_sysSetting.Unicode_File >= 0 And load_ini_sysSetting.Unicode_File < 3 Then
            Combo_unicode_ctrl(0).ListIndex = load_ini_sysSetting.Unicode_File
        End If
        '其他Unicode文本字符：
        If load_ini_sysSetting.Unicode_Str >= 0 And load_ini_sysSetting.Unicode_Str < 3 Then
            Combo_unicode_ctrl(1).ListIndex = load_ini_sysSetting.Unicode_Str
        End If
        '下载默认路径
        If load_ini_sysSetting.def_path_tf = True And load_ini_sysSetting.def_path <> "" Then
            def_path(1).Value = True
            def_path_txt = load_ini_sysSetting.def_path
        Else
            def_path(0).Value = True
        End If
        '下载时，是否以网页地址作为目录
        If load_ini_sysSetting.url_folder = True Then set_url_folder(1).Value = True
        '下载文件名已存在
        If load_ini_sysSetting.file_compare = 0 Then
            file_compare(0).Value = True
        ElseIf load_ini_sysSetting.file_compare = 2 Then
            file_compare(2).Value = True
        Else
            file_compare(1).Value = True
        End If
        '伪图检查
        If load_ini_sysSetting.fix_rar >= 0 And load_ini_sysSetting.fix_rar < 3 Then
            Combo_rar.ListIndex = load_ini_sysSetting.fix_rar
        End If
        fix_name_list Trim(load_ini_sysSetting.fix_rar_name)
    End If
    
    '脚本控制---------------------------------------
    If frameID = 0 Or frameID = 3 Then
        Select Case load_ini_sysSetting.include_script
        Case "first"
            scriptOP(1).Value = True
        Case "close"
            scriptOP(2).Value = True
        Case Else
            scriptOP(0).Value = True
        End Select
        
        Dim sys_scriptlist_str, split_i
        scriptList.ListItems.Clear
        sys_scriptlist_str = Split(OX_Check_include_scriptlist(load_ini_sysSetting.include_scriptlist, False), "|")
        For split_i = 0 To UBound(sys_scriptlist_str)
            scriptList.ListItems.Add , , Left(sys_scriptlist_str(split_i), Len(sys_scriptlist_str(split_i)) - 2)
            scriptList.ListItems(scriptList.ListItems.count).Checked = (Right(sys_scriptlist_str(split_i), 1) = "1")
        Next
    End If
    
    '代理服务器设置------------------------------------
    If frameID = 0 Or frameID = 4 Then
        Select Case load_ini_sysSetting.proxy_A_type
        Case 1
            ProxyComb(0).ListIndex = 1
        Case 2
            ProxyComb(0).ListIndex = 2
        Case Else
            ProxyComb(0).ListIndex = 0
        End Select
        Select Case load_ini_sysSetting.proxy_B_type
        Case 1
            ProxyComb(1).ListIndex = 1
        Case 2
            ProxyComb(1).ListIndex = 2
        Case Else
            ProxyComb(1).ListIndex = 0
        End Select
        
        Dim proxy_str(2) As String
        Dim split_str
        proxy_str(0) = load_ini_sysSetting.proxy_A
        proxy_str(1) = load_ini_sysSetting.proxy_A_user
        proxy_str(2) = load_ini_sysSetting.proxy_A_pw
        
        proxy_str(0) = Replace(Replace(proxy_str(0), Chr(10), ""), Chr(13), "")
        proxy_str(1) = Replace(Replace(proxy_str(1), Chr(10), ""), Chr(13), "")
        proxy_str(2) = Replace(Replace(proxy_str(2), Chr(10), ""), Chr(13), "")
        
        If Len(proxy_str(0)) > 0 Then
            proxy_txt1(0) = proxy_str(0)
            proxy_txt1(1) = proxy_str(1)
            proxy_txt1(2) = proxy_str(2)
        End If
        
        proxy_str(0) = load_ini_sysSetting.proxy_B
        proxy_str(1) = load_ini_sysSetting.proxy_B_user
        proxy_str(2) = load_ini_sysSetting.proxy_B_pw
        
        proxy_str(0) = Replace(Replace(proxy_str(0), Chr(10), ""), Chr(13), "")
        proxy_str(1) = Replace(Replace(proxy_str(1), Chr(10), ""), Chr(13), "")
        proxy_str(2) = Replace(Replace(proxy_str(2), Chr(10), ""), Chr(13), "")
        
        If Len(proxy_str(0)) > 0 Then
            proxy_txt2(0) = proxy_str(0)
            proxy_txt2(1) = proxy_str(1)
            proxy_txt2(2) = proxy_str(2)
        End If
        
        Select Case load_ini_sysSetting.web_proxy
        Case 0
            web_proxy_box.Value = 0
        Case Else
            web_proxy_box.Value = 1
        End Select
    End If
    '常规参数设置------------------------------------
    If frameID = 0 Or frameID = 5 Then
        '自动更新
        If load_ini_sysSetting.autocheck = False Then autoOp(0).Value = True
        '更新服务器设置
        update_host_Combo.List(0) = "INI设置|" & load_ini_sysSetting.update_host
        If update_host_Combo.List(0) = "INI设置|" & "http://www.shanhaijing.net/163/" Then update_host_Combo.List(0) = "默认设置|" & "http://www.shanhaijing.net/163/"
        update_host_Combo.ListIndex = 0
        '列表时显示列表
        If load_ini_sysSetting.listshow = True Then listOp(1).Value = True
        '是否最小化到系统托盘
        If load_ini_sysSetting.sysTray = False Then set_tray(0).Value = True
        '是否显示信息栏
        If load_ini_sysSetting.bottom_StatusBar = False Then set_sbar(0).Value = True
        '是否自动全部标记多选框
        If load_ini_sysSetting.check_all = False Then set_checkall(0).Value = True
    End If
    '热键与警告框------------------------------------
    If frameID = 0 Or frameID = 6 Then
        '复制链接(url && Lst)到剪贴板
        If load_ini_sysSetting.list_copy = False Then list_copy(0).Value = True
        '目录错误时，保存到download目录
        If load_ini_sysSetting.savedef = False Then saveOp(0).Value = True
        '程序执行时，打开退出询问
        If load_ini_sysSetting.askquit = False Then quitOp(0).Value = True
        '密码错误时，询问是否重填密码？
        If load_ini_sysSetting.change_psw = False Then changepsw(0).Value = True
        '下载完成后，是否出现提示框？
        If load_ini_sysSetting.openfloder = False Then askfloder(0).Value = True
    End If
    '网易相册设置------------------------------------
    If frameID = 0 Or frameID = 7 Then
        '163相册验证码
        passcode_text(0) = load_ini_sysSetting.new163passcode_def(0)
        passcode_text(1) = load_ini_sysSetting.new163passcode_def(1)
        passcode_text(2) = load_ini_sysSetting.new163passcode_def(2)
        If passcode_text(0) = "" Or passcode_text(1) = "" Or passcode_text(2) = "" Then
            passcode_text(0) = "wehi"
            passcode_text(1) = "1530930"
            passcode_text(2) = "asd"
        End If
        '是否修正163相册中文密码问题
        If load_ini_sysSetting.new163pass_rules = False Then new163passrule(0).Value = True
    End If
    '内置浏览器设置------------------------------------
    If frameID = 0 Or frameID = 8 Then
        '是否阻止浏览器弹出新开窗口
        If load_ini_sysSetting.new_ie_win = False Then ie_window(0).Value = True
        '是否用OX163打开新窗口
        If load_ini_sysSetting.ox163_ie_win = False Then ox163_window(0).Value = True
    End If
End Sub
Sub fix_name_list(ByVal fix_rar_name As String)
    Combo_rar_name.Clear
    Combo_rar_name.AddItem "添加新后缀", 0
    Combo_rar_name.ListIndex = 0
    If fix_rar_name = "" Or fix_rar_name = "-1" Then Exit Sub
    name_list = Split(fix_rar_name, "|")
    For i = 0 To UBound(name_list)
        name_list(i) = Trim(name_list(i))
        If Len(name_list(i)) > 0 And is_fileName(name_list(i)) Then Combo_rar_name.AddItem name_list(i), i + 1
    Next i
    Combo_rar_name.ListIndex = 0
End Sub

Private Sub Command1_Click()
    Dim pos_i As Integer
    pos_i = Combo_rar_name.ListIndex
    '查找重复
    For i = 1 To Combo_rar_name.ListCount - 1
        If UCase(fix_name_Text.Text) = UCase(Combo_rar_name.List(i)) Then
            MsgBox "文件后缀名重复！", vbOKOnly + vbExclamation, "警告"
            Combo_rar_name.ListIndex = pos_i
            Exit Sub
        End If
    Next i
    Combo_rar_name.ListIndex = pos_i
    '添加新后缀
    If Combo_rar_name.ListIndex = 0 Then
        If is_fileName(fix_name_Text.Text) And fix_name_Text.Text <> "" Then
            Combo_rar_name.AddItem fix_name_Text.Text
            Combo_rar_name.ListIndex = Combo_rar_name.ListCount - 1
        Else
            MsgBox "文件后缀名不正确！", vbOKOnly + vbExclamation, "警告"
        End If
    Else
    '修改后缀
        If is_fileName(fix_name_Text.Text) And fix_name_Text.Text <> "" Then
            Combo_rar_name.List(pos_i) = fix_name_Text.Text
        ElseIf fix_name_Text.Text = "" Then
            Combo_rar_name.RemoveItem Combo_rar_name.ListIndex
            Combo_rar_name.ListIndex = 0
        Else
            MsgBox "文件后缀名不正确！", vbOKOnly + vbExclamation, "警告"
        End If
    End If
End Sub


Private Sub sys_rec_Click(Index As Integer)
    Call sys_def(0)
    Call load_ini(0)
End Sub

Private Sub sys_yes_Click()
    Call sys_apply_Click
    Unload sys
End Sub

Private Sub update_host_Combo_Click()
    update_host_Text = update_host_Combo.List(update_host_Combo.ListIndex)
    update_host_Text = Mid$(update_host_Text, InStr(update_host_Text, "|") + 1)
End Sub

Private Sub update_host_Combo_KeyUp(KeyCode As Integer, Shift As Integer)
    update_host_Text = update_host_Combo.List(update_host_Combo.ListIndex)
    update_host_Text = Mid$(update_host_Text, InStr(update_host_Text, "|") + 1)
End Sub

Private Sub VS_retry_Change()
    If VS_retry.Value > 0 Then
        LB_retry.caption = VS_retry.Value & "次"
    Else
        LB_retry.caption = "无限重试"
    End If
End Sub

Private Sub VS_timeout_Change()
    LB_timeout.caption = VS_timeout.Value & "秒"
End Sub

'-----------------------------------------------------------------------------
Private Sub scriptList_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
    Dim li As ListItem
    
    Set li = scriptList.HitTest(x, Y)
    If Not li Is Nothing Then
        li.EnsureVisible
        scriptList.DropHighlight = li
    End If
End Sub
Private Sub scriptList_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        Set m_dragItem = scriptList.HitTest(x, Y)
    End If
End Sub
Private Sub scriptList_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        If Not m_dragItem Is Nothing Then
            'scriptList.DragIcon = m_dragItem.CreateDragImage
            scriptList.Drag vbBeginDrag
        End If
    End If
End Sub
Private Sub scriptList_DragDrop(Source As Control, x As Single, Y As Single)
    Dim li As ListItem
    Dim addli As ListItem
    Dim i As Integer
    Dim li_check As Boolean
    
    Set li = scriptList.HitTest(x, Y)
    
    If (Not li Is Nothing) And (Not m_dragItem Is Nothing) Then
        
        If li.Index <> m_dragItem.Index Then
            li_check = True
            li_check = scriptList.ListItems(m_dragItem.Index).Checked
            If li.Index > m_dragItem.Index Then
                scriptList.ListItems.Remove m_dragItem.Index
                Set addli = scriptList.ListItems.Add(li.Index + 1, m_dragItem.key, m_dragItem.Text)
                addli.Checked = li_check
            Else
                scriptList.ListItems.Remove m_dragItem.Index
                Set addli = scriptList.ListItems.Add(li.Index, m_dragItem.key, m_dragItem.Text)
                addli.Checked = li_check
            End If
            For i = 1 To m_dragItem.ListSubItems.count
                addli.SubItems(i) = m_dragItem.ListSubItems(i).Text
            Next i
        End If
    End If
    
    scriptList.DropHighlight = Nothing
    Set m_dragItem = Nothing
    scriptList.Refresh
End Sub

'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Private Sub web_proxy_box_Click()
    If web_proxy_box.Value <> sysSet.web_proxy And web_proxy_box.Value = 0 Then
        MsgBox "关闭该功能后，可能需要重新启动程序才能有效", vbOKOnly, "提示"
    End If
End Sub
