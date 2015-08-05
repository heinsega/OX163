VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Ctrl8dot3name 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "短路径8dot3name设置"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7140
   Icon            =   "Ctrl8dot3name.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7140
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList DRV_Image 
      Left            =   360
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":406A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":4B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":506C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":55C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":5B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":6065
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":65BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":6B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":7050
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":7597
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":7AE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":802B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ctrl8dot3name.frx":8579
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   2040
      ScaleHeight     =   5055
      ScaleWidth      =   5055
      TabIndex        =   1
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton Set_all1 
         Caption         =   "每磁盘单独设置8dot3name功能(系统默认)"
         Height          =   975
         Left            =   2520
         TabIndex        =   4
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox drvText 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "Ctrl8dot3name.frx":8ABB
         Top             =   120
         Width           =   4935
      End
      Begin VB.CommandButton Set_all 
         Caption         =   "全局开启8dot3name功能"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   $"Ctrl8dot3name.frx":8AC1
         ForeColor       =   &H000000FF&
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   3120
         Width           =   4815
      End
   End
   Begin MSComctlLib.TreeView DRV_Menu 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   8705
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      Scroll          =   0   'False
      ImageList       =   "DRV_Image"
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "Ctrl8dot3name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GetShortPathName(8dot3name) Console
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Dim drvNum As Integer

Private Sub Build_DRV_Menu()
    'On Error Resume Next
    Dim strSave As String
    Dim drvName As String
    Dim drvSplit
    Dim DriveType As Long
    'Set the graphic mode to persistent
    Me.AutoRedraw = True
    'Create a buffer to store all the drives
    strSave = String(255, Chr(0))
    'Get all the drives
    ret& = GetLogicalDriveStrings(255, strSave)
    'Extract the drives from the buffer and print them on the form
    Do While Right(strSave, 1) = Chr(0)
        strSave = Left(strSave, Len(strSave) - 1)
    Loop
    
    drvSplit = Split(strSave, Chr(0))
    drvNum = UBound(drvSplit)
    
    Call DRV_Menu.Nodes.Add(, 4, "DRV", "磁盘全局设置", 1, 1 + 7)
    
    For keer = 0 To drvNum
        
        drvName = drvSplit(keer)
        DriveType = GetDriveType(drvName)
        If DriveType < 2 Or DriveType > 6 Then DriveType = 7
        '0 不明
        '2 软驱
        '3 硬盘
        '4 网络盘
        '5 光驱
        '6 RamDisk
        Call DRV_Menu.Nodes.Add("DRV", 4, , drvName, DriveType, DriveType + 7)
        
    Next
    
    Dim nodx As Node
    For Each nodx In DRV_Menu.Nodes
        nodx.Expanded = True
    Next
    DRV_Menu.Nodes(1).Selected = True
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
    
    DRV_Menu_1
End Sub

Private Sub DRV_Menu_1()
    Select Case OX_8dot3Name_Sys()
    Case 0
        drvText = vbCrLf & "系统已对所有磁盘启用了8dot3name短路径功能" & vbCrLf & vbCrLf & "无需设置"
    Case 1
        drvText = vbCrLf & "系统已对所有磁盘禁用了8dot3name短路径功能" & vbCrLf & vbCrLf & "必须设置,才能保证程序正常下载含有Unicode字符的文件"
    Case 2
        drvText = vbCrLf & "系统已对每个盘符单独设置了8dot3name短路径功能" & vbCrLf & vbCrLf & "建议开启全局设置, 或者单独对磁盘设置该功能"
    Case 3
        drvText = vbCrLf & "除系统盘外全部磁盘均禁用了8dot3name短路径功能" & vbCrLf & vbCrLf & "建议开启全局设置, 或者单独对磁盘设置该功能"
    End Select
    drvText = drvText & vbCrLf & vbCrLf & "程序目录regfile文件夹下对应设置reg文件:" & vbCrLf & vbCrLf & "OX163_Unicode_Support(ForceOpen_ShortPathName_on_Win8aboveOS).reg" & vbCrLf & "全部启动8dot3name短路径功能" & vbCrLf & vbCrLf & "OX163_Unicode_Support(Default_ShortPathName_on_Win8aboveOS).reg" & vbCrLf & "恢复系统默认"
    '0（全部启动），1（全部禁用），2（每个盘符单独设置），3（除系统盘外全部禁用）。
End Sub

Private Sub DRV_Menu_NodeClick(ByVal Node As MSComctlLib.Node)
    Static nodeID As Long
    Dim temp_str As String, temp_str1 As String, ver As String
    If nodeID = Node.Index And Node.Index <> 1 Then Exit Sub
    nodeID = Node.Index
    If Node.Index = 1 Then
        DRV_Menu_1
    Else
        If Node.Image = 5 Then drvText = Node.Text & vbCrLf & "光驱不支持写入": Exit Sub
        Ctrl8dot3name.Enabled = False
        drvText = "正在检测" & Node.Text & "8dot3name设置..."
        temp_str = OX_8dot3Name_Dir(Node.Text)
        ver = Node.Text & vbCrLf & Mid(temp_str, InStr(temp_str, vbCrLf) + Len(vbCrLf))
        temp_str = Mid(temp_str, 1, InStr(temp_str, vbCrLf) - 1)
        temp_str1 = OX_8dot3Name_Sys
        
        If temp_str1 = 1 Or (temp_str = 1 And temp_str1 = 2) Or (temp_str1 = 3 And Left(App.Path, 2) <> Left(GetSysDir, 2)) Then
            ver = ver & vbCrLf & vbCrLf & "磁盘未启用8dot3name短路径功能：程序无法在该磁盘操作特殊unicode字符" & vbCrLf & "您可以使用""fsutil 8dot3name set 0""命令启用全局用8dot3name短路径" & vbCrLf & "您可以使用""fsutil 8dot3name set 0 " & Node.Text & """命令单独打开该磁盘8dot3name短路径"
        End If
        drvText = ver
        Ctrl8dot3name.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Build_DRV_Menu
    If sysSet.always_top = True Then Sys_on_top
End Sub

Private Sub Sys_on_top()
    Dim flags As Integer
    flags = SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags
End Sub

Private Sub Set_all_Click()
    On Error Resume Next
    If MsgBox("此方法会调用\regfile\文件夹下的OX163_Unicode_Support(ForceOpen_ShortPathName_on_Win8aboveOS).reg文件" & vbCrLf & "调用后会出现添加注册表信息的对话框,请自己阅读并选择", vbOKCancel, "提醒") = vbCancel Then Exit Sub
    err.Clear
    Shell "regedit " & App_path & "\regfile\OX163_Unicode_Support(ForceOpen_ShortPathName_on_Win8aboveOS).reg", vbNormalFocus
    If err.Number <> 0 Then MsgBox "错误:" & err.Number & vbCrLf & err.Descriptionr & vbCrLf & "您可以打开regfile目录直接操作", vbOKOnly, "提醒"
    err.Clear
End Sub
Private Sub Set_all1_Click()
    On Error Resume Next
    If MsgBox("此方法会调用\regfile\文件夹下的OX163_Unicode_Support(Default_ShortPathName_on_Win8aboveOS).reg文件" & vbCrLf & "调用后会出现添加注册表信息的对话框,请自己阅读并选择", vbOKCancel, "提醒") = vbCancel Then Exit Sub
    err.Clear
    Shell "regedit " & App_path & "\regfile\OX163_Unicode_Support(Default_ShortPathName_on_Win8aboveOS).reg", vbNormalFocus
    If err.Number <> 0 Then MsgBox "错误:" & err.Number & vbCrLf & err.Descriptionr & vbCrLf & "您可以打开regfile目录直接操作", vbOKOnly, "提醒"
    err.Clear
End Sub
