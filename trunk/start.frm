VERSION 5.00
Begin VB.Form start_ox163 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "OX163 starting page"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   Icon            =   "start.frx":0000
   LinkTopic       =   "start_ox163"
   Picture         =   "start.frx":406A
   ScaleHeight     =   3375
   ScaleWidth      =   6000
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Com5 
      Caption         =   "仅关闭本窗口"
      Height          =   420
      Left            =   4200
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Com3 
      Caption         =   $"start.frx":9B24
      Height          =   1095
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   120
      Top             =   960
   End
   Begin VB.CommandButton Com2 
      Caption         =   "关闭错误提示"
      Height          =   420
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Com1 
      Caption         =   "退出程序(QUIT)"
      Height          =   420
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox start_text 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   360
   End
End
Attribute VB_Name = "start_ox163"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ver_info = "53"

Private Sub Com1_Click()
End
End Sub

Private Sub Com2_Click()
WriteIniTF "maincenter", "err_report", False
Unload start_ox163
End Sub

Private Sub Com3_Click()
On Error Resume Next
Shell "OX163_SystemRecovery.exe"
End
End Sub



Private Sub Com5_Click()
Unload start_ox163
End Sub



'Private Sub Form_Initialize()
'InitCommonControls
'End Sub

Private Sub Form_Load()
Timer1.Interval = 100
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

Timer1.Enabled = False

Dim err_report As Boolean
Dim test_Object As Object

start_text.Text = ""

'------------------------------------------------------------------------------------------
start_text.Text = start_text.Text & vbCrLf & "检查msvbvm60.dll"

If Dir(GetSysDir & "\msvbvm60.dll") = "" Then
start_text.Text = start_text.Text & "msvbvm60.dll不存在"
FileCopy App.Path & "\msvbvm60.dll", GetSysDir & "\msvbvm60.dll"
ElseIf FileDateTime(GetSysDir & "\msvbvm60.dll") < FileDateTime(App.Path & "\msvbvm60.dll") Then
start_text.Text = start_text.Text & "msvbvm60.dll版本低"
FileCopy App.Path & "\msvbvm60.dll", GetSysDir & "\msvbvm60.dll"
End If
If Err.Number <> 0 Then
start_text.Text = start_text.Text & vbCrLf & "错误" & Err.Number & "：" & Err.Description
Err.Number = 0
Else
start_text.Text = start_text.Text & "...OK"
End If

start_text.SelStart = Len(start_text.Text)
'------------------------------------------------------------------------------------------

start_text.Text = start_text.Text & vbCrLf & "检查ole32.dll"

If Dir(GetSysDir & "\ole32.dll") = "" Then
start_text.Text = start_text.Text & "ole32.dll不存在"
FileCopy App.Path & "\ole32.dll", GetSysDir & "\ole32.dll"
ElseIf FileDateTime(GetSysDir & "\ole32.dll") < FileDateTime(App.Path & "\ole32.dll") Then
start_text.Text = start_text.Text & "ole32.dll版本低"
End If
If Err.Number <> 0 Then
start_text.Text = start_text.Text & vbCrLf & "错误" & Err.Number & "：" & Err.Description
Err.Number = 0
Else
start_text.Text = start_text.Text & "...OK"
End If

start_text.SelStart = Len(start_text.Text)
'------------------------------------------------------------------------------------------

start_text.Text = start_text.Text & vbCrLf & "检查COMDLG32.OCX" & vbCrLf & "创建CommonDialog"
Set test_Object = CreateObject("MSComDlg.CommonDialog")
If Err.Number <> 0 Then
start_text.Text = start_text.Text & vbCrLf & "错误" & Err.Number & "：" & Err.Description
Else
start_text.Text = start_text.Text & "...OK"
End If

Set test_Object = Nothing
Err.Number = 0

start_text.SelStart = Len(start_text.Text)
'------------------------------------------------------------------------------------------

start_text.Text = start_text.Text & vbCrLf & "检查MSINET.OCX" & vbCrLf & "创建Inet"
Set test_Object = CreateObject("InetCtls.Inet")
If Err.Number <> 0 Then
start_text.Text = start_text.Text & vbCrLf & "错误" & Err.Number & "：" & Err.Description
Else
start_text.Text = start_text.Text & "...OK"
End If

Set test_Object = Nothing
Err.Number = 0

start_text.SelStart = Len(start_text.Text)

'------------------------------------------------------------------------------------------
'
'start_text.Text = start_text.Text &   vbCrLf & "检查shdocvw.dll" & vbCrLf & "创建WebBrowser"
'Set test_Object = CreateObject("SHDocVwCtl.WebBrowser_V1")
'
'If Err.Number <> 0 Then
'start_text.Text = start_text.Text &   vbCrLf & "错误" & Err.Number & "：" & Err.Description
'Else
'start_text.Text = start_text.Text & "...OK"
'End If
'
'Set test_Object = Nothing
'Err.Number = 0
'
'start_text.SelStart = Len(start_text.Text)
'------------------------------------------------------------------------------------------
start_text.Text = start_text.Text & vbCrLf & "检查MSCOMCTL.OCX" & vbCrLf & "创建ListView"
Set test_Object = CreateObject("MSComctlLib.ListViewctrl")
If Err.Number <> 0 Then
start_text.Text = start_text.Text & vbCrLf & "错误" & Err.Number & "：" & Err.Description
Else
start_text.Text = start_text.Text & "...OK"
End If

Set test_Object = Nothing
Err.Number = 0

start_text.SelStart = Len(start_text.Text)

'------------------------------------------------------------------------------------------
start_text.Text = start_text.Text & vbCrLf & "检查wininet.dll" & vbCrLf & "创建InternetGetCookie"

GetCookie "http://www.163.com"
If Err.Number <> 0 Then
start_text.Text = start_text.Text & vbCrLf & "错误" & Err.Number & "：" & Err.Description
Else
start_text.Text = start_text.Text & "...OK"
End If

Err.Number = 0

start_text.SelStart = Len(start_text.Text)

'------------------------------------------------------------------------------------------
start_text.Text = start_text.Text & vbCrLf & "检查msscript.ocx" & vbCrLf & "创建ScriptControl"
Set test_Object = CreateObject("MSScriptControl.ScriptControl")
If Err.Number <> 0 Then
start_text.Text = start_text.Text & vbCrLf & "错误" & Err.Number & "：" & Err.Description
Else
start_text.Text = start_text.Text & "...OK"
End If

Set test_Object = Nothing
Err.Number = 0

start_text.SelStart = Len(start_text.Text)
'------------------------------------------------------------------------------------------
start_text.Text = start_text.Text & vbCrLf & "检查scrrun.dll" & vbCrLf & "创建FileSystemObject"
Set test_Object = CreateObject("Scripting.FileSystemObject")
If Err.Number <> 0 Then
start_text.Text = start_text.Text & vbCrLf & "错误" & Err.Number & "：" & Err.Description
Else
start_text.Text = start_text.Text & "...OK"
End If

Set test_Object = Nothing
Err.Number = 0

start_text.SelStart = Len(start_text.Text)

'------------------------------------------------------------------------------------------
start_text.Text = start_text.Text & vbCrLf & "检查文件夹"
If Dir(App.Path & "\download", vbDirectory) = "" Then
MkDir App.Path & "\download"
End If

If Dir(App.Path & "\url", vbDirectory) = "" Then
MkDir App.Path & "\url"
End If

If Err.Number <> 0 Then
start_text.Text = start_text.Text & vbCrLf & "错误：" & Err.Description
Err.Number = 0
Else
start_text.Text = start_text.Text & "...OK"
End If

start_text.SelStart = Len(start_text.Text)

'------------------------------------------------------------------------------------------
start_text.Text = start_text.Text & vbCrLf & "检查设定文档"
If Dir(App.Path & "\OX163setup.ini") = "" Then

    '默认参数
    WriteIniStr "maincenter", "ver", ver_info '默认参数
    
    WriteIniStr "maincenter", "downloadblock", "5120"
    WriteIniStr "maincenter", "time_out", "30"
    WriteIniStr "maincenter", "retry_times", "5"
    WriteIniStr "maincenter", "list_type", "1"
    WriteIniStr "maincenter", "fix_rar", "1"
    WriteIniStr "maincenter", "fix_rar_name", "RAR|ZIP|7Z|PNG|BMP"
    WriteIniTF "maincenter", "def_path_tf", False
    WriteIniStr "maincenter", "def_path", ""
    WriteIniStr "maincenter", "include_script", "delay"
    
    WriteIniStr "maincenter", "new163passcode_user", "wehi"
    WriteIniStr "maincenter", "new163passcode_album", "1530930"
    WriteIniStr "maincenter", "new163passcode_pw", "asd"
    
    WriteIniTF "maincenter", "autocheck", True
    WriteIniTF "maincenter", "askquit", True
    WriteIniTF "maincenter", "listshow", False
    WriteIniTF "maincenter", "savedef", True
    WriteIniTF "maincenter", "openfloder", True
    WriteIniTF "maincenter", "change_psw", True
    WriteIniTF "maincenter", "always_top", True
    WriteIniTF "maincenter", "new_ie_win", True
    WriteIniTF "maincenter", "ox163_ie_win", True
    WriteIniTF "maincenter", "sysTray", True
    
    WriteIniTF "maincenter", "new163pass_rules", True
    
    WriteIniTF "maincenter", "list_copy", True
    
    WriteIniStr "maincenter", "file_compare", "1"
    
    WriteIniTF "maincenter", "err_report", True
    
    WriteIniTF "maincenter", "bottom_StatusBar", True
    
    WriteIniTF "maincenter", "check_all", True
    
    WriteIniTF "maincenter", "url_folder", False
    
    
    WriteIniStr "proxyset", "proxy_A", ""
    WriteIniStr "proxyset", "proxy_A_user", ""
    WriteIniStr "proxyset", "proxy_A_pw", ""
    WriteIniStr "proxyset", "proxy_B", ""
    WriteIniStr "proxyset", "proxy_B_user", ""
    WriteIniStr "proxyset", "proxy_B_pw", ""
    WriteIniStr "proxyset", "proxy_A_type", "icUseDefault"
    WriteIniStr "proxyset", "proxy_B_type", "icUseDefault"
    
    
End If


If Err.Number <> 0 Then
start_text.Text = start_text.Text & vbCrLf & "错误：" & Err.Description
Err.Number = 0
Else
start_text.Text = start_text.Text & "...OK"
End If

start_text.SelStart = Len(start_text.Text)

'------------------------------------------------------------------------------------------

start_text.Text = start_text.Text & vbCrLf & "读取参数"


WriteIniStr "maincenter", "ver", ver_info '默认参数

sysSet.ver = CInt(GetIniStr("maincenter", "ver"))

sysSet.downloadblock = CLng(GetIniStr("maincenter", "downloadblock"))
sysSet.time_out = CInt(GetIniStr("maincenter", "time_out"))
sysSet.retry_times = CInt(GetIniStr("maincenter", "retry_times"))

sysSet.list_type = CByte(GetIniStr("maincenter", "list_type"))
sysSet.fix_rar = CByte(GetIniStr("maincenter", "fix_rar"))
sysSet.fix_rar_name = Trim(GetIniStr("maincenter", "fix_rar_name"))

sysSet.include_script = GetIniStr("maincenter", "include_script")

sysSet.new163passcode_def(0) = GetIniStr("maincenter", "new163passcode_user")
sysSet.new163passcode_def(1) = GetIniStr("maincenter", "new163passcode_album")
sysSet.new163passcode_def(2) = GetIniStr("maincenter", "new163passcode_pw")

If sysSet.new163passcode_def(0) = "" Or sysSet.new163passcode_def(1) = "" Or sysSet.new163passcode_def(2) = "" Then
sysSet.new163passcode_def(0) = "wehi"
sysSet.new163passcode_def(1) = "1530930"
sysSet.new163passcode_def(2) = "asd"
End If

sysSet.autocheck = GetIniTF("maincenter", "autocheck")
sysSet.askquit = GetIniTF("maincenter", "askquit")
sysSet.listshow = GetIniTF("maincenter", "listshow")
sysSet.savedef = GetIniTF("maincenter", "savedef")
sysSet.openfloder = GetIniTF("maincenter", "openfloder")
sysSet.change_psw = GetIniTF("maincenter", "change_psw")
sysSet.always_top = GetIniTF("maincenter", "always_top")
sysSet.new_ie_win = GetIniTF("maincenter", "new_ie_win")
sysSet.ox163_ie_win = GetIniTF("maincenter", "ox163_ie_win")
sysSet.sysTray = GetIniTF("maincenter", "sysTray")

sysSet.new163pass_rules = GetIniTF("maincenter", "new163pass_rules")

sysSet.list_copy = GetIniTF("maincenter", "list_copy")

sysSet.file_compare = CInt(GetIniStr("maincenter", "file_compare"))

err_report = GetIniTF("maincenter", "err_report")

sysSet.def_path_tf = GetIniTF("maincenter", "def_path_tf")

sysSet.bottom_StatusBar = GetIniTF("maincenter", "bottom_StatusBar")

sysSet.check_all = GetIniTF("maincenter", "check_all")

sysSet.url_folder = GetIniTF("maincenter", "url_folder")

sysSet.proxy_A = GetIniStr("proxyset", "proxy_A_type")
Select Case sysSet.proxy_A
Case "icDirect"
sysSet.proxy_A_type = 1
Case "icNamedProxy"
sysSet.proxy_A_type = 2
Case Else
sysSet.proxy_A_type = 0
End Select

sysSet.proxy_A = GetIniStr("proxyset", "proxy_B_type")
Select Case sysSet.proxy_A
Case "icDirect"
sysSet.proxy_B_type = 1
Case "icNamedProxy"
sysSet.proxy_B_type = 2
Case Else
sysSet.proxy_B_type = 0
End Select

sysSet.proxy_A = Trim(GetIniStr("proxyset", "proxy_A"))
sysSet.proxy_A_user = Trim(GetIniStr("proxyset", "proxy_A_user"))
sysSet.proxy_A_pw = GetIniStr("proxyset", "proxy_A_pw")
sysSet.proxy_B = Trim(GetIniStr("proxyset", "proxy_B"))
sysSet.proxy_B_user = Trim(GetIniStr("proxyset", "proxy_B_user"))
sysSet.proxy_B_pw = GetIniStr("proxyset", "proxy_B_pw")


'------------------------------------------------------------------------------------------
If Err.Number <> 0 Then
start_text.Text = start_text.Text & vbCrLf & "错误：" & Err.Description
Err.Number = 0
Else
start_text.Text = start_text.Text & "...OK"
End If

start_text.SelStart = Len(start_text.Text)

start_text.Text = start_text.Text & vbCrLf & "检查下载路径"
If sysSet.def_path_tf = True Then
    sysSet.def_path = GetIniStr("maincenter", "def_path")
        If Mid$(sysSet.def_path, 2, 2) <> ":\" Then GoTo reset_path
        If Right(sysSet.def_path, 1) = "\" Then sysSet.def_path = Mid$(sysSet.def_path, 1, Len(sysSet.def_path) - 1): WriteIniStr "maincenter", "def_path", sysSet.def_path
    Dim check_path
    check_path = Split(sysSet.def_path, "\")
    
        For i = 0 To UBound(check_path)
            If i > 0 Then
            sysSet.def_path = sysSet.def_path & "\" & check_path(i)
                If Dir(sysSet.def_path, vbDirectory) = "" Then
                MkDir sysSet.def_path
                End If
            Else
            sysSet.def_path = check_path(0)
            End If
        Next i
If (GetFileAttributes(sysSet.def_path) = -1) Then GoTo reset_path
Else
reset_path:
    If sysSet.def_path <> "" Then sysSet.def_path = "": WriteIniStr "maincenter", "def_path", ""
End If

'------------------------------------------------------------------------------------------
If Err.Number <> 0 Then
start_text.Text = start_text.Text & vbCrLf & "错误：" & Err.Description
Err.Number = 0
Else
start_text.Text = start_text.Text & "...OK"
End If

start_text.SelStart = Len(start_text.Text)

If InStr(start_text.Text, "错误：") > 0 And err_report = True Then
start_text.Text = start_text.Text & vbCrLf & vbCrLf & "有错误发生，可以点击上方'X (QUIT)'按钮关闭"
Com1.Visible = True
Com2.Visible = True
Com3.Visible = True
Com5.Visible = True
Else
start_text.Text = start_text.Text & vbCrLf & vbCrLf & "一切就绪,启动主程序" & vbCrLf & "如遇问题修复按钮15秒启动" & vbCrLf & "请确认网络已经连接"
End If
start_text.SelStart = Len(start_text.Text)
Timer2.Interval = 15000
Timer2.Enabled = True
Form1.Show
start_text.Enabled = True
End Sub

Private Sub Timer2_Timer()
Timer2.Interval = 0
Timer2.Enabled = False
Com1.Visible = True
Com2.Visible = True
Com3.Visible = True
Com5.Visible = True
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
End
End Sub
