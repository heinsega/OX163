VERSION 5.00
Begin VB.Form start_ox163 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "starting OX163"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   Icon            =   "start.frx":0000
   LinkTopic       =   "start_ox163"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "start.frx":406A
   ScaleHeight     =   3390
   ScaleWidth      =   6000
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3390
      Left            =   0
      Picture         =   "start.frx":91F1
      ScaleHeight     =   3390
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Left            =   0
         Top             =   1080
      End
      Begin VB.Timer Timer1 
         Left            =   0
         Top             =   720
      End
      Begin VB.TextBox start_text 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   1335
         Left            =   2760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2040
         Width           =   3135
      End
      Begin VB.CommandButton Com1 
         Caption         =   "退出程序(QUIT)"
         Height          =   420
         Left            =   4200
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Com5 
         Caption         =   "仅关闭本窗口"
         Height          =   420
         Left            =   2160
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
   End
End
Attribute VB_Name = "start_ox163"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Const WS_EX_LAYERED = &H80000
'Private Const GWL_EXSTYLE = (-20)
'Private Const LWA_ALPHA = &H2
'Private Const LWA_COLORKEY = &H1
'Private Sub Start_Form_alph()
'    BorderStyler = 0
'    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
'    a = Me.hwnd
'    rtn = rtn Or WS_EX_LAYERED
'    SetWindowLong hwnd, GWL_EXSTYLE, rtn
'    SetLayeredWindowAttributes hwnd, &HFFFFFF, 100, LWA_COLORKEY
'End Sub

Private Sub Com1_Click()
    End
End Sub

Private Sub Com5_Click()
    Unload start_ox163
End Sub

Private Sub Form_Load()
    Picture1.PaintPicture Me.Picture, 0, 0, Picture1.Width, Picture1.Height
    Timer1.Interval = 280
    Timer1.Enabled = True
    OX_G_Disable8dot3 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Timer1.Interval = 0
End Sub

Private Sub start_text_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 65 And Shift = vbCtrlMask Then
        start_text.SelStart = 0
        start_text.SelLength = Len(start_text.Text)
    End If
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    
    Static timer1_counter As Boolean
    
    Timer1.Enabled = False
    Timer1.Interval = 0
    
    If timer1_counter = False Then
        timer1_counter = True
    Else
        Exit Sub
    End If
    
    'App.PrevInstance 检查是否多开程序 F没有， T多开
    'Dim logfile As String
    'logfile = App.Path & "\OX163-" & Format(Now(), "YYYY.MM.DD-HH.MM.SS") & ".log"
    'Open logfile For Binary Access Write As #1
    'Close #1
    'App.StartLogging App.Path & logfile, vbLogAuto
    'App.LogEvent "engilsh", 4
    
    Dim test_Object As Object
    Dim start_check1, start_check2
    Dim step_counter As Integer
    Dim first_sart_up As Boolean
    
    
    '判断Non Unicode程序设置,并提示
    If GetOSLCID <> 1 Then MsgBox "Your System Lanuages for Non Unicode Program is not Simplified Chinese." _
        & vbCrLf & "If you want to get a better experience.(No distortion No unknow Error & etc.)" _
        & vbCrLf & "You should open" _
        & vbCrLf & "'Control Panel'->'Region and Language'->'Administrative'" _
        & vbCrLf & "Change 'language for non-Unicode programs' to 'Chinese(Simplified, PRC)'." _
        & vbCrLf & "When you have finished setting, you need to restart your computer.", vbOKOnly
    '--------------------------------------------------------
    start_text.Text = "启动程序:" & vbCrLf & "0.5.9以上版本默认以管理员身份启动,以应对win8.1以上系统的诸多问题"
    step_counter = 0
    err.Clear
    '-----------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "检查msvbvm60.dll"
    start_check1 = ""
    start_check2 = ""
    
    start_check1 = FileDateTime(GetSysDir & "\..\system32\msvbvm60.dll")
    If err.Number <> 0 Then
        start_text.Text = start_text.Text & vbCrLf & GetSysDir & "\..\system32\msvbvm60.dll" & "：" & err.Description
        err.Clear
    Else
        start_text.Text = start_text.Text & "...OK"
    End If
    
    start_check2 = FileDateTime(GetSysDir & "\..\sysWOW64\msvbvm60.dll")
    If err.Number <> 0 Then
        start_text.Text = start_text.Text & vbCrLf & GetSysDir & "\..\sysWOW64\msvbvm60.dll" & "：" & err.Description
        err.Clear
    Else
        start_text.Text = start_text.Text & "...OK"
    End If
    
    If start_check1 = "" And start_check2 = "" Then
        start_text.Text = start_text.Text & "系统文件夹中msvbvm60.dll可能不存在(一般情况不影响程序使用)"
    End If
    
    start_text.SelStart = Len(start_text.Text)
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "%delay_NtfsDisable8dot3Name_Checking%"
    
    start_text.SelStart = Len(start_text.Text)
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "检查scrrun.dll" & vbCrLf & "创建FileSystemObject"
    err.Clear
    Set test_Object = CreateObject("Scripting.FileSystemObject")
    If err.Number <> 0 Then
        start_text.Text = start_text.Text & vbCrLf & "Error-" & err.Number & ": " & err.Description
        start_text.Text = start_text.Text & vbCrLf & "无法创建FileSystemObject：程序可能无法操作特殊unicode字符" & vbCrLf & "您可能需要修复windows系统文件：scrrun.dll，并重新设定FileSystemObject权限"
        App_path = App.Path
    Else
        start_text.Text = start_text.Text & "...OK"
        '规格化程序所在路径的短路径名
        start_check1 = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
        App_path = test_Object.GetAbsolutePathName("")
        App_path = IIf(Right(App_path, 1) = "\", App_path, App_path & "\")
        App_path = IIf((InStr(start_check1, Chr(63)) < 1 And App_path <> start_check1), start_check1, App_path)
        App_path = GetShortName(App_path)
        start_text.Text = start_text.Text & vbCrLf & "程序主目录-> " & App.Path
        start_text.Text = start_text.Text & vbCrLf & "主目录启用值-> " & App_path
    End If
    
    Set test_Object = Nothing
    
    start_text.SelStart = Len(start_text.Text)
    '------------------------------------------------------------------------------------------
    
    first_sart_up = False
    If OX_Dirfile(App_path & "\" & App.EXEName & ".exe.manifest") = False Or OX_Dirfile(App_path & "\OX163setup.ini") = False Then first_sart_up = True
    
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "检查msinet.ocx" & vbCrLf & "创建Inet控件"
    err.Clear
    Set test_Object = CreateObject("InetCtls.Inet.1")
    If err.Number <> 0 Then
        start_text.Text = start_text.Text & vbCrLf & "Error-" & err.Number & ": " & err.Description
        start_text.Text = start_text.Text & vbCrLf & "无法创建创建Inet控件：程序可能无法下载网页与图片" & vbCrLf & "您可能需要修复windows系统文件：msinet.ocx (32位)"
    Else
        start_text.Text = start_text.Text & "...OK"
    End If
    
    Set test_Object = Nothing
    
    start_text.SelStart = Len(start_text.Text)
    
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "检查wininet.dll" & vbCrLf & "创建InternetGetCookie"
    
    err.Clear
    GetCookie "http://www.163.com"
    If err.Number <> 0 Then
        start_text.Text = start_text.Text & vbCrLf & "Error-" & err.Number & ": " & err.Description
        start_text.Text = start_text.Text & vbCrLf & "无法创建创建wininet应用：程序可能无法获取页面cookies甚至Inet控件将失效" & vbCrLf & "您可能需要修复windows系统文件：wininet.dll"
    Else
        start_text.Text = start_text.Text & "...OK"
    End If
    
    start_text.SelStart = Len(start_text.Text)
    
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "检查comdlg32.dll" & vbCrLf & "创建CommonDialog"
    err.Clear
    Set test_Object = CreateObject("MSComDlg.CommonDialog.1")
    If err.Number <> 0 Then
        start_text.Text = start_text.Text & vbCrLf & "Error-" & err.Number & ": " & err.Description
        start_text.Text = start_text.Text & vbCrLf & "无法创建创建CommonDialog对话框：程序可能无法创建文件保存或选择窗口" & vbCrLf & "您可能需要修复windows系统文件：comdlg32.dll"
    Else
        start_text.Text = start_text.Text & "...OK"
    End If
    
    Set test_Object = Nothing
    
    start_text.SelStart = Len(start_text.Text)
    
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "检查mscomctl.ocx" & vbCrLf & "创建ListView"
    err.Clear
    Set test_Object = CreateObject("MSComctlLib.ListViewctrl")
    If err.Number <> 0 Then
        start_text.Text = start_text.Text & vbCrLf & "Error-" & err.Number & ": " & err.Description
        start_text.Text = start_text.Text & vbCrLf & "无法创建创建ListView列表：程序可能无法创建下载列表" & vbCrLf & "您可能需要修复windows系统文件：mscomctl.ocx (32位)"
    Else
        start_text.Text = start_text.Text & "...OK"
    End If
    
    Set test_Object = Nothing
    
    start_text.SelStart = Len(start_text.Text)
    
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "检查msscript.ocx" & vbCrLf & "创建ScriptControl"
    err.Clear
    Set test_Object = CreateObject("MSScriptControl.ScriptControl")
    If err.Number <> 0 Then
        start_text.Text = start_text.Text & vbCrLf & "Error-" & err.Number & ": " & err.Description
        start_text.Text = start_text.Text & vbCrLf & "无法创建创建ScriptControl：程序可能无法调用外部脚本" & vbCrLf & "您可能需要修复windows系统文件：msscript.ocx (32位)"
    Else
        start_text.Text = start_text.Text & "...OK"
    End If
    
    Set test_Object = Nothing
    
    start_text.SelStart = Len(start_text.Text)
    
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "检查msado15.dll" & vbCrLf & "创建ADODB.Stream"
    err.Clear
    Set test_Object = CreateObject("ADODB.Stream")
    If err.Number <> 0 Then
        start_text.Text = start_text.Text & vbCrLf & "Error-" & err.Number & ": " & err.Description
        start_text.Text = start_text.Text & vbCrLf & "无法创建创建ADODB.Stream：程序可能无法正确识别文本字符集和加载外部脚本" & vbCrLf & "您可能需要修复Program Files\Common Files\System\ado\中的msado15.dll"
    Else
        start_text.Text = start_text.Text & "...OK"
    End If
    
    Set test_Object = Nothing
    
    start_text.SelStart = Len(start_text.Text)
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "检查文件夹"
    err.Clear
    If Dir(App_path & "\download", vbDirectory) = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立download文件夹"
        MkDir App_path & "\download"
    End If
    
    If Dir(App_path & "\url", vbDirectory) = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立url文件夹"
        MkDir App_path & "\url"
    End If
    
    If Dir(App_path & "\include", vbDirectory) = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立include文件夹"
        MkDir App_path & "\include"
    End If
    
    If Dir(App_path & "\include\sys", vbDirectory) = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立include\sys文件夹"
        MkDir App_path & "\include\sys"
    End If
    
    If Dir(App_path & "\include\custom", vbDirectory) = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立include\custom文件夹"
        MkDir App_path & "\include\custom"
    End If
    
    If Dir(App_path & "\regfile", vbDirectory) = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立regfile文件夹"
        MkDir App_path & "\regfile"
    End If
    
    If err.Number <> 0 Then
        start_text.Text = start_text.Text & vbCrLf & "Error-" & err.Number & ": " & err.Description
        start_text.Text = start_text.Text & vbCrLf & "默认文件夹错误：程序检测或创建默认文件夹失败" & vbCrLf & "请手动检测程序目录下以下文件夹是否存在:"
        start_text.Text = start_text.Text & vbCrLf & "\url" & vbCrLf & "\download" & vbCrLf & "\include" & vbCrLf & "\include\sys" & vbCrLf & "\include\custom" & vbCrLf & "\regfile"
    Else
        start_text.Text = start_text.Text & "...OK"
    End If
    
    start_text.SelStart = Len(start_text.Text)
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "检查reg文件"
    err.Clear
    start_check1 = ""
    start_check2 = ""
    start_check1 = "Windows Registry Editor Version 5.00" & vbCrLf & _
                   "[HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION]" & vbCrLf & _
                   "#dword#" & vbCrLf & vbCrLf & _
                   "[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION]" & vbCrLf & _
                   "#dword#"
    
    If Dir(App_path & "\regfile\use_IE8.reg") = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立\regfile\use_IE8.reg文件"
        start_check2 = OX_GreatTxtFile(App_path & "\regfile\use_IE8.reg", Replace(start_check1, "#dword#", """OX163.exe""=dword:00001F40"), "unicode")
    End If
    
    If Dir(App_path & "\regfile\use_IE9.reg") = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立\regfile\use_IE9.reg文件"
        start_check2 = OX_GreatTxtFile(App_path & "\regfile\use_IE9.reg", Replace(start_check1, "#dword#", """OX163.exe""=dword:00002328"), "unicode")
    End If
    
    If Dir(App_path & "\regfile\use_IE10.reg") = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立\regfile\use_IE10.reg文件"
        start_check2 = OX_GreatTxtFile(App_path & "\regfile\use_IE10.reg", Replace(start_check1, "#dword#", """OX163.exe""=dword:00002710"), "unicode")
    End If
    
    If Dir(App_path & "\regfile\use_IE11.reg") = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立\regfile\use_IE11.reg文件"
        start_check2 = OX_GreatTxtFile(App_path & "\regfile\use_IE11.reg", Replace(start_check1, "#dword#", """OX163.exe""=dword:00002AF8"), "unicode")
    End If
    
    If Dir(App_path & "\regfile\clear_OX163.reg") = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立\regfile\clear_OX163.reg文件"
        start_check2 = OX_GreatTxtFile(App_path & "\regfile\clear_OX163.reg", Replace(start_check1, "#dword#", """OX163.exe""=-"), "unicode")
    End If
    
    start_check1 = "Windows Registry Editor Version 5.00" & vbCrLf & _
                   "[HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem]" & vbCrLf & _
                   "#dword#"
                   
    If Dir(App_path & "\regfile\OX163_Unicode_Support(ForceOpen_ShortPathName_on_Win8aboveOS).reg") = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立\regfile\OX163_Unicode_Support(ForceOpen_ShortPathName_on_Win8aboveOS).reg文件"
        start_check2 = OX_GreatTxtFile(App_path & "\regfile\OX163_Unicode_Support(ForceOpen_ShortPathName_on_Win8aboveOS).reg", Replace(start_check1, "#dword#", """NtfsDisable8dot3NameCreation""=dword:0"), "unicode")
    End If
                   
    If Dir(App_path & "\regfile\OX163_Unicode_Support(Default_ShortPathName_on_Win8aboveOS).reg") = "" Then
    start_text.Text = start_text.Text & vbCrLf & "建立\regfile\OX163_Unicode_Support(Default_ShortPathName_on_Win8aboveOS).reg文件"
        start_check2 = OX_GreatTxtFile(App_path & "\regfile\OX163_Unicode_Support(Default_ShortPathName_on_Win8aboveOS).reg", Replace(start_check1, "#dword#", """NtfsDisable8dot3NameCreation""=dword:2"), "unicode")
    End If
    
    If err.Number <> 0 Or LCase(start_check2) = "false" Then
        start_text.Text = start_text.Text & vbCrLf & "Error-" & err.Number & ": " & err.Description
    Else
        start_text.Text = start_text.Text & "...OK"
    End If
    
    start_check1 = ""
    start_check2 = ""
    start_text.SelStart = Len(start_text.Text)
    
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "初始化程序默认设置"
    err.Clear
    sysSet = OX_Default_Setting
    If err.Number <> 0 Then
        start_text.Text = start_text.Text & vbCrLf & "Error-" & err.Number & ": " & err.Description
        start_text.Text = start_text.Text & vbCrLf & "初始化程序默认设置错误"
    Else
        start_text.Text = start_text.Text & "...OK"
    End If
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "检查OX163setup.ini"
    err.Clear
    If Dir(App_path & "\OX163setup.ini") = "" Then
        start_text.Text = start_text.Text & vbCrLf & "OX163setup.ini不存在"
        start_text.Text = start_text.Text & vbCrLf & "重新建立OX163setup.ini"
        '默认参数
        start_check1 = 0
        start_check1 = OX_WriteIni_Setting(sysSet)
        If Int(start_check1) <> 0 Then
            start_text.Text = start_text.Text & vbCrLf & "Error-" & start_check1 & ": " & err.Description
            start_text.Text = start_text.Text & vbCrLf & "建立OX163setup.ini发生错误，可能建立失败"
        Else
            start_text.Text = start_text.Text & "...OK"
        End If
    End If
    
    If err.Number <> 0 Then
        start_text.Text = start_text.Text & vbCrLf & "Error-" & err.Number & ": " & err.Description
        start_text.Text = start_text.Text & vbCrLf & "检查或建立OX163setup.ini文件失败" & vbCrLf & "程序可能无法调用或保存用户个人设置"
        err.Clear
    Else
        start_text.Text = start_text.Text & "...OK"
    End If
    
    start_text.SelStart = Len(start_text.Text)
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "读取OX163setup.ini"
    err.Clear
        start_check1 = 0
        start_check1 = OX_GetIni_Setting(sysSet)
        If Int(start_check1) <> 0 Then
            start_text.Text = start_text.Text & vbCrLf & "Error-" & start_check1 & ": " & err.Description
            start_text.Text = start_text.Text & vbCrLf & "读取OX163setup.ini发生错误,可能需要开启程序设置重新写入ini"
        Else
            start_text.Text = start_text.Text & "...OK"
        End If
    
    start_text.SelStart = Len(start_text.Text)
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//step." & step_counter & "//"
    start_text.Text = start_text.Text & vbCrLf & "检测OX163.exe.manifest"
    err.Clear
        start_check1 = 1
        If OX_Dirfile(App_path & "\" & App.EXEName & ".exe.manifest") = False Then start_check1 = 0: start_check1 = Set_OX_manifest: Shell App_path & "\" & App.EXEName & ".exe", vbNormalFocus: End
        
        If Int(start_check1) = 0 Then
            start_text.Text = start_text.Text & vbCrLf & "创建" & App.EXEName & ".exe.manifest文件失败!"
        Else
            start_text.Text = start_text.Text & "...OK"
        End If
    
    start_text.SelStart = Len(start_text.Text)
    '------------------------------------------------------------------------------------------
    
    step_counter = step_counter + 1: start_text.Text = start_text.Text & vbCrLf & vbCrLf & "//启动结束//"
    
    
    
    If InStr(start_text.Text, "Error-") > 0 Then
        start_text.Text = start_text.Text & vbCrLf & vbCrLf & "有错误发生，可以点击上方'X (QUIT)'按钮关闭"
    Else
        start_text.Text = start_text.Text & vbCrLf & vbCrLf & "一切就绪,启动主程序,请确认网络已连接,修复按钮15秒后启动"
    End If
    start_text.Text = start_text.Text & vbCrLf & vbCrLf & "Vista Win7 Win8 Win10下无法启动,可对程序进行如下操作:" & vbCrLf & "右键 -> 以管理员身份运行程序"
    start_text.SelStart = Len(start_text.Text)
    start_text.Enabled = True
    Timer2.Interval = 15000
    Timer2.Enabled = True
    BrowserW_url = ""
    BrowserW_load_ok = True
    windows_destop_Width = Screen.Width 'start_ox163.Width + start_ox163.Left * 2
    windows_destop_Height = Screen.Height 'start_ox163.Height + start_ox163.Top * 2
    OX_Start_log = start_text
    Load History_Logs
    'History_Logs.Hide
    'MsgBox "Screen.TwipsPerPixelX" & Screen.TwipsPerPixelX & vbCrLf & "Screen.TwipsPerPixelY" & Screen.TwipsPerPixelY
    Form1.Show
End Sub

Private Sub Timer2_Timer()
    Timer2.Interval = 0
    Timer2.Enabled = False
    Com1.Visible = True
    Com5.Visible = True
End Sub
