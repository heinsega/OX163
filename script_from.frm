VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form script_from 
   Caption         =   "OX163 Script Setting"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   Icon            =   "script_from.frx":0000
   LinkTopic       =   "script_from"
   ScaleHeight     =   4905
   ScaleWidth      =   7065
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   1800
   End
   Begin InetCtlsObjects.Inet script_load 
      Left            =   0
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin MSComctlLib.ImageList Image_over 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   14933984
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "script_from.frx":406A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "script_from.frx":40E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "script_from.frx":416D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "script_from.frx":41F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "script_from.frx":427F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Image_normal 
      Left            =   600
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   14933984
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "script_from.frx":4307
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "script_from.frx":437C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "script_from.frx":4400
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "script_from.frx":447E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "script_from.frx":4502
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   1349
      ButtonWidth     =   1455
      ButtonHeight    =   1296
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "Image_normal"
      HotImageList    =   "Image_over"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "自动更新"
            Key             =   "Update"
            Description     =   "Auto Update Script"
            Object.ToolTipText     =   "Auto Update Script"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Update1"
                  Text            =   "更新全部的脚本(&All)"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Update2"
                  Text            =   "更新选中的脚本(&Checked)"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Update3"
                  Text            =   "更新缺少的脚本(&Lack)"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "重新检查"
            Key             =   "Check_html"
            Description     =   "Check again"
            Object.ToolTipText     =   "Check again"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "停止所有"
            Key             =   "stop_script"
            Description     =   "Stop All"
            Object.ToolTipText     =   "Stop All"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "脚本主页"
            Key             =   "View_web"
            Description     =   "View Homepage"
            Object.ToolTipText     =   "View Homepage"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "脚本目录"
            Key             =   "Browse_Folder"
            Description     =   "Browse Folder"
            Object.ToolTipText     =   "Browse Folder"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   4665
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   25400
            MinWidth        =   25400
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView script_list 
      Height          =   3420
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Shift or Ctrl to MultiSelect"
      Top             =   720
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6033
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
      Icons           =   "ImageList1"
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "list_ID"
         Object.Tag             =   "sc_ID"
         Text            =   "序号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "script_name"
         Object.Tag             =   "sc_name"
         Text            =   "脚本名称"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "update_time"
         Object.Tag             =   "sc_time"
         Text            =   "更新时间"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "local_info"
         Object.Tag             =   "sc_local"
         Text            =   "本地情况"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "auto_update"
         Object.Tag             =   "sc_update"
         Text            =   "建议更新"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "script_from"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim script_txt As String
Dim script_txt_byte() As Byte
Dim script_down_ok As Boolean
Dim script_quit As Boolean
Dim strURL As String
Dim htmlCharsetType As String
Dim script_update_txt As String
Dim script_include As String
Dim local_include As String

Public Sub on_top(on_top As Boolean)
    Dim flags As Integer
    flags = SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
    If on_top = True Then
        SetWindowPos script_from.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags
    Else
        SetWindowPos script_from.hWnd, -2, 0, 0, 0, 0, flags
    End If
End Sub


Private Sub Form_Load()
    On Error Resume Next
    on_top sysSet.always_top
    Toolbar.Buttons(1).Enabled = False
    Toolbar.Buttons(2).Enabled = False
    Form_Resize
    script_quit = True
    script_down_ok = True
    htmlCharsetType = "GB2312"
    script_from.caption = script_from.caption & " (" & sysSet.update_host & ")"
    Timer1.Enabled = True
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'Static max_size As Boolean
    If script_from.WindowState <> 1 Then
        If script_from.Width < 5000 Then script_from.Width = 5000
        If script_from.Height < 4000 Then script_from.Height = 4000
        script_list.Width = script_from.ScaleWidth
        script_list.Top = Toolbar.Top + Toolbar.Height + 50
        script_list.Left = Toolbar.Left
        script_list.Height = script_from.ScaleHeight - script_list.Top - StatusBar.Height
        script_list.ColumnHeaders.Item(3).Width = (script_list.Width - 2800) / 2
        script_list.ColumnHeaders.Item(4).Width = (script_list.Width - 2800) / 2 - 500
        StatusBar.Panels(2).Width = script_list.Width
        'frame_resize
    End If
End Sub


Private Sub Check_script()
    Dim update_split
    Dim include_split
    Dim def_thing
    Dim file_time
    
    StatusBar.Panels(2) = "Now Checking"
    script_load.Cancel
    script_down_ok = False
    strURL = Trim$(sysSet.update_host & "script_update.vbs?ntime=" & OX_ntime(OX_ntime_Now, OX_ntime_Hex))
    
    script_download
    
    Do While script_down_ok = False
        If script_quit = True Then Exit Sub
        DoEvents
    Loop
    script_update_txt = script_txt
    
    script_load.Cancel
    script_down_ok = False
    
    update_split = Split(script_update_txt, vbCrLf)
    
    script_list.ListItems.Clear
    If script_quit = True Then Exit Sub
    
    For i = 0 To UBound(update_split)
        DoEvents
        
        def_thing = ""
        
        include_split = Split(update_split(i), "|")
        
        '序号
        script_list.ListItems.Add i + 1, , Format$(i + 1, "000")
        '脚本名称
        script_list.ListItems.Item(i + 1).ListSubItems.Add , , include_split(0)
        '更新时间
        script_list.ListItems.Item(i + 1).ListSubItems.Add , , include_split(1)
        
        If Dir(App_path & "\include\sys\" & include_split(0)) <> "" Then
            file_time = FileDateTime(App_path & "\include\sys\" & include_split(0))
            If DateDiff("s", include_split(1), file_time) < 0 Then
                def_thing = "10"
            ElseIf FileLen(App_path & "\include\sys\" & include_split(0)) <> include_split(2) Then
                def_thing = "10"
            Else
                def_thing = "11"
            End If
        Else
            def_thing = "00"
        End If
        '本地情况
        script_list.ListItems.Item(i + 1).ListSubItems.Add , , def_thing
        '建议更新
        script_list.ListItems.Item(i + 1).ListSubItems.Add , , ""
    Next i
    
    local_include = ""
    StatusBar.Panels(2) = ""
    file_time = 0
    
    '自动勾选需要更新的脚本
    For i = 0 To UBound(update_split)
        DoEvents
        script_list.ListItems(i + 1).ListSubItems(3).Text = static_str(script_list.ListItems(i + 1).ListSubItems(3).Text)
        If (script_list.ListItems(i + 1).ListSubItems(3).Text Like "?0*") Then
            script_list.ListItems(i + 1).ListSubItems(4).Text = "YES"
            script_list.ListItems(i + 1).Checked = True
            file_time = file_time + 1
        Else
            script_list.ListItems(i + 1).ListSubItems(4).Text = "NO"
        End If
    Next i
    StatusBar.Panels(1) = script_list.ListItems.count & "(Files) / " & file_time & "(Need Update)"
End Sub

Private Function static_str(ByVal str_temp)
    '000,该文件不存在/include.txt需要更新
    '001,该文件不存在/include.txt不需要更新
    '100,该文件需要更新/include.txt需要更新
    '101,该文件需要更新/include.txt不需要更新
    '110,该文件存在/include.txt需要更新
    '111,该文件存在/include.txt不需要更新
    If Len(str_temp) = 2 Then str_temp = str_temp & "0"
    Select Case str_temp
    Case "000"
        static_str = "000,该文件不存在/include.txt需要更新"
    Case "001"
        static_str = "001,该文件不存在/include.txt不需要更新"
    Case "100"
        static_str = "100,该文件需要更新/include.txt需要更新"
    Case "101"
        static_str = "101,该文件需要更新/include.txt不需要更新"
    Case "110"
        static_str = "110,该文件存在/include.txt需要更新"
    Case "111"
        static_str = "111,该文件存在/include.txt不需要更新"
    Case Else
        static_str = "情况不明"
    End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
    If script_quit = False And sysSet.askquit = True Then
        If MsgBox("正在执行操作，是否退出？", vbYesNo + vbDefaultButton2, "退出询问") = vbYes Then Cancel = True: Exit Sub
    End If
    Call load_in_Script_Code
    script_quit = True
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Timer1.Enabled = False
    
    Toolbar.Buttons(1).Enabled = False
    Toolbar.Buttons(2).Enabled = False
    Toolbar.Buttons(4).Enabled = True
    script_quit = False
    
    Check_script
    
    script_quit = True
    Toolbar.Buttons(1).Enabled = True
    Toolbar.Buttons(2).Enabled = True
    Toolbar.Buttons(4).Enabled = False
    
    If Form1.form_quit = False Then
        If MsgBox("正在执行下载，更新脚本可能会有潜在危险" & vbCrLf & "是否继续执行脚本更新？", vbYesNo + vbDefaultButton2 + vbExclamation, "警告：") = vbNo Then Unload script_from
    End If
    
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    
    If Dir(App_path & "\include", vbDirectory) = "" Then MkDir App_path & "\include"
    If Dir(App_path & "\include\sys", vbDirectory) = "" Then MkDir App_path & "\include\sys"
    
    Select Case Button.Index
    Case 1
        Update_script "auto"
    Case 2
        Timer1.Enabled = False
        
        Toolbar.Buttons(1).Enabled = False
        Toolbar.Buttons(2).Enabled = False
        Toolbar.Buttons(4).Enabled = True
        script_quit = False
        
        Check_script
        
        script_quit = True
        Toolbar.Buttons(1).Enabled = True
        Toolbar.Buttons(2).Enabled = True
        Toolbar.Buttons(4).Enabled = False
    Case 4
        script_quit = True
    Case 6
        ShellExecute 0&, vbNullString, StrConv(sysSet.update_host & "?key=3", vbUnicode), vbNullString, vbNullString, vbNormalFocus
    Case 7
        Shell "explorer.exe " & App_path & "\include\sys", vbNormalFocus
    End Select
End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error Resume Next
    Select Case ButtonMenu.Index
    Case 1
        Update_script "all"
        
    Case 2
        Update_script "checked"
        
    Case 3
        Update_script "lack"
    End Select
End Sub

Public Sub Update_script(update_type As String)
    On Error Resume Next
    
    Toolbar.Buttons(1).Enabled = False
    Toolbar.Buttons(2).Enabled = False
    Toolbar.Buttons(4).Enabled = True
    script_quit = False
    
    Dim include_txt As String
    Dim fs_file, fs_filename
    Dim no_include As Boolean
    
    include_txt = StatusBar.Panels(2)
    
    '--------------------------------------------------------
    StatusBar.Panels(2) = "Checking Script"
    '--------------------------------------------------------
    Select Case update_type
    Case "auto"
        For i = 1 To script_list.ListItems.count
            DoEvents
            If script_list.ListItems(i).ListSubItems(4).Text = "YES" Then
                script_list.ListItems(i).Checked = True
            Else
                script_list.ListItems(i).Checked = False
            End If
        Next i
        
    Case "all"
        For i = 1 To script_list.ListItems.count
            DoEvents
            script_list.ListItems(i).Checked = True
        Next i
        
    Case "checked"
        
    Case "lack"
        For i = 1 To script_list.ListItems.count
            DoEvents
            If (script_list.ListItems(i).ListSubItems(3).Text Like "0*") Then
                script_list.ListItems(i).Checked = True
            Else
                script_list.ListItems(i).Checked = False
            End If
        Next i
        
    End Select
    '--------------------------------------------------------
    StatusBar.Panels(2) = "Update Script"
    '--------------------------------------------------------
    For i = 1 To script_list.ListItems.count
        DoEvents
        If script_list.ListItems(i).Checked = True Then
            
            StatusBar.Panels(2) = "Update Script: " & script_list.ListItems(i).ListSubItems(1).Text
            
            script_load.Cancel
            script_down_ok = False
            strURL = Trim$(sysSet.update_host & script_list.ListItems(i).ListSubItems(1).Text & "?ntime=" & OX_ntime(OX_ntime_Now, OX_ntime_Hex))
            
            script_download
            
            Do While script_down_ok = False
                If script_quit = True Then Exit Sub
                DoEvents
            Loop
            
            fs_file = FreeFile
            fs_filename = App_path & "\include\sys\" & script_list.ListItems(i).ListSubItems(1).Text
            Kill fs_filename
            Open fs_filename For Binary Access Write As #fs_file
            Put #fs_file, , script_txt_byte
            Close #fs_file
            
            '写入指定格式的ansi文件
            'Call OX_GreatTxtFile(App_path & "\include\sys\" & script_list.ListItems(i).ListSubItems(1).Text, script_txt, htmlCharsetType)
            
            'FSO方式在非简体环境中会写入错误字段
            'Set fso = CreateObject("Scripting.FileSystemObject")
            'Set file = fso.CreateTextFile(App_path & "\include\sys\" & script_list.ListItems(i).ListSubItems(1).Text, True, False)
            'file.Write script_txt
            'file.Close
            
        End If
    Next i
    
    StatusBar.Panels(2) = ""
    Check_script
    
    script_quit = True
    Toolbar.Buttons(1).Enabled = True
    Toolbar.Buttons(2).Enabled = True
    Toolbar.Buttons(4).Enabled = False
End Sub

Private Sub script_load_StateChanged(ByVal State As Integer)
    If script_quit = True Then script_load.Cancel
    DoEvents
    
    On Error Resume Next
    Dim vtData As Variant '数据变量
    Dim binBuffer() As Byte
    Dim firt_byte As Boolean
    Dim buff() As Byte
    
    Select Case State
        
    Case icResponseCompleted
        
        firt_byte = False
        
        Do   '从缓冲区读取数据
            DoEvents
            vtData = script_load.GetChunk(51200, icByteArray)
            binBuffer = vtData
            If firt_byte = False Then
                buff = vtData
                firt_byte = True
            Else
                buff = UniteByteArray(buff, binBuffer)
            End If
        Loop Until (LenB(vtData) = 0)
        
        script_txt = ""
        script_txt = bin2str(buff)
        script_txt_byte = Null
        script_txt_byte = buff
        
        script_down_ok = True
    Case icError
        '与主机通信出错
        Call script_download
    End Select
    
End Sub

Public Sub script_download()
    DoEvents
    '文件大小值复位
    On Error GoTo err_ctrl
    
    '定义ITC控件使用的协议为HTTP协议
    'script_load.Protocol = icHTTP
    
    '调用Execute方法向Web服务器发送HTTP请求
    script_load.Execute Trim$(strURL), "GET"
    Exit Sub
    
err_ctrl:
    script_load.Cancel
    
    script_down_ok = True
End Sub

Private Function bin2str(ByVal binstr)
    On Error Resume Next
    Const adTypeBinary = 1
    Const adTypeText = 2
    Dim BytesStream, StringReturn
    Set BytesStream = CreateObject("ADODB.Stream") '建立一个流对象
    With BytesStream
        
        .Type = adTypeBinary
        .Open
        .Write binstr
        .Position = 0
        .Type = adTypeText
        .Charset = htmlCharsetType
        StringReturn = .ReadText
        .Close
        
    End With
    Set BytesStream = Nothing
    bin2str = StringReturn
End Function


Private Function UniteByteArray(bBa1() As Byte, bBa2() As Byte) As Byte()
    On Error Resume Next
    Dim bUb() As Byte
    Dim iUbd1 As Double
    Dim iUbd2 As Double
    Dim i As Single
    
    iUbd1 = UBound(bBa1)
    iUbd2 = UBound(bBa2)
    ReDim bUb(0 To iUbd1 + iUbd2 + 1) As Byte
    For i = 0 To iUbd1
        DoEvents
        bUb(i) = bBa1(i)
    Next i
    For i = iUbd1 + 1 To UBound(bUb)
        DoEvents
        bUb(i) = bBa2(i - iUbd1 - 1)
    Next i
    UniteByteArray = bUb
End Function

