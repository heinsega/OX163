VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OX_History_Logs 
   Caption         =   "History Logs"
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15495
   Icon            =   "History_Logs.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9585
   ScaleWidth      =   15495
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Frame History_Logs_Frame 
      Caption         =   "程序运行记录"
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13695
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   2655
         TabIndex        =   4
         Top             =   240
         Width           =   2655
         Begin VB.CheckBox HL_Auto_Cls 
            Caption         =   "记录上限10000条"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   0
            Value           =   1  'Checked
            Width           =   2295
         End
      End
      Begin MSComctlLib.ImageList OX_HL_ImageList2 
         Left            =   1440
         Top             =   6840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   21
         ImageHeight     =   21
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "History_Logs.frx":406A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "History_Logs.frx":40EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "History_Logs.frx":416E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "History_Logs.frx":41EA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList OX_HL_ImageList1 
         Left            =   600
         Top             =   6840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   21
         ImageHeight     =   21
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "History_Logs.frx":4266
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "History_Logs.frx":42E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "History_Logs.frx":436A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "History_Logs.frx":43E6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView listView 
         Height          =   3420
         Index           =   0
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Shift or Ctrl to MultiSelect"
         Top             =   720
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6033
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "LV_time"
            Object.Tag             =   "LV_time"
            Text            =   "时间"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "LV_text"
            Object.Tag             =   "LV_text"
            Text            =   "程序信息"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "LV_url"
            Object.Tag             =   "LV_url"
            Text            =   "当前链接"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "LV_path"
            Object.Tag             =   "LV_path"
            Text            =   "当前路径"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "LV_script"
            Object.Tag             =   "LV_script"
            Text            =   "当前脚本"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.Toolbar OXH_tool 
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   714
         ButtonWidth     =   2540
         ButtonHeight    =   714
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "OX_HL_ImageList1"
         HotImageList    =   "OX_HL_ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "运行历史  "
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HL_LV_cls1"
                     Object.Tag             =   "HL_LV_cls1"
                     Text            =   "清空列表"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "浏览器历史 "
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HL_LV_cls2"
                     Object.Tag             =   "HL_LV_cls2"
                     Text            =   "清空列表"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         MousePointer    =   1
      End
      Begin MSComctlLib.ListView listView 
         Height          =   6420
         Index           =   1
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Shift or Ctrl to MultiSelect"
         Top             =   720
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   11324
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "LV_time"
            Object.Tag             =   "LV_time"
            Text            =   "时间"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "LV_text"
            Object.Tag             =   "LV_text"
            Text            =   "浏览信息"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "LV_url"
            Object.Tag             =   "LV_url"
            Text            =   "浏览链接"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "LV_path"
            Object.Tag             =   "LV_path"
            Text            =   "页面标题"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "LV_script"
            Object.Tag             =   "LV_script"
            Text            =   "浏览状态"
            Object.Width           =   3528
         EndProperty
      End
   End
End
Attribute VB_Name = "OX_History_Logs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> 1 Then
        History_Logs_Frame.Width = Me.ScaleWidth - 240
        History_Logs_Frame.Height = Me.ScaleHeight - 240
        For i = 0 To 1
            listView(i).Width = History_Logs_Frame.Width - 240
            listView(i).Height = History_Logs_Frame.Height - 840
            If listView(i).Width > 10400 Then
                listView(i).ColumnHeaders.Item(1).Width = 2000
                listView(i).ColumnHeaders.Item(4).Width = 2000
                listView(i).ColumnHeaders.Item(5).Width = 2000
                listView(i).ColumnHeaders.Item(2).Width = (listView(i).Width - 6000) * 0.5
                listView(i).ColumnHeaders.Item(3).Width = listView(i).Width - 6400 - listView(i).ColumnHeaders.Item(2).Width
            End If
        Next
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OX_History_Logs.Hide
    Cancel = True
End Sub

Private Sub listView_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim copy_txt() As String, copy_txt_join As String
    If KeyCode = 65 And Shift = vbCtrlMask Then
        listView(Index).Enabled = False
        listView(Index).Visible = False
        For i = 1 To listView(Index).ListItems.count
            DoEvents
            listView(Index).ListItems(i).Selected = True
        Next
        listView(Index).Visible = True
        listView(Index).Enabled = True
        listView(Index).SetFocus
    ElseIf KeyCode = 67 And Shift = vbCtrlMask Then
        GoTo list_copy
    End If
    Exit Sub
    '--------------------------------------------------
list_copy:
    listView(Index).Enabled = False
    ReDim copy_txt(listView(Index).ListItems.count - 1)
    For i = 1 To listView(Index).ListItems.count
        DoEvents
        If listView(Index).ListItems(i).Selected = True Then
            copy_txt(i) = ""
            copy_txt(i) = copy_txt(i) & listView(Index).ListItems(i).Text & vbCrLf
            copy_txt(i) = copy_txt(i) & listView(Index).ListItems(i).ListSubItems(1).Text & vbCrLf
            copy_txt(i) = copy_txt(i) & listView(Index).ListItems(i).ListSubItems(2).Text & vbCrLf
            copy_txt(i) = copy_txt(i) & listView(Index).ListItems(i).ListSubItems(3).Text & vbCrLf
            copy_txt(i) = copy_txt(i) & listView(Index).ListItems(i).ListSubItems(4).Text & vbCrLf
            copy_txt(i) = copy_txt(i) & "------------------------------" & vbCrLf
        End If
    Next
    copy_txt_join = Trim(Join(copy_txt, ""))
    If copy_txt_join <> "" Then
        Clipboard.Clear
        Clipboard.SetText copy_txt_join
    End If
    listView(Index).Enabled = True
    listView(Index).SetFocus
    Exit Sub
End Sub

Private Sub OXH_tool_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Index
    Case 1
        OXH_tool.Buttons(1).Image = 2
        OXH_tool.Buttons(2).Image = 3
        listView(1).Visible = False
        listView(0).Visible = True
    Case 2
        OXH_tool.Buttons(1).Image = 1
        OXH_tool.Buttons(2).Image = 4
        listView(0).Visible = False
        listView(1).Visible = True
    End Select
End Sub

Public Sub OX_HL_listView_Add(ByVal Index As Integer, ByVal OX_HLLV_s1 As String, ByVal OX_HLLV_s2 As String, ByVal OX_HLLV_s3 As String, ByVal OX_HLLV_s4 As String)
    On Error Resume Next
    listView(Index).ListItems.Item(1).Selected = False
    listView(Index).ListItems.Add 1, , Now()
    listView(Index).ListItems.Item(1).ListSubItems.Add , , OX_HLLV_s1
    listView(Index).ListItems.Item(1).ListSubItems.Add , , OX_HLLV_s2
    listView(Index).ListItems.Item(1).ListSubItems.Add , , OX_HLLV_s3
    listView(Index).ListItems.Item(1).ListSubItems.Add , , OX_HLLV_s4
    listView(Index).ListItems.Item(1).Selected = True
    History_Logs_Frame.caption = "程序运行记录(" & listView(0).ListItems.count & "/" & listView(1).ListItems.count & ")"
    Form1.StatusBar.Panels(2).Text = listView(0).ListItems.count & "/" & listView(1).ListItems.count
    If HL_Auto_Cls.Value = 1 Then
    listView(Index).Visible = False
        Do While listView(Index).ListItems.count > 10000
            listView(Index).ListItems.Remove listView(Index).ListItems.count
        Loop
    listView(Index).Visible = True
    End If
End Sub

Private Sub OXH_tool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error Resume Next
    Select Case ButtonMenu.key
    Case "HL_LV_cls1"
        listView(0).ListItems.Clear
        History_Logs_Frame.caption = "程序运行记录(" & listView(0).ListItems.count & "/" & listView(1).ListItems.count & ")"
        Form1.StatusBar.Panels(2).Text = listView(0).ListItems.count & "/" & listView(1).ListItems.count
    Case "HL_LV_cls2"
        listView(1).ListItems.Clear
        History_Logs_Frame.caption = "程序运行记录(" & listView(0).ListItems.count & "/" & listView(1).ListItems.count & ")"
        Form1.StatusBar.Panels(2).Text = listView(0).ListItems.count & "/" & listView(1).ListItems.count
    End Select
End Sub
