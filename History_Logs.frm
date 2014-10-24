VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OX_History_Logs 
   Caption         =   "History Logs"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7740
   Icon            =   "History_Logs.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6285
   ScaleWidth      =   7740
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame History_Logs_Frame 
      Caption         =   "程序运行记录"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   609
         ButtonWidth     =   2011
         ButtonHeight    =   609
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "asdasda"
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "cccc"
               Style           =   2
            EndProperty
         EndProperty
         MousePointer    =   1
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5415
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   9551
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "OX_History_Logs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
OX_History_Logs.Hide
Cancel = True
End Sub
