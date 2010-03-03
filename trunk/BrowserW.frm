VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form BrowserW 
   BorderStyle     =   0  'None
   Caption         =   "Browser Windows"
   ClientHeight    =   90
   ClientLeft      =   -105
   ClientTop       =   -105
   ClientWidth     =   90
   Enabled         =   0   'False
   LinkTopic       =   "BrowserW"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   2295
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin SHDocVwCtl.WebBrowser WebBrowser 
         Height          =   4935
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   6735
         ExtentX         =   11880
         ExtentY         =   8705
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
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
End
Attribute VB_Name = "BrowserW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BrowserW_load_ok As Boolean



Private Sub Form_Load()
On Error Resume Next
BrowserW_load_ok = False

BrowserW.Height = 0
BrowserW.Width = 0
BrowserW.Top = 0
BrowserW.Left = 0

'BrowserW.Height = 5000
'BrowserW.Width = 5000
'BrowserW.Top = 1
'BrowserW.Left = 1
'Picture1.Visible = True
'Picture1.Enabled = True
'Me.Enabled = True


BrowserW_load_ok = True
End Sub


Private Sub WebBrowser_FileDownload(Cancel As Boolean)
On Error Resume Next
Cancel = True
End Sub

Private Sub WebBrowser_NewWindow2(ppDisp As Object, Cancel As Boolean)
On Error Resume Next
Cancel = True
End Sub

