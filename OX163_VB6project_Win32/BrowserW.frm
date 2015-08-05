VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form BrowserW 
   BorderStyle     =   0  'None
   Caption         =   "Browser Windows"
   ClientHeight    =   1845
   ClientLeft      =   -105
   ClientTop       =   -105
   ClientWidth     =   2460
   Enabled         =   0   'False
   LinkTopic       =   "BrowserW"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   2460
   ShowInTaskbar   =   0   'False
   Begin VB.Timer BrowserW_Timer 
      Interval        =   1
      Left            =   2040
      Top             =   1440
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   2295
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin SHDocVwCtl.WebBrowser WebBrowser 
         Height          =   4575
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10575
         ExtentX         =   18653
         ExtentY         =   8070
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
         Location        =   ""
      End
   End
End
Attribute VB_Name = "BrowserW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BrowserW_Timer_Timer()
    On Error Resume Next
    
    BrowserW_Timer.Enabled = False
    
    WebBrowser.Silent = True
    WebBrowser.Document.Open
    WebBrowser.Document.Write ""
    WebBrowser.Document.Close
    '
    BrowserW_load_ok = True
    'BrowserW.Hide
End Sub

Private Sub Form_Load()
    On Error Resume Next
    BrowserW_load_ok = False
    
    BrowserW.Height = 0
    BrowserW.Width = 0
    BrowserW.Top = 0
    BrowserW.Left = 0
    BrowserW.Enabled = False
    
    
    'BrowserW.Height = 5200
    'BrowserW.Width = 10000
    'Picture1.Height = 5200
    'Picture1.Width = 10000
    'BrowserW.Top = 1
    'BrowserW.Left = 1
    'Picture1.Visible = True
    'Picture1.Enabled = True
    'Me.Enabled = True
    
    BrowserW_Timer.Enabled = True
    
End Sub


Private Sub WebBrowser_DownloadComplete()
    On Error Resume Next
    WebBrowser.Stop
End Sub


Private Sub WebBrowser_NewWindow2(ppDisp As Object, Cancel As Boolean)
    On Error Resume Next
    Cancel = True
End Sub


Private Sub WebBrowser_StatusTextChange(ByVal Text As String)
    On Error Resume Next
    Static count_http As Byte
    If InStr(Text, "http://") > 0 And InStr(Text, BrowserW_url) <= 0 And count_http > 10 Then
        count_http = 0
        WebBrowser.Stop
    Else
        count_http = count_http + 1
    End If
End Sub
