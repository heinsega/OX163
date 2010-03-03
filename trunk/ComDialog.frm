VERSION 5.00
Begin VB.Form ComDialog 
   BorderStyle     =   0  'None
   Caption         =   "ComDialog"
   ClientHeight    =   0
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   0
   LinkTopic       =   "Form2"
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Visible         =   0   'False
End
Attribute VB_Name = "ComDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
If Form1.WindowState = 0 And sysSet.always_top = True Then
Me.Top = Form1.Top
Me.Left = Form1.Left
flags = SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
SetWindowPos ComDialog.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags
End If
End Sub
