Attribute VB_Name = "OX_MouseWheel"
'HKEY_CURRENT_USER\ControlPanel\Desktop\LogPixels
'win10预览版缩放级别>100%,即LogPixels>96时定位偏移，不支持滚轮
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion>6.3为win10
Option Explicit
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = -4&
Private Const WM_MOUSEWHEEL = &H20A
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private OldWindowProc As Long '用来保存系统默认的窗口消息处理函数的地址
Private OX_hwndBox() As String   '用来保存控件的句柄

'自定义的消息处理函数
Private Function OX_MouseWheel(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    If msg = WM_MOUSEWHEEL Then
        '下面得到鼠标位置处的对象的句柄
        Dim CurPoint As POINTAPI, hwndUnderCursor As Long
        GetCursorPos CurPoint
        hwndUnderCursor = WindowFromPoint(CurPoint.X, CurPoint.Y)
        '如果鼠标位于supermap内部，则对鼠标滚轮事件进行处理
        If OX_IsEheelArea(hwndUnderCursor) Then
            If wParam = -7864320 Then '向下滚动
                sys.FrameL1_bgvs.Value = sys.FrameL1_bgvs.Value + 1
               
            ElseIf wParam = 7864320 Then '向上滚动
                sys.FrameL1_bgvs.Value = sys.FrameL1_bgvs.Value - 1
            End If
        End If
    Else
        '调用summap的默认窗口消息处理函数
        OX_MouseWheel = CallWindowProc(OldWindowProc, hwnd, msg, wParam, lParam)
    End If
End Function

'设置响应滚轮的控件句柄集合
Public Sub OX_SetWheelArea(ByVal hwnd As String)
Dim split_hwnd
split_hwnd = Split(hwnd, ",")
OX_hwndBox = split_hwnd
End Sub

'设置响应滚轮的控件句柄
Private Function OX_IsEheelArea(ByVal hwnd As Long) As Boolean
Dim i As Integer
OX_IsEheelArea = False
For i = 0 To UBound(OX_hwndBox)
If CStr(hwnd) = Trim(OX_hwndBox(i)) Then OX_IsEheelArea = True: Exit Function
Next
End Function
Public Sub OX_SetWheelStart(ByVal hwnd As Long)
    OX_SetWheelArea hwnd
    '保存默认窗口消息处理函数地址
    OldWindowProc = GetWindowLong(hwnd, GWL_WNDPROC)
    '将smMap控件的消息处理函数指定为自定义函数NewWindowProc
    Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf OX_MouseWheel)
End Sub


