Attribute VB_Name = "CMDresult"
Option Explicit

Private Const STARTF_USESHOWWINDOW     As Long = &H1
Private Const STARTF_USESTDHANDLES     As Long = &H100
Private Const SW_HIDE                  As Integer = 0

Private Type SECURITY_ATTRIBUTES
    nLength                                As Long
    lpSecurityDescriptor                   As Long
    bInheritHandle                         As Long
End Type
Private Type STARTUPINFO
    cb                                     As Long
    lpReserved                             As String
    lpDesktop                              As String
    lpTitle                                As String
    dwX                                    As Long
    dwY                                    As Long
    dwXSize                                As Long
    dwYSize                                As Long
    dwXCountChars                          As Long
    dwYCountChars                          As Long
    dwFillAttribute                        As Long
    dwFlags                                As Long
    wShowWindow                            As Integer
    cbReserved2                            As Integer
    lpReserved2                            As Long
    hStdInput                              As Long
    hStdOutput                             As Long
    hStdError                              As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess                               As Long
    hThread                                As Long
    dwProcessId                            As Long
    dwThreadId                             As Long
End Type

Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'---------------------------------------------------
' Call this sub to execute and capture a console app.
' Ex: Call CMDcapture("ping localhost", Text1)
Private Function CMDcapture(ByVal cmd As String, Optional ByVal start_dir As String = vbNullString) As String
Const BUFSIZE         As Long = 1024 * 10
Dim hPipeRead         As Long
Dim hPipeWrite        As Long
Dim sa                As SECURITY_ATTRIBUTES
Dim si                As STARTUPINFO
Dim pi                As PROCESS_INFORMATION
Dim baOutput(BUFSIZE) As Byte
Dim result            As String
Dim lBytesRead        As Long
Dim time1 As Date

    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1
    End With

    If CreatePipe(hPipeRead, hPipeWrite, sa, 0) = 0 Then
        CMDcapture = "error creating handles"
        Exit Function
    End If

    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE
        .hStdOutput = hPipeWrite
        .hStdError = hPipeWrite
    End With
    
    If CreateProcess(vbNullString, cmd, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, start_dir, si, pi) Then
        time1 = Now()
        Do
        DoEvents
        Loop While DateDiff("s", time1, Now()) < 0.5
        
        CloseHandle hPipeWrite
        CloseHandle pi.hThread
        hPipeWrite = 0
        Do
            DoEvents
            If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0 Then
                Exit Do
            End If
            result = Left(StrConv(baOutput(), vbUnicode), lBytesRead)
        Loop
        CloseHandle pi.hProcess
    Else
        result = "error creating process"
    End If

    CloseHandle hPipeRead
    CloseHandle hPipeWrite
    result = Left(result, InStr(result, Chr(0)) - 1)

    CMDcapture = result
End Function

Public Function OX_8dot3Name_Dir(ByVal OX_8dot3driver As String) As String
'1表示禁用，0表示启用。
On Error GoTo ErrHandler
Dim result As String
OX_8dot3Name_Dir = "-1"
OX_8dot3driver = Left(OX_8dot3driver, 2)
If OX_8dot3driver Like "?:" Then
result = CMDcapture("fsutil 8dot3name query " & OX_8dot3driver)
    If result Like "Disable8dot3 ?* 0 (8dot3*" Or result Like "?*: 0 (8dot3*" Then '卷状态为: 0 (8dot3 名称创建已启用)。win10
        OX_8dot3Name_Dir = "0"
    ElseIf result Like "Disable8dot3 ?* 1 (8dot3*" Or result Like "?*: 1 (8dot3*" Then
        OX_8dot3Name_Dir = "1"
    End If
End If
ErrHandler:
OX_8dot3Name_Dir = OX_8dot3Name_Dir & vbCrLf & result
End Function

Public Function OX_8dot3Name_Sys() As String
'0（全部启动），1（全部禁用），2（每个盘符单独设置），3（除系统盘外全部禁用）。
On Error Resume Next
Dim result As String
OX_8dot3Name_Sys = "2"

err.Clear
    Dim OX_8dot3Name_reg
        Set OX_8dot3Name_reg = CreateObject("WScript.Shell")
        OX_8dot3Name_Sys = OX_8dot3Name_reg.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem\NtfsDisable8dot3NameCreation")

If err.Number = 0 And (OX_8dot3Name_Sys = "0" Or OX_8dot3Name_Sys = "1" Or OX_8dot3Name_Sys = "2" Or OX_8dot3Name_Sys = "3") Then Exit Function

OX_8dot3Name_Sys = "2"
result = CMDcapture("fsutil 8dot3name query") '注册表状态为: 2 (按卷设置 - 默认值)。
    If result Like "NtfsDisable8dot3NameCreation ?* 0 (*" Or result Like "?*: 0 (*" Then
        OX_8dot3Name_Sys = "0"
    ElseIf result Like "NtfsDisable8dot3NameCreation ?* 1 (*" Or result Like "?*: 1 (*" Then
        OX_8dot3Name_Sys = "1"
    ElseIf result Like "NtfsDisable8dot3NameCreation ?* 2 (*" Or result Like "?*: 2 (*" Then
        OX_8dot3Name_Sys = "2"
    ElseIf result Like "NtfsDisable8dot3NameCreation ?* 3 (*" Or result Like "?*: 3 (*" Then
        OX_8dot3Name_Sys = "3"
    End If

End Function
