Attribute VB_Name = "OX_variable"
Public Const title_info = "OX163 plus(0.5.5build110416)"
Public Const ver_info = "55"
'------------------------------------------------------------------------------------
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209

Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_HIDE = 0

Public Const WM_SYSCOMMAND = &H112
Public Const SC_RESTORE = &HF120&

'窗口最前端参数----------------------------------------
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 256
End Type

Public TrayI As NOTIFYICONDATA


'BrowserW传递、判断参数----------------------------------------
Public BrowserW_url As String
Public BrowserW_load_ok As Boolean

'外部脚本脚本头（包括必要参数以及函数）-------------------
Type include_ScriptCode
    OX163_vbs_var As String
    OX163_vbs_fn As String
    OX163_js_var As String
    OX163_js_fn As String
End Type

'脚本调用传递参数
Public in_Script_Code As include_ScriptCode
'浏览器脚本控制参数
Public OX163_WebBrowser_scriptCode As String

'全局程序组目录
Public App_path As String

'系统参数------------------------------------------------
Type sysSetting
    '版本
    ver As Integer
    '下载区块
    downloadblock As Long
    '检查更新
    autocheck As Boolean
    '执行时退出询问
    askquit As Boolean
    '执行时显示列表
    listshow As Boolean
    '保存到默认文件夹
    savedef As Boolean
    '下载后打开文件夹
    openfloder As Boolean
    '密码错误时，再次询问密码
    change_psw As Boolean
    '窗口总在最前面
    always_top As Boolean
    '阻止弹出窗口
    new_ie_win As Boolean
    '弹出窗口用OX163打开
    ox163_ie_win As Boolean
    '链接超时
    time_out As Integer
    '重试次数
    retry_times As Integer
    '输出下载列表方式
    list_type As Byte
    '自动校正伪图
    fix_rar As Byte
    '伪图文件名列表
    fix_rar_name As String
    '系统托盘
    sysTray As Boolean
    '是否开启默认路径
    def_path_tf As Boolean
    '默认路径
    def_path As String
    '外部脚本执行方式
    include_script As String
    'ctrl+c等操作设定
    list_copy As Boolean
    '已下载文件比较
    file_compare As Integer
    '底部信息栏
    bottom_StatusBar As Boolean
    '新163相册验证码测试
    new163passcode_def(2) As String
    '列表后是否全选
    check_all As Boolean
    '代理服务器A for start fast
    proxy_A As String
    proxy_A_user As String
    proxy_A_pw As String
    '代理服务器B for inet1 and header ckeck
    proxy_B As String
    proxy_B_user As String
    proxy_B_pw As String
    '代理服务器使用方式
    proxy_A_type As Byte
    proxy_B_type As Byte
    '是否建立URL文件夹
    url_folder As Boolean
    '使用新163相册中文密码规则
    new163pass_rules As Boolean
End Type

Public sysSet As sysSetting
