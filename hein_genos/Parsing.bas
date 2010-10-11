Attribute VB_Name = "Parsing"
Option Explicit

Public Type ScriptInfo
    'tom.vbs|vbscript|GB2312|album|http://photo.tom.com/pim.php?*
    FileName As String
    Language As String
    Encoding As String
    HandleType As String
    Criteria As String
End Type

Public Function ParseInclude(ByVal sourceString As String) As ScriptInfo
    'include中的内容
    'tom.vbs|vbscript|GB2312|album|http://photo.tom.com/pim.php?*
    '判断完include之后发送到list_album_script或者list_photo_script中的内容
    'eht.vbs|vbscript|UTF-8|album|http://g.e-hentai.org/g/132226/e6f6d72be1/
End Function

