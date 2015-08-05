Attribute VB_Name = "OX_manifest"
'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
'<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
'  <assemblyIdentity
'    version = "1.0.0.0"
'    processorArchitecture = "X86"
'    Name = "OX163"
'    type="win32"
'    />
'  <description>OX163 Windows Resource Manifest</description>
'  <!-- Identify the application Windows Style: XP and above -->
'    <dependency>
'        <dependentAssembly>
'            <assemblyIdentity
'                type="win32"
'                Name = "Microsoft.Windows.Common-Controls"
'                version = "6.0.0.0"
'                processorArchitecture = "X86"
'                publicKeyToken = "6595b64144ccf1df"
'                Language = "*"
'             />
'        </dependentAssembly>
'    </dependency>
'
'<!-- Identify the application as DPI-aware: Vista and above -->
'  <asmv3:application xmlns:asmv3="urn:schemas-microsoft-com:asm.v3">
'      <asmv3:windowsSettings xmlns="http://schemas.microsoft.com/SMI/2005/WindowsSettings">
'        <dpiAware>true</dpiAware>
'      </asmv3:windowsSettings>
'  </asmv3:application>
'
'  <!-- Identify the application security requirements: Vista and above -->
'  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v2">
'      <security>
'        <requestedPrivileges>
'          <requestedExecutionLevel
'            Level = "asInvoker"
'            uiAccess = "false"
'            />
'        </requestedPrivileges>
'      </security>
'  </trustInfo>
'
'</assembly>

Public Function Set_OX_manifest(Optional ByVal SetMani_tf1 As Boolean = True, Optional ByVal SetMani_tf2 As Boolean = True, Optional ByVal SetMani_tf3 As Boolean = True) As String
Dim manifest_str As String, manifest_name As String
Set_OX_manifest = ""
manifest_name = App_path & "\" & App.EXEName & ".exe.manifest"
If OX_Dirfile(manifest_name) Then Set_OX_manifest = CInt(OX_DelFile(manifest_name))
If SetMani_tf1 = False And SetMani_tf2 = False And SetMani_tf3 = False Then Exit Function

manifest_str = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
manifest_str = manifest_str & "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf
manifest_str = manifest_str & "<assemblyIdentity version=""1.0.0.0"" processorArchitecture=""X86"" name=""OX163"" type=""win32"" />" & vbCrLf
manifest_str = manifest_str & "<description>OX163 Windows Resource Manifest</description>" & vbCrLf
If SetMani_tf1 = True Then manifest_str = manifest_str & "<!-- Identify the application Windows Style: XP and above -->" & vbCrLf & "<dependency><dependentAssembly><assemblyIdentity type=""win32"" name=""Microsoft.Windows.Common-Controls"" version=""6.0.0.0"" processorArchitecture=""X86"" publicKeyToken=""6595b64144ccf1df"" language=""*"" /></dependentAssembly></dependency>" & vbCrLf
If SetMani_tf2 = True Then manifest_str = manifest_str & "<!-- Identify the application as DPI-aware: Vista and above -->" & vbCrLf & "<asmv3:application xmlns:asmv3=""urn:schemas-microsoft-com:asm.v3""><asmv3:windowsSettings xmlns=""http://schemas.microsoft.com/SMI/2005/WindowsSettings""><dpiAware>true</dpiAware></asmv3:windowsSettings></asmv3:application>" & vbCrLf
If SetMani_tf3 = True Then manifest_str = manifest_str & "<!-- Identify the application security requirements: Vista and above -->" & vbCrLf & "<trustInfo xmlns=""urn:schemas-microsoft-com:asm.v2""><security><requestedPrivileges><requestedExecutionLevel level=""requireAdministrator"" uiAccess=""false"" /></requestedPrivileges></security></trustInfo>" & vbCrLf
manifest_str = manifest_str & "</assembly>"

Set_OX_manifest = CInt(OX_GreatTxtFile(manifest_name, manifest_str, "UTF-8"))
If Set_OX_manifest = 0 Then Set_OX_manifest = manifest_str
End Function
