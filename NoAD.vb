Imports System.Security.AccessControl    'Used by: SetAccessControl

Module NoAD
    Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, ByVal lpVolumeSerialNumber As Integer, ByVal lpMaximumComponentLength As Integer, ByVal lpFileSystemFlags As Integer, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Integer) As Integer
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal Msg As UInt32, ByVal wParam As Integer, ByVal lParam As IntPtr) As Integer

#Region "DllErrorInfo"
    Friend Declare Function GetLastError Lib "kernel32 " () As Integer
    Private Declare Auto Function FormatMessage Lib "kernel32" (ByVal dwFlags As FormatMessageFlags, ByRef lpSource As Integer, ByVal dwMessageId As Integer, ByVal dwLanguageId As Languages, ByVal lpBuffer As String, ByVal nSize As Integer, ByRef Arguments As Integer) As Integer
    Private Enum FormatMessageFlags As Integer
        FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100I
        FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000I
        FORMAT_MESSAGE_FROM_HMODULE = &H800I
        FORMAT_MESSAGE_FROM_STRING = &H400I
        FORMAT_MESSAGE_FROM_SYSTEM = &H1000I
        FORMAT_MESSAGE_IGNORE_INSERTS = &H200I
        Zero = 0I
        FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFFI
    End Enum
    Private Enum Languages As Short
        LANG_NEUTRAL = &H0S
        SUBLANG_DEFAULT = &H1S
    End Enum
    ''' <summary>Gets information about error with specified number</summary>
    ''' <param name="ErrN">Number of error</param>
    ''' <returns>Description of error</returns>
    Friend Function LastDllErrorInfo(ByVal ErrN As Integer) As String
        Dim Buffer As String
        Buffer = Space(200)
        FormatMessage(FormatMessageFlags.FORMAT_MESSAGE_FROM_SYSTEM, 0, ErrN, Languages.LANG_NEUTRAL, Buffer, 200, 0)
        Return Buffer.Trim
    End Function
#End Region

    ''' <summary>
    ''' 读取注册表的键值
    ''' </summary>
    ''' <param name="KeyPath">键值的完整路径</param>
    ''' <param name="OutputKeyValue">函数成功时将重新设置此变量为键值的数据</param>
    ''' <returns>成功返回0，否则返回错误代码（-1表示HKEY_USERS之类的名称错误）</returns>
    ''' <remarks>ReadRegedit("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\CommonFilesDir", val) 
    ''' #[returns 0; val=C:\Program Files\Common Files]#</remarks>
    Friend Function ReadRegedit(ByVal KeyPath As String, ByRef OutputKeyValue As String) As Integer
        'Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\AFK Bot", KeyName, DefaultSettings)
        On Error GoTo ErrRR
        Dim key As Microsoft.Win32.RegistryKey
        Select Case Strings.Left(KeyPath, Strings.InStr(KeyPath, "\") - 1)
            Case "HKEY_CLASSES_ROOT"
                key = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(Strings.Mid(KeyPath, Strings.InStr(KeyPath, "\") + 1, Strings.InStrRev(KeyPath, "\") - Strings.InStr(KeyPath, "\") - 1))
            Case "HKEY_CURRENT_USER"
                key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(Strings.Mid(KeyPath, Strings.InStr(KeyPath, "\") + 1, Strings.InStrRev(KeyPath, "\") - Strings.InStr(KeyPath, "\") - 1))
            Case "HKEY_LOCAL_MACHINE"
                key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(Strings.Mid(KeyPath, Strings.InStr(KeyPath, "\") + 1, Strings.InStrRev(KeyPath, "\") - Strings.InStr(KeyPath, "\") - 1))
            Case "HKEY_USERS"
                key = Microsoft.Win32.Registry.Users.OpenSubKey(Strings.Mid(KeyPath, Strings.InStr(KeyPath, "\") + 1, Strings.InStrRev(KeyPath, "\") - Strings.InStr(KeyPath, "\") - 1))
            Case "HKEY_CURRENT_CONFIG"
                key = Microsoft.Win32.Registry.CurrentConfig.OpenSubKey(Strings.Mid(KeyPath, Strings.InStr(KeyPath, "\") + 1, Strings.InStrRev(KeyPath, "\") - Strings.InStr(KeyPath, "\") - 1))
            Case Else
                Return -1
        End Select
        OutputKeyValue = CType(key.GetValue(Strings.Right(KeyPath, Strings.Len(KeyPath) - Strings.InStrRev(KeyPath, "\"))), String)
        Exit Function
ErrRR:
        Return Err.Number
    End Function

    ''' <summary>
    ''' 读取文件版本
    ''' </summary>
    ''' <param name="DllExePath">目标DLL或EXE的完整路径</param>
    ''' <param name="OutputVersionValue">函数成功时将重新设置此变量为文件版本</param>
    ''' <returns>成功返回0，否则返回错误代码（可用ErrorToString(returns)来获取错误代码的信息）</returns>
    ''' <remarks>GetFileVersion(Application.ExecutablePath, ver) #[returns 0; ver=1.0.0.0]#</remarks>
    Friend Function GetFileVersion(ByVal DllExePath As String, ByRef OutputVersionValue As String) As Integer
        On Error GoTo ErrGFV
        If System.IO.File.Exists(DllExePath) Then
            Dim info As FileVersionInfo = FileVersionInfo.GetVersionInfo(DllExePath)
            OutputVersionValue = info.FileVersion
        End If
        Exit Function
ErrGFV:
        Return Err.Number
    End Function

    ''' <summary>
    ''' 读取文件修改日期
    ''' </summary>
    ''' <param name="FilePath">目标文件的完整路径</param>
    ''' <param name="OutputDateValue">函数成功时将重新设置此变量为文件修改时间</param>
    ''' <returns>成功返回0，否则返回错误代码（可用ErrorToString(returns)来获取错误代码的信息）</returns>
    ''' <remarks>GetFileLastWriteTime(Application.ExecutablePath, ver) #[returns 0; ver.Year=2012]#</remarks>
    Friend Function GetFileLastWriteTime(ByVal FilePath As String, ByRef OutputDateValue As Date) As Integer
        On Error GoTo ErrGFLWT
        Dim objFileInfo As New IO.FileInfo(FilePath)
        OutputDateValue = objFileInfo.LastWriteTime
        Exit Function
ErrGFLWT:
        Return Err.Number
    End Function

    ''' <summary>
    ''' 检测进程是否正在运行
    ''' </summary>
    ''' <param name="EXEName">进程名称(不包含.exe，忽略大小写)</param>
    ''' <returns>该进程是否运行</returns>
    ''' <remarks>ProcessIfExist(My.Application.Info.AssemblyName) #[returns True]#</remarks>
    Friend Function ProcessIfExist(ByVal exeName As String) As Boolean
        For Each pro As Process In Process.GetProcesses
            If pro.ProcessName.ToLower = exeName.ToLower Then Return True
        Next
    End Function

    ''' <summary>
    ''' 设置目标路径的ACL权限，并可附带属性设置
    ''' </summary>
    ''' <param name="AimPath">目标文件(夹)的路径</param>
    ''' <param name="UserName">需要修改权限的用户名</param>
    ''' <param name="AccessControl">完全控制或完全拒绝</param>
    ''' <param name="SetAttributes">设置目标属性（仅当AccessControl=true时有效）</param>
    ''' <param name="ForceCreatFile">若目标不存在，true=创建文件，false=创建文件夹</param>
    ''' <param name="FunctionDetail">出错时，会重新设置此变量为具体错误信息（EN）</param>
    ''' <returns>成功返回0，否则返回错误代码（可用ErrorToString(returns)来获取错误代码的信息）</returns>
    ''' <remarks>目标D:\desktop\a\1.txt，用户everyone，完全控制，属性为readonly，创建文件，返回错误信息errmsg
    ''' 前提：1.txt文件不存在，程序没有权限在文件夹a中创建文件 required: @@Imports System.Security.AccessControl@@
    ''' SetAccessControl("D:\desktop\a\1.txt", "everyone", True, IO.FileAttributes.ReadOnly, True, errmsg)
    ''' #[returns 5]# errmsg=。。。
    ''' </remarks>
    Friend Function SetAccessControl(ByVal AimPath As String, ByVal UserName As String, ByVal AccessControl As Boolean, Optional ByVal SetAttributes As IO.FileAttributes = 0, Optional ByVal ForceCreatFile As Boolean = False, Optional ByRef FunctionDetail As String = "") As Integer
        On Error GoTo ErrSAC
        If System.IO.File.Exists(AimPath) = False And System.IO.Directory.Exists(AimPath) = False Then
            If ForceCreatFile Then
                IO.File.Create(AimPath)
            Else
                IO.Directory.CreateDirectory(AimPath)
            End If
        End If
        Dim AimPathInfo As IO.DirectoryInfo = New IO.DirectoryInfo(AimPath)
        Dim AimPathSecurity As DirectorySecurity = AimPathInfo.GetAccessControl()
        If AccessControl Then
            AimPathSecurity.RemoveAccessRuleAll(New FileSystemAccessRule(UserName, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Deny))
            AimPathSecurity.AddAccessRule(New FileSystemAccessRule(UserName, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
        Else
            AimPathSecurity.AddAccessRule(New FileSystemAccessRule(UserName, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Deny))
        End If
        AimPathInfo.SetAccessControl(AimPathSecurity)
        If AccessControl Then AimPathInfo.Attributes = SetAttributes
        Exit Function
ErrSAC:
        FunctionDetail = "Funtion variables: AimPath=<" & AimPath & "> UserName=<" & UserName & "> AccessControl=<" & AccessControl & ">" & vbCrLf & "Source: " & Err.Source
        Return Err.Number
    End Function

    ''' <summary>
    ''' 检测目标盘符是否为NTFS文件格式
    ''' </summary>
    ''' <param name="Disk">目标硬盘的盘符（例如：c）</param>
    ''' <param name="ForceExitIfnotNTFS">若目标盘符不是NTFS格式或不存在，并且此变量=true时，程序msgbox然后强行退出</param>
    ''' <returns>若目标盘符是NTFS格式，返回true，否则返回false</returns>
    ''' <remarks>IsNTFS("c",true) #[returns true]#</remarks>
    Friend Function IsNTFS(ByVal Disk As String, Optional ByVal ForceExitIfnotNTFS As Boolean = False) As Boolean
        Dim sbVol As New String(" ", 255), sbFil As New String(" ", 255)
        If System.IO.Directory.Exists(Disk & ":\") Then
            GetVolumeInformation(Disk & ":\", sbVol, 255, 0, 0, 0, sbFil, 255)
            If Strings.Left(sbFil, 4) = "NTFS" Then
                Return True
            Else
                If ForceExitIfnotNTFS Then
                    MsgBox("""" & Disk.ToUpper & ":\"" is not an NTFS file system, program cannot run properly.", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "NTFS Required!")
                    'MessageBox.Show("""" & Disk.ToUpper & ":\"" is not an NTFS file system, program cannot run properly.", "NTFS Required!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Process.GetCurrentProcess().Kill()
                    'Application.Exit()
                End If
            End If
        Else
            If ForceExitIfnotNTFS Then Process.GetCurrentProcess().Kill()
        End If
    End Function

    ''' <summary>
    ''' 当前程序是否以admin身份运行
    ''' </summary>
    ''' <returns>true=admin运行，否则false</returns>
    ''' <remarks>IsRunAsAdmin() #[returns false]#</remarks>
    Friend Function IsRunAsAdmin() As Boolean
        Dim principal As New Security.Principal.WindowsPrincipal(Security.Principal.WindowsIdentity.GetCurrent)
        Return principal.IsInRole(Security.Principal.WindowsBuiltInRole.Administrator)
    End Function

    ''' <summary>
    ''' 当前程序提权/目标btn添加UAC图标。函数仅提权，不会结束当前进程。
    ''' </summary>
    ''' <param name="Sender">true=提权，false=无操作，当参数为button时将修改其UAC图标！请根据需求：退出程序！</param>
    ''' <param name="Arguments">命令行参数</param>
    ''' <param name="ForceRunFunction">true=无论进程是否已提权，强制执行，false=若程序已提权，则不进行任何操作</param>
    ''' <returns>成功返回0，否则返回错误代码（-1表示OS小于Vista）</returns>
    ''' <remarks>ElevateMyself(True) #[returns 0]#</remarks>
    Friend Function ElevateMyself(ByVal Sender As Object, Optional ByVal Arguments As String = "", Optional ByVal ForceRunFunction As Boolean = False) As Integer
        If IsRunAsAdmin() = False Or ForceRunFunction Then
            If Sender.GetType.Name = "Boolean" Then
                If Sender Then
                    Dim proc As New ProcessStartInfo
                    proc.UseShellExecute = True
                    proc.WorkingDirectory = Environment.CurrentDirectory
                    proc.FileName = Application.ExecutablePath
                    proc.Arguments = Arguments
                    'proc.FileName = GetAppPath()
                    proc.Verb = "runas"
                    Try
                        Process.Start(proc)
                    Catch
                        Return Err.Number
                    End Try
                End If
            ElseIf Sender.GetType.Name = "Button" Then
                If (Environment.OSVersion.Version.Major >= 6) Then
                    Sender.FlatStyle = 3       'FlatStyle.System = 3
                    Try
                        SendMessage(Sender.Handle, &H160C, 0, New IntPtr(1))
                        'Const BCM_SETSHIELD As UInt32 = &H160C
                    Catch
                        Return Err.Number
                    End Try
                Else
                    Return -1
                End If
            End If
        End If
    End Function

    '''' <summary>
    '''' 获取当前exe路径
    '''' </summary>
    '''' <returns>返回调用此DLL的EXE完整路径</returns>
    '''' <remarks>GetAppPath() #[returns D:\Desktop\test.exe]#</remarks>
    'Friend Function GetAppPath() As String
    '    Dim strPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
    '    Dim i As Integer = strPath.LastIndexOf("\ ")
    '    If i > 0 Then
    '        strPath = Strings.Left(strPath, i)
    '    End If
    '    Return strPath
    'End Function

End Module
