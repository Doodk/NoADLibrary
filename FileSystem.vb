Imports System.IO
Imports System.Security.AccessControl    'Used by: SetAccessControl

Module FileSystem

#Region "    检测目标盘符是否为NTFS文件格式    "
    Private Declare Auto Function GetVolumeInformation Lib "kernel32" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, ByVal lpVolumeSerialNumber As Integer, ByVal lpMaximumComponentLength As Integer, ByVal lpFileSystemFlags As Integer, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Integer) As Integer
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
#End Region

    ''' <summary>
    ''' 检查目标文件(夹)是否有写入权限
    ''' </summary>
    ''' <param name="PathName">目标路径</param>
    ''' <returns>有权限返回true，否则false</returns>
    ''' <remarks>CheckACLPermission("C:\") #[returns False]#</remarks>
    Friend Function CheckACLPermission(ByVal PathName As String) As Boolean
        If File.Exists(PathName) Then
            Try
                Dim FlStr As FileStream
                FlStr = New FileStream(PathName, FileMode.Open)
                FlStr.Close()
                Return True
            Catch
                Return False
            End Try
        ElseIf Directory.Exists(PathName) Then
            Try
                Dim FlStr As FileStream
                PathName &= "\" & Guid.NewGuid.ToString
                FlStr = New FileStream(PathName, FileMode.CreateNew)
                FlStr.Close()
                Kill(PathName)
                Return True
            Catch
                Return False
            End Try
        End If
    End Function

    ''' <summary>
    ''' 设置目标路径的ACL权限，并可附带属性设置
    ''' </summary>
    ''' <param name="AimPath">目标文件(夹)的路径</param>
    ''' <param name="AccessControl">完全控制或完全拒绝</param>
    ''' <param name="UserName">需要修改权限的用户名</param>
    ''' <param name="SetAttributes">设置目标属性（仅当AccessControl=true时有效）</param>
    ''' <param name="ForceCreatFile">若目标不存在，true=创建文件，false=创建文件夹</param>
    ''' <param name="FunctionDetail">出错时，会重新设置此变量为具体错误信息（EN）</param>
    ''' <returns>成功返回0，否则返回错误代码（可用ErrorToString(returns)来获取错误代码的信息）</returns>
    ''' <remarks>目标D:\desktop\a\1.txt，用户everyone，完全控制，属性为readonly，创建文件，返回错误信息errmsg
    ''' 前提：1.txt文件不存在，程序没有权限在文件夹a中创建文件 required: @@Imports System.Security.AccessControl@@
    ''' SetAccessControl("D:\desktop\a\1.txt", "everyone", True, IO.FileAttributes.ReadOnly, True, errmsg)
    ''' #[returns 5]# errmsg=。。。
    ''' </remarks>
    Friend Function SetAccessControl(ByVal AimPath As String, ByVal AccessControl As Boolean, Optional ByVal UserName As String = "Authenticated Users", Optional ByVal SetAttributes As IO.FileAttributes = 0, Optional ByVal ForceCreatFile As Boolean = False, Optional ByRef FunctionDetail As String = "") As Integer
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
    ''' 读取文件版本
    ''' </summary>
    ''' <param name="DllExePath">目标DLL或EXE的完整路径</param>
    ''' <param name="OutputVersionValue">函数成功时将重新设置此变量为文件版本Major.Minor.Build.Private</param>
    ''' <returns>成功返回0，否则返回错误代码（可用ErrorToString(returns)来获取错误代码的信息）</returns>
    ''' <remarks>GetFileVersion(Application.ExecutablePath, ver) #[returns 0; ver.FileVersion=1.0.0.0]#</remarks>
    Friend Function GetFileVersion(ByVal DllExePath As String, ByRef OutputVersionValue As FileVersionInfo) As Integer
        On Error GoTo ErrGFV
        If File.Exists(DllExePath) Then
            Dim info As FileVersionInfo = FileVersionInfo.GetVersionInfo(DllExePath)
            OutputVersionValue = info
            'info.FileMajorPart & "." & info.FileMinorPart & "." & info.FileBuildPart & "." & info.FilePrivatePart
            'info.FileVersion
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

#Region "    锁定(文件占用)目标文件夹中所有内容    "
    'Friend FileList() As String
    'Private tmp_times As Integer
    'Friend Sub SearchFiles(ByVal strDirectory As String)
    '    Erase FileList
    '    tmp_times = 0
    '    Main_SearchFiles(strDirectory)
    'End Sub
    Friend Sub LockFiles(ByVal strDirectory As String)
        Dim mFileInfo As System.IO.FileInfo
        Dim mDir As System.IO.DirectoryInfo
        Dim mDirInfo As New System.IO.DirectoryInfo(strDirectory)
        For Each mFileInfo In mDirInfo.GetFiles()
            'ReDim Preserve FileList(tmp_times)
            'FileList(tmp_times) = mFileInfo.FullName
            'tmp_times += 1
            Do_LockFiles(mFileInfo.FullName)
        Next
        For Each mDir In mDirInfo.GetDirectories
            LockFiles(mDir.FullName)
        Next
    End Sub
    Private Sub Do_LockFiles(ByVal strDirectory As String)
        Dim val_LockFile As IO.FileStream
        val_LockFile = New IO.FileStream(strDirectory, IO.FileMode.Open)
    End Sub
#End Region

    ''' <summary>
    ''' 提取字符串中的数字(0-9)
    ''' </summary>
    ''' <param name="MixedString">输入的字符串</param>
    ''' <returns>纯数字字符串</returns>
    ''' <remarks>GetNum("这个是22") #[returns 22]#</remarks>
    Friend Function GetNum(ByVal MixedString As String) As Integer
        On Error Resume Next
        Dim i As Integer, Tmp As String = "0"
        For i = 1 To Len(MixedString)
            If InStr(1, "0123456789", Mid(MixedString, i, 1)) Then Tmp = Tmp & Mid(MixedString, i, 1)
        Next
        GetNum = Val(Tmp)
    End Function

    ''' <summary>
    ''' 读取注册表的键值
    ''' </summary>
    ''' <param name="KeyPath">键值的完整路径</param>
    ''' <param name="OutputKeyValue">函数成功时将重新设置此变量为键值的数据</param>
    ''' <returns>成功返回0，否则返回错误代码（-1表示HKEY_USERS之类的名称错误）</returns>
    ''' <remarks>ReadRegedit("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\CommonFilesDir", val) 
    ''' #[returns 0; val=C:\Program Files\Common Files]#</remarks>
    Friend Function ReadRegedit(ByVal KeyPath As String, ByRef OutputKeyValue As String) As Integer
        Try
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
        Catch
            Return Err.Number
        End Try
    End Function

End Module
