Imports System.Runtime.InteropServices

Public NotInheritable Class Interaction

    Private Sub New()
    End Sub

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=1)>
    Public Structure SHFILEINFO
        Public hIcon As IntPtr ' : icon
        Public iIcon As Int32 ' : icondex
        Public dwAttributes As Int32 ' : SFGAO_ flags
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
        Public szTypeName As String
    End Structure

    Public Const SHGFI_ICON = &H100
    Public Const SHGFI_SMALLICON = &H1 '16x16
    Public Const SHGFI_LARGEICON = &H0    ' Large icon 32x32
    Public Const SHGFI_EXTRALARGEICON = &H2    ' Extra Large icon 48x48
    Public Const SHGFI_JUMBOICON = &H4    ' Jumbo icon solo Vista e sucessivi 256x256
    Public Const SHGFI_TYPENAME As Int32 = &H400

    ' Callers require Unmanaged permission   
    Public Shared Function StrCmpLogicalW(strA As String, strB As String) As Integer
        Return NativeMethods.StrCmpLogicalW(strA, strB)
    End Function

    Public Shared Function NetWkstaGetInfo(server As String, level As Integer, ByRef info As IntPtr) As Integer
        Return NativeMethods.NetWkstaGetInfo(server, level, info)
    End Function

    Public Shared Function NetApiBufferFree(pBuf As IntPtr) As Integer
        Return NativeMethods.NetApiBufferFree(pBuf)
    End Function

    Public Declare Auto Function SHGetFileInfo Lib "shell32.dll" (ByVal pszPath As String, ByVal dwFileAttributes As Integer, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Integer, ByVal uFlags As Integer) As IntPtr


    Private Declare Auto Function ShellExecuteEx Lib "shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Boolean
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure SHELLEXECUTEINFO
        Public cbSize As Integer
        Public fMask As UInteger
        Public hwnd As IntPtr
        <MarshalAs(UnmanagedType.LPTStr)>
        Public lpVerb As String
        <MarshalAs(UnmanagedType.LPTStr)>
        Public lpFile As String
        <MarshalAs(UnmanagedType.LPTStr)>
        Public lpParameters As String
        <MarshalAs(UnmanagedType.LPTStr)>
        Public lpDirectory As String
        Public nShow As Integer
        Public hInstApp As IntPtr
        Public lpIDList As IntPtr
        <MarshalAs(UnmanagedType.LPTStr)>
        Public lpClass As String
        Public hkeyClass As IntPtr
        Public dwHotKey As UInteger
        Public hIcon As IntPtr
        Public hProcess As IntPtr
    End Structure
    Private Const SW_SHOW As Integer = 5
    Private Const SEE_MASK_INVOKEIDLIST As UInteger = 12

    Public Shared Sub ShowFileProperties(Filename As String)
        Dim info As New SHELLEXECUTEINFO()
        info.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(info)
        info.lpVerb = "properties"
        info.lpFile = Filename
        info.nShow = SW_SHOW
        info.fMask = SEE_MASK_INVOKEIDLIST
        ShellExecuteEx(info)
    End Sub
End Class

Friend NotInheritable Class NativeMethods

    Private Sub New()
    End Sub

    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True, SetLastError:=False)>
    Friend Shared Function StrCmpLogicalW(strA As String, strB As String) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Friend Shared Function MessageBeep(ByVal uType As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode)>
    Friend Shared Function NetWkstaGetInfo(server As String, level As Integer, ByRef info As IntPtr) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode)>
    Friend Shared Function NetApiBufferFree(pBuf As IntPtr) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Friend Structure WKSTA_INFO_100
        Public wki100_platform_id As Integer
        <MarshalAs(UnmanagedType.LPWStr)>
        Public wki100_computername As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public wki100_langroup As String
        Public wki100_ver_major As Integer
        Public wki100_ver_minor As Integer
    End Structure





End Class

