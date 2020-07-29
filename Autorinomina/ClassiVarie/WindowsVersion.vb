Imports System.Runtime.InteropServices

Module WindowsVersion
    Public Function CheckOSVersion() As String
        Dim OS_VerMajor As Integer = Environment.OSVersion.Version.Major
        'Dim OS_VerMinor As Integer = Environment.OSVersion.Version.Minor
        Dim API_VerMajor, API_VerMinor As Integer
        Dim Result As Integer = 0

        'OSVersion function is deprecated on Windows 8/ 8.1 / 10 etc use API workaround
        Try
            Dim success As Integer
            Dim pInfoBuffer As IntPtr = IntPtr.Zero
            success = Interaction.NetWkstaGetInfo(vbNullChar, 100, pInfoBuffer)
            If success = 0 Then
                Dim info As NativeMethods.WKSTA_INFO_100 = Marshal.PtrToStructure(pInfoBuffer, GetType(NativeMethods.WKSTA_INFO_100))

                API_VerMajor = info.wki100_ver_major
                API_VerMinor = info.wki100_ver_minor

                Interaction.NetApiBufferFree(pInfoBuffer)

                Result = API_VerMajor ' & "." & API_VerMinor

            Else
                'error rertrive info
                Result = OS_VerMajor ' & "." & OS_VerMinor
            End If

        Catch ex As Exception
            'api error? use OS version
            Result = OS_VerMajor ' & "." & OS_VerMinor
        End Try


        Select Case Result
            Case 5 '5.1 = XP || 5.2 = 2003 and XP64bit
                Return "XP"
            Case 6 '6.0 = Vista || 6.1 = Win2008R2 / Win7 || 6.2 = Win8 / Win2012 || 6.3 = Win8.1 / Win2012R2
                Return "Win7_8"
            Case 10, Is > 10  '10.0 = Win10
                Return "Win10"
            Case Else
                Return "Win7_8"
        End Select

    End Function




End Module
