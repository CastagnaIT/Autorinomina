Imports System.Runtime.InteropServices
Imports AutoRinomina.Interaction

Module GetAssociatedIconFiles
    Private nIndex = 0

    Public Function EstraiIconaAssociata(ByVal indirizzofile As String, ByVal SmallImg As Boolean) As Object()
        Dim hImgSmall As IntPtr  'The handle to the system image list.


        Dim shfi As New SHFILEINFO
        SHGetFileInfo(indirizzofile, 0, shfi, Marshal.SizeOf(shfi), SHGFI_TYPENAME Or SHGFI_ICON Or SHGFI_EXTRALARGEICON Or SHGFI_LARGEICON)
        hImgSmall = shfi.hIcon

        Dim icon3 As System.Drawing.Icon
        Dim BS As BitmapSource = Nothing
        Try
            icon3 = System.Drawing.Icon.FromHandle(shfi.hIcon)
            BS = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(icon3.Handle, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
            BS.Freeze()
        Catch ex As Exception

        End Try

        Dim Temp_type As String = ""
        Try
            Temp_type = shfi.szTypeName
            If Temp_type = "" Then
                Temp_type = IO.Path.GetExtension(indirizzofile).Remove(0, 1).ToLower
            End If
        Catch ex As Exception
            Temp_type = IO.Path.GetExtension(indirizzofile).Remove(0, 1).ToLower
        End Try

        Return {BS, Temp_type} 'bImg
    End Function

End Module

