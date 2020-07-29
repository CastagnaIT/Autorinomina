Public Class WND_About
    Private Sub WND_About_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        LB_Version.Content = Localization.Resource_WND_About.Version & Space(1) & My.Application.Info.Version.ToString
    End Sub

    Private Sub HY_Homepage_Click(sender As Object, e As RoutedEventArgs)
        Try
            System.Diagnostics.Process.Start("http://www.autorinomina.it/")
        Catch ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, Me.Title)
        End Try
    End Sub

    Private Sub HY_Support_click(sender As Object, e As RoutedEventArgs)
        Try
            System.Diagnostics.Process.Start("MAILTO:info@autorinomina.it")
        Catch ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, Me.Title)
        End Try
    End Sub

    Private Sub HY_License_Click(sender As Object, e As RoutedEventArgs)
        Try
            System.Diagnostics.Process.Start(IO.Path.Combine(My.Application.Info.DirectoryPath, "License.txt"))
        Catch ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, Me.Title)
        End Try
    End Sub

    Private Sub WND_About_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Me.PreviewKeyDown
        If e.Key = Key.Escape Then
            Close()
        End If
    End Sub
End Class
