Class Application
    Private Sub Application_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
        Try
            If IO.File.Exists(IO.Path.Combine(My.Application.Info.DirectoryPath, "settings.xml")) Then
                DataPath = AppDomain.CurrentDomain.BaseDirectory
            Else
                DataPath = IO.Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData).ToString, "Autorinomina")
            End If

            If IO.Directory.Exists(IO.Path.Combine(DataPath, "cache")) = False Then IO.Directory.CreateDirectory(IO.Path.Combine(DataPath, "cache"))
            Debug.Print(DataPath)
            CaricaValoriFallback()
            XMLSettings_LoadFile()
            XMLStrutture_LoadFile()

            If Not String.IsNullOrEmpty(XMLSettings_Read("Language")) Then
                Globalization.CultureInfo.DefaultThreadCurrentCulture = New Globalization.CultureInfo(XMLSettings_Read("Language"))
                Globalization.CultureInfo.DefaultThreadCurrentUICulture = New Globalization.CultureInfo(XMLSettings_Read("Language"))
            End If

        Catch ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & Space(1) & ex.Message, MsgBoxStyle.Critical, "Autorinomina")
            End
        End Try
    End Sub

    ' Gli eventi a livello di applicazione, ad esempio Startup, Exit e DispatcherUnhandledException,
    ' possono essere gestiti in questo file.

End Class
