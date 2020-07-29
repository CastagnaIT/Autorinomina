Imports System.Globalization
Imports System.IO
Imports System.Reflection

Public Class WND_Languages
    Dim Coll_Lingue As New CollItemsLanguages
    Dim Dict_DatiAutori As New Dictionary(Of String, String())

    Protected Overrides Sub OnSourceInitialized(e As EventArgs)
        CustomWindow(Me, True)
    End Sub

    Private Sub GetDatiAutori()
        Dict_DatiAutori.Add("SystemDefault", New String() {"", ""})
        Dict_DatiAutori.Add("Italian", New String() {"Stefano Gottardo", "1.0"})
        Dict_DatiAutori.Add("English", New String() {"Stefano Gottardo", "1.0"})
    End Sub

    Public Function GetSupportedCulture() As IEnumerable(Of CultureInfo)
        'Get all culture 
        Dim culture As CultureInfo() = CultureInfo.GetCultures(CultureTypes.AllCultures)

        'Find the location where application installed.
        Dim exeLocation As String = Path.GetDirectoryName(Uri.UnescapeDataString(New UriBuilder(Assembly.GetExecutingAssembly().CodeBase).Path))

        'Return all culture for which satellite folder found with culture code.
        Return culture.Where(Function(cultureInfo__1) Directory.Exists(Path.Combine(exeLocation, cultureInfo__1.Name)))
    End Function

    Private Sub WND_Languages_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Dim LingueSupportate = GetSupportedCulture()

        GetDatiAutori()

        Coll_Lingue.Add(New ItemLanguage(Localization.Resource_WND_Languages.Desc_SystemDefaultLanguage, Dict_DatiAutori.Item("SystemDefault")(0), Dict_DatiAutori.Item("SystemDefault")(1), ""))
        For Each lingua In LingueSupportate
            If lingua.TwoLetterISOLanguageName = "iv" Then Continue For 'iv = invariant language, è la lingua fallback nel caso mancano risorse (stringhe mancanti o file lingua mancante) si può ricopiare dall'inglese
            Coll_Lingue.Add(New ItemLanguage(lingua.EnglishName + " [" + lingua.TwoLetterISOLanguageName + "]", Dict_DatiAutori.Item(lingua.EnglishName)(0), Dict_DatiAutori.Item(lingua.EnglishName)(1), lingua.TwoLetterISOLanguageName))
        Next

        DG_Lingue.ItemsSource = Coll_Lingue

        For Each cl As ItemLanguage In DG_Lingue.Items
            If XMLSettings_Read("Language").Equals(cl.EnglishName) Then
                DG_Lingue.SelectedItem = cl
            End If
        Next
    End Sub

    Private Sub BTN_Applica_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Applica.Click
        If DG_Lingue.SelectedItems.Count = 0 Then
            MsgBox(Autorinomina.Localization.Resource_WND_Languages.Msg_SelectLanguage, MsgBoxStyle.Exclamation)
            Return
        End If

        Dim itemSel As ItemLanguage = DG_Lingue.SelectedItem
        XMLSettings_Save("Language", itemSel.EnglishName)

        MsgBox(Localization.Resource_WND_Languages.Msg_RebootApp, MsgBoxStyle.Information)

        Close()
    End Sub

    Private Sub BTN_Annulla_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Annulla.Click
        Close()
    End Sub

    Private Sub BTN_VuoiTradurre_Click(sender As Object, e As RoutedEventArgs) Handles BTN_VuoiTradurre.Click
        Try
            System.Diagnostics.Process.Start("http://www.autorinomina.it/index.php/contribuisci/come-puoi-aiutare")
        Catch ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, BTN_VuoiTradurre.Content.ToString)
        End Try
    End Sub
End Class
