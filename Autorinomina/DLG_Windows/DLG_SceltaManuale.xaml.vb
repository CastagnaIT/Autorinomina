Public Class DLG_SceltaManuale

    Dim ARCore_SerieTv As New AR_Core_SerieTv(Me)
    Private Collection_Temp As New ObjectModel.ObservableCollection(Of Coll_Temp)()

    Class Coll_Temp
        Public Sub New(ByVal _RisultatoView As String, ByVal _Risultato As String, ByVal _Metodo As String)
            RisultatoView = _RisultatoView
            Risultato = _Risultato
            Metodo = _Metodo
        End Sub

        Public Property RisultatoView As String
        Public Property Risultato As String
        Public Property Metodo As String
    End Class


    Public Property Path As String

    Public Property ResultData As String() = Nothing
    Public Property ApplicaTutti As Boolean = False

    Protected Overrides Sub OnSourceInitialized(e As EventArgs)
        CustomWindow(Me, True)
    End Sub

    Public Sub New(ByVal _path As String)

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        Path = _path
    End Sub

    Private Sub Window_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Dim NomeDelFile As String = IO.Path.GetFileName(Path)
        Image1.Source = EstraiIconaAssociata(Path, False)(0)

        TextBlock2.Text = Localization.Resource_Common_Dialogs.DLG_ManualFilter_FilePath & Space(1) & IO.Path.GetDirectoryName(Path)
        TextBlock2.ToolTip = IO.Path.GetDirectoryName(Path)
        TextBlock1.Text = IO.Path.GetFileName(Path)
        TextBlock1.ToolTip = IO.Path.GetFileName(Path)
        Label11.Content = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_ManualFilter_FileSize & " - " & FormatBYTE(New IO.FileInfo(Path).Length)


        Dim Filtri As String() = {" ", ".", "_", "-", "_."}

        For Each TipoFiltro As String In Filtri
            Dim ResultTemp As String = ""
            Try
                ResultTemp = ARCore_SerieTv.RunCore(Path, 0, True, 0, TipoFiltro)
            Catch ex As Exception

            Finally
                If Not String.IsNullOrEmpty(ResultTemp) Then Collection_Temp.Add(New Coll_Temp(ResultTemp, "", TipoFiltro))
            End Try
        Next

        If Collection_Temp.Count = 0 Then
            Collection_Temp.Add(New Coll_Temp(AutoRinomina.Localization.Resource_Common_Dialogs.DLG_ManualFilter_NoResults, NomeDelFile, "#"))
        End If

        LB_RisultatiFiltri.ItemsSource = Collection_Temp
        LB_RisultatiFiltri.SelectedIndex = 0
        LB_RisultatiFiltri.Focus()
    End Sub


    Private Sub BTN_Comferma_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BTN_Comferma.Click
        Me.DialogResult = True
        ResultData = {"True", CType(LB_RisultatiFiltri.SelectedItem, Coll_Temp).Metodo, CB_ApplicaTutti.IsChecked.ToString}
        Close()
    End Sub

    Private Sub BTN_Salta_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BTN_Salta.Click
        Me.DialogResult = False
        ResultData = {"False", "", CB_ApplicaTutti.IsChecked.ToString} 'il filtro vuoto "" annulla il filtro e restituisce validazione False
        Close()
    End Sub

    Private Sub BTN_Interrompi_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BTN_Interrompi.Click
        Me.DialogResult = False
        ResultData = {Nothing}
        Close()
    End Sub
End Class
