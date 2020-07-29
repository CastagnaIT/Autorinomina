Public Class DLG_ModificaManuale
    Public Property Path As String
    Public Property ResultData As String() = Nothing

    Protected Overrides Sub OnSourceInitialized(e As EventArgs)
        CustomWindow(Me, True)
    End Sub

    Private Sub Window_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Dim NomeDelFile As String = IO.Path.GetFileName(Path)
        Image1.Source = EstraiIconaAssociata(Path, False)(0)

        TextBlock2.Text = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_ManualFilter_FilePath & Space(1) & IO.Path.GetDirectoryName(Path)
        TextBlock2.ToolTip = IO.Path.GetDirectoryName(Path)
        TextBlock1.Text = IO.Path.GetFileName(Path)
        TextBlock1.ToolTip = IO.Path.GetFileName(Path)
        Label11.Content = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_ManualFilter_FileSize & " - " & FormatBYTE(New IO.FileInfo(Path).Length)

        TB_NewFilename.Text = IO.Path.GetFileNameWithoutExtension(Path)
        TB_NewFilename.SelectAll()
        TB_NewFilename.Focus()
    End Sub

    Public Sub New(ByVal _path As String)

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        Path = _path
    End Sub

    Private Sub BTN_Conferma_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BTN_Conferma.Click
        Me.DialogResult = True
        ResultData = {"True", TB_NewFilename.Text & IO.Path.GetExtension(Path)}
        Close()
    End Sub

    Private Sub BTN_Salta_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BTN_Salta.Click
        Me.DialogResult = False
        ResultData = {"False", ""}
        Close()
    End Sub

    Private Sub TB_NewFilename_PreviewTextInput(ByVal sender As System.Object, ByVal e As System.Windows.Input.TextCompositionEventArgs) Handles TB_NewFilename.PreviewTextInput
        e.Handled = ValidateNoSpecialChar(e.Text, False)
    End Sub

    Private Sub BTN_Interrompi_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Interrompi.Click
        Me.DialogResult = False
        ResultData = {Nothing}
        Close()
    End Sub
End Class
