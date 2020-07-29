Public Class DLG_IncollaWikipedia
    Dim Elenco As String()
    Public Property NColonna As Integer = -1
    Public Property IncollaNomeColonna As Boolean = False

    Protected Overrides Sub OnSourceInitialized(e As EventArgs)
        CustomWindow(Me, True)
    End Sub

    Public Sub New(_Elenco As String())

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        Elenco = _Elenco
    End Sub

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded

        For x As Integer = 1 To Elenco.Count
            LB_Elenco.Items.Add("[Col. " & x & "] " & Elenco(x - 1).Trim)
        Next


        LB_Elenco.SelectedIndex = 0
        LB_Elenco.Focus()
    End Sub


    Private Sub BTN_Ok_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Ok.Click
        If LB_Elenco.SelectedIndex = -1 Then
            MsgBox(Localization.Resource_Common_Dialogs.DLG_PasteWiki_SelectHeaderColumn, MsgBoxStyle.Exclamation, BTN_Ok.Content.ToString)
            Return
        End If

        NColonna = LB_Elenco.SelectedIndex
        IncollaNomeColonna = CB_NomeColonna.IsChecked

        Me.DialogResult = True
        Close()
    End Sub

    Private Sub BTN_Annulla_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Annulla.Click
        Me.DialogResult = False
        Close()
    End Sub

    Private Sub LB_Item_MouseDoubleClick(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
        BTN_Ok_Click(Me, Nothing)
    End Sub

End Class
