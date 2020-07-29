Public Class DLG_Aggiungi_TXT
    Public Property Testo As String = ""
    Public Property CreaNuovo As Boolean = False

    Protected Overrides Sub OnSourceInitialized(e As EventArgs)
        CustomWindow(Me, True)
    End Sub

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        If Not String.IsNullOrEmpty(Testo) Then
            Me.Title = "Modifica testo"
            BTN_Aggiungi.Content = "Modifica"

            BTN_AggiungiNuovo.Visibility = Windows.Visibility.Collapsed
            Dim NuoviMargini As Thickness = BTN_Aggiungi.Margin
            NuoviMargini.Right = 139
            BTN_Aggiungi.Margin = NuoviMargini
        End If

        TB_Testo.Text = Testo
        TB_Testo.Focus()
    End Sub

    Public Sub New(Optional _Testo As String = "")

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        Testo = _Testo
    End Sub

    Private Sub BTN_Annulla_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Annulla.Click
        Me.DialogResult = False
        Close()
    End Sub

    Private Sub BTN_Aggiungi_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Aggiungi.Click
        Testo = TB_Testo.Text
        Me.DialogResult = True
    End Sub

    Private Sub BTN_AggiungiNuovo_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_AggiungiNuovo.Click
        Testo = TB_Testo.Text
        CreaNuovo = True
        Me.DialogResult = True
    End Sub

    Private Sub TB_Testo_PreviewTextInput(sender As Object, e As System.Windows.Input.TextCompositionEventArgs) Handles TB_Testo.PreviewTextInput
        e.Handled = ValidateNoSpecialChar(e.Text, False)
    End Sub
End Class
