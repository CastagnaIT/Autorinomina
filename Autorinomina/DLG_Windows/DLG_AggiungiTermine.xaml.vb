Imports System.Data

Public Class DLG_AggiungiTermine
    Dim DisabilitaAggiungiNuovo As Boolean
    Dim CollTermini As CollItemsTermini
    Dim TipoWND As String
    Dim CategoriaAttuale As String

    Protected Overrides Sub OnSourceInitialized(e As EventArgs)
        CustomWindow(Me, True)
    End Sub

    Public Sub New(_TipoWND As String, ByRef _CollTermini As CollItemsTermini, _CategoriaAttuale As String)

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        TipoWND = _TipoWND
        CollTermini = _CollTermini
        CategoriaAttuale = _CategoriaAttuale
    End Sub

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        If TipoWND = "BlackList" Then
            LB_Hint.Content = Autorinomina.Localization.Resource_Common_Dialogs.DLG_InsertTermine_BlackList_Tips
            LB_Sostituisci.Visibility = Windows.Visibility.Collapsed
            TB_TermineSostituto.IsEnabled = False
            TB_TermineSostituto.Visibility = Windows.Visibility.Collapsed
            Grid_Contenuto.RowDefinitions(1).Height = New GridLength(1)
        Else
            LB_Hint.Content = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_InsertTermine_Substitutions_Tips
        End If

        RB_Case_NonSensitivo.IsChecked = True

        CB_Termine.Focus()

        If CategoriaAttuale.Equals("CATEGORIA_serietv") Then CB_Termine.ItemsSource = Coll_BlackList_New
    End Sub

    Private Sub BTN_Annulla_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Annulla.Click
        Me.DialogResult = False
        Close()
    End Sub

    Private Sub BTN_Aggiungi_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Aggiungi.Click
        If TipoWND = "BlackList" Then
            CollTermini.Add(New ItemTermine(CB_Termine.Text, Nothing, RB_Case_Sensitivo.IsChecked.ToString))
            If CategoriaAttuale.Equals("CATEGORIA_serietv") Then
                Coll_BlackList_New.Remove(CB_Termine.Text) 'rimuovi dalla lista se fa parte dei termini suggeriti
            End If
        Else
            CollTermini.Add(New ItemTermine(CB_Termine.Text, TB_TermineSostituto.Text, RB_Case_Sensitivo.IsChecked.ToString))
        End If
        Me.DialogResult = True
        Close()
    End Sub

    Private Sub BTN_AggiungiNuovo_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_AggiungiNuovo.Click
        If TipoWND = "BlackList" Then
            CollTermini.Add(New ItemTermine(CB_Termine.Text, Nothing, RB_Case_Sensitivo.IsChecked.ToString))
            If CategoriaAttuale.Equals("CATEGORIA_serietv") Then
                Coll_BlackList_New.Remove(CB_Termine.Text) 'rimuovi dalla lista se fa parte dei termini suggeriti
                CB_Termine.Items.Refresh() 'gli items della cb non vengono aggiornati perchè il List<> non supporta OnPropertyChanged come le collection
            End If
        Else
            CollTermini.Add(New ItemTermine(CB_Termine.Text, TB_TermineSostituto.Text, RB_Case_Sensitivo.IsChecked.ToString))
        End If

        CB_Termine.Text = ""
        TB_TermineSostituto.Text = ""
        CB_Termine.Focus()
    End Sub

    Private Sub CB_Termine_TextChanged(sender As Object, e As TextChangedEventArgs)
        Dim bindExpr As BindingExpression = BindingOperations.GetBindingExpression(CB_Termine, ComboBox.TextProperty)
        Dim bindExprBase As BindingExpressionBase = BindingOperations.GetBindingExpressionBase(CB_Termine, ComboBox.TextProperty)
        Dim validationError As New ValidationError(New ExceptionValidationRule(), bindExpr)

        Dim ItemEsistente As Boolean = False
        Dim TextSensitive As String = ""
        If RB_Case_Sensitivo.IsChecked Then
            ItemEsistente = (CollTermini.Where(Function(item) item.Termine.Equals(CB_Termine.Text)).Count > 0)
            TextSensitive = RB_Case_Sensitivo.Content.ToString
        Else
            ItemEsistente = (CollTermini.Where(Function(item) item.Termine.Equals(CB_Termine.Text, StringComparison.OrdinalIgnoreCase)).Count > 0)
            TextSensitive = RB_Case_NonSensitivo.Content.ToString
        End If

        If ItemEsistente Then
            validationError.ErrorContent = Autorinomina.Localization.Resource_Common_Dialogs.DLG_InsertTermine_TermineAlreadyExists
            Validation.MarkInvalid(bindExprBase, validationError)
            BTN_Aggiungi.IsEnabled = False
            BTN_AggiungiNuovo.IsEnabled = False
        Else
            Validation.ClearInvalid(bindExprBase)
            BTN_Aggiungi.IsEnabled = True
            BTN_AggiungiNuovo.IsEnabled = True
        End If

    End Sub
End Class
