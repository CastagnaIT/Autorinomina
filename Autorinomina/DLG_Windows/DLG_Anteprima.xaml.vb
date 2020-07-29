Imports System.ComponentModel

Public Class DLG_Anteprima
    Public Shadows DialogResult As Boolean '<< fix errore DialogResult in chiusura della finestra mentre si annulla il BW
    Public Property ErrorExceptionMessage As String = Nothing
    Public Property PreviewErrorsResults As Integer() = Nothing
    Dim AR_Preview As New AR_Core_RunPreview(Me)
    Dim _CollectionFiles As ItemCollection

    Public Sub New(CollectionFiles As ItemCollection)

        ' La chiamata è richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        _CollectionFiles = CollectionFiles
    End Sub

    Private Sub BTN_Annulla_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Annulla.Click
        AR_Preview.StopPreview()
        BTN_Annulla.IsEnabled = False
    End Sub

    Private Sub DLG_Anteprima_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        AddHandler AR_Preview.ProgressChanged, AddressOf ProgressChanged
        AddHandler AR_Preview.RunCompleted, AddressOf RunCompleted

        AR_Preview.StartPreview(_CollectionFiles, Nothing, False)
    End Sub

    Private Sub ProgressChanged(sender As Object, e As String)
        TB_Percentuale.Text = e
    End Sub

    Private Sub RunCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs)
        If (e.Error IsNot Nothing) Then
            Me.DialogResult = False
            ErrorExceptionMessage = e.Error.Message
        Else
            If e.Cancelled Then
                Me.DialogResult = False
            Else
                PreviewErrorsResults = e.Result
                Me.DialogResult = True
            End If
        End If

        Close()
    End Sub

End Class
