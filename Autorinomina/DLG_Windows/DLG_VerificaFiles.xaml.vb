Imports System.ComponentModel
Imports System.IO

Public Class DLG_VerificaFiles
    Public Shadows DialogResult As Boolean '<< fix errore DialogResult in chiusura della finestra mentre si annulla il BW
    Public Property Message As String = Nothing
    Public Property ResultCollItems As CollItemsFilesData
    WithEvents BW_Verifica As New BackgroundWorker
    Dim FilesToCheck As Array

    Public Sub New(_filesToCheck As Array)

        ' La chiamata è richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        FilesToCheck = _filesToCheck
    End Sub

    Private Sub DLG_Verifica_ContentRendered(sender As Object, e As System.EventArgs) Handles Me.ContentRendered
        BW_Verifica.WorkerSupportsCancellation = True
        BW_Verifica.WorkerReportsProgress = True
        BW_Verifica.RunWorkerAsync(FilesToCheck)
    End Sub

    Private Sub BW_Verifica_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BW_Verifica.DoWork
        Dim worker As BackgroundWorker = sender

        Dim CategoriaSelezionata As String = XMLSettings_Read("CategoriaSelezionata")
        Dim FilesDaVerificare As Array = e.Argument()
        Dim nSequenza As Integer = 1
        Dim nTotale As Integer = Coll_Files.Count

        Dim GetExtsResult() As Object = GetExtensions(CategoriaSelezionata)
        Dim Estensioni As String = GetExtsResult(0)
        Dim EstensioniList As List(Of String) = GetExtsResult(1)
        Dim EsploraSottocartelle As Boolean = XMLSettings_Read("VerificaFilesIncludiSottocartelle")
        Dim ListaFiles As New ArrayList

        Dim ResultColl As New CollItemsFilesData

        'Aggiungo i files nella lista, solo se corrispondono alle estensioni della categoria
        Try
            For Each FS_item As String In FilesDaVerificare
                If worker.CancellationPending Then
                    e.Cancel = True
                    Return
                End If

                If (New IO.FileInfo(FS_item).Attributes And IO.FileAttributes.Directory) = IO.FileAttributes.Directory Then
                    'se è una directory

                    Dim Ricerca As New FileSearch(FS_item, EsploraSottocartelle)
                    Ricerca.Search(FS_item, Estensioni)
                    ListaFiles.AddRange(Ricerca.Files)

                Else
                    'se è un file

                    If EstensioniList.Contains(IO.Path.GetExtension(FS_item).ToLower.Replace(".", "")) OrElse (EstensioniList.Count = 1 AndAlso EstensioniList.Contains("*;")) Then
                        ListaFiles.Add(New IO.FileInfo(FS_item))
                    End If

                End If
            Next
        Catch ex As Exception
            e.Result = {Nothing, "", ex.Message}
            Return
        End Try

        If ListaFiles.Count = 0 AndAlso FilesDaVerificare.Length > 0 Then
            e.Result = {Nothing, AutoRinomina.Localization.Resource_Common_Dialogs.DLG_VerificationFiles_WrongCategory.Replace("\n", Environment.NewLine), ""}
            Return
        End If

        Try
            If ListaFiles.Count > 0 Then
                'Copio l'elenco definitivo dei files
                Dim IncludiFileNascosti As Boolean = Boolean.Parse(XMLSettings_Read("VerificaFilesIncludiNascosti"))
                Dim EscludiFileDB_CategoriaGenerica As Boolean = Boolean.Parse(XMLSettings_Read("VerificaFiles_CATEGORIA_generica_EscludiDB"))
                Dim IndexGiaPresente As Integer = Coll_Files.Count
                Dim nListaFilesTotale As Integer = ListaFiles.Count

                For x As Integer = 0 To ListaFiles.Count - 1
                    Dim Index As Integer = x
                    If worker.CancellationPending Then e.Cancel = True : Return
                    Dim PercParziale As Double = (Index / nListaFilesTotale) * 100
                    worker.ReportProgress(PercParziale)

                    Dim FileItinerato As IO.FileInfo = ListaFiles(Index)

                    If Coll_Files.FirstOrDefault(Function(i) IO.Path.Combine(i.Percorso, i.NomeFile).Equals(FileItinerato.FullName)) Is Nothing Then

                        If CategoriaSelezionata = "CATEGORIA_generica" Then
                            If FileItinerato.Extension.ToLower.Equals(".db") AndAlso EscludiFileDB_CategoriaGenerica = False Then Continue For
                        End If
                        If (FileItinerato.Attributes And IO.FileAttributes.Hidden) = IO.FileAttributes.Hidden AndAlso IncludiFileNascosti = False Then Continue For

                        ResultColl.Add(New ItemFileData(FileItinerato.FullName, x + IndexGiaPresente))
                    End If
                Next
            End If

        Catch ex As Exception
            e.Result = {Nothing, "", ex.Message}
            Return
        End Try

        e.Result = {ResultColl, "", ""}
    End Sub

    Private Sub BW_Verifica_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BW_Verifica.ProgressChanged
        TB_Percentuale.Text = e.ProgressPercentage & " %"
    End Sub

    Private Sub BW_Verifica_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BW_Verifica.RunWorkerCompleted
        If (e.Error IsNot Nothing) Then
            Me.DialogResult = False
        Else
            If e.Cancelled Then
                Me.DialogResult = False
            Else
                If Not String.IsNullOrEmpty(e.Result(2).ToString) Then
                    Message = e.Result(2).ToString
                    Me.DialogResult = False
                Else
                    Message = e.Result(1).ToString
                    Me.ResultCollItems = e.Result(0)
                    Me.DialogResult = True
                End If
            End If
        End If

        Close()
    End Sub

    Private Sub BTN_Annulla_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Annulla.Click
        BW_Verifica.CancelAsync()
        BTN_Annulla.IsEnabled = False
    End Sub

End Class
