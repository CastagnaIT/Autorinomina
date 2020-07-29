Imports System.ComponentModel

Public Class DLG_Rinomina
    Public Shadows DialogResult As Boolean '<< fix errore DialogResult in chiusura della finestra mentre si annulla il BW
    'Public Property ErrorExceptionMessage As String = Nothing
    WithEvents BW_Rinomina As New BackgroundWorker
    Dim _Coll_Files As Object
    Dim _RipristinaFiles As Boolean
    Public Property RenameErrorsResults As Integer() = Nothing

    Private Delegate Function Delegate_GeneralFunc(ByVal ParameterToPass As String, ByVal ParameterToPass2 As String) As String()

    Private Function Open_DLGFileEsistente(ByVal PathFile As String, ByVal nuovoNome As String) As String()
        Dim WND As New DLG_FileEsistente(PathFile, nuovoNome)
        WND.ShowDialog()
        Return WND.ResultData
    End Function

    Public Sub New(ByRef Coll_Files As Object, ByVal RipristinaFiles As Boolean)

        ' La chiamata è richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        _Coll_Files = Coll_Files
        _RipristinaFiles = RipristinaFiles
    End Sub

    Private Sub DLG_Rinomina_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        If _RipristinaFiles Then
            TBlock_Desc.Text = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_Rename_Desc_Restore
            Me.Title = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_Rename_Desc_Restore
        Else
            TBlock_Desc.Text = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_Rename_Desc_Rename
            Me.Title = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_Rename_Desc_Rename
        End If
    End Sub

    Private Sub DLG_Anteprima_ContentRendered(sender As Object, e As System.EventArgs) Handles Me.ContentRendered
        BW_Rinomina.WorkerSupportsCancellation = True
        BW_Rinomina.WorkerReportsProgress = True
        BW_Rinomina.RunWorkerAsync()
    End Sub

    Private Sub BW_Rinomina_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BW_Rinomina.DoWork
        Dim worker As BackgroundWorker = sender

        Dim Val(ENUM_RENAME_RESULT.Count - 1) As Integer
        Dim nSequenza As Integer = 1
        Dim nTotale As Integer = _Coll_Files.Count
        Dim CategoriaSelezionata As String = XMLSettings_Read("CategoriaSelezionata")
        'inizio var software esterno
        Dim PrgEsterno_Attivato As Boolean = Boolean.Parse(XMLSettings_Read("ExternalSoftware_Enabled"))
        Dim PrgEsterno_RunMinimized As Boolean = Boolean.Parse(XMLSettings_Read("ExternalSoftware_RunMinimized"))
        Dim PrgEsterno_Category As String = XMLSettings_Read("ExternalSoftware_Category")
        Dim PrgEsterno_ExePath As String = XMLSettings_Read("ExternalSoftware_ExePath")
        Dim PrgEsterno_CommandLine As String = XMLSettings_Read("ExternalSoftware_CommandLine")
        'fine var software esterno

        For Each IFD As ItemFileData In _Coll_Files
            If worker.CancellationPending Then e.Cancel = True : Return
            Dim PercParziale As Double = (nSequenza / nTotale) * 100
            worker.ReportProgress(PercParziale)

            nSequenza += 1

            Dim NomeOriginale As String = ""
            Dim NomeRinominato As String = ""

            If (_RipristinaFiles) Then
                NomeOriginale = IFD.NomeFileRinominato
                NomeRinominato = IFD.NomeFile

                'se il nuovo nome non c'è salta il file
                If String.IsNullOrEmpty(NomeOriginale) Then
                    IFD.StatoInfo = AutoRinomina.Localization.Resource_Common.StateFileInfo_NotRestored
                    IFD.Stato = FileDataInfoStato.CONFLITTO
                    Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI) += 1
                    Continue For
                End If
            Else
                NomeOriginale = IFD.NomeFile
                NomeRinominato = IFD.NomeFileRinominato

                'se il nuovo nome non c'è salta il file
                If String.IsNullOrEmpty(NomeRinominato) Then
                    IFD.StatoInfo = AutoRinomina.Localization.Resource_Common.StateFileInfo_NotRenamed
                    IFD.Stato = FileDataInfoStato.CONFLITTO
                    Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI) += 1
                    Continue For
                End If
            End If


            'se non esiste più il file ripristinato probabilmente è gia stato ripristinato
            If Not IO.File.Exists(IO.Path.Combine(IFD.Percorso, NomeOriginale)) Then
                If IO.File.Exists(IO.Path.Combine(IFD.Percorso, NomeRinominato)) Then

                    If _RipristinaFiles Then
                        IFD.StatoInfo = AutoRinomina.Localization.Resource_Common.StateFileInfo_AlreadyRestored
                        IFD.Stato = FileDataInfoStato.RIPRISTINATO
                    Else
                        IFD.StatoInfo = AutoRinomina.Localization.Resource_Common.StateFileInfo_AlreadyRenamed
                        IFD.Stato = FileDataInfoStato.RINOMINATO
                    End If
                    Continue For
                Else
                    IFD.StatoInfo = AutoRinomina.Localization.Resource_Common.StateFileInfo_NotExist
                    IFD.Stato = FileDataInfoStato.ERRORE
                    Val(ENUM_RENAME_RESULT.ERRORI) += 1
                    Continue For
                End If
            End If


            'controllo se il nuovo nome del file esiste già
            If String.IsNullOrEmpty(NomeRinominato) Then
                'non ha il nuovo nome
                IFD.StatoInfo = AutoRinomina.Localization.Resource_Common.StateFileInfo_NameMissing
                IFD.Stato = FileDataInfoStato.ERRORE
                Val(ENUM_RENAME_RESULT.ERRORI) += 1
                Continue For
            Else
                'procedo

                If Not NomeOriginale.Equals(NomeRinominato, StringComparison.OrdinalIgnoreCase) Then
                    If IO.File.Exists(IO.Path.Combine(IFD.Percorso, NomeRinominato)) Then
                        'il nome di destinazione esiste già
                        Dim del As New Delegate_GeneralFunc(AddressOf Open_DLGFileEsistente)
                        Dim Result() As String = Me.Dispatcher.Invoke(del, IO.Path.Combine(IFD.Percorso, NomeOriginale), NomeRinominato)

                        Select Case Result(0)
                            Case "0"
                                'sovrascrivi
                                IO.File.Delete(IO.Path.Combine(IFD.Percorso, NomeRinominato))

                            Case "1"
                                'non rinominare
                                NomeRinominato = Nothing

                            Case "2"
                                'mantieni entrambi i file
                                NomeRinominato = Result(1)

                        End Select
                    End If
                End If
            End If

            If String.IsNullOrEmpty(NomeRinominato) Then
                'si è scelto di non rinominare
                IFD.StatoInfo = AutoRinomina.Localization.Resource_Common.StateFileInfo_NotRenamedUserChoice
                IFD.Stato = FileDataInfoStato.NULLO
                Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI_BYUSER) += 1
            Else
                'rinomino
                Try
                    If Not NomeOriginale.Equals(NomeRinominato, StringComparison.OrdinalIgnoreCase) Then
                        My.Computer.FileSystem.RenameFile(IO.Path.Combine(IFD.Percorso, NomeOriginale), NomeRinominato)
                    Else
                        'controllo se i nomi hanno caratteri maiuscoli/minuscoli diversi
                        If Not NomeOriginale.Equals(NomeRinominato) Then
                            My.Computer.FileSystem.RenameFile(IO.Path.Combine(IFD.Percorso, NomeOriginale), "TMP_" & NomeRinominato)
                            My.Computer.FileSystem.RenameFile(IO.Path.Combine(IFD.Percorso, "TMP_" & NomeRinominato), NomeRinominato)
                        End If
                    End If

                    If _RipristinaFiles = False AndAlso PrgEsterno_Attivato AndAlso (PrgEsterno_Category.Equals(CategoriaSelezionata) OrElse String.IsNullOrEmpty(PrgEsterno_Category)) Then
                        EseguiSoftwareEsterno(PrgEsterno_ExePath, PrgEsterno_CommandLine, IO.Path.Combine(IFD.Percorso, NomeRinominato), PrgEsterno_RunMinimized)
                    End If


                    If NomeOriginale.Equals(NomeRinominato) Then
                        IFD.Stato = FileDataInfoStato.CONFLITTO
                        IFD.StatoInfo = Autorinomina.Localization.Resource_Common.StateFileInfo_SkippedSameName
                        Val(ENUM_RENAME_RESULT.CONFLITTI) += 1
                    Else
                        If _RipristinaFiles Then
                            IFD.Stato = FileDataInfoStato.RIPRISTINATO
                            IFD.StatoInfo = AutoRinomina.Localization.Resource_Common.StateFileInfo_Restored
                        Else
                            IFD.Stato = FileDataInfoStato.RINOMINATO
                            IFD.StatoInfo = AutoRinomina.Localization.Resource_Common.StateFileInfo_Renamed
                        End If
                    End If

                Catch ex As Exception

                    IFD.Stato = FileDataInfoStato.ERRORE
                    IFD.StatoInfo = ex.Message
                    Val(ENUM_RENAME_RESULT.ERRORI) += 1
                End Try
            End If
        Next

        e.Result = Val
    End Sub

    Private Sub EseguiSoftwareEsterno(ExePath As String, CommandLine As String, FilePath As String, RunMinimized As Boolean)
        Dim process As System.Diagnostics.Process = Nothing
        Dim processStartInfo As New System.Diagnostics.ProcessStartInfo()

        Try
            processStartInfo.FileName = ExePath

            '  If System.Environment.OSVersion.Version.Major >= 6 Then ' Windows Vista or higher
            'processStartInfo.Verb = "runas"
            'Else
            '' No need to prompt to run as admin
            'End If

            processStartInfo.Arguments = CommandLine.Replace("@", FilePath)
            If RunMinimized Then
                processStartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized
            Else
                processStartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
            End If
            processStartInfo.UseShellExecute = True

            process = System.Diagnostics.Process.Start(processStartInfo)
            process.WaitForExit()

        Catch ex As Exception
            Throw New Exception(ex.Message, ex.InnerException)

        Finally
            If Not (process Is Nothing) Then
                process.Dispose()
            End If
        End Try
    End Sub


    Private Sub BW_Rinomina_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BW_Rinomina.ProgressChanged
        TB_Percentuale.Text = e.ProgressPercentage & " %"
    End Sub

    Private Sub BW_Rinomina_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BW_Rinomina.RunWorkerCompleted
        If (e.Error IsNot Nothing) Then
            Me.DialogResult = False
            'ErrorExceptionMessage = e.Error.Message
        Else
            If e.Cancelled Then
                Me.DialogResult = False
            Else
                RenameErrorsResults = e.Result
                Me.DialogResult = True
            End If
        End If

        Close()
    End Sub

    Private Sub BTN_Annulla_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Annulla.Click
        BW_Rinomina.CancelAsync()
        BTN_Annulla.IsEnabled = False
    End Sub

End Class
