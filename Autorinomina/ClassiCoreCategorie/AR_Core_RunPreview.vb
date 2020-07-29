Imports System.ComponentModel

Public Class AR_Core_RunPreview
    WithEvents BW_Anterprima As New BackgroundWorker
    Private _WindowParent As Window
    Private _LivePreview As Boolean
    Private _Coll_Files As ItemCollection
    Private _Coll_Files_Selected As IList
    Private _OverrideFileCount As Integer

    Public Sub New(WindowParent As Window)
        _WindowParent = WindowParent
    End Sub

    Public Event ProgressChanged As EventHandler(Of String)
    Public Event RunCompleted As EventHandler(Of System.ComponentModel.RunWorkerCompletedEventArgs)


    Public Sub StopPreview()
        BW_Anterprima.CancelAsync()
    End Sub

    Public Sub StartPreview(ByRef Coll_Files As ItemCollection, Optional ByRef Coll_Files_Selected As IList = Nothing, Optional LivePreview As Boolean = False)
        _Coll_Files = Coll_Files
        _Coll_Files_Selected = Coll_Files_Selected
        _LivePreview = LivePreview
        If LivePreview Then Coll_Files_OnPropertyChanged_Enabled = True

        BW_Anterprima.WorkerSupportsCancellation = True
        BW_Anterprima.WorkerReportsProgress = True
        BW_Anterprima.RunWorkerAsync()
    End Sub

    Private Sub BW_Anterprima_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BW_Anterprima.DoWork
        Dim worker As BackgroundWorker = sender

        Dim CategoriaSelezionata As String = XMLSettings_Read("CategoriaSelezionata")
        Dim PreviewCheckLocalFile As Boolean = XMLSettings_Read("PreviewCheckLocalFile")
        Dim Val(ENUM_PREVIEW_RESULT.Count - 1) As Integer
        Dim nTotale As Integer = _Coll_Files.Count

        Dim AR_Core As AR_Core = Nothing

        Select Case CategoriaSelezionata
            Case "CATEGORIA_serietv"
                AR_Core = New AR_Core_SerieTv(_WindowParent)

            Case "CATEGORIA_video"
                AR_Core = New AR_Core_Video(_WindowParent)

            Case "CATEGORIA_immagini"
                AR_Core = New AR_Core_Immagini(_WindowParent)

            Case "CATEGORIA_audio"
                AR_Core = New AR_Core_Audio(_WindowParent)

            Case "CATEGORIA_generica"
                AR_Core = New AR_Core_Generica(_WindowParent)
        End Select

        Dim Temp_Files_Processed As New List(Of String)
        Dim TotCountColl As Integer
        If _Coll_Files_Selected IsNot Nothing Then
            TotCountColl = IIf(_Coll_Files_Selected.Count > 3, 3, _Coll_Files_Selected.Count - 1) 'LIMITO IL LIVE PREVIEW A 3 FILES
        Else
            TotCountColl = _Coll_Files.Count - 1

            Coll_BlackList_New.Clear()
            If CategoriaSelezionata.Equals("CATEGORIA_serietv") Then
                If XMLSettings_Read("PreviewHintsNewBlackListWords").Equals("True") Then CType(AR_Core, AR_Core_SerieTv).RicercaNuoveParoleBlackList(Coll_BlackList_New, XMLSettings_Read("PreviewHintsNewBlackListWords_Sensibility"))
            End If
        End If

        For n As Integer = 0 To TotCountColl
            If worker.CancellationPending Then e.Cancel = True : Return

            Dim N_Index As Integer
            Dim IFD As ItemFileData
            If _Coll_Files_Selected IsNot Nothing Then
                N_Index = _Coll_Files.IndexOf(_Coll_Files_Selected(n))
                IFD = _Coll_Files_Selected(n)
            Else
                N_Index = n
                IFD = _Coll_Files(n)
            End If

            Dim PercParziale As Double = ((N_Index + 1) / nTotale) * 100
            Dim result As String = ""
            worker.ReportProgress(PercParziale)

            IFD.StatoInfo = Nothing 'reset a causa della livepreview
            IFD.Stato = FileDataInfoStato.NULLO 'reset a causa della livepreview

            Try
                'Dim SW As Stopwatch = Stopwatch.StartNew
                result = AR_Core.RunCore(IO.Path.Combine(IFD.Percorso, IFD.NomeFile), N_Index, _LivePreview, nTotale.ToString.Length)
                'SW.Stop()
                'Debug.Print(SW.ElapsedMilliseconds)
            Catch uie As UserInterruptedException
                'interruzione da parte dell'utente
                IFD.StatoInfo = Localization.Resource_Common.StateFileInfo_UserInterrupt
                IFD.Stato = FileDataInfoStato.ERRORE
                e.Cancel = True
                Return

            Catch arce As ARCoreException
                'errori gestibili nell'elaborazione dei dati
                IFD.StatoInfo = arce.Message
                IFD.Stato = FileDataInfoStato.ERRORE
                IFD.NomeFileRinominato = ""
                Val(ENUM_PREVIEW_RESULT.ERRORI) += 1

            Catch ex As Exception
                'errori gravi
                IFD.StatoInfo = ex.Message
                IFD.Stato = FileDataInfoStato.ERRORE
                Throw New Exception(ex.Message, ex.InnerException)

            Finally

                If String.IsNullOrEmpty(result) Then
                    If IFD.Stato = FileDataInfoStato.NULLO Then
                        IFD.StatoInfo = Localization.Resource_Common.StateFileInfo_NoPreviewBySettings
                        IFD.Stato = FileDataInfoStato.ANTEPRIMA_ERRORE
                        IFD.NomeFileRinominato = ""
                        Val(ENUM_PREVIEW_RESULT.NOANTEPRIMA) += 1
                    End If
                Else
                    Dim filenameAttuale As String = result

                    'controllo se è già stato rinominato un file (localmente) con lo stesso nome
                    If Temp_Files_Processed.Contains(IO.Path.Combine(IFD.Percorso, filenameAttuale)) Then
                        Dim N_Duplicate As Integer = 1
                        While Temp_Files_Processed.Contains(IO.Path.Combine(IFD.Percorso, IO.Path.GetFileNameWithoutExtension(filenameAttuale)) & " (" & N_Duplicate & ")" & IO.Path.GetExtension(filenameAttuale))
                            N_Duplicate += 1
                        End While

                        IFD.NomeFileRinominato = IO.Path.GetFileNameWithoutExtension(filenameAttuale) & " (" & N_Duplicate & ")" & IO.Path.GetExtension(filenameAttuale)
                        IFD.StatoInfo = Localization.Resource_Common.StateFileInfo_FileAlreadyRenamedWith
                        IFD.Stato = FileDataInfoStato.CONFLITTO
                        Val(ENUM_PREVIEW_RESULT.CONFLITTI) += 1
                    Else

                        If PreviewCheckLocalFile Then
                            'controllo se esiste già un file con lo stesso nome nella cartella di origine
                            If IO.File.Exists(IO.Path.Combine(IFD.Percorso, filenameAttuale)) Then

                                Dim N_Duplicate As Integer = 1
                                While IO.File.Exists(IO.Path.Combine(IFD.Percorso, IO.Path.GetFileNameWithoutExtension(filenameAttuale)) & " (" & N_Duplicate & ")" & IO.Path.GetExtension(filenameAttuale))
                                    N_Duplicate += 1
                                End While

                                'ricontrollo se c'è già localmente
                                Dim EsisteLocalmente As Boolean = False
                                While Temp_Files_Processed.Contains(IO.Path.Combine(IFD.Percorso, IO.Path.GetFileNameWithoutExtension(filenameAttuale)) & " (" & N_Duplicate & ")" & IO.Path.GetExtension(filenameAttuale))
                                    N_Duplicate += 1
                                    EsisteLocalmente = True
                                End While

                                IFD.NomeFileRinominato = IO.Path.GetFileNameWithoutExtension(filenameAttuale) & " (" & N_Duplicate & ")" & IO.Path.GetExtension(filenameAttuale)
                                If EsisteLocalmente Then
                                    IFD.StatoInfo = Localization.Resource_Common.StateFileInfo_FileAlreadyRenamedWith
                                Else
                                    IFD.StatoInfo = Localization.Resource_Common.StateFileInfo_FileAlreadyExist
                                End If
                                IFD.Stato = FileDataInfoStato.CONFLITTO
                                Val(ENUM_PREVIEW_RESULT.CONFLITTI) += 1
                            Else
                                If _LivePreview = False Then IFD.StatoInfo = Localization.Resource_Common.StateFileInfo_PreviewOK
                                IFD.NomeFileRinominato = filenameAttuale
                                IFD.Stato = FileDataInfoStato.ANTEPRIMA_OK
                            End If
                        Else
                            If _LivePreview = False Then IFD.StatoInfo = Localization.Resource_Common.StateFileInfo_PreviewOK
                            IFD.NomeFileRinominato = filenameAttuale
                            IFD.Stato = FileDataInfoStato.ANTEPRIMA_OK
                        End If

                    End If


                    Temp_Files_Processed.Add(IO.Path.Combine(IFD.Percorso, IFD.NomeFileRinominato))
                    End If
            End Try

        Next

        e.Result = Val
    End Sub


    Private Sub BW_Anterprima_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BW_Anterprima.ProgressChanged
        RaiseEvent ProgressChanged(Me, e.ProgressPercentage & " %")
    End Sub

    Private Sub BW_Anterprima_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BW_Anterprima.RunWorkerCompleted
        If _LivePreview Then Coll_Files_OnPropertyChanged_Enabled = False
        RaiseEvent RunCompleted(Me, e)
    End Sub

End Class
