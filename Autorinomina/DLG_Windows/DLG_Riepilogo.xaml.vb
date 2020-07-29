Public Class DLG_Riepilogo
    Dim _Val As Integer()
    Dim _TipoRiepilogo As Byte

    Protected Overrides Sub OnSourceInitialized(e As EventArgs)
        CustomWindow(Me, True)
    End Sub

    Public Sub New(TipoRiepilogo As Byte, Val As Integer())
        _Val = Val
        _TipoRiepilogo = TipoRiepilogo

        ' La chiamata è richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
    End Sub

    Private Sub DLG_Riepilogo_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Dim Riepilogo_Txt As String = ""
        Dim Riepilogo_Values As String = ""
        Dim FileElaboratiPositivi As Integer


        If _TipoRiepilogo = 0 Then
            'anteprima
            FileElaboratiPositivi = Coll_Files.Count - (_Val(ENUM_PREVIEW_RESULT.ERRORI) + _Val(ENUM_PREVIEW_RESULT.CONFLITTI) + _Val(ENUM_PREVIEW_RESULT.NOANTEPRIMA))

            TB_Desc1.Text = Localization.Resource_Common_Dialogs.DLG_Riepilogo_Desc4
            Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_PreviewMsg1 & vbCrLf & vbCrLf
            Riepilogo_Values &= FileElaboratiPositivi & "/" & Coll_Files.Count & vbCrLf & vbCrLf

            If _Val(ENUM_PREVIEW_RESULT.ERRORI) > 0 Then
                Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_PreviewMsg2 & vbCrLf
                Riepilogo_Values &= _Val(ENUM_PREVIEW_RESULT.ERRORI) & vbCrLf
            End If

            If _Val(ENUM_PREVIEW_RESULT.CONFLITTI) > 0 Then
                Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_PreviewMsg3 & vbCrLf
                Riepilogo_Values &= _Val(ENUM_PREVIEW_RESULT.CONFLITTI) & vbCrLf
            End If

            If _Val(ENUM_PREVIEW_RESULT.NOANTEPRIMA) > 0 Then
                Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_PreviewMsg4 & vbCrLf
                Riepilogo_Values &= _Val(ENUM_PREVIEW_RESULT.NOANTEPRIMA) & vbCrLf
            End If

            BTN_Chiudi.Visibility = Visibility.Collapsed

        ElseIf _TipoRiepilogo = 1 Then
            'rinomina
            FileElaboratiPositivi = Coll_Files.Count - (_Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI) + _Val(ENUM_RENAME_RESULT.CONFLITTI) + _Val(ENUM_RENAME_RESULT.ERRORI) + _Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI_BYUSER))

            TB_Desc1.Text = Localization.Resource_Common_Dialogs.DLG_Riepilogo_Desc2

            Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_RenameMsg1 & vbCrLf & vbCrLf
            Riepilogo_Values &= FileElaboratiPositivi & "/" & Coll_Files.Count & vbCrLf & vbCrLf

            If _Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI_BYUSER) > 0 Then
                Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_RenameMsg2 & vbCrLf
                Riepilogo_Values &= _Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI_BYUSER) & vbCrLf
            End If

            If _Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI) > 0 Then
                Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_RenameMsg3 & vbCrLf
                Riepilogo_Values &= _Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI) & vbCrLf
            End If

            If _Val(ENUM_RENAME_RESULT.CONFLITTI) > 0 Then
                Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_RenameMsg4 & vbCrLf
                Riepilogo_Values &= _Val(ENUM_RENAME_RESULT.CONFLITTI) & vbCrLf
            End If

            If _Val(ENUM_RENAME_RESULT.ERRORI) > 0 Then
                Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_RenameMsg5 & vbCrLf
                Riepilogo_Values &= _Val(ENUM_RENAME_RESULT.ERRORI) & vbCrLf
            End If

        Else
            'ripristina
            FileElaboratiPositivi = Coll_Files.Count - (_Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI) + _Val(ENUM_RENAME_RESULT.CONFLITTI) + _Val(ENUM_RENAME_RESULT.ERRORI) + _Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI_BYUSER))

            TB_Desc1.Text = Localization.Resource_Common_Dialogs.DLG_Riepilogo_Desc1

            Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_RestoreMsg1 & vbCrLf & vbCrLf
            Riepilogo_Values &= FileElaboratiPositivi & "/" & Coll_Files.Count & vbCrLf & vbCrLf

            If _Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI_BYUSER) > 0 Then
                Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_RestoreMsg2 & vbCrLf
                Riepilogo_Values &= _Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI_BYUSER) & vbCrLf
            End If

            If _Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI) > 0 Then
                Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_RestoreMsg3 & vbCrLf
                Riepilogo_Values &= _Val(ENUM_RENAME_RESULT.NONRINOMINATI_RIPRISTINATI) & vbCrLf
            End If

            If _Val(ENUM_RENAME_RESULT.CONFLITTI) > 0 Then
                Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_RestoreMsg4 & vbCrLf
                Riepilogo_Values &= _Val(ENUM_RENAME_RESULT.CONFLITTI) & vbCrLf
            End If

            If _Val(ENUM_RENAME_RESULT.ERRORI) > 0 Then
                Riepilogo_Txt &= Localization.Resource_Common_Dialogs.DLG_Riepilogo_RestoreMsg5 & vbCrLf
                Riepilogo_Values &= _Val(ENUM_RENAME_RESULT.ERRORI) & vbCrLf
            End If

        End If

        TB_Desc2.Text = Localization.Resource_Common_Dialogs.DLG_Riepilogo_Desc3

        TB_Details.Text = Riepilogo_Txt
        TB_Values.Text = Riepilogo_Values

        If FileElaboratiPositivi = Coll_Files.Count Then
            TB_Desc2.Visibility = Visibility.Collapsed
        Else
            IMG_Stato.Source = New BitmapImage(New Uri("/Autorinomina;component/Immagini/Riepilogo_problems.png", UriKind.Relative))
        End If

    End Sub

    Private Sub BTN_VisualizzaRisultati_Click(sender As Object, e As RoutedEventArgs) Handles BTN_VisualizzaRisultati.Click
        Me.DialogResult = True
        Close()
    End Sub

    Private Sub BTN_Chiudi_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Chiudi.Click
        Me.DialogResult = False
        Close()
    End Sub

End Class
