Public Class WND_ElencoSequenziale
    Dim FilePath As String = DataPath & "\" & XMLSettings_Read("CategoriaSelezionata") & "_elencosequenziale.txt"

    Private Sub CheckStatus()
        Dim Temp As Integer
        For Each item As String In TB_Contenuto.Text.Split(vbCrLf)
            If item.Trim.Length > 0 Then
                Temp += 1
            End If
        Next

        Dim TXT As String = ""
        Dim TextColor As Brush = Brushes.Black
        If Coll_Files.Count = 0 Then
            TextColor = Brushes.Red
            TXT = Localization.Resource_WND_ElencoSequenziale.Status_nofile
        Else
            If (Coll_Files.Count - Temp) = 0 Then
                'l'elenco è completo
                TextColor = Brushes.Green
                TXT = Localization.Resource_WND_ElencoSequenziale.Status_full
            ElseIf (Coll_Files.Count - Temp) < 0 Then
                'ci sono righe in eccesso
                TextColor = Brushes.Red
                TXT = String.Format(Localization.Resource_WND_ElencoSequenziale.Status_surplus, (Temp - Coll_Files.Count))
            ElseIf Temp = 0 Then
                'non c'è nessuna riga nell'elenco
                TextColor = Brushes.Blue
                TXT = Localization.Resource_WND_ElencoSequenziale.Status_empty
            Else
                'mancano delle righe
                TextColor = Brushes.Red
                TXT = String.Format(Localization.Resource_WND_ElencoSequenziale.Status_missing, (Coll_Files.Count - Temp))
            End If
        End If

        LB_StatusDescrizione.Content = TXT
        LB_StatusDescrizione.Foreground = TextColor
    End Sub

    Private Sub PasteBinding_CanExecute(ByVal sender As Object, ByVal e As DataObjectPastingEventArgs)
        Dim ObjTxt As TextBox = CType(sender, TextBox)

        If e.SourceDataObject.GetDataPresent(System.Windows.DataFormats.Text, True) Then
            Dim Txt As String = e.SourceDataObject.GetData(System.Windows.DataFormats.Text)
            Dim OldCaretIndex As Integer = ObjTxt.CaretIndex

            If ObjTxt.SelectedText <> "" Then
                ObjTxt.SelectedText = ValidateChar_Replace(Txt)
            Else
                ObjTxt.Text = ObjTxt.Text.Insert(ObjTxt.CaretIndex, ValidateChar_Replace(Txt))
                ObjTxt.CaretIndex += OldCaretIndex + Txt.Length
            End If

        End If

        e.CancelCommand()
        e.Handled = True
    End Sub

    Private Sub TB_Contenuto_TextChanged(sender As Object, e As TextChangedEventArgs) Handles TB_Contenuto.TextChanged
        CheckStatus()
    End Sub

    Private Sub TB_Contenuto_PreviewTextInput(sender As Object, e As TextCompositionEventArgs) Handles TB_Contenuto.PreviewTextInput
        e.Handled = ValidateNoSpecialChar(e.Text, False)
    End Sub

    Private Sub WND_ElencoSequenziale_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        DataObject.AddPastingHandler(TB_Contenuto, New DataObjectPastingEventHandler(AddressOf PasteBinding_CanExecute))

        If IO.File.Exists(FilePath) Then
            Dim reader As IO.StreamReader = New IO.StreamReader(FilePath, System.Text.ASCIIEncoding.Unicode)
            Dim ContenutoTXT As String = ""
            Try
                ContenutoTXT = reader.ReadToEnd()
            Catch
                ContenutoTXT = ""
            Finally
                TB_Contenuto.Text = ContenutoTXT
                reader.Close()
            End Try
        End If

        CheckStatus()

        TB_Contenuto.Focus()
    End Sub

    Private Sub BTN_Apri_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Apri.Click
        Dim DlgOpen As New Microsoft.Win32.OpenFileDialog
        DlgOpen.DefaultExt = "*.txt"
        DlgOpen.Filter = "Text file (*.TXT)|*.txt|All files (*.*)|*.*"
        DlgOpen.FilterIndex = 0
        DlgOpen.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer).ToString
        DlgOpen.Multiselect = False
        DlgOpen.Title = BTN_Apri.Content.ToString

        If DlgOpen.ShowDialog = True Then

            Dim reader As IO.StreamReader = New IO.StreamReader(DlgOpen.FileName, System.Text.ASCIIEncoding.Default)
            Dim ContenutoTXT As String = ""
            Try
                ContenutoTXT = reader.ReadToEnd()
            Catch ex As Exception
                MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Critical, BTN_Apri.Content.ToString)
                ContenutoTXT = ""
            Finally
                reader.Close()
            End Try

            TB_Contenuto.Text = ValidateChar_Replace(ContenutoTXT)
            CheckStatus()
        End If
    End Sub

    Private Sub BTN_IncollaWiki_Click(sender As Object, e As RoutedEventArgs) Handles BTN_IncollaWiki.Click
        If Clipboard.ContainsText Then
            If Clipboard.GetText.Trim.Length = 0 Then Return

            Dim Contenuto_Clipboard As String = Clipboard.GetText
            Dim Cont_CR_Split As String() = Contenuto_Clipboard.Split(vbCrLf)

            Dim WND As New DLG_IncollaWikipedia(Cont_CR_Split(0).Split(vbTab))
            WND.Owner = Me
            WND.ShowDialog()

            If WND.NColonna <> -1 Then
                Dim IniziaDa As Integer = IIf(WND.IncollaNomeColonna, 1, 0)
                Dim ErroriVerificati As Integer = 0
                Dim Contenuto As String = ""

                For x As Integer = IniziaDa To Cont_CR_Split.Count - 1

                    Try
                        Contenuto &= ValidateChar_Replace(Cont_CR_Split(x).Split(vbTab)(WND.NColonna).Trim & vbCrLf)

                    Catch ex As Exception
                        ErroriVerificati += 1
                    End Try
                Next

                TB_Contenuto.Text = Contenuto
                If ErroriVerificati > 0 Then
                    MsgBox(Localization.Resource_WND_ElencoSequenziale.WikiPaste_error, MsgBoxStyle.Exclamation, BTN_IncollaWiki.Content.ToString)
                End If
            End If

        End If
    End Sub

    Private Sub BTN_Pulisci_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Pulisci.Click
        TB_Contenuto.Text = ""
    End Sub

    Private Sub BTN_IncollaWiki_Help_Click(sender As Object, e As RoutedEventArgs) Handles BTN_IncollaWiki_Help.Click
        Try
            System.Diagnostics.Process.Start("http://www.autorinomina.it/index.php/tutorials/elenco-dei-tutorial/16-come-utilizzare-una-tabella-come-elenco-per-rinominare")
        Catch ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, BTN_IncollaWiki_Help.Content.ToString)
        End Try
    End Sub

    Private Sub BTN_Salva_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Salva.Click
        Try
            Dim objStreamWriter = New IO.StreamWriter(FilePath, False, Text.Encoding.Unicode)
            objStreamWriter.Write(TB_Contenuto.Text)
            objStreamWriter.Close()

        Catch ex As Exception
            MsgBox(Localization.Resource_Common.Exception_General & vbCrLf & ex.Message, MsgBoxStyle.Critical, BTN_Salva.Content.ToString)
        End Try

        Close()
    End Sub

    Private Sub BTN_Annulla_Click(sender As Object, e As RoutedEventArgs) Handles BTN_Annulla.Click
        Close()
    End Sub
End Class
