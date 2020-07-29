Public Class DLG_CartellePreferite
    Private Collection_Files_Temp As New ObjectModel.ObservableCollection(Of CollFiles_Temp)()
    Public Property ElencoCartelle As Array
    Dim _PathCartella As String
    Dim _NomeCartella As String
    Dim _XMLPath_FavoriteNumber As String

    Class CollFiles_Temp
        Public Sub New(_NomeItem As String, _Percorso As String)
            NomeItem = _NomeItem
            Percorso = _Percorso
        End Sub

        Public Property NomeItem As String
        Public Property Percorso As String
    End Class

    Protected Overrides Sub OnSourceInitialized(e As EventArgs)
        CustomWindow(Me, True)
    End Sub

    Public Sub New(PathCartella As String, NomeCartella As String, XMLPath_FavoriteNumber As String)

        ' La chiamata è richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        _PathCartella = PathCartella
        _XMLPath_FavoriteNumber = XMLPath_FavoriteNumber

        If String.IsNullOrEmpty(PathCartella) Then
            _NomeCartella = NomeCartella
        Else
            _NomeCartella = New IO.DirectoryInfo(PathCartella).Name
        End If
    End Sub

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        VisualizzaCartelle()

        LB_Folder.ItemsSource = Collection_Files_Temp
    End Sub

    Private Sub VisualizzaCartelle()
        Collection_Files_Temp.Clear()

        Me.Title = Autorinomina.Localization.Resource_Common_Dialogs.DLG_FavoriteFolder_Title & " '" & _NomeCartella & "'"

        If IO.Directory.Exists(_PathCartella) Then

            Collection_Files_Temp.Add(New CollFiles_Temp(AutoRinomina.Localization.Resource_Common_Dialogs.DLG_FavoriteFolder_RowTextFolder & vbTab & vbTab & ">> " & New IO.DirectoryInfo(_PathCartella).Name, _PathCartella))

            For Each Path As String In IO.Directory.GetDirectories(_PathCartella)
                Collection_Files_Temp.Add(New CollFiles_Temp(AutoRinomina.Localization.Resource_Common_Dialogs.DLG_FavoriteFolder_RowTextSubfolder & vbTab & ">> " & New IO.DirectoryInfo(Path).Name, Path))
            Next
        End If

        LB_BackgroundTXT()
    End Sub

    Private Sub SelectAllExecuted()
        LB_Folder.SelectAll()
    End Sub

    Private Sub CheckBox_Checked(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim CB As CheckBox = sender
        LB_Folder.SelectedItems.Add(CB.DataContext)
    End Sub

    Private Sub CheckBox_Unchecked(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim CB As CheckBox = sender
        LB_Folder.SelectedItems.Remove(CB.DataContext)
    End Sub

    Private Sub BTN_Aggiungi_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Aggiungi.Click
        If LB_Folder.SelectedItems.Count = 0 Then Return

        Dim ArrayL As New ArrayList
        For Each item As CollFiles_Temp In LB_Folder.SelectedItems
            ArrayL.Add(item.Percorso)
        Next
        ElencoCartelle = ArrayL.ToArray

        Me.DialogResult = True
        Close()
    End Sub

    Private Sub BTN_Annulla_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Annulla.Click
        Me.DialogResult = False
        Close()
    End Sub

    Private Sub BTN_Cambia_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BTN_Cambia.Click
        Using DLG As New Windows.Forms.FolderBrowserDialog

            DLG.Description = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_FavoriteFolder_Msg_ChangePath
            DLG.ShowNewFolderButton = False
            DLG.RootFolder = Environment.SpecialFolder.Desktop
            If DLG.ShowDialog = Forms.DialogResult.OK Then

                XMLSettings_Save("FavoriteFolderPath_" & _XMLPath_FavoriteNumber, DLG.SelectedPath)
                _PathCartella = DLG.SelectedPath
                _NomeCartella = New IO.DirectoryInfo(_PathCartella).Name

                VisualizzaCartelle()
            End If

        End Using
    End Sub

    Private Sub Image_MouseDown(sender As System.Object, e As System.Windows.Input.MouseButtonEventArgs)

        Dim process As System.Diagnostics.Process = Nothing
        Dim processStartInfo As New System.Diagnostics.ProcessStartInfo()

        processStartInfo.FileName = "explorer.exe"

        '  If System.Environment.OSVersion.Version.Major >= 6 Then ' Windows Vista or higher
        'processStartInfo.Verb = "runas"
        'Else
        '' No need to prompt to run as admin
        'End If
        Dim Img As Image = sender
        Dim CP As ContentPresenter = Img.TemplatedParent

        If CP.Content Is Nothing Then Exit Sub

        processStartInfo.Arguments = CType(CP.Content, CollFiles_Temp).Percorso
        processStartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
        processStartInfo.UseShellExecute = True

        Try
            process = System.Diagnostics.Process.Start(processStartInfo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Title)
        Finally

            If Not (process Is Nothing) Then
                process.Dispose()
            End If

        End Try

        e.Handled = True
    End Sub

    Private Sub LB_BackgroundTXT()
        If Collection_Files_Temp.Count = 0 Then
            Dim VB As New VisualBrush
            VB.TileMode = TileMode.None
            VB.Stretch = Stretch.None
            Dim R As New Run
            R.Foreground = Brushes.Gray
            R.Text = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_FavoriteFolder_EmptyFolder
            R.FontStyle = FontStyles.Italic
            R.FontSize = 13
            VB.Visual = New TextBlock(R)
            LB_Folder.Background = VB
        Else
            LB_Folder.ClearValue(ListBox.BackgroundProperty)
        End If
    End Sub
End Class
