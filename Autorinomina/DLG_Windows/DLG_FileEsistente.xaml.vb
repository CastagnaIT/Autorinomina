Imports System.Runtime.InteropServices

Public Class DLG_FileEsistente
    Public Shadows DialogResult As Boolean '<< fix errore DialogResult in chiusura della finestra mentre si annulla il BW
    Dim Path As String
    Dim NuovoNome As String
    Public Property ResultData As String() = Nothing

    Protected Overrides Sub OnSourceInitialized(e As EventArgs)
        CustomWindow(Me, True)
    End Sub

    Public Sub New(ByVal _PathCompletoOriginale As String, ByVal _nuovoNome As String)

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        Path = _PathCompletoOriginale
        NuovoNome = _nuovoNome
    End Sub

    Private Sub Window_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded

        Dim NomeDelFile As String = IO.Path.GetFileName(Path)
        Image1.Source = EstraiIconaAssociata(Path, False)(0)

        TextBlock2.Text = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_ManualFilter_FilePath & Space(1) & IO.Path.GetDirectoryName(Path)
        TextBlock2.ToolTip = IO.Path.GetDirectoryName(Path)
        TextBlock1.Text = IO.Path.GetFileName(Path)
        TextBlock1.ToolTip = IO.Path.GetFileName(Path)
        Label11.Content = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_ManualFilter_FileSize & " - " & FormatBYTE(New IO.FileInfo(Path).Length)


        Dim N_Duplicate As Integer = 1
        While IO.File.Exists(IO.Path.Combine(IO.Path.GetDirectoryName(Path), "\" & IO.Path.GetFileNameWithoutExtension(NuovoNome) & " (" & N_Duplicate & ")" & IO.Path.GetExtension(NuovoNome)))
            N_Duplicate += 1
        End While
        Dim NewCheckedFilename As String = IO.Path.GetFileNameWithoutExtension(NuovoNome) & " (" & N_Duplicate & ")" & IO.Path.GetExtension(NuovoNome)
        TextBlock8.Text = AutoRinomina.Localization.Resource_Common_Dialogs.DLG_FileEsistente_Btn_KeepBoth_Desc & vbCrLf & NewCheckedFilename
        TextBlock8.Tag = IO.Path.Combine(IO.Path.GetDirectoryName(Path), NewCheckedFilename)

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button1.Click
        ResultData = {"0", ""}

        Me.DialogResult = True
        Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button2.Click
        ResultData = {"1", ""}

        Me.DialogResult = True
        Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button3.Click
        ResultData = {"2", IO.Path.Combine(IO.Path.GetDirectoryName(Path), TextBlock8.Tag.ToString)}

        Me.DialogResult = True
        Close()
    End Sub



End Class
