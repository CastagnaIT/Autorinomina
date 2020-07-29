Public Class SuggestionGuide
    Public Property CurrentPage As Integer = 0

    Public Sub CambiaPaginaGuida(Numero As Integer)
        Select Case Numero
            Case 1
                SP_Help_0.Visibility = Visibility.Hidden
                IMG_Help_1.Visibility = Visibility.Visible
                TB_Help_1.Visibility = Visibility.Visible

            Case 2
                IMG_Help_1.Visibility = Visibility.Hidden
                TB_Help_1.Visibility = Visibility.Hidden

                IMG_Help_2.Visibility = Visibility.Visible
                TB_Help_2.Visibility = Visibility.Visible

            Case 3
                IMG_Help_2.Visibility = Visibility.Hidden
                TB_Help_2.Visibility = Visibility.Hidden

                IMG_Help_3.Visibility = Visibility.Visible
                TB_Help_3.Visibility = Visibility.Visible

            Case 4
                IMG_Help_3.Visibility = Visibility.Hidden
                TB_Help_3.Visibility = Visibility.Hidden

                IMG_Help_4.Visibility = Visibility.Visible
                TB_Help_4.Visibility = Visibility.Visible
        End Select
        CurrentPage = Numero
    End Sub
End Class
