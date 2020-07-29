Public Class NumeriRomaniConverter

    Private ReadOnly ListaRomani As String = "IVXLCDM"
    Private mNumero As String
    Private Valore As Integer

    Public Sub New()
    End Sub

    Public Function CheckRomani(ByVal pNumero As String) As Boolean


        Dim x As Integer = 0

        For Each c As Char In pNumero.ToCharArray()


            If ListaRomani.IndexOf(c) <> -1 Then

                x += 1

            End If
        Next

        Return (x = pNumero.Length)

    End Function

    Public Function CheckBase10(ByVal pNumero As String) As Boolean


        Dim x As Integer = 0

        Try

            x = Integer.Parse(pNumero)
            Valore = x

        Catch

            Return False
        End Try

        Return ((x > 0) AndAlso (x < 4000))

    End Function

    Public ReadOnly Property Numero() As String

        Get
            Return mNumero
        End Get
    End Property

    Public Function ToRomani(ByVal pNumero As String) As String

        Dim risultato As String = ""

        If Not CheckBase10(pNumero) Then

            Return ""
        End If

        pNumero = Valore.ToString()

        For k As Integer = 0 To pNumero.Length - 1


            Dim s As String = pNumero.Substring(k, 1)
            Dim tempo As String = pNumero.Substring(k)
            Dim c As Char = " "c
            Dim x As Integer = Integer.Parse(s)

            Select Case tempo.Length
                Case 4

                    c = "M"c
                    Exit Select

                Case 3

                    c = "C"c
                    Exit Select

                Case 2

                    c = "X"c
                    Exit Select

                Case 1

                    c = "I"c
                    Exit Select

            End Select

            Dim z As Integer = "CXI".IndexOf(c)

            If x < 4 Then

                risultato += "".PadLeft(x, c)
            ElseIf x = 4 Then

                risultato += New String() {"CD", "XL", "IV"}(z)

            ElseIf x = 9 Then

                risultato += New String() {"CM", "XC", "IX"}(z)
            Else

                risultato += New String() {"D", "L", "V"}(z)

                If x > 5 Then

                    x -= 5
                    risultato += "".PadLeft(x, c)

                End If

            End If
        Next

        mNumero = risultato

        Return risultato
    End Function

    Public Function ToBase10(ByVal pNumero As String) As String

        If Not CheckRomani(pNumero) Then

            Return ""
        End If

        Dim c As Char() = pNumero.ToCharArray()
        Dim s As String = ""
        Dim somma As Integer = 0
        Dim tmp As Integer = 0

        For k As Integer = 0 To c.Length - 1


            Dim z As String = c(k).ToString()
            Dim p As String = s & z
            Dim i As Integer = ListaRomani.IndexOf(z)

            If i = -1 Then

                Continue For
            End If

            i = New Integer() {1, 5, 10, 50, 100, 500,
             1000}(i)

            If (tmp <> 0) AndAlso (c(k) <> c(k - 1)) Then

                Dim j As Integer = "IV-IX-XL-XC-CD-CM".IndexOf(p)

                If j <> -1 Then

                    j = "VXLCDM".IndexOf(z)
                    somma += New Integer() {4, 9, 40, 90, 400, 900}(j)
                    tmp = 0

                    s = ""
                Else

                    somma += tmp
                    s = z

                    tmp = i

                End If
            Else

                s = p
                tmp += i

            End If
        Next

        somma += tmp
        mNumero = somma.ToString()

        Return Numero
    End Function

End Class



