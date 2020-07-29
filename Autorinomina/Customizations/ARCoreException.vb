Public Class ARCoreException
    Inherits Exception

    Public Sub New()
        'no code needed
    End Sub

    Public Sub New(ByVal message As String)
        MyBase.New(message)
    End Sub

End Class
