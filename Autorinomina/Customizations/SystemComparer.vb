Imports System.ComponentModel

Public Class SystemComparer
    Implements IComparer

    Private direction As SByte
    Public SortMemberPath As String

    Public Sub New(_direction As ListSortDirection, _SortMemberPath As String)
        direction = IIf(_direction = ListSortDirection.Ascending, 1, -1)
        SortMemberPath = _SortMemberPath
    End Sub

    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        Dim ReflProprery As Reflection.PropertyInfo = CType(x, ItemFileData).GetType().GetProperty(SortMemberPath)

        Dim ValX As String = ReflProprery.GetValue(CType(x, ItemFileData), Nothing)
        Dim ValY As String = ReflProprery.GetValue(CType(y, ItemFileData), Nothing)

        Return Interaction.StrCmpLogicalW(ValX, ValY) * direction
    End Function
End Class



