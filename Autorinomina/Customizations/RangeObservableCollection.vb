Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel

Public Class RangeObservableCollection(Of T)
    Inherits ObservableCollection(Of T)
    Private _suppressNotification As Boolean
    Public Shadows Event CollectionChanged As NotifyCollectionChangedEventHandler

    Protected Overrides Sub OnCollectionChanged(e As NotifyCollectionChangedEventArgs)
        If Not _suppressNotification Then
            MyBase.OnCollectionChanged(e)
        End If
    End Sub

    Protected Sub LaunchCollectionChanged()
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
    End Sub

    Protected Overridable Sub OnCollectionChangedMultiItem(e As NotifyCollectionChangedEventArgs)
        Dim handlers As NotifyCollectionChangedEventHandler = AddressOf LaunchCollectionChanged

        If handlers IsNot Nothing Then
            For Each handler As NotifyCollectionChangedEventHandler In handlers.GetInvocationList()
                If TypeOf handler.Target Is CollectionView Then
                    DirectCast(handler.Target, CollectionView).Refresh()
                Else
                    handler(Me, e)
                End If
            Next
        End If
    End Sub

    Public Sub AddRange(list As IEnumerable(Of T))
        If list Is Nothing Then
            Throw New ArgumentNullException("list")
        End If
        _suppressNotification = True
        For Each item As T In list
            Add(item)
        Next
        _suppressNotification = False

        Dim obEvtArgs As New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Add, TryCast(list, System.Collections.IList))
        OnCollectionChangedMultiItem(obEvtArgs)
    End Sub

    Public Sub RemoveRange(list As IEnumerable(Of T))
        If list Is Nothing Then
            Throw New ArgumentNullException("list")
        End If
        If list.Count = 0 Then Return

        _suppressNotification = True
        Dim removeList As New List(Of T)
        For x As Integer = list.Count To 0 Step -1
            Remove(list(x))
            removeList.Add(list(x))
        Next
        _suppressNotification = False

        Dim obEvtArgs As New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Remove, TryCast(removeList, System.Collections.IList), 0)
        OnCollectionChangedMultiItem(obEvtArgs)

        OnPropertyChanged(New PropertyChangedEventArgs("Count"))
        OnPropertyChanged(New PropertyChangedEventArgs("Item[]"))
        OnCollectionChanged(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
    End Sub

End Class
