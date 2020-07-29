Imports System.Windows.Media.Animation

Namespace AnimatedStackPanel
    Public Class AnimatedStackPanel
        Inherits StackPanel
        Private _AnimationLength As TimeSpan = TimeSpan.FromMilliseconds(200)

        Protected Overrides Function MeasureOverride(availableSize As Size) As Size
            Dim infiniteSize As New Size(Double.PositiveInfinity, Double.PositiveInfinity)
            Dim curX As Double = 0, curY As Double = 0, curLineHeight As Double = 0
            For Each child As UIElement In Children
                child.Measure(infiniteSize)

                If curX + child.DesiredSize.Width > availableSize.Width Then
                    'Wrap to next line
                    curY += curLineHeight
                    curX = 0
                    curLineHeight = 0
                End If

                curX += child.DesiredSize.Width
                If child.DesiredSize.Height > curLineHeight Then
                    curLineHeight = child.DesiredSize.Height
                End If
            Next

            curY += curLineHeight

            Dim resultSize As New Size()
            resultSize.Width = If(Double.IsPositiveInfinity(availableSize.Width), curX, availableSize.Width)
            resultSize.Height = If(Double.IsPositiveInfinity(availableSize.Height), curY, availableSize.Height)

            Return resultSize
        End Function

        Protected Overrides Function ArrangeOverride(finalSize As Size) As Size
            If Me.Children Is Nothing OrElse Me.Children.Count = 0 Then
                Return finalSize
            End If

            Dim trans As TranslateTransform = Nothing
            Dim curX As Double = 0, curY As Double = 0, curLineHeight As Double = 0

            For Each child As UIElement In Children
                Dim TransVuoto As Boolean = False
                trans = TryCast(child.RenderTransform, TranslateTransform)
                If trans Is Nothing Then
                    child.RenderTransformOrigin = New Point(0, 0)
                    trans = New TranslateTransform()
                    child.RenderTransform = trans
                    TransVuoto = True
                End If

                If curX + child.DesiredSize.Width > finalSize.Width Then
                    'Wrap to next line
                    curY += curLineHeight
                    curX = 0
                    curLineHeight = 0
                End If



                child.Arrange(New Rect(0, 0, child.DesiredSize.Width, child.DesiredSize.Height))

                If TransVuoto Then
                    trans.X = curX
                    trans.Y = curY
                    child.BeginAnimation(UIElement.OpacityProperty, New DoubleAnimation(0, 1, _AnimationLength + _AnimationLength), HandoffBehavior.Compose)
                Else
                    trans.BeginAnimation(TranslateTransform.XProperty, New DoubleAnimation(curX, _AnimationLength), HandoffBehavior.Compose)
                    trans.BeginAnimation(TranslateTransform.YProperty, New DoubleAnimation(curY, _AnimationLength), HandoffBehavior.Compose)
                End If

                curX += child.DesiredSize.Width
                If child.DesiredSize.Height > curLineHeight Then
                    curLineHeight = child.DesiredSize.Height
                End If
            Next

            Return finalSize
        End Function
    End Class
End Namespace
