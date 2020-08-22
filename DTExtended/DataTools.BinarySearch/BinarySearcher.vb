Option Explicit On
Option Compare Binary

Imports System.Threading
Imports System.Reflection
Imports DataTools.Persistence
Imports System.Collections.ObjectModel
Imports System.Windows.Forms

Namespace BinarySearch

    Public Enum BinarySearchContextType
        Text
        [Property]
        AllProperties
        [Object]
        Number
    End Enum

    Public Interface IBinarySearchContext

        Property Type As BinarySearchContextType
        Property PropertyInfo As PropertyInfo

    End Interface

    Public Class BinarySearchContext
        Implements IBinarySearchContext

        Public Property PropertyInfo As PropertyInfo Implements IBinarySearchContext.PropertyInfo

        Public Property Type As BinarySearchContextType Implements IBinarySearchContext.Type

        Public Sub New(contextType As BinarySearchContextType)
            Me.Type = contextType
        End Sub

        Public Sub New(contextType As BinarySearchContextType, prop As PropertyInfo)
            Me.Type = contextType
            Me.PropertyInfo = prop
        End Sub

    End Class

    Public Interface IBinarySearcher(Of T)

        Property Comparer As IComparer(Of T)

        Property Comparison As Comparison(Of T)

        Property Context As IList(Of T)

        Property AltContext As ObservableCollection(Of T)

        Property SearchContext As IBinarySearchContext

        Function SearchFor(value As Object, Optional returnNear As Boolean = False) As T

        Function SearchForIdx(value As Object) As Integer

    End Interface

    Public Class BinarySearcher(Of T)
        Implements IBinarySearcher(Of T)

        Public Property Comparer As IComparer(Of T) Implements IBinarySearcher(Of T).Comparer

        Public Property Comparison As Comparison(Of T) Implements IBinarySearcher(Of T).Comparison

        Public Property Context As IList(Of T) Implements IBinarySearcher(Of T).Context

        Public Property AltContext As ObservableCollection(Of T) Implements IBinarySearcher(Of T).AltContext

        Public Function SearchForIdx(value As Object) As Integer Implements IBinarySearcher(Of T).SearchForIdx
            Dim index As Integer = -1

            Select Case SearchContext.Type
                Case BinarySearchContextType.AllProperties
                    SearchAllProps(value, index)

                Case BinarySearchContextType.Object
                    SearchObject(value, index)

                Case BinarySearchContextType.Text
                    SearchText(value, index)

                Case BinarySearchContextType.Property
                    SearchProp(value, index)

                Case BinarySearchContextType.Number
                    SearchNumber(value, index)

            End Select

            Return index
        End Function

        Public Function SearchFor(value As Object, Optional returnNear As Boolean = False) As T Implements IBinarySearcher(Of T).SearchFor
            Dim index As Integer = -1

            Select Case SearchContext.Type
                Case BinarySearchContextType.AllProperties
                    Return SearchAllProps(value, index, returnNear)

                Case BinarySearchContextType.Object
                    Return SearchObject(value, index, returnNear)

                Case BinarySearchContextType.Text
                    Return SearchText(value, index, returnNear)

                Case BinarySearchContextType.Property
                    Return SearchProp(value, index, , returnNear)

                Case BinarySearchContextType.Number
                    Return SearchNumber(value, index, returnNear)

            End Select

            Return Nothing
        End Function

        Public Async Function SearchForAsync(value As Object) As Tasks.Task(Of T)
            Dim out As Object = Nothing

            Await Tasks.Task.Run(Sub()
                                     out = SearchFor(value)
                                 End Sub)

            Return out
        End Function

        Public Async Function SearchForIdxAsync(value As Object) As Tasks.Task(Of T)
            Dim out As Object = Nothing

            Await Tasks.Task.Run(Sub()
                                     out = SearchForIdx(value)
                                 End Sub)

            Return out
        End Function

        Protected Overridable Function SearchNumber(value As T, ByRef index As Integer, Optional returnNear As Boolean = False) As T
            index = -1

            Dim activeContext As ICollection(Of T)

            If Context Is Nothing Then
                activeContext = AltContext
            Else
                activeContext = Context
            End If

            Dim count As Integer = activeContext.Count
            Dim currentIdx As Integer = activeContext.Count / 2

            Dim startPos As Integer = 0
            Dim endPos As Integer = count - 1

            Dim valDiff As Integer
            Dim nextCenter As Integer

            Dim propVal As Object
            If activeContext.Count = 0 Then Return Nothing

            Do
                propVal = (activeContext(currentIdx))

                valDiff = value - propVal

                If valDiff = 0 Then
                    index = currentIdx
                    Return activeContext(currentIdx)
                End If

                If valDiff < 0 Then
                    endPos = currentIdx - 1
                    If endPos < startPos Then Exit Do

                    nextCenter = (endPos - startPos) / 2
                    If nextCenter = currentIdx Then Exit Do
                    currentIdx = nextCenter
                Else
                    startPos = currentIdx + 1
                    If startPos >= count Then Exit Do

                    nextCenter = (startPos + ((endPos - startPos) / 2))
                    If nextCenter = currentIdx Then Exit Do
                    currentIdx = nextCenter
                End If

            Loop

            If returnNear Then Return activeContext(currentIdx)

            Return Nothing

        End Function

        Protected Overridable Function SearchText(value As Object, ByRef index As Integer, Optional returnNear As Boolean = False) As T
            index = -1

            Dim activeContext As ICollection(Of T)

            If Context Is Nothing Then
                activeContext = AltContext
            Else
                activeContext = Context
            End If

            Dim count As Integer = activeContext.Count
            Dim currentIdx As Integer = activeContext.Count / 2

            Dim startPos As Integer = 0
            Dim endPos As Integer = count - 1

            Dim valDiff As Integer
            Dim nextCenter As Integer

            Dim propVal As Object
            If activeContext.Count = 0 Then Return Nothing

            Do
                propVal = (activeContext(currentIdx))

                valDiff = String.Compare(value.ToString, propVal.ToString)

                If valDiff = 0 Then
                    index = currentIdx
                    Return activeContext(currentIdx)
                End If

                If valDiff < 0 Then
                    endPos = currentIdx - 1
                    If endPos < startPos Then Exit Do

                    nextCenter = (endPos - startPos) / 2
                    If nextCenter = currentIdx Then Exit Do
                    currentIdx = nextCenter
                Else
                    startPos = currentIdx + 1
                    If startPos >= count Then Exit Do

                    nextCenter = (startPos + ((endPos - startPos) / 2))
                    If nextCenter = currentIdx Then Exit Do
                    currentIdx = nextCenter
                End If

            Loop

            If returnNear Then Return activeContext(currentIdx)

            Return Nothing

        End Function

        Protected Overridable Function SearchObject(value As Object, ByRef index As Integer, Optional returnNear As Boolean = False) As T
            index = -1

            Dim activeContext As ICollection(Of T)

            If Context Is Nothing Then
                activeContext = AltContext
            Else
                activeContext = Context
            End If

            If activeContext.Count = 0 Then Return Nothing

            Dim count As Integer = activeContext.Count
            Dim currentIdx As Integer = activeContext.Count / 2

            Dim cbs As Integer = count - currentIdx

            Dim startPos As Integer = 0
            Dim endPos As Integer = count - 1

            Dim valDiff As Integer
            Dim nextCenter As Integer

            Dim propVal As Object

            Do
                propVal = (activeContext(currentIdx))

                If value.GetType() = propVal.GetType() Then

                    If Comparer Is Nothing Then Return Nothing

                    valDiff = Comparer.Compare(value, propVal)

                End If

                If valDiff = 0 Then
                    index = currentIdx
                    Return activeContext(currentIdx)
                End If

                If valDiff < 0 Then
                    endPos = currentIdx - 1
                    If endPos < startPos Then Exit Do

                    nextCenter = (endPos - startPos) / 2
                    If nextCenter = currentIdx Then Exit Do
                    currentIdx = nextCenter
                Else
                    startPos = currentIdx + 1
                    If startPos >= count Then Exit Do

                    nextCenter = (startPos + ((endPos - startPos) / 2))
                    If nextCenter = currentIdx Then Exit Do
                    currentIdx = nextCenter
                End If

            Loop

            If returnNear Then Return activeContext(currentIdx)

            Return Nothing

        End Function

        Protected Overridable Function SearchAllProps(value As Object, ByRef index As Integer, Optional returnNear As Boolean = False) As T
            index = -1

            Dim activeContext As ICollection(Of T)

            If Context Is Nothing Then
                activeContext = AltContext
            Else
                activeContext = Context
            End If

            Dim count As Integer = activeContext.Count
            Dim currentIdx As Integer = activeContext.Count / 2

            Dim startPos As Integer = 0
            Dim endPos As Integer = count - 1

            Dim valDiff As Integer
            Dim nextCenter As Integer

            Dim propVal As List(Of KeyValuePair(Of String, Object))

            Do
                propVal = PropsWithValues(activeContext(currentIdx))

                For Each kv In propVal

                    If value.GetType() = kv.Value.GetType() Then

                        If TypeOf value Is String Then
                            valDiff = String.Compare(value, kv.Value)
                            Exit For
                        ElseIf TypeOf value Is Date Then
                            valDiff = Date.Compare(value, kv.Value)
                            Exit For
                        ElseIf IsNumeric(value) Then
                            If CDbl(value) < CDbl(kv.Value) Then
                                valDiff = -1
                            ElseIf CDbl(value) > CDbl(kv.Value) Then
                                valDiff = 1
                            Else
                                valDiff = 0
                            End If

                            Exit For
                        Else
                            Throw New ArgumentException("Cannot determine appropriate search method", "value")
                        End If

                    Else
                        valDiff = String.Compare(value.ToString, kv.Value.ToString)
                    End If
                Next

                If valDiff = 0 Then
                    index = currentIdx
                    Return activeContext(currentIdx)
                End If

                If valDiff < 0 Then
                    endPos = currentIdx - 1
                    If endPos < startPos Then Exit Do

                    nextCenter = (endPos - startPos) / 2
                    If nextCenter = currentIdx Then Exit Do
                    currentIdx = nextCenter
                Else
                    startPos = currentIdx + 1
                    If startPos >= count Then Exit Do

                    nextCenter = (startPos + ((endPos - startPos) / 2))
                    If nextCenter = currentIdx Then Exit Do
                    currentIdx = nextCenter
                End If

            Loop

            If returnNear Then Return activeContext(currentIdx)

            Return Nothing
        End Function

        Protected Overridable Function SearchProp(value As Object, ByRef index As Integer, Optional key As String = Nothing, Optional returnNear As Boolean = False) As T
            index = -1

            Dim activeContext As ICollection(Of T)

            If Context Is Nothing Then
                activeContext = AltContext
            Else
                activeContext = Context
            End If

            Dim prop As PropertyInfo = SearchContext.PropertyInfo
            If prop Is Nothing Then Return Nothing

            Dim count As Integer = activeContext.Count
            Dim currentIdx As Integer = activeContext.Count / 2

            Dim startPos As Integer = 0
            Dim endPos As Integer = count - 1

            Dim valDiff As Integer
            Dim nextCenter As Integer

            Dim propVal As Object

            If value.GetType() <> prop.PropertyType Then
                Throw New ArgumentException("Property type and search paramters do not match")
            End If

            If activeContext.Count = 0 Then Return Nothing

            Do
                Try

                    propVal = prop.GetValue(activeContext(currentIdx))

                    If TypeOf value Is String Then
                        valDiff = String.Compare(value, propVal)
                    ElseIf TypeOf value Is Date Then
                        valDiff = Date.Compare(value, propVal)
                    ElseIf IsNumeric(value) Then
                        If CDbl(value) < CDbl(propVal) Then
                            valDiff = -1
                        ElseIf CDbl(value) > CDbl(propVal) Then
                            valDiff = 1
                        Else
                            valDiff = 0
                        End If
                    Else
                        valDiff = String.Compare(value.ToString, propVal.ToString)
                    End If

                    If valDiff = 0 Then
                        index = currentIdx
                        Return activeContext(currentIdx)
                    End If

                    If valDiff < 0 Then
                        endPos = currentIdx - 1
                        If endPos < startPos Then Exit Do

                        nextCenter = (endPos - startPos) / 2
                        If nextCenter = currentIdx Then Exit Do
                        currentIdx = nextCenter
                    Else
                        startPos = currentIdx + 1
                        If startPos >= count Then Exit Do

                        nextCenter = (startPos + ((endPos - startPos) / 2))
                        If nextCenter = currentIdx Then Exit Do
                        currentIdx = nextCenter
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message & " : Index: " & currentIdx)
                End Try

            Loop

            If returnNear Then Return activeContext(currentIdx)

            Return Nothing
        End Function

        Public Overridable Property SearchContext As IBinarySearchContext Implements IBinarySearcher(Of T).SearchContext

        Public Sub Sort()
            If Context IsNot Nothing Then

                If TypeOf Context Is List(Of T) Then
                    CType(Context, List(Of T)).Sort(Comparer)

                ElseIf Context.GetType.IsArray Then
                    Array.Sort(CType(Context, T()), Comparer)
                End If

            End If
        End Sub

        Public Sub New(comparer As IComparer(Of T), context As IList(Of T), sortFirst As Boolean)
            Me.Comparer = comparer
            Me.Context = context

            If GetType(T).IsPrimitive Or GetType(T).IsEnum Then
                Me.SearchContext = New BinarySearchContext(BinarySearchContextType.Number)
            Else
                Me.SearchContext = New BinarySearchContext(BinarySearchContextType.Text)
            End If

            If sortFirst Then Sort()
        End Sub

        Public Sub New(comparer As IComparer(Of T), context As IList(Of T), searchContext As IBinarySearchContext, sortFirst As Boolean)
            Me.New(comparer, context, sortFirst)
            Me.SearchContext = searchContext
        End Sub

        Public Sub New(comparison As Comparison(Of T), context As IList(Of T), sortFirst As Boolean)
            Me.Comparison = comparison
            Me.Comparer = New ComparisonProxy(Of T)(comparison)

            Me.Context = context

            If GetType(T).IsPrimitive Or GetType(T).IsEnum Then
                Me.SearchContext = New BinarySearchContext(BinarySearchContextType.Number)
            Else
                Me.SearchContext = New BinarySearchContext(BinarySearchContextType.Text)
            End If

            If sortFirst Then Sort()
        End Sub

        Public Sub New(comparison As Comparison(Of T), context As IList(Of T), searchContext As IBinarySearchContext, sortFirst As Boolean)
            Me.Comparison = comparison
            Me.Comparer = New ComparisonProxy(Of T)(comparison)

            Me.Context = context

            If GetType(T).IsPrimitive Or GetType(T).IsEnum Then
                Me.SearchContext = New BinarySearchContext(BinarySearchContextType.Number)
            Else
                Me.SearchContext = New BinarySearchContext(BinarySearchContextType.Text)
            End If

            If sortFirst Then Sort()

            Me.SearchContext = searchContext
        End Sub

        Public Sub New(context As IList(Of T), searchContext As IBinarySearchContext, sortFirst As Boolean)
            Me.Comparison = Nothing
            Me.Comparer = Nothing

            Me.Context = context
            If GetType(T).IsPrimitive Or GetType(T).IsEnum Then
                Me.SearchContext = New BinarySearchContext(BinarySearchContextType.Number)
            Else
                Me.SearchContext = New BinarySearchContext(BinarySearchContextType.Text)
            End If

            If sortFirst Then Sort()

            Me.SearchContext = searchContext
        End Sub

        Protected Sub New()

        End Sub

    End Class

    Public Class TextBinarySearcher
        Inherits BinarySearcher(Of String)

        Public Sub New(context As String(), sortFirst As Boolean)
            MyBase.New()
            MyBase.Context = New List(Of String)
            CType(MyBase.Context, List(Of String)).AddRange(context)
            MyBase.Comparer = StringComparer.Create(Application.CurrentCulture, False)

            If sortFirst Then Sort()
        End Sub

        Public Sub New(context As List(Of String), sortFirst As Boolean)
            MyBase.New(StringComparer.Create(Application.CurrentCulture, False), context, sortFirst)
        End Sub

    End Class

    Public Class NumericBinarySearcher
        Inherits BinarySearcher(Of Double)

        Public Sub New(context As String(), sortFirst As Boolean)
            MyBase.New()
            MyBase.Context = New List(Of Double)
            CType(MyBase.Context, List(Of Double)).AddRange(context)

            MyBase.Comparer = New ComparisonProxy(Of Double)(Function(x As Double, y As Double) As Integer
                                                                 Return x - y
                                                             End Function)

            If sortFirst Then Sort()
        End Sub

        Public Sub New(context As List(Of Double), sortFirst As Boolean)
            MyBase.New(StringComparer.Create(Application.CurrentCulture, False), context, Nothing, sortFirst)
        End Sub

    End Class

    Friend Class ComparisonProxy(Of T)
        Implements IComparer(Of T)

        Private _Comparison As Comparison(Of T) = Nothing

        Public Sub New(c As Comparison(Of T))
            _Comparison = c
        End Sub

        Public Function Compare(x As T, y As T) As Integer Implements IComparer(Of T).Compare
            Return _Comparison(x, y)
        End Function
    End Class

End Namespace