'' Undoable State Object
'' Copyright (C) 2013-2015 Nathaniel N. Moschkin
'' All Rights Reserved

Imports System
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Collections.Generic
Imports System.Collections
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports DataTools.Memory

Imports System.Runtime.ConstrainedExecution

Namespace Persistence

    ''' <summary>
    ''' Either an undoer for another object, or the base object
    ''' for an undoable object. Works only with public properties of classes that implement INotifyPropertyChanged.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class UndoableStateObject
        Inherits CriticalFinalizerObject

#Region "INotifyPropertyChanged"
        Implements INotifyPropertyChanged

        Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

        Protected _Changed As Boolean = False

        '' a value to indicate we are restoring objects to another state, and to ignore calls to the TargetChangeEvent method.
        Protected _Restoring As Boolean = False

        Public Property Changed As Boolean
            Get
                Return _Changed
            End Get
            Set(value As Boolean)
                If value <> _Changed Then
                    _Changed = value
                    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Changed"))
                End If
            End Set
        End Property

        Protected Overloads Sub OnPropertyChanged(propertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
            Changed = True
        End Sub

        Protected Overloads Sub OnPropertyChanged(e As PropertyChangedEventArgs)
            RaiseEvent PropertyChanged(Me, e)
            Changed = True
        End Sub

        ''' <summary>
        ''' Called when a property on the target has changed.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Protected Sub TargetChangeEvent(sender As Object, e As PropertyChangedEventArgs) Handles _target.PropertyChanged
            If Not _Restoring AndAlso Not _ManualSnapshot AndAlso e.PropertyName <> "Changed" AndAlso _
                ((_IgnoreProperties Is Nothing) OrElse (_IgnoreProperties.Contains(e.PropertyName) = False)) Then Snapshot()
        End Sub

#End Region

        Protected _Parent As UndoableStateObject = Nothing
        Protected _pi As PropertyInfo() = Nothing

        Protected WithEvents _target As INotifyPropertyChanged = Me

        Protected _UndoableTypes() As System.Type = {}
        Protected _Undoables As New List(Of UndoableStateObject)

        Protected _States As New List(Of Object())
        Protected _Pos As Integer = -1

        Protected _MaxSnapshots As Integer = 100

        Protected _ManualSnapshot As Boolean = False
        Protected _Initialized As Boolean = False

        Protected _IgnoreProperties As String()
        Protected _NoCloneTypes As System.Type()

        ''' <summary>
        ''' A list of property names to ignore.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IgnoreProperties As String()
            Get
                Return _IgnoreProperties
            End Get
            Set(value As String())
                _IgnoreProperties = value
                SetInitialState()
                OnPropertyChanged("IgnoreProperties")
            End Set
        End Property

        ''' <summary>
        ''' A list of potentially cloneable types not to clone.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property NoCloneTypes As System.Type()
            Get
                Return _NoCloneTypes
            End Get
            Set(value As System.Type())
                _NoCloneTypes = value
                SetInitialState()
                OnPropertyChanged("NoCloneTypes")
            End Set
        End Property

        ''' <summary>
        ''' Returns true if SetInitialState has been called.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Initialized As Boolean
            Get
                Return _Initialized
            End Get
        End Property

        ''' <summary>
        ''' Returns the parent object of this undoable object.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Parent As UndoableStateObject
            Get
                Return _Parent
            End Get
        End Property

        ''' <summary>
        ''' Controls whether the undoer takes change snapshots automatically.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ManualSnapShot As Boolean
            Get
                Return _ManualSnapshot
            End Get
            Set(value As Boolean)
                _ManualSnapshot = value
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the current position in the undo stack.
        ''' Setting the position applies that position's state.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable Property Position As Integer
            Get
                Return _Pos
            End Get
            Set(value As Integer)

                If _Pos < 0 OrElse _Pos > _States.Count - 1 Then
                    Throw New ArgumentOutOfRangeException()
                End If

                _Pos = value
                ApplyPosition()
            End Set
        End Property

        ''' <summary>
        ''' Specifies the target object.  The default is self.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable Property Target As INotifyPropertyChanged
            Get
                Return _target
            End Get
            Set(value As INotifyPropertyChanged)
                If value IsNot _target Then
                    _target = value
                    _Pos = -1
                    _States.Clear()
                    SetInitialState()
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the list of types that we can instantiate an undoer for.
        ''' Setting this property completely wipes the undo buffer.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property UndoableTypes As System.Type()
            Get
                Return _UndoableTypes
            End Get
            Set(value As System.Type())
                _UndoableTypes = value
                MakeUndoables()
                ClearUndoBuffer()
            End Set
        End Property

        ''' <summary>
        ''' Clears the undo buffer
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub ClearUndoBuffer()
            'If disposedValue Then Throw New ObjectDisposedException("UndoableStateObject")

            If _States Is Nothing Then Return

            _States.Clear()
            For Each u In _Undoables
                If Not u.IsInvalid Then
                    u.ClearUndoBuffer()
                End If
            Next
            Snapshot()
        End Sub

        ''' <summary>
        ''' Maximum number of object states to keep on hand.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable Property MaxSnapshots As Integer
            Get
                Return _MaxSnapshots
            End Get
            Set(value As Integer)
                _MaxSnapshots = value
                TrimSnapshots()

                For Each u In _Undoables
                    If u IsNot Nothing Then
                        u.MaxSnapshots = value
                    End If
                Next
            End Set
        End Property

        ''' <summary>
        ''' Returns true if actions can be undone.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable Function CanUndo() As Boolean
            If _Parent IsNot Nothing Then Return _Parent.CanUndo
            Return (_Pos > 0)
        End Function

        ''' <summary>
        ''' Returns true if actions can be redone.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable Function CanRedo() As Boolean
            If _Parent IsNot Nothing Then Return _Parent.CanRedo
            Return (_Pos <> -1) And (_Pos < (_States.Count - 1))
        End Function

        ''' <summary>
        ''' Trim the collection of snapshots to the maximum limit by truncating from the front.
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overridable Sub TrimSnapshots()

            '' initial state not set!
            If _Pos = -1 Then Return

            If _States.Count > _MaxSnapshots Then
                Dim j As Integer = _States.Count - _MaxSnapshots

                _Pos -= j

                For i = 1 To j
                    _States.RemoveAt(0)
                Next
            End If

        End Sub

        ''' <summary>
        ''' Restore the last snapshot.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable Function Undo() As Boolean
            If _Parent IsNot Nothing Then Return _Parent.Undo

            '' initial state not set!
            If _Pos = -1 Then Return False

            If Not CanUndo() Then Return False

            _Pos -= 1
            For Each u In _Undoables
                If u IsNot Nothing Then
                    u._Pos -= 1
                End If
            Next

            ApplyPosition()

            Return True
        End Function

        ''' <summary>
        ''' Apply the next snapshot.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable Function Redo() As Boolean
            If _Parent IsNot Nothing Then Return _Parent.Redo

            '' initial state not set!
            If _Pos = -1 Then Return False

            If Not CanRedo() Then Return False

            _Pos += 1
            For Each u In _Undoables
                If u IsNot Nothing Then
                    u._Pos += 1
                End If
            Next

            ApplyPosition()

            Return True
        End Function

        ''' <summary>
        ''' Take a snapshot of the current state of the object.
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub Snapshot()
            If Not Initialized Then Return

            '' The parent has control over all.
            If _Parent IsNot Nothing Then
                _Parent.ApplyPosition()
                Return
            End If

            '' this sets the initial state automatically.

            Dim fs As New List(Of Object)
            Dim o As Object = Nothing

            _Restoring = True

            Dim i As Integer = 0
            For Each pe In _pi

                '' Snapshot each one.
                If _Undoables(i).IsInvalid = False Then
                    _Undoables(i).Snapshot()
                Else
                    '' let's take advantage of the presumed deep-copy capabilities of an ICloneable object.
                    'If pe.PropertyType <> GetType(String) AndAlso pe.PropertyType.GetInterface("ICloneable") IsNot Nothing AndAlso _
                    '    ((_NoCloneTypes Is Nothing) OrElse (_NoCloneTypes.Contains(pe.PropertyType) = False)) Then

                    '    o = pe.GetValue(_target)
                    '    If o IsNot Nothing Then fs.Add(CType(o, ICloneable).Clone)
                    '    o = Nothing
                    'Else
                    fs.Add(pe.GetValue(_target))
                    ' End If

                End If
            Next

            _Restoring = False

            _States.Add(fs.ToArray)
            _Pos = _States.Count - 1

            TrimSnapshots()

        End Sub

        ''' <summary>
        ''' Apply the values at the current snapshot position to the target object.
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub ApplyPosition()
            If Not Initialized Then Return

            '' the parent has control over all.
            If _Parent IsNot Nothing Then
                _Parent.ApplyPosition()
                Return
            End If

            '' initial state not set!
            If _Pos = -1 Then Return

            Dim fs As Object() = _States(_Pos)
            Dim i As Integer = 0

            _Restoring = True

            For Each pe In _pi
                '' our subordinate ApplyPosition's are taking care of our Undoable types.
                If _Undoables(i).IsInvalid = False Then
                    _Undoables(i).ApplyPosition()
                Else
                    pe.SetValue(_target, fs(i))
                End If

                i += 1
            Next

            _Restoring = False
        End Sub

        Private Shared ReadOnly InvalidObject = New UndoableStateObject(Nothing, {GetType(Object)})

        Protected ReadOnly Property IsInvalid As Boolean
            Get
                If _UndoableTypes IsNot Nothing AndAlso _UndoableTypes.Count > 0 AndAlso _UndoableTypes(0) = GetType(Object) Then Return True Else Return False
            End Get
        End Property

        ''' <summary>
        ''' Creates the UndoableStateObjects for all of the properties found to be of an undoable type.
        ''' </summary>
        ''' <remarks></remarks>
        Protected Sub MakeUndoables()
            
            Dim i As Integer, _
                c As Integer = _pi.Count - 1

            _Undoables.Clear()

            Dim z As Integer = 0

            For i = 0 To c
                If _UndoableTypes.Contains(_pi(i).PropertyType) Then
                    '' check for suitability:

                    If (_pi(i).PropertyType.IsPrimitive Or _pi(i).PropertyType.IsValueType Or _pi(i).PropertyType.IsArray Or _pi(i).PropertyType.IsEnum) Or (Not _pi(i).PropertyType.IsClass) Then
                        Throw New ArgumentException("Undoable types must be classes with properties.", _pi(i).PropertyType.FullName)
                    End If

                    If _pi(i).PropertyType.GetInterface("INotifyPropertyChanged") Is Nothing Then
                        Throw New ArgumentException("Undoable types must implement INotifyPropertyChanged", _pi(i).PropertyType.FullName)
                    End If

                    '' we made it, let's create a new little undoer for this property.
                    Dim u As New UndoableStateObject(_pi(i).GetValue(_target), Me)

                    u._MaxSnapshots = _MaxSnapshots
                    _Undoables.Add(u)
                Else
                    _Undoables.Add(InvalidObject)
                End If
            Next
        End Sub

        ''' <summary>
        ''' Sets the initial state of the buffer.
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overridable Sub SetInitialState()
            If IsInvalid Then Return

            Dim pi() As PropertyInfo = _target.GetType.GetProperties(BindingFlags.Public Or BindingFlags.Instance)

            If pi Is Nothing Then
                Throw New ArgumentException("This class cannot be undoable.  It has no public properties.")
            End If

            Dim lm As New List(Of PropertyInfo)

            For Each p In pi
                If (p.SetMethod IsNot Nothing) AndAlso _
                    ((_IgnoreProperties Is Nothing) OrElse (_IgnoreProperties.Contains(p.Name) = False)) Then

                    lm.Add(p)
                End If
            Next

            _pi = lm.ToArray

            MakeUndoables()

            _Initialized = True
            ClearUndoBuffer()

        End Sub

        ''' <summary>
        ''' Creates a new object with the specified target.
        ''' The target MUST implement INotifyPropertyChanged.
        ''' </summary>
        ''' <param name="targetObject">The object to monitor.</param>
        ''' <remarks></remarks>
        Public Sub New(targetObject As INotifyPropertyChanged)
            _target = targetObject
            SetInitialState()
        End Sub

        ''' <summary>
        ''' Creates a new object with the specified target.
        ''' The target MUST implement INotifyPropertyChanged.
        ''' </summary>
        ''' <param name="targetObject">The object to monitor.</param>
        ''' <remarks></remarks>
        Public Sub New(targetObject As INotifyPropertyChanged, ignoreProperties() As String)
            _target = targetObject
            _IgnoreProperties = ignoreProperties
            SetInitialState()
        End Sub

        ''' <summary>
        ''' Creates a new object with the specified target.
        ''' The target MUST implement INotifyPropertyChanged.
        ''' </summary>
        ''' <param name="targetObject">The object to monitor.</param>
        ''' <remarks></remarks>
        Public Sub New(targetObject As INotifyPropertyChanged, ignoreProperties() As String, noClone() As System.Type)
            _target = targetObject
            _IgnoreProperties = ignoreProperties
            _NoCloneTypes = noClone
            SetInitialState()
        End Sub

        Protected Sub New(targetObject As INotifyPropertyChanged, parent As UndoableStateObject)
            _Parent = parent

            _UndoableTypes = _Parent._UndoableTypes
            _target = targetObject

            SetInitialState()
        End Sub

        ''' <summary>
        ''' Creates a new object with the specified target.
        ''' The target MUST implement INotifyPropertyChanged.
        ''' </summary>
        ''' <param name="targetObject">The object to monitor.</param>
        ''' <param name="undoableTypes">An array of object types that can have their own UndoableStateObject's.</param>
        ''' <remarks></remarks>
        Public Sub New(targetObject As INotifyPropertyChanged, undoableTypes() As System.Type)
            _target = targetObject
            _UndoableTypes = undoableTypes

            SetInitialState()
        End Sub

        '#Region "IDisposable Support"

        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.

        '    If _States IsNot Nothing Then

        '        If _States.Count > 0 Then
        '            For Each o In _States
        '                Erase o
        '            Next

        '            _States.Clear()
        '        End If

        '        _States = Nothing
        '    End If

        '    If _Undoables IsNot Nothing Then

        '        If _Undoables.Count > 0 Then
        '            _Undoables.Clear()
        '        End If

        '        _Undoables = Nothing
        '    End If

        '    _Pos = -1
        '    _UndoableTypes = Nothing
        '    _target = Nothing

        '    MyBase.Finalize()
        'End Sub

        '#End Region

    End Class

End Namespace