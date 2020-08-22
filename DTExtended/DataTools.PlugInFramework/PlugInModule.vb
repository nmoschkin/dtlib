''
'' Plug-In Module Framework
'' Copyright (C) 2015 Nathan Moschkin
''
Option Explicit On

Imports System.Windows
Imports System.Windows.Controls
Imports System.Threading
Imports System.ComponentModel

Namespace PlugInFramework

    ''' <summary>
    ''' Encapsulates a common plug-in module interface.
    ''' </summary>
    Public Interface IPlugInModule

        Inherits INotifyPropertyChanged, INotifyStatusProgress

        Event RequestFocus(sender As Object, e As EventArgs)

        Event RequestBlur(sender As Object, e As EventArgs)

        Property Changed As Boolean

        Property Title As String

        Property DenyProcessing As Boolean

        ReadOnly Property IsProcessing As Boolean

        ReadOnly Property PlugInMenuItem As MenuItem

        Sub CancelTasks()

        Function Initialize() As Boolean

        Function Shutdown() As Boolean

    End Interface

    ''' <summary>
    ''' File transfer states for classes and events related to <see cref="IFileTransferPlugInModule"/>
    ''' </summary>
    Public Enum FileTransferModuleEventStates

        ''' <summary>
        ''' There are no active file transfers.
        ''' </summary>
        NotTransferring

        ''' <summary>
        ''' A file transfer is starting.
        ''' </summary>
        Starting

        ''' <summary>
        ''' Currently uploading a file.
        ''' </summary>
        Uploading

        ''' <summary>
        ''' Currently downloading a file.
        ''' </summary>
        Downloading

        ''' <summary>
        ''' Transfer complete.
        ''' </summary>
        Complete

        ''' <summary>
        ''' An error occured.
        ''' </summary>
        TransferError
    End Enum

    ''' <summary>
    ''' EventArgs class to be used with file transfers in classes that implement the <see cref="IFileTransferPlugInModule"/>
    ''' </summary>
    Public Class FileTransferModuleEventArgs
        Inherits EventArgs

        Public ReadOnly Property FileSize As Long

        Public ReadOnly Property Progress As Long

        Public ReadOnly Property Filename As String

        Public ReadOnly Property State As FileTransferModuleEventStates

        Public ReadOnly Property TransferStart As Date

        Public ReadOnly Property TransferStop As Date

        Public ReadOnly Property TransferElapsed As TimeSpan

        Public ReadOnly Property TransferRate As Single

        Public ReadOnly Property ErrorCode As Integer

        Public ReadOnly Property ErrorMessage As String

        Public Property Cancel As Boolean = False

        ''' <summary>
        ''' Creates as new file transfer event object with the specified parameters.
        ''' If either elapsed time or transfer rate are omitted, they are calculated automatically
        ''' based on the stop time.  If stop time is omitted, Date.Now is used.
        ''' </summary>
        ''' <param name="filename">The name of the file being transferred.</param>
        ''' <param name="state">The current transfer state.</param>
        ''' <param name="filesize">The total size of the file.</param>
        ''' <param name="progress">The number of bytes transmitted, so far.</param>
        ''' <param name="start">The start date/time of the transfer.</param>
        ''' <param name="[stop]">The optional stop date/time of the transfer.  If this field is omitted, the current time is used.</param>
        ''' <param name="elapsed">The total elapsed time of the transfer.  If this field is omitted, it is automatically calculated.</param>
        ''' <param name="rate">The current transfer rate. If this field is omitted, it is automatically calculated in bytes per second.</param>
        ''' <param name="errorcode">Optional error code.</param>
        ''' <param name="errormsg">Optional error message.</param>
        Public Sub New(filename As String,
                       state As FileTransferModuleEventStates,
                       filesize As Long,
                       progress As Long,
                       start As Date,
                       Optional [stop] As Date = #1/1/1899#,
                       Optional elapsed As TimeSpan = Nothing,
                       Optional rate As Single = 0.0,
                       Optional errorcode As Integer = 0,
                       Optional errormsg As String = Nothing
                       )

            _Filename = filename
            _State = state
            _FileSize = filesize
            _Progress = progress
            _TransferStart = start
            _TransferStop = [stop]

            If [stop] = #1/1/1899# Then [stop] = Date.Now

            If elapsed.Ticks = 0 Then
                elapsed = New TimeSpan([stop].Ticks - start.Ticks)
            End If

            _TransferElapsed = elapsed

            If rate <= 0 Then
                rate = (progress / elapsed.TotalSeconds)
            End If

            _TransferRate = rate

            _ErrorCode = errorcode
            _ErrorMessage = errormsg

        End Sub

    End Class

    Public Interface IFileTransferPlugInModule
        Inherits IPlugInModule

        Event FileTransferProgress(sender As Object, e As FileTransferModuleEventArgs)

        ReadOnly Property TransferFilename As String

        ReadOnly Property TransferState As FileTransferModuleEventStates

        ReadOnly Property TransferFileSize As Long

        ReadOnly Property TransferProgress As Long

        ReadOnly Property TransferErrorCode As Integer

        ReadOnly Property TransferErrorMessage As String

        ReadOnly Property TransferStart As Date

        ReadOnly Property TransferStop As Date

        ReadOnly Property TransferElapsed As TimeSpan

        ReadOnly Property TransferRate As Single

    End Interface

    ''' <summary>
    ''' Base class for a plug-in module.
    ''' </summary>
    Public MustInherit Class PlugInModule
        Implements IPlugInModule

        Protected _ProcThread As Thread
        Protected _Changed As Boolean
        Protected _Title As String
        Protected _Suspend As Boolean
        Protected _Owner As DependencyObject

        Public Event Finished As INotifyStatusProgress.FinishedEventHandler Implements INotifyStatusProgress.Finished
        Public Event Progress As INotifyStatusProgress.ProgressEventHandler Implements INotifyStatusProgress.Progress
        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
        Public Event Starting As INotifyStatusProgress.StartingEventHandler Implements INotifyStatusProgress.Starting
        Public Event RequestFocus As IPlugInModule.RequestFocusEventHandler Implements IPlugInModule.RequestFocus
        Public Event RequestBlur As IPlugInModule.RequestBlurEventHandler Implements IPlugInModule.RequestBlur

#Region "Instantiation"

        ''' <summary>
        ''' Create a new object with the specified Dispatch owner.
        ''' </summary>
        ''' <param name="owner">Dispatch owner.</param>
        Public Sub New(owner As DependencyObject)
            _Owner = owner
        End Sub

#End Region

#Region "Methods"

        ''' <summary>
        ''' Returns true of the task thread is running.
        ''' </summary>
        ''' <returns></returns>
        Protected Function IsThreadRunning() As Boolean
            If _ProcThread IsNot Nothing AndAlso _ProcThread.ThreadState = ThreadState.Running Then Return Not DenyProcessing Else Return False
        End Function

        ''' <summary>
        ''' Cancel all currently running threads and tasks in this module.
        ''' </summary>
        Public Sub CancelTasks() Implements IPlugInModule.CancelTasks
            If _ProcThread IsNot Nothing Then
                _ProcThread.Abort()
            End If
        End Sub

        Public MustOverride Function Initialize() As Boolean Implements IPlugInModule.Initialize

        Public MustOverride Function Shutdown() As Boolean Implements IPlugInModule.Shutdown

#End Region

#Region "Properties"

        ''' <summary>
        ''' Gets or sets the DependencyObject that owns the Dispatcher.
        ''' </summary>
        ''' <returns></returns>
        Public Overridable Property DispatchOwner As DependencyObject
            Get
                Return _Owner
            End Get
            Set(value As DependencyObject)
                _Owner = value
                OnPropertyChanged("Owner")
            End Set
        End Property

        ''' <summary>
        ''' Returns true if there is an active process in this module.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property IsProcessing As Boolean Implements IPlugInModule.IsProcessing
            Get
                Return IsThreadRunning()
            End Get
        End Property

        ''' <summary>
        ''' Set this value to True to prevent this module from starting any processes.
        ''' </summary>
        ''' <returns></returns>
        Public Property DenyProcessing As Boolean Implements IPlugInModule.DenyProcessing
            Get
                Return _Suspend
            End Get
            Set(value As Boolean)
                _Suspend = value
                OnPropertyChanged("DenyProcessing")
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the title text of this module.
        ''' </summary>
        ''' <returns></returns>
        Public Property Title As String Implements IPlugInModule.Title
            Get
                Return _Title
            End Get
            Set(value As String)
                _Title = value
                OnPropertyChanged("Title")
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets a value indicating that data has changed.
        ''' </summary>
        ''' <returns></returns>
        Public Property Changed As Boolean Implements IPlugInModule.Changed
            Get
                Return _Changed
            End Get
            Set(value As Boolean)
                If _Changed = value Then Return

                _Changed = value

                If value Then
                    If _Title.Substring(0, 1) <> "*" Then
                        Title = "*" & _Title
                    End If
                Else
                    If _Title.IndexOf("*") = 0 Then
                        Title = _Title.Substring(1)
                    End If
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets the menu dropdown for this module.
        ''' </summary>
        ''' <returns></returns>
        Public MustOverride ReadOnly Property PlugInMenuItem As MenuItem Implements IPlugInModule.PlugInMenuItem

#End Region

#Region "Status Update Shortcut Functions"

        '' Just raising the events
        Protected Sub OnStart(msg As String, Optional pos As Integer = 0, Optional count As Integer = 0, Optional data As Object = Nothing, Optional mode As StatusModes = StatusModes.ProgressBar)
            RaiseEvent Starting(Me, New StatusProgressEventArgs(StatusCodes.Starting, msg, pos, count,,, data, mode))
        End Sub

        Protected Sub OnProgress(msg As String, Optional pos As Integer = 0, Optional count As Integer = 0, Optional data As Object = Nothing, Optional mode As StatusModes = StatusModes.ProgressBar)
            RaiseEvent Progress(Me, New StatusProgressEventArgs(StatusCodes.Running, msg, pos, count,,, data, mode))
        End Sub

        Protected Sub OnStop(msg As String)
            RaiseEvent Finished(Me, New StatusProgressEventArgs(StatusCodes.Starting, msg, 0, 0))
        End Sub

        Protected Sub OnAbort(msg As String)
            RaiseEvent Finished(Me, New StatusProgressEventArgs(StatusCodes.Aborted, msg, 0, 0))
        End Sub

        '' Using Invoke
        Protected Sub InvokeStart(msg As String, Optional pos As Integer = 0, Optional count As Integer = 0, Optional data As Object = Nothing, Optional mode As StatusModes = StatusModes.ProgressBar)
            Invoke(Sub() RaiseEvent Starting(Me, New StatusProgressEventArgs(StatusCodes.Starting, msg, pos, count,,, data, mode)))
        End Sub

        Protected Sub InvokeProgress(msg As String, Optional pos As Integer = 0, Optional count As Integer = 0, Optional data As Object = Nothing, Optional mode As StatusModes = StatusModes.ProgressBar)
            Invoke(Sub() RaiseEvent Progress(Me, New StatusProgressEventArgs(StatusCodes.Running, msg, pos, count,,, data, mode)))
        End Sub

        Protected Sub InvokeStop(msg As String)
            Invoke(Sub() RaiseEvent Finished(Me, New StatusProgressEventArgs(StatusCodes.Starting, msg, 0, 0)))
        End Sub

        Protected Sub InvokeAbort(msg As String)
            Invoke(Sub() RaiseEvent Finished(Me, New StatusProgressEventArgs(StatusCodes.Aborted, msg, 0, 0)))
        End Sub

        '' Using BeginInvoke
        Protected Sub BeginInvokeStart(msg As String, Optional pos As Integer = 0, Optional count As Integer = 0, Optional data As Object = Nothing, Optional mode As StatusModes = StatusModes.ProgressBar)
            BeginInvoke(Sub() RaiseEvent Starting(Me, New StatusProgressEventArgs(StatusCodes.Starting, msg, pos, count,,, data, mode)))
        End Sub

        Protected Sub BeginInvokeProgress(msg As String, Optional pos As Integer = 0, Optional count As Integer = 0, Optional data As Object = Nothing, Optional mode As StatusModes = StatusModes.ProgressBar)
            BeginInvoke(Sub() RaiseEvent Progress(Me, New StatusProgressEventArgs(StatusCodes.Running, msg, pos, count,,, data, mode)))
        End Sub

        Protected Sub BeginInvokeStop(msg As String)
            BeginInvoke(Sub() RaiseEvent Finished(Me, New StatusProgressEventArgs(StatusCodes.Starting, msg, 0, 0)))
        End Sub

        Protected Sub BeginInvokeAbort(msg As String)
            BeginInvoke(Sub() RaiseEvent Finished(Me, New StatusProgressEventArgs(StatusCodes.Aborted, msg, 0, 0)))
        End Sub

#End Region

#Region "Misc Interface Implementations"

        Protected Overloads Sub OnPropertyChanged(prop As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(prop))
        End Sub

        Public Sub OnProgress(sender As Object, e As StatusProgressEventArgs) Implements INotifyStatusProgress.OnProgress
            RaiseEvent Progress(sender, e)
        End Sub

        Public Sub OnFinished(sender As Object, e As StatusProgressEventArgs) Implements INotifyStatusProgress.OnFinished
            RaiseEvent Finished(sender, e)
        End Sub

        Public Sub OnStarting(sender As Object, e As StatusProgressEventArgs) Implements INotifyStatusProgress.OnStarting
            RaiseEvent Starting(sender, e)
        End Sub

        Public Sub Invoke(func As [Delegate]) Implements INotifyStatusProgress.Invoke
            _Owner.Dispatcher.Invoke(func)
        End Sub

        Public Sub BeginInvoke(func As [Delegate]) Implements INotifyStatusProgress.BeginInvoke
            _Owner.Dispatcher.BeginInvoke(func)
        End Sub

#End Region

    End Class

End Namespace