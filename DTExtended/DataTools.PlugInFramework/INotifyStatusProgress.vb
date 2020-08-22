
'' The INotifyStatusProgress framework.
'' Copyright (C) 2015 Nathan Moschkin
'' All Rights Reserved.


Namespace PlugInFramework

    ''' <summary>
    ''' Status display characteristics.
    ''' </summary>
    <Flags>
    Public Enum StatusModes
        StaticDisplay = 1
        ProgressBar = 2
        ProgressStatic = 4
        Marquis = 8
    End Enum

    ''' <summary>
    ''' Work progress status codes.
    ''' </summary>
    Public Enum StatusCodes
        Aborted = -1
        Stopped = 0
        Starting = 1
        Running = 2
        Paused = 3
    End Enum

    ''' <summary>
    ''' Provides an common interface for classes that do long-running out-of-process work.
    ''' </summary>
    Public Interface INotifyStatusProgress

        ''' <summary>
        ''' Raised while the work is progressing.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Event Progress(sender As Object, e As StatusProgressEventArgs)

        ''' <summary>
        ''' Raised when work is completed.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Event Finished(sender As Object, e As StatusProgressEventArgs)

        ''' <summary>
        ''' Raised when work is starting.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Event Starting(sender As Object, e As StatusProgressEventArgs)

        ''' <summary>
        ''' Call this while the work is progressing.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Sub OnProgress(sender As Object, e As StatusProgressEventArgs)

        ''' <summary>
        ''' Call this when the work is complete.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Sub OnFinished(sender As Object, e As StatusProgressEventArgs)

        ''' <summary>
        ''' Call this when the work is starting.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Sub OnStarting(sender As Object, e As StatusProgressEventArgs)

        ''' <summary>
        ''' Call this to invoke a method on another thread or the Dispatcher thread.
        ''' The implementing class must provide for the Dispatcher or other means of invoking a method.
        ''' </summary>
        ''' <param name="func">The delegate function to invoke.</param>
        Sub Invoke(func As [Delegate])


        ''' <summary>
        ''' Call this to begin the process of invoking a method on another thread or the Dispatcher thread.
        ''' The implementing class must provide for the Dispatcher or other means of invoking a method.
        ''' </summary>
        ''' <param name="func">The delegate function to invoke.</param>
        Sub BeginInvoke(func As [Delegate])

    End Interface

    ''' <summary>
    ''' Arguments to be used when calling events implemented in <see cref="INotifyStatusProgress" />
    ''' </summary>
    Public Class StatusProgressEventArgs
        Inherits EventArgs

        Friend _pos As Integer
        Friend _count As String
        Friend _msg As String
        Friend _data As Object
        Friend _msgcode As Integer
        Friend _detail As String
        Friend _mode As StatusModes
        Friend _scode As StatusCodes = StatusCodes.Stopped

        ''' <summary>
        ''' Gets or sets a value indicating a desire to abort the current process.
        ''' </summary>
        ''' <returns></returns>
        Public Property Abort As Boolean

        ''' <summary>
        ''' Gets the current status display characteristics.
        ''' </summary>
        ''' <returns></returns>
        Public Property Mode As StatusModes
            Get
                Return _mode
            End Get
            Set(value As StatusModes)
                _mode = value
            End Set
        End Property

        ''' <summary>
        ''' Gets the current status code.
        ''' </summary>
        ''' <returns></returns>
        Public Property StatusCode As StatusCodes
            Get
                Return _scode
            End Get
            Set(value As StatusCodes)
                _scode = value
            End Set
        End Property

        ''' <summary>
        ''' Gets the numeric message code.
        ''' </summary>
        ''' <returns></returns>
        Public Property MessageCode As Integer
            Get
                Return _msgcode
            End Get
            Set(value As Integer)
                _msgcode = value
            End Set
        End Property

        ''' <summary>
        ''' Gets the current position.
        ''' </summary>
        ''' <returns></returns>
        Public Property Position As Integer
            Get
                Return _pos
            End Get
            Set(value As Integer)
                _pos = value
            End Set
        End Property

        ''' <summary>
        ''' Gets the total number of elements.
        ''' </summary>
        ''' <returns></returns>
        Public Property Count As Integer
            Get
                Return _count
            End Get
            Set(value As Integer)
                _count = value
            End Set
        End Property

        ''' <summary>
        ''' Gets the status message text.
        ''' </summary>
        ''' <returns></returns>
        Public Property Message As String
            Get
                Return _msg
            End Get
            Set(value As String)
                _msg = value
            End Set
        End Property

        ''' <summary>
        ''' Gets the status detail text.
        ''' </summary>
        ''' <returns></returns>
        Public Property Detail As String
            Get
                Return _detail
            End Get
            Set(value As String)
                _detail = value
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets miscellaneous data.
        ''' </summary>
        ''' <returns></returns>
        Public Property Data As Object
            Get
                Return _data
            End Get
            Set(value As Object)
                _data = value
            End Set
        End Property

        Public Sub New()

        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="mode">A StatusModes value indicating the desired display characteristics.</param>
        ''' <param name="msg">The textual status message.</param>
        ''' <param name="count">The total number of elements.</param>
        ''' <param name="detail">Optional detailed textual status message.</param>
        ''' <param name="code">Optional numeric message code.</param>
        ''' <param name="data">Optional miscellaneous data.</param>
        Public Sub New(mode As StatusModes, msg As String, count As Integer, Optional detail As String = Nothing, Optional code As Integer = 0, Optional data As Object = Nothing)

            _scode = StatusCodes.Starting
            _mode = mode
            _msg = msg
            _pos = 0
            _count = count
            _data = data
            _msgcode = code
            _detail = detail

        End Sub

        ''' <summary>
        ''' Creates a new StatusProgressEventArgs object with the specified parameters.
        ''' </summary>
        ''' <param name="status">A StatusCodes value indicating the type of status update.</param>
        ''' <param name="msg">The textual status message.</param>
        ''' <param name="pos">The numeric position.</param>
        ''' <param name="count">The total number of elements.</param>
        ''' <param name="detail">Optional detailed textual status message.</param>
        ''' <param name="code">Optional numeric message code.</param>
        ''' <param name="data">Optional miscellaneous data.</param>
        ''' <param name="mode">Optional StatusModes value indicating the desired display characteristics.</param>
        Public Sub New(status As StatusCodes, msg As String, pos As Integer, count As Integer, Optional detail As String = Nothing, Optional code As Integer = 0, Optional data As Object = Nothing, Optional mode As StatusModes = StatusModes.ProgressBar)
            _scode = status
            _mode = mode
            _msg = msg
            _pos = pos
            _count = count
            _data = data
            _msgcode = code
            _detail = detail
        End Sub


    End Class

End Namespace