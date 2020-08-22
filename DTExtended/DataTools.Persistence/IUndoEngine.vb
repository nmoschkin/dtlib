Namespace Persistence

    Public Enum UndoableActionTypes As Integer
        Typing = &H1
        Drawing = &H2
        Move = &H4
        Size = &H8
        Cut = &H10
        Paste = &H20
        Add = &H40
        Insert = &H80
        Remove = &H100

        SelectionChange = &H200
        PropertyChange = &H400

        Basic = &HFFF

        ElementPropertyChange = &H1400
        ElementAdd = &H1040
        ElementInsert = &H1080
        ElementRemove = &H1100

        Element = &H1000

        Left = &H10000
        Right = &H20000
        Up = &H40000
        Down = &H80000
        Top = &H100000
        Bottom = &H200000
        Width = &H400000
        Height = &H800000
        Horizontal = &H1000000
        Vertical = &H2000000
        Edge = &H4000000

        Layout = &HFFF0000

        Major = &H10000000

    End Enum

    Public Interface IUndoable
        Property Guid As System.Guid

    End Interface

    Public Interface IUndoEngine
        Inherits IUndoable

        Property PropertyTreatment(propertyName As String) As PropertyTreatments

        Property UndoBufferSize As Integer
        Property UndoEnabled As Boolean

        ReadOnly Property CanRedo As Boolean
        ReadOnly Property CanUndo As Boolean

        Sub InitializeUndo(Optional ByVal enableUndo As Boolean = False)
        Sub Record(action As UndoableActionTypes, ParamArray affectedObjects() As IUndoable)

        Function Undo() As Boolean
        Function Redo() As Boolean

    End Interface

End Namespace
