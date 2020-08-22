Namespace Info

    <HideModuleName>
    Public Module MsgBoxResults

        Public Const vbYes = 2
    End Module

    Public Enum YesNoDefault As Byte
        [Default] = 0
        No = 1
        Yes = 2
    End Enum

    Public Enum YesNoAlwaysNeverDefault As Byte
        [Default] = 0
        Yes = 1
        No = 2
        Always = 3
        Never = 4
    End Enum

    Public Enum YesNoAlwaysDefault As Byte
        [Default] = 0
        Yes = 1
        No = 2
        Always = 3
    End Enum

    Public Enum YesNoNeverDefault As Byte
        [Default] = 0
        Yes = 1
        No = 2
        Never = 4
    End Enum

    Public Enum YesNoAlwaysNever As Byte
        Yes = 1
        No = 2
        Always = 3
        Never = 4
    End Enum

    Public Enum YesNoAlways As Byte
        Yes = 1
        No = 2
        Always = 3
    End Enum

    Public Enum YesNoNever As Byte
        Yes = 1
        No = 2
        Never = 4
    End Enum

End Namespace