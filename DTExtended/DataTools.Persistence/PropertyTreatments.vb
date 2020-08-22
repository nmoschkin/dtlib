Imports System.ComponentModel

Namespace Persistence

#Region "Property Treatments"

    Public Class PropertyTreatmentEventArgs
        Inherits EventArgs

        Private m_OwnerType As System.Type
        Private m_PropertyName As String
        Private m_Treatment As PropertyTreatments

        Public Property Treatment() As PropertyTreatments
            Get
                Return m_Treatment
            End Get
            Friend Set(ByVal value As PropertyTreatments)
                m_Treatment = value
            End Set
        End Property

        Public Property PropertyName() As String
            Get
                Return m_PropertyName
            End Get
            Friend Set(ByVal value As String)
                m_PropertyName = value
            End Set
        End Property

        Public Property OwnerType() As System.Type
            Get
                Return m_OwnerType
            End Get
            Friend Set(ByVal value As System.Type)
                m_OwnerType = value
            End Set
        End Property

        Friend Sub New()

        End Sub

        Friend Sub New(ownerType As System.Type, propertyName As String, treatment As PropertyTreatments)
            m_OwnerType = ownerType
            m_PropertyName = propertyName
            m_Treatment = treatment
        End Sub

    End Class

    Public Enum PropertyTreatments As Integer
        None = 0
        Ignore = 1
        Force = 2
    End Enum

    <Serializable()>
    Public Class PropertyTreatmentObject

        Private m_OwnerType As System.Type
        Private m_PropertyName As String
        Private m_Treatment As DataTools.Persistence.PropertyTreatments = PropertyTreatments.None

        Public Property Treatment() As DataTools.Persistence.PropertyTreatments
            Get
                Return m_Treatment
            End Get
            Set(ByVal value As DataTools.Persistence.PropertyTreatments)
                m_Treatment = value
            End Set
        End Property

        Public Property PropertyName() As String
            Get
                Return m_PropertyName
            End Get
            Set(ByVal value As String)
                m_PropertyName = value
            End Set
        End Property

        Public Property OwnerType() As System.Type
            Get
                Return m_OwnerType
            End Get
            Set(ByVal value As System.Type)
                m_OwnerType = value
            End Set
        End Property

        Friend Sub New()

        End Sub

        Friend Sub New(ownerType As System.Type, propertyName As String, treatment As PropertyTreatments)
            m_OwnerType = ownerType
            m_PropertyName = propertyName
            m_Treatment = treatment
        End Sub

    End Class

    <Serializable()>
    Public Class PropertyTreatmentObjects
        Inherits CollectionBase

        Public Event Changed(sender As Object, e As PropertyTreatmentEventArgs)

        Public ReadOnly Property InnerArray() As PropertyTreatmentObject()
            Get
                Return CType(InnerList.ToArray(GetType(PropertyTreatmentObject)), PropertyTreatmentObject())
            End Get
        End Property

        Default Public Property Item(index As Integer) As PropertyTreatmentObject
            Get
                Return CType(List(index), PropertyTreatmentObject)
            End Get
            Set(value As PropertyTreatmentObject)
                List(index) = CType(value, PropertyTreatmentObject)
            End Set
        End Property

        Private Function FindTreatmentIndex(ownerType As System.Type, propertyName As String) As Integer
            Dim objTreat As PropertyTreatmentObject
            Dim n As Integer = 0

            For Each objTreat In List
                If objTreat.OwnerType = ownerType AndAlso objTreat.PropertyName = propertyName Then Return n
                n += 1
            Next

            Return -1
        End Function

        Private Function FindTreatmentObject(ownerType As System.Type, propertyName As String) As PropertyTreatmentObject
            Dim objTreat As PropertyTreatmentObject

            For Each objTreat In List
                If (objTreat.OwnerType = ownerType OrElse objTreat.OwnerType Is Nothing) AndAlso (objTreat.PropertyName = propertyName) Then Return objTreat
            Next

            Return Nothing
        End Function

        Public Function GetTreatment(ownerType As System.Type, propertyName As String) As PropertyTreatments
            Dim objTreat As PropertyTreatmentObject = FindTreatmentObject(ownerType, propertyName)
            If objTreat Is Nothing Then Return PropertyTreatments.None

            Return objTreat.Treatment
        End Function

        Public Sub SetTreatment(ownerType As System.Type, propertyName As String, treatment As PropertyTreatments)
            Dim objTreat As PropertyTreatmentObject = FindTreatmentObject(ownerType, propertyName)

            If objTreat Is Nothing Then
                Dim objProps As PropertyDescriptorCollection = TypeDescriptor.GetProperties(ownerType),
                objProp As PropertyDescriptor

                objProp = objProps.Find(propertyName, False)
                If objProp Is Nothing Then
                    Throw New ArgumentNullException("Property not found.")
                Else
                    If objProp.IsReadOnly = True AndAlso treatment <> PropertyTreatments.Force Then
                        Throw New ReadOnlyException("Property is read-only or has no publicly accessable set method.")
                    End If
                End If

                objTreat = New PropertyTreatmentObject(ownerType, propertyName, treatment)
                List.Add(objTreat)

                RaiseEvent Changed(Me, New PropertyTreatmentEventArgs(ownerType, propertyName, treatment))
                Return
            End If

            objTreat.Treatment = treatment
            RaiseEvent Changed(Me, New PropertyTreatmentEventArgs(ownerType, propertyName, treatment))
        End Sub

        Public Overloads Function Add(ownerType As System.Type, propertyName As String, treatment As PropertyTreatments) As Integer
            SetTreatment(ownerType, propertyName, treatment)
            Return List.Count - 1
        End Function

        Friend Overloads Function Add(value As PropertyTreatmentObject) As Integer

            Return List.Add(value)

        End Function 'Add

        Public Overloads Function IndexOf(value As PropertyTreatmentObject) As Integer
            Return List.IndexOf(value)
        End Function 'IndexOf(value)

        Public Overloads Function IndexOf(ownerType As System.Type, propertyName As String) As Integer
            Return FindTreatmentIndex(ownerType, propertyName)
        End Function

        Friend Sub Insert(index As Integer, value As PropertyTreatmentObject)
            List.Insert(index, value)

        End Sub 'Insert

        Friend Sub Remove(value As PropertyTreatmentObject)
            List.Remove(value)
        End Sub 'Remove

        Friend Shadows Sub Clear()
            MyBase.Clear()
        End Sub

        Public Overloads Function Contains(value As PropertyTreatmentObject) As Boolean
            ' If value is not of type PropertyTreatmentObject, this will return false.
            Return List.Contains(value)
        End Function 'Contains

        Public Overloads Function Contains(ownerType As System.Type, propertyName As String) As Boolean
            If (FindTreatmentObject(ownerType, propertyName) IsNot Nothing) Then Return True
            Return False
        End Function

        Protected Overrides Sub OnInsert(index As Integer, value As Object)
            ' Insert additional code to be run only when inserting values.
        End Sub 'OnInsert

        Protected Overrides Sub OnRemove(index As Integer, value As Object)

            ' Insert additional code to be run only when removing values.
        End Sub 'OnRemove

        Protected Overrides Sub OnSet(index As Integer, oldValue As Object, newValue As Object)
            ' Insert additional code to be run only when setting values.
        End Sub 'OnSet

        Protected Overrides Sub OnValidate(value As Object)
            If Not GetType(PropertyTreatmentObject).IsAssignableFrom(value.GetType()) Then
                Throw New ArgumentException("value must be of type Object.", "value")
            End If
        End Sub 'OnValidate 

        Public Sub New()
            '' initialization code goes here
        End Sub

    End Class

#End Region


End Namespace