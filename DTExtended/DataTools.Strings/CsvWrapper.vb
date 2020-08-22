''
'' Hugely useful, ultra-secret CSV Wrappers. 
'' Copyright (C) 2012-2015 Nathaniel Moschkin.
'' All Rights Reserved.

'' Commercial use prohibited without express authorization from Nathaniel Moschkin.

Imports System
Imports System.IO
Imports System.Text
Imports System.Reflection
Imports DataTools
Imports DataTools.BinarySearch
Imports DataTools.ByteOrderMark
Imports DataTools.Memory
Imports System.Threading
Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.Globalization

#Region "ICsvRow"

''' <summary>
''' Common interface for CSV records.
''' </summary>
''' <remarks></remarks>
Public Interface ICsvRow

    ''' <summary>
    ''' Get the values for a CSV record.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetValues() As String()

    ''' <summary>
    ''' Set the values for a CSV record.
    ''' </summary>
    ''' <param name="vals"></param>
    ''' <remarks></remarks>
    Sub SetValues(vals As String())

    ''' <summary>
    ''' Get the raw text for a CSV record.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Text As String

End Interface

#End Region

#Region "Enums and Sundry"

''' <summary>
''' Describes a column type for sorting.
''' </summary>
''' <remarks></remarks>
Public Enum ColumnType

    ''' <summary>
    ''' No type, auto, or default.
    ''' </summary>
    ''' <remarks></remarks>
    None

    ''' <summary>
    ''' Explicity textual column type.
    ''' </summary>
    ''' <remarks></remarks>
    Text

    ''' <summary>
    ''' Explicitly numeric column type.
    ''' </summary>
    ''' <remarks></remarks>
    Number
End Enum

''' <summary>
''' Specifies which kind of members of a class in a collection to import.
''' </summary>
''' <remarks></remarks>
<Flags, DefaultValue(1)>
Public Enum ImportFlags

    ''' <summary>
    ''' Any public property.
    ''' </summary>
    ''' <remarks></remarks>
    Any = 0

    ''' <summary>
    ''' Only browsable public properties.
    ''' </summary>
    ''' <remarks></remarks>
    Browsable = 1

    ''' <summary>
    ''' Only DataMember public properties.
    ''' </summary>
    ''' <remarks></remarks>
    DataMember = 2

    ''' <summary>
    ''' Include non-public properties.
    ''' </summary>
    ''' <remarks></remarks>
    NonPublic = 4

    ''' <summary>
    ''' Use descriptions, where available.
    ''' </summary>
    ''' <remarks></remarks>
    Descriptions = &H80

End Enum

#End Region

#Region "CSV Wrapper Classes"

''' <summary>
''' Encapsulates an entire CSV file, including all headers and record data, loading and saving data, and sorting and searching data.
''' </summary>
''' <remarks></remarks>
Public Class CsvWrapper

    Implements IEnumerable(Of CsvRow), ICollection(Of CsvRow), INotifyCollectionChanged, INotifyPropertyChanged

    Public Event CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Implements INotifyCollectionChanged.CollectionChanged

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    Friend _Lines As String()
    Friend _Cols As CsvRow

#Region "Data Import From ICollection"

    Public Function ImportCollection(Of T)(col As ICollection(Of T), Optional flags As ImportFlags = ImportFlags.Browsable Or ImportFlags.Descriptions) As Boolean
        Try
            '' will contain column names
            Dim scol As New List(Of String)

            '' collection of property info
            Dim mi() As PropertyInfo

            '' valid properties will wind up here.
            Dim ml As New List(Of PropertyInfo)

            '' check for public and non public orientations
            If flags And ImportFlags.NonPublic Then
                mi = GetType(T).GetProperties(BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.NonPublic)
            Else
                mi = GetType(T).GetProperties(BindingFlags.Public Or BindingFlags.Instance)
            End If

            Dim dm As Boolean, _
                br As Boolean

            Dim mt As PropertyInfo = Nothing
            Dim lfin As New List(Of PropertyInfo)
            Dim s As String

            For Each x In mi

                br = Persistence.HasAttribute(x, GetType(BrowsableAttribute))
                dm = Persistence.HasAttribute(x, GetType(System.Runtime.Serialization.DataMemberAttribute))

                '' check for browsable requirement
                '' check for datamember requirement
                If (((flags And ImportFlags.Browsable) = 0) OrElse (br = True)) AndAlso _
                   (((flags And ImportFlags.DataMember) = 0) OrElse (dm = True)) Then

                    ml.Add(x)
                End If

            Next

            If flags And ImportFlags.Descriptions Then

                '' with the Description directive, where this is no description, those columns are excluded.
                '' to include columns without a description directive, do not specify the Description flag.
                For Each x In ml
                    s = Persistence.GetDescription(x)
                    If s IsNot Nothing Then
                        scol.Add(s)
                        lfin.Add(x)
                    End If
                Next
            Else
                For Each x In ml
                    scol.Add(x.Name)
                    lfin.Add(x)
                Next
            End If

            '' all set, let's make the data.
            Me.Clear()

            _Cols = New CsvRow(Me)
            _Cols.List = scol

            scol.Clear()
            Dim o As Object

            For Each rec In col
                For Each p In lfin
                    o = p.GetValue(rec)
                    If o IsNot Nothing Then
                        scol.Add(o.ToString)
                    Else
                        scol.Add("")
                    End If
                Next

                Me.AddRow(New CsvRow(Me, scol.ToArray))
                scol.Clear()
            Next

        Catch ex As Exception

            Return False

        End Try

        Return True

    End Function


#End Region

#Region "File Handling"

    Public Property FileName As String

    ''' <summary>
    ''' Open the CSV document named in the Filename property.
    ''' </summary>
    ''' <returns>True if successful.</returns>
    ''' <remarks></remarks>
    Public Function OpenDocument() As Boolean
        If String.IsNullOrWhiteSpace(Me.FileName) OrElse File.Exists(Me.FileName) = False Then Return False
        Return OpenDocument(FileName)
    End Function

    ''' <summary>
    ''' Open a CSV document.
    ''' </summary>
    ''' <param name="fileName">Full path of the file to read.</param>
    ''' <returns>True if successful.</returns>
    ''' <remarks></remarks>
    Public Function OpenDocument(fileName As String) As Boolean
        If File.Exists(fileName) = False Then Return False

        Try
            Dim fs As New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read)
            OpenDocument = OpenDocument(fs)
            If OpenDocument Then
                Me.FileName = fileName
                Return True
            End If

        Catch ex As Exception
            MsgBox("Cannot open '" & IO.Path.GetFileName(fileName) & "': " & ex.Message)
        End Try

        Return False

    End Function

    ''' <summary>
    ''' Open a CSV document from a stream.
    ''' </summary>
    ''' <param name="stream"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function OpenDocument(stream As System.IO.Stream) As Boolean
        Dim b() As Byte
        Try
            ReDim b(CInt(stream.Length) - 1)
            stream.Read(b, 0, b.Length)

            Lines = CType(GetLines(SafeTextRead(b)), StringBlob)
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    ''' <summary>
    ''' Save a CSV document using the current Filename property.
    ''' </summary>
    ''' <returns>True if successful.</returns>
    ''' <remarks></remarks>
    Public Function SaveDocument() As Boolean
        If String.IsNullOrWhiteSpace(Me.FileName) Then Return False
        Return SaveDocument(FileName)
    End Function

    ''' <summary>
    ''' Save a CSV document to the specified file.
    ''' </summary>
    ''' <param name="fileName">Full path of the file to save.</param>
    ''' <returns>True if successful.</returns>
    ''' <remarks></remarks>
    Public Function SaveDocument(fileName As String) As Boolean
        Dim fs As New FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None)
        SaveDocument = SaveDocument(fs)
        If SaveDocument Then
            Me.FileName = fileName
        End If
    End Function

    ''' <summary>
    ''' Save a CSV document to a stream.
    ''' </summary>
    ''' <param name="stream">Stream to receive the document.</param>
    ''' <returns>True if successful.</returns>
    ''' <remarks></remarks>
    Public Function SaveDocument(stream As System.IO.Stream) As Boolean
        Return SaveDocument(stream, True)
    End Function

    ''' <summary>
    ''' Save a CSV document to a stream.
    ''' </summary>
    ''' <param name="stream">Stream to receive the document.</param>
    ''' <param name="closeStream">True to close the stream after writing.</param>
    ''' <returns>True if successful.</returns>
    ''' <remarks></remarks>
    Public Function SaveDocument(stream As System.IO.Stream, closeStream As Boolean) As Boolean
        Try
            Dim b() As Byte = SafeTextWrite(CType(Lines, StringBlob).ToFormattedString(StringBlobFormats.CrLf))
            stream.Write(b, 0, b.Length)

            If closeStream Then stream.Close()
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

#End Region

#Region "Content Management"

    ''' <summary>
    ''' Returns the names of all the columns in this csv document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ColumnNames As CsvRow
        Get
            Return _Cols
        End Get
        Set(value As CsvRow)
            _Cols = value
            OnPropertyChanged("ColumnNames")
            RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the text lines of a CSV file.
    ''' For setting, column names must be present.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Lines As String()
        Get
            Dim l As New List(Of String)
            If _Cols IsNot Nothing Then
                l.Add(_Cols.Text)
            End If
            l.AddRange(_Lines)
            Return l.ToArray
        End Get
        Set(value As String())
            Dim i As Integer
            Dim c As Integer = value.Count - 1

            ReDim _Lines(c - 1)

            For i = 1 To c
                _Lines(i - 1) = value(i)
            Next

            _Cols = value(0)
            OnPropertyChanged("Lines")
            RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
        End Set
    End Property

    ''' <summary>
    ''' Add a raw CSV record to the CSV file.
    ''' </summary>
    ''' <param name="text"></param>
    ''' <remarks></remarks>
    Public Sub AddLine(text As String)

        If _Lines Is Nothing Then
            ReDim _Lines(0)
            _Lines(0) = text
            Return
        End If

        Dim c As Integer = _Lines.Count
        ReDim Preserve _Lines(c)

        _Lines(c) = text
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Add))
    End Sub

    ''' <summary>
    ''' Gets or sets a CsvRow object to the specified row index.
    ''' </summary>
    ''' <param name="index">Index at which to get or set the object.</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Default Public Property Row(index As Integer) As CsvRow
        Get
            Return New CsvRow(Me, _Lines(index))
        End Get
        Set(value As CsvRow)
            _Lines(index) = CType(value, String)
            RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Replace))
        End Set
    End Property

    ''' <summary>
    ''' Returns the entire, formatted text string for a CSV file.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Text As String
        Get
            Return CType(Lines, StringBlob).ToFormattedString(StringBlobFormats.CrLf)
        End Get
        Set(value As String)
            Lines = GetLines(value)
            OnPropertyChanged("Text")
            RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the text of the column names row (the first row.)
    ''' </summary>
    ''' <returns></returns>
    Public Property ColumnText As String
        Get
            Return CStr(New CsvRow(ColumnNames))
        End Get
        Set(value As String)
            _Cols = New CsvRow(value)
        End Set
    End Property

    ''' <summary>
    ''' Returns all the values for the specified column as an array of strings.
    ''' </summary>
    ''' <param name="index">The 0-based numeric index of the column to retrieve.</param>
    ''' <returns></returns>
    Public Function GetColumnData(index As Integer) As String()
        Dim lns As New List(Of String)

        For Each r As CsvRow In Me
            lns.Add(r.List(index))
        Next

        Return lns.ToArray
    End Function

    ''' <summary>
    ''' Returns all the values for the specified column as an array of strings.
    ''' </summary>
    ''' <param name="columnName">The name of the column to retrieve.</param>
    ''' <returns></returns>
    Public Function GetColumnData(columnName As String) As String()
        Dim x As Integer = GetColumnIndex(columnName)
        Dim lns As New List(Of String)

        If x = -1 Then Return Nothing

        For Each r As CsvRow In Me
            lns.Add(r.List(x))
        Next

        Return lns.ToArray
    End Function

    ''' <summary>
    ''' Find the index of the specified column; not case-sensitive.
    ''' </summary>
    ''' <param name="columnName"></param>
    ''' <returns>The column index or -1 if no such column was found.</returns>
    Public Function GetColumnIndex(columnName As String) As Integer
        Dim i As Integer,
            c As Integer = _Cols.Count - 1

        For i = 0 To c

            If _Cols.List(i) = columnName Then Return i

        Next

        Return -1
    End Function

    ''' <summary>
    ''' Gets the maximum number of columns based on the current raw string record set.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function MaxCols() As Integer
        Dim cc As Integer = 0

        For Each c As CsvRow In _Lines
            If c.Count > cc Then cc = c.Count
        Next

        If _Cols.Count > cc Then cc = _Cols.Count

        Return cc
    End Function

    ''' <summary>
    ''' Truncate the CSV record set to the specified index.
    ''' </summary>
    ''' <param name="lastIndex">Last item index to preserve.</param>
    ''' <remarks></remarks>
    Public Sub Truncate(lastIndex As Integer)

        Dim a As New List(Of String)

        Dim i As Integer
        lastIndex = Math.Min(lastIndex, _Lines.Count - 1)

        For i = 0 To lastIndex
            a.Add(_Lines(i))
        Next

        _Lines = a.ToArray
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Remove))

    End Sub

#End Region

#Region "Export"


    ''' <summary>
    ''' Extract specific columns from this CSV file to a new CSV file.
    ''' </summary>
    ''' <param name="columns">Numerical list of columns, starting with 0.</param>
    ''' <returns></returns>
    Public Function ExtractColumns(ParamArray columns() As Integer) As CsvWrapper

        If (columns Is Nothing OrElse columns.Count = 0) Then Return Nothing

        Dim i As Integer,
            c As Integer = columns.Length


        Dim d As Integer

        Dim csvOut As New CsvWrapper

        Dim st As New List(Of String)


        For i = 0 To c - 1
            d = columns(i)

            st.Add(Me.ColumnNames(d))


        Next

        csvOut.ColumnNames = st.ToArray


        For Each r As CsvRow In Me
            st.Clear()



            For i = 0 To c - 1
                d = columns(i)


                st.Add(r.Columns(d))
            Next


            csvOut.AddRow(New CsvRow(st.ToArray))

        Next


        Return csvOut

    End Function

    '''' <summary>
    '''' Extract specific columns from this CSV file to a new CSV file.
    '''' </summary>
    '''' <param name="columns">Alphabetical list of columns, starting with A.</param>
    '''' <returns></returns>
    'Public Function ExtractColumns(ParamArray columns() As String) As CsvWrapper


    '    Dim n As New List(Of Integer)
    '    Dim i As Integer

    '    For Each s As String In columns

    '        i = ColumnFromLetters(s)
    '        n.Add(i)

    '    Next

    '    Return ExtractColumns(n.ToArray)

    'End Function


    '''' <summary>
    '''' Gets a column number from a column letter combination akin to Excel
    '''' </summary>
    '''' <param name="letters">String of letters.</param>
    '''' <returns></returns>
    'Public Shared Function ColumnFromLetters(letters As String) As Integer

    '    Dim l As New List(Of Char)
    '    Dim x As Integer = 0
    '    Dim fres As Integer = 0
    '    Dim c As Integer = 0

    '    l.AddRange(letters)

    '    For Each s As String In l


    '        s = s.ToUpper
    '        x = (AscW(s) - AscW("A")) + (26 * If(c = 0, 0, 1))
    '        fres += x
    '        c += 1

    '    Next



    '    Return fres

    'End Function


#End Region

#Region "ICollection implementation"

    ''' <summary>
    ''' Add a CsvRow object to the CSV file.
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Public Sub AddRow(row As CsvRow) Implements ICollection(Of CsvRow).Add
        AddLine(row.Text)
    End Sub

    ''' <summary>
    ''' Copies the contents of this record set to an array of CsvRows
    ''' </summary>
    ''' <param name="array">The array to copy to.</param>
    ''' <param name="arrayIndex">The index within the array at which to start the copying.</param>
    ''' <remarks></remarks>
    Public Sub CopyTo(array() As CsvRow, arrayIndex As Integer) Implements ICollection(Of CsvRow).CopyTo
        Dim c As Integer = arrayIndex
        Dim d As Integer = RowCount - 1

        For i = 0 To d
            array(c) = Row(i)
            c += 1
        Next

    End Sub

    ''' <summary>
    ''' Gets the number of rows in the current record set.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property RowCount As Integer Implements ICollection(Of CsvRow).Count
        Get
            If _Lines Is Nothing Then Return 0
            Return _Lines.Count
        End Get
    End Property

    ''' <summary>
    ''' Gets a value indicating that the list is read-only.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of CsvRow).IsReadOnly
        Get
            Return False
        End Get
    End Property

    ''' <summary>
    ''' Clear the entire record set, excluding column names.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear() Implements ICollection(Of CsvRow).Clear
        _Lines = {}
    End Sub

    ''' <summary>
    ''' Returns the number of records in the CSV file.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Count As Integer
        Get
            If _Lines Is Nothing Then Return 0
            Return _Lines.Count
        End Get
    End Property


    ''' <summary>
    ''' Determines whether or not the CSV file contains the specified CsvRow object.
    ''' The exact text of a rendered CSV record must match the exact text of the rendered item for this function to succeed.
    ''' </summary>
    ''' <param name="item">A CsvRow object to scan for.</param>
    ''' <returns>True if the row is found.</returns>
    ''' <remarks></remarks>
    Public Function Contains(item As CsvRow) As Boolean Implements ICollection(Of CsvRow).Contains

        Dim i As Integer,
            c As Integer = _Lines.Count - 1

        Dim t As String = item.Text

        For i = 0 To c
            If _Lines(i) = t Then
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Remove a CsvRow item from the current record set.
    ''' The exact text of a rendered CSV record must match the exact text of the rendered item for this function to succeed.
    ''' </summary>
    ''' <param name="item">A CsvRow object to scan for.</param>
    ''' <returns>True if a row was removed.</returns>
    ''' <remarks></remarks>
    Public Function Remove(item As CsvRow) As Boolean Implements ICollection(Of CsvRow).Remove

        Dim i As Integer,
            c As Integer = _Lines.Count - 1

        Dim b As Boolean = False

        Dim ln As New List(Of String)

        For i = 0 To c
            If item.Text <> _Lines(i) Then
                ln.Add(_Lines(i))
            Else
                b = True
            End If
        Next

        _Lines = ln.ToArray
        If b Then RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Remove))

        Return b
    End Function

#End Region

#Region "Sort and Search by Column"

    ''' <summary>
    ''' Saves an index from a sort.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class IndexSaverObject
        Public Property Text As String
        Public Property Index As Integer

        Friend Sub New()

        End Sub
    End Class

    ''' <summary>
    ''' Sorts the entire CSV record set by the specified column, in ascending order.  
    ''' To sort in another order, or in a method other than a string comparison, pass an IComparer object.
    ''' </summary>
    ''' <param name="columnIndex">The column index to sort by.</param>
    ''' <param name="comparer">The optional IComparer object to use for comparisons.</param>
    ''' <returns>True if successful.  False may indicate there are no records or the columnIndex value is greater than the greatest available column index for this record set or less than zero.</returns>
    ''' <remarks></remarks>
    Public Function SortByColumn(columnIndex As Integer, Optional comparer As IComparer(Of IndexSaverObject) = Nothing, Optional desc As Boolean = False) As Boolean
        If _Lines Is Nothing OrElse _Lines.Count = 0 Then Return False
        If Me.ColumnNames Is Nothing OrElse columnIndex >= Me.ColumnNames.Count OrElse columnIndex < 0 Then Return False

        Dim cl As New List(Of IndexSaverObject)
        Dim ll As IndexSaverObject

        Dim e As Integer

        For e = 0 To Me.RowCount - 1

            ll = New IndexSaverObject

            ll.Text = Me.Row(e).List(columnIndex)
            ll.Index = e

            cl.Add(ll)
            ll = Nothing
        Next

        If comparer Is Nothing Then

            cl.Sort(New Comparison(Of IndexSaverObject) _
                    (Function(x As IndexSaverObject, y As IndexSaverObject) As Integer
                         Return If(desc, -String.Compare(x.Text, y.Text), String.Compare(x.Text, y.Text))
                     End Function))

        Else
            cl.Sort(comparer)
        End If

        Dim lns As New List(Of String)

        For Each cc In cl
            lns.Add(_Lines(cc.Index))
        Next

        _Lines = lns.ToArray

        GC.Collect(0)
        Return True
    End Function

    ''' <summary>
    ''' Searches a sorted record set by columnIndex for the specified value using a binary search algorithm.
    ''' </summary>
    ''' <param name="value">The value to search for.</param>
    ''' <param name="columnIndex">The column to search.</param>
    ''' <param name="noPreSort">Specifies that the binary searcher will not attempt to pre-sort the list by the specified column.</param>
    ''' <param name="comparer">The optional IComparer object to use for comparisons during sort.</param>
    ''' <returns>The index of the found record or -1 if no record was found.</returns>
    ''' <remarks></remarks>
    Public Function BinarySearchByColumn(value As String, columnIndex As Integer, Optional noPreSort As Boolean = True, Optional comparer As IComparer(Of IndexSaverObject) = Nothing) As Integer

        If Me.Count <= 0 Then Return -1

        If Me.Count <= 4 Then
            Dim x As Integer = 0

            For Each r As CsvRow In Me
                If r.List(columnIndex) = value Then Return x
            Next
        End If

        If Not noPreSort Then
            Me.SortByColumn(columnIndex, comparer)
        End If

        Dim count As Integer = Me.Count
        Dim currentIdx As Integer = count / 2

        Dim startPos As Integer = 0
        Dim endPos As Integer = count - 1

        Dim valDiff As Integer
        Dim nextCenter As Integer

        Dim propVal As String

        Do
            propVal = (Me.Row(currentIdx).List(columnIndex))
            valDiff = String.Compare(value, propVal)

            If valDiff = 0 Then
                Return currentIdx
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

        Return -1
    End Function

#End Region

#Region "Enumerator"

    Public Function GetEnumerator() As IEnumerator(Of CsvRow) Implements IEnumerable(Of CsvRow).GetEnumerator
        Return New CsvRowEnumerator(Me)
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return New CsvRowEnumerator(Me)
    End Function
#End Region

#Region "PropertyChanged"

    Protected Overloads Sub OnPropertyChanged(e As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(e))
    End Sub

    Protected Overloads Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

#End Region

End Class

Public Class CsvKeyValue
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    Private _Key As String
    Private _Value As String

    Private _Row As CsvRow
    Private _Index As Integer

    Private _UpdateParentRow As Boolean

    ''' <summary>
    ''' Specify whether or not to update the row data of the row object
    ''' that owns this key/value pair if the value is changed.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property UpdateParentRow As Boolean
        Get
            Return _UpdateParentRow
        End Get
        Set(value As Boolean)
            _UpdateParentRow = True
            OnPropertyChanged("UpdateParentRow")
        End Set
    End Property

    ''' <summary>
    ''' The index of this column.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Property Index As Integer
        Get
            Return _Index
        End Get
        Set(value As Integer)
            _Index = value
            OnPropertyChanged("Index")
        End Set
    End Property

    ''' <summary>
    ''' The name of this column.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Key As String
        Get
            Return _Key
        End Get
        Set(value As String)
            _Key = value
            OnPropertyChanged("Key")
        End Set
    End Property

    ''' <summary>
    ''' The value of this column.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Value As String
        Get
            Return _Value
        End Get
        Set(value As String)
            If _Value = value Then Return
            _Value = value

            If _UpdateParentRow Then
                _Row.List(_Index) = value
            End If

            OnPropertyChanged("Value")
        End Set
    End Property

    ''' <summary>
    ''' Find the index of a key given its name.
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FindKeyNumber(key As String) As Integer
        Dim s() As String = _Row.Parent.ColumnNames

        Dim i As Integer, _
            c As Integer = s.Count - 1

        For i = 0 To c
            If key = s(i) Then Return i
        Next

        Return -1
    End Function

    ''' <summary>
    ''' Find the name of a key given its index.
    ''' </summary>
    ''' <param name="index"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FindKeyName(index As Integer) As String
        Return _Row.Parent.ColumnNames(index)
    End Function


    ''' <summary>
    ''' Initialize a new Key/Value pair based on key name and value.
    ''' 0 based column index will be determined automatically.
    ''' </summary>
    ''' <param name="key">Column name.</param>
    ''' <param name="value">Value of column.</param>
    ''' <remarks></remarks>
    Friend Sub New(key As String, value As String)
        Me.Index = FindKeyNumber(key)
        Me.Key = key
        Me.Value = value
    End Sub

    ''' <summary>
    ''' Initialize a new Key/Value pair based on column index and value.
    ''' Column name will be determined automatically.
    ''' </summary>
    ''' <param name="index">0 Based column index.</param>
    ''' <param name="value">Value of column.</param>
    ''' <remarks></remarks>
    Friend Sub New(index As Integer, value As String)
        Me.Key = FindKeyNumber(index)
        Me.Index = index
        Me.Value = value
    End Sub


    Protected Overloads Sub OnPropertyChanged(e As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(e))
    End Sub

    Protected Overloads Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub


End Class

Public Class CsvRow
    Implements ICsvRow, IEnumerable(Of String), INotifyPropertyChanged, ICloneable

    Protected _Fields As New List(Of String)

    Private _parent As CsvWrapper

    ''' <summary>
    ''' Gets the parent CSV file.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Parent As CsvWrapper
        Get
            Return _parent
        End Get
        Friend Set(value As CsvWrapper)
            _parent = value
            OnPropertyChanged("Parent")
        End Set
    End Property

    ''' <summary>
    ''' Get the CSV-formatted text string for the entire row.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Text As String Implements ICsvRow.Text
        Get
            Dim sb As New StringBuilder, _
                c As Integer = 0

            For Each s In _Fields
                If c > 0 Then sb.Append(",")

                sb.Append("""")

                s = s.Replace(vbCrLf, vbCr)
                s = s.Replace("""", """""")

                sb.Append(s)
                sb.Append("""")

                c += 1
            Next

            Return sb.ToString
        End Get
        Set(value As String)

            If String.IsNullOrWhiteSpace(value) Then Return
            _Fields.Clear()

            '' let's just normalize our line breaks

            '' make everything one thing.
            value = value.Replace(vbCrLf, vbCr)
            value = value.Replace(vbLf, vbCr)

            '' change that one thing to the appropriate thing.
            value = value.Replace(vbCr, vbCrLf)

            Dim sa As String() = (BatchParse(value, ",", True, True, """"c, """"c, False, False, True))
            _Fields.AddRange(sa)

            OnPropertyChanged("Text")
        End Set
    End Property

    ''' <summary>
    ''' Get all values for this record as a list.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property List As List(Of String)
        Get
            Return _Fields
        End Get
        Set(value As List(Of String))
            _Fields = value
            OnPropertyChanged("List")
        End Set
    End Property

    ''' <summary>
    ''' Get all values for this record.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Columns() As String()
        Get
            Return _Fields.ToArray
        End Get
        Set(value As String())
            _Fields.Clear()
            _Fields.AddRange(value)
            OnPropertyChanged("Columns")
        End Set
    End Property

    ''' <summary>
    ''' Get all values for this record.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetValues() As String() Implements ICsvRow.GetValues
        GetValues = Columns
    End Function

    ''' <summary>
    ''' Set all values for this record.
    ''' </summary>
    ''' <param name="vals"></param>
    ''' <remarks></remarks>
    Public Sub SetValues(vals() As String) Implements ICsvRow.SetValues
        Columns = vals
    End Sub


    ''' <summary>
    ''' Returns the column index from the given name. 
    ''' The CsvRow object must be associated with a parent CsvWrapper with a list of column names for this function to succeed.
    ''' </summary>
    ''' <param name="name">Case-insensitive name of the column to retrieve the index for.</param>
    ''' <returns>Column index, or -1 if not found.</returns>
    ''' <remarks></remarks>
    Public Function GetColumnIndex(name As String) As Integer
        If _parent Is Nothing Then Return -1

        Dim i As Integer = 0

        For Each s In _parent.ColumnNames
            If name.Trim.ToUpper = s.Trim.ToUpper Then
                Return i
            End If
            i += 1
        Next

        Return -1
    End Function


    ''' <summary>
    ''' Returns the value for the given column name. 
    ''' The CsvRow object must be associated with a parent CsvWrapper with a list of column names for this function to succeed.
    ''' </summary>
    ''' <param name="name">Case-insensitive name of the column to retrieve the data for.</param>
    ''' <returns>Data, or null if not found.</returns>
    ''' <remarks></remarks>
    Public Function GetColumnData(name As String) As String
        If _parent Is Nothing Then Return -1

        Dim i As Integer = 0

        For Each s In _parent.ColumnNames
            If name.Trim.ToUpper = s.Trim.ToUpper Then
                Return Me.List(i)
            End If
            i += 1
        Next

        Return Nothing
    End Function


    ''' <summary>
    ''' Get key/value pairs for all contents of this record.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetKeyValues() As CsvKeyValue()
        Dim i As Integer, _
            c As Integer = _Fields.Count - 1

        Dim out As New List(Of CsvKeyValue)

        For i = 0 To c
            out.Add(New CsvKeyValue(i, _Fields(i)))
        Next

        Return out.ToArray
    End Function


    ''' <summary>
    ''' Normalize the data to number of columns (or the number of columns in the parent)
    ''' </summary>
    ''' <param name="c">Optional number of columns.</param>
    ''' <remarks></remarks>
    Public Sub Normalize(Optional c As Integer = 0)
        Dim d As Integer

        If c <= 0 Then
            If _parent IsNot Nothing Then
                c = _parent.ColumnNames.Count
            Else
                Return
            End If
        End If

        If _Fields.Count < c Then
            d = _Fields.Count
            Do Until d = c
                _Fields.Add("")
                d += 1
            Loop
        Else
            Do Until _Fields.Count = c
                _Fields.RemoveAt(_Fields.Count - 1)
            Loop
        End If

    End Sub

    ''' <summary>
    ''' Create a new row.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        _Fields = New List(Of String)
    End Sub

    ''' <summary>
    ''' Create a new row with the specified parent.
    ''' </summary>
    ''' <param name="parent">Parent CsvWrapper object.</param>
    ''' <remarks></remarks>
    Public Sub New(parent As CsvWrapper)
        _Fields = New List(Of String)
        _parent = parent
    End Sub


    ''' <summary>
    ''' Create a new row with the specified raw CSV row text.
    ''' </summary>
    ''' <param name="text">Raw CSV row text.</param>
    ''' <remarks></remarks>
    Public Sub New(text As String)
        Me.Text = text
    End Sub

    ''' <summary>
    ''' Create a new row with the specified parent and raw CSV row text.
    ''' </summary>
    ''' <param name="parent">The parent CsvWrapper object.</param>
    ''' <param name="text">Raw CSV row text.</param>
    ''' <remarks></remarks>
    Public Sub New(parent As CsvWrapper, text As String)
        Me.Text = text
        _parent = parent
        Normalize()
    End Sub

    ''' <summary>
    ''' Create a new row with the specified parent and column values.
    ''' </summary>
    ''' <param name="parent">The parent CsvWrapper object.</param>
    ''' <param name="values">A list of string values to fill the content of columns.</param>
    ''' <remarks></remarks>
    Public Sub New(parent As CsvWrapper, ParamArray values() As String)
        _parent = parent
        Dim l As New List(Of String)

        For Each v In values
            l.Add(v)
        Next

        Me._Fields = l
        Normalize()
    End Sub

    ''' <summary>
    ''' Create a new row with the specified column values.
    ''' </summary>
    ''' <param name="values">A list of string values to fill the content of columns.</param>
    ''' <remarks></remarks>
    Public Sub New(ParamArray values() As String)
        Dim l As New List(Of String)

        For Each v In values
            l.Add(v)
        Next

        Me._Fields = l
        Normalize()
    End Sub

#Region "Operators"

    Public Shared Widening Operator CType(operand As String) As CsvRow
        Return New CsvRow(operand)
    End Operator

    Public Shared Widening Operator CType(operand As CsvRow) As String
        Return operand.Text
    End Operator

    Public Shared Narrowing Operator CType(operand As CsvRow) As String()
        Return operand.GetValues
    End Operator

    Public Shared Narrowing Operator CType(operand As String()) As CsvRow
        Dim c As New CsvRow
        c._Fields.AddRange(operand)
        Return c
    End Operator

#End Region

#Region "IClonable"

    Public Function Clone() As Object Implements ICloneable.Clone
        Dim r As New CsvRow
        For Each s In _Fields
            r._Fields.Add(s)
        Next

        r.Parent = _parent
        Return r
    End Function

    Public Function Clone(parent As CsvWrapper) As CsvRow
        Dim r As New CsvRow
        For Each s In _Fields
            r._Fields.Add(s)
        Next

        r.Parent = parent
        Return r
    End Function

#End Region

#Region "IEnumerable"

    Public Function GetEnumerator() As IEnumerator(Of String) Implements IEnumerable(Of String).GetEnumerator
        Return New StringEnumerator(_Fields)
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return New StringEnumerator(_Fields)
    End Function

#End Region

#Region "INotifyPropertyChanged"

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    Protected Overloads Sub OnPropertyChanged(e As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(e))
    End Sub

    Protected Overloads Sub OnPropertyChanged(e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

#End Region

    Public Overrides Function ToString() As String
        Return Text
    End Function

End Class

#End Region

#Region "CsvRow Enumerator"

Public Class CsvRowEnumerator
    Implements IEnumerator(Of CsvRow)

    Dim subj As CsvWrapper
    Dim pos As Integer = -1

    Friend Sub New(subject As CsvWrapper)
        subj = subject
    End Sub

    Public ReadOnly Property Current As CsvRow Implements IEnumerator(Of CsvRow).Current
        Get
            Current = subj(pos)
        End Get
    End Property

    Public ReadOnly Property Current1 As Object Implements IEnumerator.Current
        Get
            Current1 = subj(pos)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        pos += 1
        If pos >= subj.RowCount Then Return False
        Return True
    End Function

    Public Sub Reset() Implements IEnumerator.Reset
        pos = -1
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class

#End Region

#Region "String Enumerator"

Public Class StringEnumerator
    Implements IEnumerator(Of String)

    Dim subj As List(Of String)
    Dim pos As Integer = -1

    Friend Sub New(subject As List(Of String))
        subj = subject
    End Sub

    Public ReadOnly Property Current As String Implements IEnumerator(Of String).Current
        Get
            Current = subj(pos)
        End Get
    End Property

    Public ReadOnly Property Current1 As Object Implements IEnumerator.Current
        Get
            Current1 = subj(pos)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        pos += 1
        If pos >= subj.Count Then Return False
        Return True
    End Function

    Public Sub Reset() Implements IEnumerator.Reset
        pos = -1
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class

#End Region

#Region "Class Writer"

''' <summary>
''' Provides tools for generating VB.NET class code from a CSV record set.
''' </summary>
Public NotInheritable Class CSVClassWriter

    Private Shared _Dict As BinarySearch.TextBinarySearcher

    Private Sub New()

    End Sub

    ''' <summary>
    ''' Generate a VB.NET Class from a CSV record set, with type auto-determination.
    ''' </summary>
    ''' <param name="csv">The record set to base the new class on.</param>
    ''' <param name="name">The name of the new class.</param>
    ''' <param name="showSave">Specifies whether to show the save file dialog when complete.</param>
    ''' <param name="propChange">Specifies the class should implement INotifyPropertyChanged.</param>
    ''' <param name="useDictionary">Specifies whether to use the English dictionary to analyze column names for proper camel casing.</param>
    ''' <param name="setterType">Specifies the default scope of property setters; acceptable values are Public, Private, or Friend.</param>
    ''' <param name="knownTypes">Specifies a collection of known types for specific columns.</param>
    ''' <param name="makeReadOnly">Specifies that the class will be generated read-only. If this value is True setterType is ignored.</param>
    ''' <returns></returns>
    Public Shared Function WriteClass(csv As CsvWrapper, Optional name As String = "Class1", Optional showSave As Boolean = False, Optional propChange As Boolean = True, Optional useDictionary As Boolean = True, Optional setterType As String = "Public", Optional knownTypes As IEnumerable(Of KeyValuePair(Of Integer, System.Type)) = Nothing, Optional makeReadOnly As Boolean = False) As String

        Dim pcol As New List(Of System.Type)
        Dim b As New StringBuilder
        Dim output As String
        Dim pname As String
        Dim cnames As New List(Of String)

        Dim i As Integer,
            c As Integer = csv.ColumnNames.Count - 1

        Dim ind As Integer = 1
        Dim data() As String

        If makeReadOnly Then propChange = False

        Select Case setterType.ToLower
            Case "public"
                setterType = ""
            Case "private"
                setterType = "Private "
            Case "friend"
                setterType = "Friend "
            Case Else
                Throw New ArgumentException("Invalid scope.", "setterType")

        End Select

        For i = 0 To c
            cnames.Add(csv.ColumnNames(i))
        Next

        If useDictionary Then
            For i = 0 To c
                cnames(i) = SmartCamel(cnames(i))
            Next
        End If

        For i = 0 To c
            data = csv.GetColumnData(i)
            pcol.Add(DetermineLikelyType(data))
        Next

        If knownTypes IsNot Nothing Then
            For Each kt In knownTypes

                If kt.Key >= 0 AndAlso kt.Key <= c Then
                    pcol(kt.Key) = kt.Value
                End If
            Next
        End If


        If propChange Then
            b.AppendLine("Imports System.ComponentModel")
            b.AppendLine()
        End If

        b.AppendLine("Public Class " & name)

        If propChange Then
            b.AppendLine(New String(" ", 4 * ind) & "Implements INotifyPropertyChanged")
            b.AppendLine(New String(" ", 4 * ind) & "Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged")
            b.AppendLine()
        End If

        b.AppendLine()

        For i = 0 To c

            b.AppendLine(New String(" ", 4 * ind) & "Private _" & cnames(i) & " As " & pcol(i).Name)

        Next

        b.AppendLine()

        If propChange Then
            b.AppendLine(New String(" ", 4 * ind) & "Public Sub OnPropertyChanged(propName As String)")
            b.AppendLine(New String(" ", 12 * ind) & "RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))")
            b.AppendLine(New String(" ", 4 * ind) & "End Sub")

            b.AppendLine()
        End If


        For i = 0 To c

            If pcol(i).Namespace <> "System" Then
                pname = pcol(i).FullName
            Else
                pname = pcol(i).Name
            End If

            If makeReadOnly Then

                b.AppendLine(New String(" ", 4 * ind) & "PUblic ReadOnly Property " & cnames(i) & " As " & pname)
            Else
                b.AppendLine(New String(" ", 4 * ind) & "PUblic Property " & cnames(i) & " As " & pname)

            End If

            b.AppendLine(New String(" ", 8 * ind) & "Get")
            b.AppendLine(New String(" ", 12 * ind) & "Return _" & cnames(i))
            b.AppendLine(New String(" ", 8 * ind) & "End Get")

            If Not makeReadOnly Then
                b.AppendLine(New String(" ", 8 * ind) & setterType & "Set(value As " & pname & ")")

                b.AppendLine(New String(" ", 12 * ind) & "_" & cnames(i) & " = value")


                If propChange Then
                    b.AppendLine(New String(" ", 12 * ind) & "OnPropertyChanged(""" & cnames(i) & """)")
                End If

                b.AppendLine(New String(" ", 8 * ind) & "End Set")
            End If

            b.AppendLine(New String(" ", 4 * ind) & "End Property")
            b.AppendLine()

        Next

        '' write th setters

        b.AppendLine(New String(" ", 4 * ind) & "PUblic Sub SetRecord(data As CsvRow)")
        b.AppendLine()


        For i = 0 To c

            If pcol(i).Namespace <> "System" Then
                pname = pcol(i).FullName
            Else
                pname = pcol(i).Name
            End If

            If pcol(i) IsNot GetType(String) AndAlso pcol(i).GetMethod("Parse", {GetType(String)}) IsNot Nothing Then
                b.AppendLine(New String(" ", 8 * ind) & "_" & cnames(i) & " = " & pname & ".Parse(data.List(" & i & "))")
            Else
                b.AppendLine(New String(" ", 8 * ind) & "_" & cnames(i) & " = data.List(" & i & ")")
            End If
        Next

        b.AppendLine()
        b.AppendLine(New String(" ", 4 * ind) & "End Sub")
        b.AppendLine()

        If Not makeReadOnly Then
            b.AppendLine(New String(" ", 4 * ind) & "PUblic Sub GetRecord(ByRef data As CsvRow)")
            b.AppendLine(New String(" ", 8 * ind) & "Dim output As New List(Of String)")
            b.AppendLine()
            b.AppendLine(New String(" ", 8 * ind) & "If data Is Nothing Then data = New CsvRow()")


            For i = 0 To c
                b.AppendLine(New String(" ", 8 * ind) & "Output.Add(_" & cnames(i) & ".ToString())")
            Next

            b.AppendLine()
            b.AppendLine(New String(" ", 8 * ind) & "data.List = output")

            b.AppendLine()
            b.AppendLine(New String(" ", 4 * ind) & "End Sub")
            b.AppendLine()

        End If

        '' write the constructors

        b.AppendLine(New String(" ", 4 * ind) & "PUblic Sub New(data As CsvRow)")
        b.AppendLine(New String(" ", 8 * ind) & "SetRecord(data)")
        b.AppendLine(New String(" ", 4 * ind) & "End Sub")
        b.AppendLine()


        b.AppendLine("End Class")

        output = b.ToString

        '' let's do some VB cleanup
        output = output.Replace("Int32", "Integer").Replace("Int64", "Long")
        output = output.Replace("As DateTime", "As Date")

        If showSave Then
            Dim dlg As New System.Windows.Forms.SaveFileDialog

            dlg.Title = "Save New VB Class"
            dlg.Filter = "Visual Basic Code Files (*.vb)|*.vb|All Files (*.*)|*.*"
            dlg.FileName = name & ".vb"


            If dlg.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                File.WriteAllText(dlg.FileName, output)
            End If
        End If

        Return output
    End Function

    ''' <summary>
    ''' Attempt to determine the property type from one cell of data.  
    ''' This is not recommended. It would be better to use a sample column of data for more accurate results.
    ''' </summary>
    ''' <param name="sampleData"></param>
    ''' <returns></returns>
    Public Shared Function DetermineLikelyType(sampleData As String) As System.Type
        Return DetermineLikelyType({sampleData})
    End Function

    ''' <summary>
    ''' Attempt to determine the property type of a column based on its data.
    ''' </summary>
    ''' <param name="sampleData"></param>
    ''' <returns></returns>
    Public Shared Function DetermineLikelyType(sampleData As String()) As System.Type

        Dim sng As Single,
            dbl As Double,
            lng As Long,
            int As Integer,
            ulng As ULong

        Dim dat As Date
        Dim g As Guid

        Dim maxInt As Integer

        Dim scan As String
        Dim soFar As System.Type = GetType(String)

        For Each scan In sampleData
            scan = scan.Trim

            If Date.TryParse(scan, dat) Then
                soFar = GetType(Date)
                Continue For
            End If

            If Guid.TryParse(scan, g) Then
                soFar = GetType(Guid)
                Continue For
            End If

            If IsNumber(scan) Then
                If scan.IndexOf(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator) >= 0 Then

                    Single.TryParse(scan, sng)
                    Double.TryParse(scan, dbl)

                    If dbl > Single.MaxValue OrElse dbl < Single.MinValue OrElse (sng <> dbl) Then
                        soFar = GetType(Double)
                    Else
                        soFar = GetType(Single)
                    End If

                ElseIf soFar IsNot GetType(Double) AndAlso soFar IsNot GetType(Single) Then

                    If Long.TryParse(scan, lng) Then

                        If lng <= UInteger.MaxValue AndAlso lng >= UInteger.MinValue Then
                            soFar = GetType(UInteger)
                        Else
                            soFar = GetType(Long)
                        End If

                        If Integer.TryParse(scan, int) Then
                            If maxInt < int Then maxInt = int

                            If maxInt = 0 OrElse maxInt = 1 Then
                                soFar = GetType(Boolean)
                            Else
                                soFar = GetType(Integer)
                            End If
                        End If

                    ElseIf ULong.TryParse(scan, ulng) Then

                        soFar = GetType(ULong)

                    End If
                End If

                Continue For
            End If

            Select Case scan.ToLower
                Case "yes", "oui", "si", "sí", "ya", "ja", "true", "vrai", "verdad", "on",
                     "no", "non", "nein", "faux", "falso", "false", "off"

                    soFar = GetType(Boolean)

            End Select

            soFar = GetType(String)

        Next

        Return soFar
    End Function


    ''' <summary>
    ''' Uses the complete English dictionary to look for words in a string and turn that string
    ''' into a properly camel-cased string.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="wordList"></param>
    ''' <returns></returns>
    Public Shared Function SmartCamel(value As String, Optional wordList As String() = Nothing) As String
        If value Is Nothing Then Return Nothing
        If value.ToLower = "timestamp" Then Return "TimeStamp"

        Dim i As Integer,
            j As Integer = 0,
            c As Integer = value.Length

        Dim t As String
        Dim w As New List(Of String)
        Dim orig As String = value
        Dim tc As Integer = 0
        Dim ov As String

        Dim brk() As String

        value = value.Replace("_", " ")
        value = value.Replace("-", " ")

        brk = BatchParse(value, " ")


        For Each value In brk
            ov = value
            i = value.Length
            tc = 0
            j = 0

            Do

                t = value.Substring(j, i).ToLower
                If _Dict.SearchFor(t) IsNot Nothing Then

                    j = t.Length
                    If t.Length = value.Length Then
                        w.Add(CamelCase(t))
                        tc = 0
                        Exit Do
                    End If

                    value = value.Substring(j)
                    i = value.Length
                    c = value.Length
                    w.Add(CamelCase(t))
                    tc += t.Length
                    j = 0

                    Continue Do
                End If

                i -= 1
                If i < 1 Then Exit Do
            Loop

            If tc Then
                w.Add(CamelCase(ov.Substring(tc - 1)))
            End If
        Next

        If w.Count = 0 Then Return orig

        t = ""
        For Each s In w
            t &= s
        Next

        Return t
    End Function



    Shared Sub New()
        Dim l As New List(Of String)

        Dim s() As String = File.ReadAllLines("Dictionary.txt")
        l.AddRange(s)

        _Dict = New TextBinarySearcher(l, False)
    End Sub



End Class


#End Region
