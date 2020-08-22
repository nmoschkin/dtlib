Imports System
Imports System.IO
Imports DataTools.Interop.Native
Imports System.Collections.ObjectModel
Imports System.Collections

Namespace Desktop

    ''' <summary>
    ''' Provides a file-system-locked object to represent the contents of a directory.
    ''' </summary>
    Public Class DirectoryObject
        Implements ICollection(Of FileObject)

        Private _Files As New List(Of FileObject)
        Private _Path As String

        ''' <summary>
        ''' Refresh the contents of the directory.
        ''' </summary>
        Public Sub Refresh()
            _Files.Clear()

            Dim f() As String = IO.Directory.GetFiles(_Path)
            Dim fobj As FileObject

            For Each file In f

                fobj = New FileObject(file)
                fobj.Parent = Me

                _Files.Add(fobj)
            Next
        End Sub


        ''' <summary>
        ''' Get a directory object from a file object.
        ''' The <see cref="FileObject"/> instance passed is retained in the returned <see cref="DirectoryObject"/> object.
        ''' </summary>
        ''' <param name="fileObj">A <see cref="FileObject"/></param>
        ''' <returns>A <see cref="DirectoryObject"/>.</returns>
        Public Shared Function FromFileObject(fileObj As FileObject) As DirectoryObject

            Dim dir As New DirectoryObject(fileObj.Directory)

            Dim i As Integer,
                c As Integer = dir.Count - 1

            Dim nlist As New List(Of FileObject)

            For i = 0 To c
                If dir(i).Filename = fileObj.Filename Then
                    nlist.Add(fileObj)
                    fileObj.Parent = dir
                Else
                    nlist.Add(dir(i))
                End If
            Next

            dir._Files = nlist
            Return dir
        End Function

        ''' <summary>
        ''' Returns the full path of the directory.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Directory As String
            Get
                Return _Path
            End Get
        End Property

        ''' <summary>
        ''' Returns an item in the collection.
        ''' </summary>
        ''' <param name="index"></param>
        ''' <returns></returns>
        Default Public ReadOnly Property Item(index As Integer) As FileObject
            Get
                Return _Files(index)
            End Get
        End Property

        ''' <summary>
        ''' Returns the number of files in the directory.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Count As Integer Implements ICollection(Of FileObject).Count
            Get
                Return _Files.Count
            End Get
        End Property

        Private ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of FileObject).IsReadOnly
            Get
                Return True
            End Get
        End Property

        ''' <summary>
        ''' Create a new file-system-linked directory object 
        ''' </summary>
        ''' <param name="path"></param>
        Public Sub New(path As String)

            If IO.Directory.Exists(path) = False Then
                Throw New DirectoryNotFoundException("Directory Not Found: " & path)
            End If

            Refresh()
        End Sub

        Friend Sub Add(item As FileObject) Implements ICollection(Of FileObject).Add
            _Files.Add(item)
        End Sub

        Friend Sub Clear() Implements ICollection(Of FileObject).Clear
            _Files.Clear()
        End Sub

        Public Function Contains(fileName As String, Optional ByRef obj As FileObject = Nothing) As Boolean
            fileName = fileName.ToLower
            For Each f In _Files
                If f.Name.ToLower = fileName Then
                    obj = f
                    Return True
                End If
            Next

            Return False
        End Function

        Public Function Contains(item As FileObject) As Boolean Implements ICollection(Of FileObject).Contains
            Return _Files.Contains(item)
        End Function

        Public Sub CopyTo(array() As FileObject, arrayIndex As Integer) Implements ICollection(Of FileObject).CopyTo
            _Files.CopyTo(array, arrayIndex)
        End Sub

        Friend Function Remove(item As FileObject) As Boolean Implements ICollection(Of FileObject).Remove
            Return _Files.Remove(item)
        End Function

        Public Function GetEnumerator() As IEnumerator(Of FileObject) Implements IEnumerable(Of FileObject).GetEnumerator
            Return New DirEnumer(Me)
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return New DirEnumer(Me)
        End Function


        Private Class DirEnumer
            Implements IEnumerator(Of FileObject)

            Private _obj As DirectoryObject
            Private _pos As Integer = -1


            Public Sub New(obj As DirectoryObject)
                _obj = obj
            End Sub

            Public ReadOnly Property Current As FileObject Implements IEnumerator(Of FileObject).Current
                Get
                    Return _obj(_pos)
                End Get
            End Property

            Private ReadOnly Property IEnumerator_Current As Object Implements IEnumerator.Current
                Get
                    Return _obj(_pos)
                End Get
            End Property

            Public Sub Reset() Implements IEnumerator.Reset
                _pos = -1
            End Sub

            Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
                _pos += 1
                If (_pos >= _obj.Count) Then Return False
                Return True
            End Function

#Region "IDisposable Support"
            Private disposedValue As Boolean ' To detect redundant calls

            ' IDisposable
            Protected Overridable Sub Dispose(disposing As Boolean)
                If Not disposedValue Then
                    If disposing Then
                        _obj = Nothing
                        _pos = -1
                    End If

                End If
                disposedValue = True
            End Sub

            ' This code added by Visual Basic to correctly implement the disposable pattern.
            Public Sub Dispose() Implements IDisposable.Dispose
                Dispose(True)
            End Sub
#End Region
        End Class


    End Class

    ''' <summary>
    ''' Provides a file-system-locked object to represent a file.
    ''' </summary>
    Public Class FileObject

        Private _Filename As String

        Private _Type As SystemFileType
        Private _Parent As DirectoryObject


        ''' <summary>
        ''' Returns the parent directory object if one exists.
        ''' </summary>
        ''' <returns></returns>
        Public Property Parent As DirectoryObject
            Get
                Return _Parent
            End Get
            Friend Set(value As DirectoryObject)
                _Parent = value
            End Set
        End Property


        ''' <summary>
        ''' Return the file type object.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property TypeObject As SystemFileType
            Get
                Return _Type
            End Get
        End Property

        ''' <summary>
        ''' Returns the file type description
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property FileType As String
            Get
                If _Type Is Nothing Then Return "Unknown"
                Return _Type.Description
            End Get
        End Property

        ''' <summary>
        ''' Returns the file type icon
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property FileTypeIcon As System.Drawing.Icon
            Get
                If _Type Is Nothing Then Return Nothing
                Return _Type.DefaultIcon
            End Get
        End Property

        ''' <summary>
        ''' Gets or sets the file attributes.
        ''' </summary>
        ''' <returns></returns>
        Public Property Attributes As IO.FileAttributes
            Get
                Return GetFileAttributes(_Filename)
            End Get
            Set(value As IO.FileAttributes)
                SetFileAttributes(_Filename, value)
            End Set
        End Property

        ''' <summary>
        ''' Get or set the creation time of the file.
        ''' </summary>
        ''' <returns></returns>
        Public Property CreationTime As Date
            Get
                Dim c As Date, a As Date, m As Date

                GetFileTime(_Filename, c, a, m)
                Return c
            End Get
            Set(value As Date)
                Dim c As Date, a As Date, m As Date

                GetFileTime(_Filename, c, a, m)
                SetFileTime(_Filename, value, a, m)
            End Set
        End Property

        ''' <summary>
        ''' Get or set the last write time of the file.
        ''' </summary>
        ''' <returns></returns>
        Public Property LastWriteTime As Date
            Get
                Dim c As Date, a As Date, m As Date

                GetFileTime(_Filename, c, a, m)
                Return m
            End Get
            Set(value As Date)
                Dim c As Date, a As Date, m As Date

                GetFileTime(_Filename, c, a, m)
                SetFileTime(_Filename, c, a, value)
            End Set
        End Property

        ''' <summary>
        ''' Get or set the last access time of the file.
        ''' </summary>
        ''' <returns></returns>
        Public Property LastAccessTime As Date
            Get
                Dim c As Date, a As Date, m As Date

                GetFileTime(_Filename, c, a, m)
                Return a
            End Get
            Set(value As Date)
                Dim c As Date, a As Date, m As Date

                GetFileTime(_Filename, c, a, m)
                SetFileTime(_Filename, c, value, m)
            End Set
        End Property

        ''' <summary>
        ''' Get the size of the file, in bytes.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Size As Long
            Get
                Return GetFileSize(_Filename)
            End Get
        End Property


        ''' <summary>
        ''' Get the full path of the file.
        ''' </summary>
        ''' <returns></returns>
        Public Property Filename As String
            Get
                Return _Filename
            End Get
            Friend Set(value As String)

                If _Filename IsNot Nothing Then
                    If Not Utility.MoveFile(_Filename, value) Then
                        Throw New AccessViolationException("Unable to rename/move file.")
                    End If
                ElseIf Not File.Exists(value) Then
                    Throw New FileNotFoundException("File Not Found: " & Filename)
                End If

                _Filename = value
                Refresh()
            End Set
        End Property

        ''' <summary>
        ''' Get the containing directory of the file.
        ''' </summary>
        ''' <returns></returns>
        Public Property Directory As String
            Get
                Return Path.GetDirectoryName(_Filename)
            End Get
            Set(value As String)
                If Not Move(value) Then
                    Throw New AccessViolationException("Unable to move file.")
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets the name of the file.
        ''' </summary>
        ''' <returns></returns>
        Public Property Name As String
            Get
                Return Path.GetFileName(_Filename)
            End Get
            Friend Set(value As String)
                If Not Rename(value) Then
                    Throw New AccessViolationException("Unable to rename file.")
                End If
            End Set
        End Property


        ''' <summary>
        ''' Attempt to move the file to a new directory.
        ''' </summary>
        ''' <param name="newDirectory">Destination directory.</param>
        ''' <returns>True if successful.</returns>
        Public Function Move(newDirectory As String) As Boolean

            If newDirectory.Substring(newDirectory.Length - 1, 1) = "\" Then newDirectory = newDirectory.Substring(0, newDirectory.Length - 1)
            If Not IO.Directory.Exists(newDirectory) Then Return False

            Dim p As String = Path.GetFileName(_Filename)
            Dim f = newDirectory & "\" & p

            If Utility.MoveFile(_Filename, f) Then

                _Filename = f
                Refresh()

                Return True
            Else
                Return False
            End If

        End Function

        ''' <summary>
        ''' Attempt to rename the file.
        ''' </summary>
        ''' <param name="newName">The new name of the file.</param>
        ''' <returns>True if successful</returns>
        Public Function Rename(newName As String) As Boolean

            Dim p As String = Path.GetDirectoryName(_Filename)
            Dim f As String = p & "\" & newName

            If Not Utility.MoveFile(_Filename, f) Then
                Return False
            End If

            _Filename = f
            Refresh()
            Return True

        End Function

        ''' <summary>
        ''' Create a new FileObject from the given filename. 
        ''' If the file does not exist, an exception will be thrown.
        ''' </summary>
        ''' <param name="filename"></param>
        Public Sub New(filename As String)

            If File.Exists(filename) = False Then
                Throw New FileNotFoundException("File Not Found: " & filename)
            End If

            _Filename = filename
            Refresh()

        End Sub

        ''' <summary>
        ''' Create a blank file object.  
        ''' </summary>
        Friend Sub New()

        End Sub

        ''' <summary>
        ''' Refresh the state of the file object from the disk.
        ''' </summary>
        ''' <returns></returns>
        Public Function Refresh() As Boolean
            If Not File.Exists(_Filename) Then Return False

            _Type = SystemFileType.FromExtension(Path.GetExtension(_Filename),, StandardIcons.Icon16)

            ' if we are no longer in the directory of the original parent, set to null
            If _Parent IsNot Nothing Then
                If _Parent.Directory.ToLower <> Me.Directory.ToLower Then
                    _Parent.Remove(Me)
                    _Parent = Nothing
                End If
            End If

            Return True
        End Function

        Public Overrides Function ToString() As String
            Return Filename
        End Function

    End Class

End Namespace
