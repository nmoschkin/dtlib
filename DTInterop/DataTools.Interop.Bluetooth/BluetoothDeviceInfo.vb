'' ************************************************* ''
'' DataTools Visual Basic Utility Library - Interop
''
'' Module: BlueTooth information (TO DO)
'' 
'' Copyright (C) 2011-2020 Nathan Moschkin
'' All Rights Reserved
''
'' Licensed Under the Microsoft Public License   
'' ************************************************* ''


Imports DataTools.Memory
Imports DataTools.Interop.Native
Imports System.ComponentModel
Imports System.Reflection
Imports System.Collections.ObjectModel
Imports System.Windows.Media.Imaging
Imports DataTools.Interop.Printers
Imports DataTools.SystemInfo
Imports DataTools.Interop.Desktop


'' TO DO!  
#Region "Bluetooth Device Info"

Public Class BluetoothDeviceInfo
    Inherits DeviceInfo

    Public Overrides ReadOnly Property UIDescription As String
        Get
            If String.IsNullOrEmpty(FriendlyName) Then Return MyBase.UIDescription Else Return FriendlyName
        End Get
    End Property

    Public Overrides Function ToString() As String
        If String.IsNullOrEmpty(FriendlyName) Then Return MyBase.ToString Else Return FriendlyName
    End Function

End Class

#End Region
