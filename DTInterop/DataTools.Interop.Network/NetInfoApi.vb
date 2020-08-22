'' ************************************************* ''
'' DataTools Visual Basic Utility Library - Interop
''
'' Module: NetInfoApi
''         Windows Networking Api
''
''         Enums are documented in part from the API documentation at MSDN.
''
'' Copyright (C) 2011-2020 Nathan Moschkin
'' All Rights Reserved
''
'' Licensed Under the Microsoft Public License   
'' ************************************************* ''


Imports DataTools.Interop.Native
Imports System.Runtime.InteropServices
Imports DataTools.Memory
Imports DataTools.SystemInfo

Namespace Network

    ''' <summary>
    ''' Server access authorization flags for the current principal.
    ''' </summary>
    ''' <remarks></remarks>
    <Flags>
    Public Enum AuthFlags

        ''' <summary>
        ''' Printing authorization.
        ''' </summary>
        ''' <remarks></remarks>
        Print = 1

        ''' <summary>
        ''' Communications authorization.
        ''' </summary>
        ''' <remarks></remarks>
        Communications = 2

        ''' <summary>
        ''' File server authorization.
        ''' </summary>
        ''' <remarks></remarks>
        Server = 4

        ''' <summary>
        ''' User account remote access authorization.
        ''' </summary>
        ''' <remarks></remarks>
        Accounts = 8
    End Enum

    ''' <summary>
    ''' The join status for a computer on a Workgroup or Domain.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum NetworkJoinStatus
        ''' <summary>
        ''' Join status is unknown.
        ''' </summary>
        ''' <remarks></remarks>
        Unknown = 0

        ''' <summary>
        ''' Computer is not joined to a network.
        ''' </summary>
        ''' <remarks></remarks>
        Unjoined

        ''' <summary>
        ''' Computer is joined to a Workgroup.
        ''' </summary>
        ''' <remarks></remarks>
        Workgroup

        ''' <summary>
        ''' Computer is joined to a full domain.
        ''' </summary>
        ''' <remarks></remarks>
        Domain
    End Enum

    ''' <summary>
    ''' Extended username formats.  In Workgroups, only NameSamCompatible is supported.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ExtendedNameFormat
        ''' <summary>
        ''' Unknown/invalid format.
        ''' </summary>
        ''' <remarks></remarks>
        NameUnknown = 0

        ''' <summary>
        ''' Fully-qualified domain name.
        ''' </summary>
        ''' <remarks></remarks>
        NameFullyQualifiedDN = 1

        ''' <summary>
        ''' Windows-networking SAM-compatible MACHINE\User formatted string.  This parameter is valid for Workgroups.
        ''' </summary>
        ''' <remarks></remarks>
        NameSamCompatible = 2

        ''' <summary>
        ''' Display name.
        ''' </summary>
        ''' <remarks></remarks>
        NameDisplay = 3

        ''' <summary>
        ''' The user's unique Id.
        ''' </summary>
        ''' <remarks></remarks>
        NameUniqueId = 6

        ''' <summary>
        ''' The canonical name of the user. This usually means all high-bit Unicode characters have been converted to their lower-order representations
        ''' and the string has been normalized to UTF-8 or ASCII as much as possible, including case-hardening.
        ''' </summary>
        ''' <remarks></remarks>
        NameCanonical = 7

        ''' <summary>
        ''' The user principal.
        ''' </summary>
        ''' <remarks></remarks>
        NameUserPrincipal = 8

        ''' <summary>
        ''' Extended formatting canonical name.
        ''' </summary>
        ''' <remarks></remarks>
        NameCanonicalEx = 9

        ''' <summary>
        ''' Service principal.
        ''' </summary>
        ''' <remarks></remarks>
        NameServicePrincipal = 10

        ''' <summary>
        ''' The Dns domain.
        ''' </summary>
        ''' <remarks></remarks>
        NameDnsDomain = 12
    End Enum

    ''' <summary>
    ''' User privilege levels.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum UserPrivilege
        Guest
        User
        Administrator
    End Enum

    Public Module NetInfoApi

        Declare Unicode Function GetUserNameEx Lib "Secur32.dll" Alias "GetUserNameExW" (NameFormat As ExtendedNameFormat,
                                                                                    <MarshalAs(UnmanagedType.LPWStr)> lpNameBuffer As String,
                                                                                    ByRef lpnSize As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean

        Declare Unicode Function GetUserNameEx Lib "Secur32.dll" Alias "GetUserNameExW" (NameFormat As ExtendedNameFormat,
                                                                                    lpNameBuffer As IntPtr,
                                                                                    ByRef lpnSize As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean

        Declare Unicode Function GetUserName Lib "advapi32.dll" Alias "GetUserNameW" (lpNameBuffer As IntPtr,
                                                                                    ByRef lpnSize As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean

        Public Function CurrentUserFullName(Optional type As ExtendedNameFormat = ExtendedNameFormat.NameDisplay)

            Dim lps As MemPtr
            Dim cb As Integer = 10240

            lps.SetLength(10240)
            lps.ZeroMemory()
            If GetUserNameEx(type, lps.Handle, cb) Then
                CurrentUserFullName = lps.ToString
            Else
                CurrentUserFullName = Nothing
            End If

            lps.Free()
        End Function

        Public Function CurrentUserName() As String

            Dim lps As MemPtr
            Dim cb As Integer = 10240

            lps.SetLength(10240)
            lps.ZeroMemory()
            If GetUserName(lps.Handle, cb) Then
                CurrentUserName = lps.ToString
            Else
                CurrentUserName = Nothing
            End If

            lps.Free()

        End Function

        ''' <summary>
        ''' Windows API SERVER_INFO_100 structure.  Name and platform Id for a computer.
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure ServerInfo100

            Public PlatformId As Integer

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Name As String

            Public Overrides Function ToString() As String
                Return Name
            End Function

        End Structure

        ''' <summary>
        ''' Windows Networking SID Types.
        ''' </summary>
        ''' <remarks></remarks>
        <Flags()>
        Public Enum SidUsage
            ''' <summary>
            ''' A user SID.
            ''' </summary>
            ''' <remarks></remarks>
            SidTypeUser = 1

            ''' <summary>
            ''' A group SID.
            ''' </summary>
            ''' <remarks></remarks>
            SidTypeGroup

            ''' <summary>
            ''' A domain SID.
            ''' </summary>
            ''' <remarks></remarks>
            SidTypeDomain

            ''' <summary>
            ''' An alias SID.
            ''' </summary>
            ''' <remarks></remarks>
            SidTypeAlias

            ''' <summary>
            ''' A SID for a well-known group.
            ''' </summary>
            ''' <remarks></remarks>
            SidTypeWellKnownGroup

            ''' <summary>
            ''' A SID for a deleted account.
            ''' </summary>
            ''' <remarks></remarks>
            SidTypeDeletedAccount

            ''' <summary>
            ''' A SID that is not valid.
            ''' </summary>
            ''' <remarks></remarks>
            SidTypeInvalid

            ''' <summary>
            ''' A SID of unknown type.
            ''' </summary>
            ''' <remarks></remarks>
            SidTypeUnknown

            ''' <summary>
            ''' A SID for a computer.
            ''' </summary>
            ''' <remarks></remarks>
            SidTypeComputer

            ''' <summary>
            ''' A mandatory integrity label SID.
            ''' </summary>
            ''' <remarks></remarks>
            SidTypeLabel
        End Enum

        ''' <summary>
        ''' Windows Networking server/computer types.
        ''' </summary>
        ''' <remarks></remarks>
        <Flags()>
        Public Enum ServerTypes
            ''' <summary>
            ''' A workstation.
            ''' </summary>
            ''' <remarks></remarks>
            Workstation = &H1

            ''' <summary>
            ''' A server.
            ''' </summary>
            ''' <remarks></remarks>
            Server = &H2

            ''' <summary>
            ''' A server running with Microsoft SQL Server.
            ''' </summary>
            ''' <remarks></remarks>
            SqlServer = &H4

            ''' <summary>
            ''' A primary domain controller.
            ''' </summary>
            ''' <remarks></remarks>
            DomainController = &H8

            ''' <summary>
            ''' A backup domain controller.
            ''' </summary>
            ''' <remarks></remarks>
            BackupDomainController = &H10

            ''' <summary>
            ''' A server running the Timesource service.
            ''' </summary>
            ''' <remarks></remarks>
            TimeSource = &H20

            ''' <summary>
            ''' A server running the Apple Filing Protocol (AFP) file service.
            ''' </summary>
            ''' <remarks></remarks>
            AFPServer = &H40

            ''' <summary>
            ''' A Novell server.
            ''' </summary>
            ''' <remarks></remarks>
            Novell = &H80

            ''' <summary>
            ''' A LAN Manager 2.x domain member.
            ''' </summary>
            ''' <remarks></remarks>
            DomainMember = &H100

            ''' <summary>
            ''' A server that shares a print queue.
            ''' </summary>
            ''' <remarks></remarks>
            PrintQueueServer = &H200

            ''' <summary>
            ''' A server that runs a dial-in service.
            ''' </summary>
            ''' <remarks></remarks>
            DialInServer = &H400

            ''' <summary>
            ''' A Xenix or Unix server.
            ''' </summary>
            ''' <remarks></remarks>
            XenixServer = &H800

            ''' <summary>
            ''' A workstation or server.
            ''' </summary>
            ''' <remarks></remarks>
            WindowsNT = &H1000

            ''' <summary>
            ''' A computer that runs Windows for Workgroups.
            ''' </summary>
            ''' <remarks></remarks>
            WindowsForWorkgroups = &H2000

            ''' <summary>
            ''' A server that runs the Microsoft File and Print for NetWare service.
            ''' </summary>
            ''' <remarks></remarks>
            NetwareFilePrintServer = &H4000

            ''' <summary>
            ''' Any server that is not a domain controller.
            ''' </summary>
            ''' <remarks></remarks>
            NTServer = &H8000

            ''' <summary>
            ''' A computer that can run the browser service.
            ''' </summary>
            ''' <remarks></remarks>
            PotentialBrowser = &H10000

            ''' <summary>
            ''' A server running a browser service as backup.
            ''' </summary>
            ''' <remarks></remarks>
            BackupBrowser = &H20000

            ''' <summary>
            ''' A server running the master browser service.
            ''' </summary>
            ''' <remarks></remarks>
            MasterBrowser = &H40000

            ''' <summary>
            ''' A server running the domain master browser.
            ''' </summary>
            ''' <remarks></remarks>
            DomainMasterBrowser = &H80000

            ''' <summary>
            ''' A computer that runs OSF.
            ''' </summary>
            ''' <remarks></remarks>
            OSFServer = &H100000

            ''' <summary>
            ''' A computer that runs VMS.
            ''' </summary>
            ''' <remarks></remarks>
            VMSServer = &H200000

            ''' <summary>
            ''' A computer that runs Windows.
            ''' </summary>
            ''' <remarks></remarks>
            Windows = &H400000

            ''' <summary>
            ''' A server that is the root of a DFS tree.
            ''' </summary>
            ''' <remarks></remarks>
            DFSRootServer = &H800000

            ''' <summary>
            ''' A server cluster available in the domain.
            ''' </summary>
            ''' <remarks></remarks>
            NTServerCluster = &H1000000

            ''' <summary>
            ''' A server that runs the Terminal Server service.
            ''' </summary>
            ''' <remarks></remarks>
            TerminalServer = &H2000000

            ''' <summary>
            ''' Cluster virtual servers available in the domain.
            ''' </summary>
            ''' <remarks></remarks>
            NTVSServerCluster = &H4000000

            ''' <summary>
            ''' A server that runs the DCE Directory and Security Services or equivalent.
            ''' </summary>
            ''' <remarks></remarks>
            DCEServer = &H10000000

            ''' <summary>
            ''' A server that is returned by an alternate transport.
            ''' </summary>
            ''' <remarks></remarks>
            AlternateTransport = &H20000000

            ''' <summary>
            ''' A server that is maintained by the browser.
            ''' </summary>
            ''' <remarks></remarks>
            LocalListOnly = &H40000000

            ''' <summary>
            ''' A primary domain.
            ''' </summary>
            ''' <remarks></remarks>
            DomainEnum = &H80000000I
        End Enum


        ''' <summary>
        ''' Windows API SERVER_INFO_101 structure.  Contains extended, vital information
        ''' about a computer on the network.
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure ServerInfo101

            Public PlatformId As Integer

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Name As String
            Public VersionMajor As Integer
            Public VersionMinor As Integer
            Public Type As ServerTypes

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Comment As String

            Public Overrides Function ToString() As String
                Return Name
            End Function
        End Structure

        ''' <summary>
        ''' Windows API NET_DISPLAY_MACHINE structure.
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure NetDisplayMachine

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Name As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Comment As String

            Public Flags As Integer
            Public UserId As Integer
            Public NextIndex As Integer

            Public Overrides Function ToString() As String
                Return Name
            End Function
        End Structure

        ''' <summary>
        ''' Windows API LOCALGROUP_INFO_1 structure.  Basic local group information.
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure LocalGroupInfo1

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Name As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Comment As String

            Public Overrides Function ToString() As String
                Return Name
            End Function
        End Structure

        ''' <summary>
        ''' Windows API GROUP_INFO_0 structure.  Returns only the group name.
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure GroupInfo0

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Name As String

            Public Overrides Function ToString() As String
                Return Name
            End Function
        End Structure

        ''' <summary>
        ''' Windows API GROUP_INFO_1 structure.  Basic group information and attributes.
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure GroupInfo1

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Name As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Comment As String

            Public GroupId As IntPtr
            Public Attributes As Integer

            Public Overrides Function ToString() As String
                Return Name
            End Function
        End Structure

        ''' <summary>
        ''' Windows API GROUP_INFO_2 structure.  Basic group information and attributes.
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure GroupInfo2

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Name As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Comment As String

            Public GroupId As Integer
            Public Attributes As Integer

            Public Overrides Function ToString() As String
                Return Name
            End Function
        End Structure

        ''' <summary>
        ''' Windows API GROUP_INFO_3 structure.  Basic group information and attributes.
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure GroupInfo3

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Name As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Comment As String

            Public GroupId As IntPtr
            Public Attributes As Integer

            Public Overrides Function ToString() As String
                Return Name
            End Function
        End Structure

        ''' <summary>
        ''' Windows API USER_INFO_1 structure.  A moderately verbose report
        ''' version of users on a system.
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure UserInfo1

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Name As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Password As String

            Public PasswordAge As FriendlySeconds

            Public Priv As UserPrivilege

            <MarshalAs(UnmanagedType.LPWStr)>
            Public HomeDir As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Commant As String

            Public Flags As Integer

            <MarshalAs(UnmanagedType.LPWStr)>
            Public ScriptPath As String

            Public Overrides Function ToString() As String
                Return Name
            End Function
        End Structure

        ''' <summary>
        ''' Windows API USER_INFO_11 structure.  A very verbose 
        ''' report version of users on a system.
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure UserInfo11
            'LPWSTR usri11_name;
            '  LPWSTR usri11_comment;
            '  LPWSTR usri11_usr_comment;
            '  LPWSTR usri11_full_name;

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Name As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Comment As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public UserComment As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public FullName As String

            Public Priv As UserPrivilege

            Public AuthFlags As AuthFlags

            Public PasswordAge As FriendlySeconds

            <MarshalAs(UnmanagedType.LPWStr)>
            Public HomeDir As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Parms As String

            Public LastLogon As FriendlyUnixTime

            Public LastLogout As FriendlyUnixTime

            Public BadPwCount As Integer

            Public NumLogons As Integer

            <MarshalAs(UnmanagedType.LPWStr)>
            Public LogonServer As String

            Public CountryCode As Integer

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Workstations As String

            Public MaxStorage As Integer

            Public UnitsPerWeek As Integer

            Public LogonHours As MemPtr

            Public CodePage As Integer

            Public Overrides Function ToString() As String
                Return FullName
            End Function

        End Structure

        ''' <summary>
        ''' Windows 8 Identity Structure for Windows Logon
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure UserInfo24
            <MarshalAs(UnmanagedType.Bool)>
            Public InternetIdentity As Boolean

            Public Flags As Integer

            <MarshalAs(UnmanagedType.LPWStr)>
            Public InternetProviderName As String

            <MarshalAs(UnmanagedType.LPWStr)>
            Public InternetPrincipalName As String

            Public UserSid As IntPtr

            Public Overrides Function ToString() As String
                Return InternetPrincipalName
            End Function
        End Structure

        <Flags>
        Public Enum NET_API_STATUS As UInteger
            NERR_Success = 0
            NERR_InvalidComputer = 2351
            NERR_NotPrimary = 2226
            NERR_SpeGroupOp = 2234
            NERR_LastAdmin = 2452
            NERR_BadPassword = 2203
            NERR_PasswordTooShort = 2245
            NERR_UserNotFound = 2221
            ERROR_ACCESS_DENIED = 5
            ERROR_NOT_ENOUGH_MEMORY = 8
            ERROR_INVALID_PARAMETER = 87
            ERROR_INVALID_NAME = 123
            ERROR_INVALID_LEVEL = 124
            ERROR_SESSION_CREDENTIAL_CONFLICT = 1219
        End Enum

        Public Declare Function NetUserGetInfo Lib "netapi32.dll" _
        (<MarshalAs(UnmanagedType.LPWStr)> servername As String,
         <MarshalAs(UnmanagedType.LPWStr)> username As String,
         level As Integer,
         ByRef bufptr As MemPtr) As NET_API_STATUS


        Declare Unicode Function NetServerEnum Lib "netapi32.dll" _
        Alias "NetServerEnum" (<MarshalAs(UnmanagedType.LPWStr)> servername As String,
                              level As Integer,
                              ByRef bufptr As MemPtr,
                              prefmaxlen As Integer,
                              ByRef entriesread As Integer,
                              ByRef totalentries As Integer,
                              serverType As ServerTypes,
                              <MarshalAs(UnmanagedType.LPWStr)> domain As String,
                              ByRef resume_handle As IntPtr) As NET_API_STATUS

        Declare Unicode Function NetServerGetInfo Lib "netapi32.dll" _
        Alias "NetServerGetInfo" (<MarshalAs(UnmanagedType.LPWStr)> servername As String,
                              level As Integer,
                              ByRef bufptr As MemPtr) As NET_API_STATUS



        Declare Unicode Function NetLocalGroupEnum Lib "netapi32.dll" _
        Alias "NetLocalGroupEnum" (<MarshalAs(UnmanagedType.LPWStr)> servername As String,
                              level As Integer,
                              ByRef bufptr As MemPtr,
                              prefmaxlen As Integer,
                              ByRef entriesread As Integer,
                              ByRef totalentries As Integer,
                              ByRef resume_handle As IntPtr) As NET_API_STATUS


        Declare Unicode Function NetGroupEnum Lib "netapi32.dll" _
        Alias "NetGroupEnum" (<MarshalAs(UnmanagedType.LPWStr)> servername As String,
                              level As Integer,
                              ByRef bufptr As MemPtr,
                              prefmaxlen As Integer,
                              ByRef entriesread As Integer,
                              ByRef totalentries As Integer,
                              ByRef resume_handle As IntPtr) As NET_API_STATUS

        Declare Unicode Function NetUserEnum Lib "netapi32.dll" _
        Alias "NetUserEnum" (<MarshalAs(UnmanagedType.LPWStr)> servername As String,
                              level As Integer,
                              filter As Integer,
                              ByRef bufptr As MemPtr,
                              prefmaxlen As Integer,
                              ByRef entriesread As Integer,
                              ByRef totalentries As Integer,
                              ByRef resume_handle As IntPtr) As NET_API_STATUS

        Declare Function NetApiBufferFree Lib "netapi32.dll" (buffer As IntPtr) As NET_API_STATUS
        Declare Function NetApiBufferAllocate Lib "netapi32.dll" (bytecount As Integer, ByRef buffer As IntPtr) As NET_API_STATUS

        Declare Unicode Function NetGetJoinInformation Lib "Netapi32.dll" (
                                ByVal lpServer As String,
                                ByRef lpNameBuffer As IntPtr,
                                ByRef bufferType As NetworkJoinStatus) _
                                As NET_API_STATUS

        Declare Function NetGroupGetUsers Lib "netapi32.dll" (<MarshalAs(UnmanagedType.LPWStr)> servername As String,
                                        <MarshalAs(UnmanagedType.LPWStr)> groupname As String,
                                        level As Integer,
                                        ByRef bufptr As MemPtr,
                                        prefmaxlen As Integer,
                                        ByRef entriesread As Integer,
                                        ByRef totalentries As Integer,
                                        ByRef ResumeHandle As IntPtr) As NET_API_STATUS

        Declare Function NetLocalGroupGetMembers Lib "netapi32.dll" (<MarshalAs(UnmanagedType.LPWStr)> servername As String,
                                        <MarshalAs(UnmanagedType.LPWStr)> groupname As String,
                                        level As Integer,
                                        ByRef bufptr As MemPtr,
                                        prefmaxlen As Integer,
                                        ByRef entriesread As Integer,
                                        ByRef totalentries As Integer,
                                        ByRef ResumeHandle As IntPtr) As NET_API_STATUS



        <StructLayout(LayoutKind.Sequential)>
        Public Structure UserGroup0
            <MarshalAs(UnmanagedType.LPWStr)>
            Public Name As String
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Public Structure UserLocalGroup1
            Public Sid As IntPtr
            Public SidUsage As SidUsage

            <MarshalAs(UnmanagedType.LPWStr)>
            Public Name As String
        End Structure

        ''' <summary>
        ''' Enumerates all computers visible to the specified computer on the specified domain.
        ''' </summary>
        ''' <param name="computer">Optional computer name.  The local computer is assumed if this parameter is null.</param>
        ''' <param name="domain">Optional domain name.  The primary domain of the specified computer is assumed if this parameter is null.</param>
        ''' <returns>An array of ServerInfo1 objects.</returns>
        ''' <remarks></remarks>
        Public Function EnumServers(Optional computer As String = Nothing, Optional domain As String = Nothing) As ServerInfo101()

            Dim adv As MemPtr
            Dim mm As MemPtr = MemPtr.Empty
            Dim en As Integer = 0
            Dim ten As Integer = 0
            Dim svr() As ServerInfo101 = Nothing
            Dim i As Integer = 0
            Dim c As Integer

            NetServerEnum(Nothing, 101, mm, -1, en, ten, ServerTypes.WindowsNT, Nothing, IntPtr.Zero)
            adv = mm

            c = ten - 1
            ReDim svr(c)
            For i = 0 To c
                svr(i) = Marshal.PtrToStructure(adv.Handle, GetType(ServerInfo101))
                adv += Marshal.SizeOf(svr(i))
            Next

            mm.NetFree()
            EnumServers = svr

        End Function

        ''' <summary>
        ''' Gets network information for the specified computer.
        ''' </summary>
        ''' <param name="computer">Computer for which to retrieve the information.</param>
        ''' <param name="info">A ServerInfo101 structure that receives the information.</param>
        ''' <remarks></remarks>
        Public Sub GetServerInfo(computer As String, ByRef info As ServerInfo101)
            Dim mm As MemPtr = MemPtr.Empty

            NetServerGetInfo(computer, 101, mm)
            info = Marshal.PtrToStructure(mm.Handle, GetType(ServerInfo101))
            mm.NetFree()
        End Sub

        ''' <summary>
        ''' Enumerates the groups on a system.
        ''' </summary>
        ''' <param name="computer">The computer to enumerate.</param>
        ''' <returns>An array of GroupInfo2 structures.</returns>
        ''' <remarks></remarks>
        Public Function EnumGroups(computer As String) As GroupInfo2()

            Dim mm As MemPtr = MemPtr.Empty
            Dim en As Integer = 0
            Dim ten As Integer = 0
            Dim adv As MemPtr
            Dim grp() As GroupInfo2

            NetGroupEnum(computer, 2, mm, -1, en, ten, IntPtr.Zero)
            adv = mm

            Dim i As Integer
            Dim c As Integer = ten - 1

            ReDim grp(c)

            For i = 0 To c
                grp(i) = Marshal.PtrToStructure(adv.Handle, GetType(GroupInfo2))
                adv += Marshal.SizeOf(grp(i))
            Next
            EnumGroups = grp
            mm.NetFree()
        End Function

        ''' <summary>
        ''' Enumerates local groups for the specified computer.
        ''' </summary>
        ''' <param name="computer">The computer to enumerate.</param>
        ''' <returns>An array of LocalGroupInfo1 structures.</returns>
        ''' <remarks></remarks>
        Public Function EnumLocalGroups(computer As String) As LocalGroupInfo1()
            Dim mm As MemPtr = MemPtr.Empty
            Dim en As Integer = 0
            Dim ten As Integer = 0
            Dim adv As MemPtr
            Dim grp() As LocalGroupInfo1

            NetLocalGroupEnum(computer, 1, mm, -1, en, ten, IntPtr.Zero)
            adv = mm

            Dim i As Integer
            Dim c As Integer = ten - 1

            ReDim grp(c)

            For i = 0 To c
                grp(i) = Marshal.PtrToStructure(adv.Handle, GetType(LocalGroupInfo1))
                adv += Marshal.SizeOf(grp(i))
            Next
            EnumLocalGroups = grp
            mm.NetFree()

        End Function

        ''' <summary>
        ''' Enumerates users of a given group.
        ''' </summary>
        ''' <param name="computer">Computer for which to retrieve the information.</param>
        ''' <param name="group">Group to enumerate.</param>
        ''' <returns>An array of user names.</returns>
        ''' <remarks></remarks>
        Public Function GroupUsers(computer As String, group As String) As String()
            Dim mm As New MemPtr
            Dim op As MemPtr

            Dim cbt As Integer = 0
            Dim cb As Integer = 0
            Dim s() As String = Nothing
            Try
                If NetGroupGetUsers(computer, group, 0, mm, -1, cb, cbt, IntPtr.Zero) = NET_API_STATUS.NERR_Success Then
                    op = mm
                    Dim z As UserGroup0

                    Dim i As Integer
                    ReDim s(cb - 1)
                    For i = 0 To cb - 1
                        z = Marshal.PtrToStructure(mm.Handle, GetType(UserGroup0))
                        s(i) = z.Name
                        mm += Marshal.SizeOf(z)
                    Next
                End If

            Catch ex As Exception
                Throw New NativeException
            End Try

            op.NetFree()
            GroupUsers = s

        End Function

        ''' <summary>
        ''' Gets the members of the specified local group on the specified machine.
        ''' </summary>
        ''' <param name="computer">The computer for which to retrieve the information.</param>
        ''' <param name="group">The name of the group to enumerate.</param>
        ''' <param name="SidType">The type of group members to return.</param>
        ''' <returns>A list of group member names.</returns>
        ''' <remarks></remarks>
        Public Function LocalGroupUsers(computer As String, group As String, Optional SidType As SidUsage = SidUsage.SidTypeUser) As String()
            Dim mm As New MemPtr
            Dim op As MemPtr
            Dim x As Integer = 0
            Dim cbt As Integer = 0
            Dim cb As Integer = 0
            Dim s() As String = Nothing
            Try
                If NetLocalGroupGetMembers(computer, group, 1, mm, -1, cb, cbt, IntPtr.Zero) = NET_API_STATUS.NERR_Success Then
                    If cb = 0 Then
                        mm.NetFree()
                        Return Nothing
                    End If

                    op = mm
                    Dim z As UserLocalGroup1

                    Dim i As Integer
                    ReDim s(cb - 1)
                    For i = 0 To cb - 1
                        z = Marshal.PtrToStructure(mm.Handle, GetType(UserLocalGroup1))
                        If z.SidUsage = SidType Then
                            s(x) = z.Name
                            mm += Marshal.SizeOf(z)
                            x += 1
                        End If
                    Next
                    ReDim Preserve s(x - 1)
                End If

            Catch ex As Exception
                Throw New NativeException
            End Try

            op.NetFree()
            LocalGroupUsers = s

        End Function

        ''' <summary>
        ''' Grab the join status for the specified computer.
        ''' </summary>
        ''' <param name="joinStatus">Receives the NetworkJoinStatus value.</param>
        ''' <param name="Computer">Optional name of a computer for which to retrieve the NetworkJoinStatus information.  If this parameter is null, the local computer is assumed.</param>
        ''' <returns>The name of the current domain or workgroup for the specified computer.</returns>
        ''' <remarks></remarks>
        Public Function GrabJoin(ByRef joinStatus As NetworkJoinStatus, Optional ByVal Computer As String = Nothing) As String
            Try
                Dim mm As MemPtr
                mm.Alloc(1024)

                NetGetJoinInformation(Computer, mm, joinStatus)

                GrabJoin = mm
                mm.NetFree()
            Catch ex As Exception
                Throw New NativeException
            End Try

        End Function

        ''' <summary>
        ''' Enumerate users into a USER_INFO_11 structure.
        ''' </summary>
        ''' <param name="machine">Computer on which to perform the enumeration.  If this parameter is null, the local machine is assumed.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function EnumUsers11(Optional machine As String = Nothing) As UserInfo11()

            Try

                Dim cb As Integer = 0
                Dim er As Integer = 0

                Dim rh As MemPtr = 0
                Dim te As Integer = 0

                Dim buff As MemPtr

                Dim usas() As UserInfo11
                Try
                    cb = Marshal.SizeOf(New UserInfo11)

                Catch ex As Exception
                    MsgBox(ex.Message & vbCrLf & vbCrLf & "Stack Trace: " & ex.StackTrace, MsgBoxStyle.Exclamation)
                End Try

                Dim err = NetUserEnum(machine, 11, 0, buff, -1, er, te, IntPtr.Zero)
                rh = buff
                ReDim usas(te - 1)

                For i = 0 To te - 1
                    usas(i) = Marshal.PtrToStructure(buff.Handle, GetType(UserInfo11))
                    usas(i).LogonHours = IntPtr.Zero
                    buff += cb
                Next

                EnumUsers11 = usas

                rh.NetFree()

            Catch ex As Exception
                Throw New NativeException
            End Try

        End Function

        ''' <summary>
        ''' For Windows 8, retrieves the user's Microsoft login account information.
        ''' </summary>
        ''' <param name="machine">Computer on which to perform the enumeration.  If this parameter is null, the local machine is assumed.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function EnumUsers24(Optional machine As String = Nothing) As UserInfo24()
            Try

                Dim rh As MemPtr = 0
                Dim i As Integer = 0
                Dim uorig() As UserInfo11 = EnumUsers11()
                Dim usas() As UserInfo24
                Dim c As Integer = uorig.Length - 1

                ReDim usas(c)

                For i = 0 To c
                    NetUserGetInfo(machine, uorig(i).Name, 24, rh)
                    usas(i) = Marshal.PtrToStructure(rh.Handle, GetType(UserInfo24))
                    rh.NetFree()
                Next

                EnumUsers24 = usas

            Catch ex As Exception
                Throw New NativeException
            End Try

        End Function

    End Module

End Namespace