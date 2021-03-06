﻿
'' ************************************************* ''
'' DataTools Visual Basic Utility Library - Interop
''
'' Module: WinHooks
''         Windows Hook Functions
''
'' Some enum documentation copied from the MSDN (and in some cases, updated).
'' Some classes and interfaces were ported from the WindowsAPICodePack.
'' 
'' Copyright (C) 2011-2020 Nathan Moschkin
'' All Rights Reserved
''
'' Licensed Under the Microsoft Public License   
'' ************************************************* ''

Option Explicit On

Imports System
Imports System.IO
Imports System.Collections.Specialized
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Drawing
Imports System.Reflection
Imports System.Text
Imports System.Numerics
Imports System.Threading
Imports Microsoft.Win32

Imports DataTools.Memory
Imports System.Linq.Expressions
Imports DataTools.Interop.My.Resources

Namespace Native

    Public Enum WindowHookTypes

        ''' <summary>
        ''' Installs a hook procedure that monitors messages before the system sends them To the destination window procedure. For more information, see the CallWndProc hook procedure.
        ''' </summary>
        CallWndProc = 4

        ''' <summary>
        ''' Installs a hook procedure that monitors messages after they have been processed by the destination window procedure. For more information, see the CallWndRetProc hook procedure.
        ''' </summary>
        CallWndProcReturning = 12

        ''' <summary>
        ''' Installs a hook procedure that receives notifications useful To a CBT application. For more information, see the CBTProc hook procedure.
        ''' </summary>
        Cbt = 5

        ''' <summary>
        ''' Installs a hook procedure useful For debugging other hook procedures. For more information, see the DebugProc hook procedure.
        ''' </summary>
        Debug = 9

        ''' <summary>
        ''' Installs a hook procedure that will be called When the application's foreground thread is about to become idle. This hook is useful for performing low priority tasks during idle time. For more information, see the ForegroundIdleProc hook procedure.
        ''' </summary>
        ForegroundIdle = 11

        ''' <summary>
        ''' Installs a hook procedure that monitors messages posted To a message queue. For more information, see the GetMsgProc hook procedure.
        ''' </summary>
        GetMessage = 3

        ''' <summary>
        ''' Installs a hook procedure that posts messages previously recorded by a WH_JOURNALRECORD hook procedure. For more information, see the JournalPlaybackProc hook procedure.
        ''' </summary>
        JournalPlayback = 1

        ''' <summary>
        ''' Installs a hook procedure that records input messages posted To the system message queue. This hook Is useful For recording macros. For more information, see the JournalRecordProc hook procedure.
        ''' </summary>
        JournalRecord = 0

        ''' <summary>
        ''' Installs a hook procedure that monitors keystroke messages. For more information, see the KeyboardProc hook procedure.
        ''' </summary>
        Keyboard = 2

        ''' <summary>
        ''' Installs a hook procedure that monitors low-level keyboard input events. For more information, see the LowLevelKeyboardProc hook procedure.
        ''' </summary>
        KeyboardLowLevel = 13

        ''' <summary>
        ''' Installs a hook procedure that monitors mouse messages. For more information, see the MouseProc hook procedure.
        ''' </summary>
        Mouse = 7

        ''' <summary>
        ''' Installs a hook procedure that monitors low-level mouse input events. For more information, see the LowLevelMouseProc hook procedure.
        ''' </summary>
        MouseLowLevel = 14

        ''' <summary>
        ''' Installs a hook procedure that monitors messages generated As a result Of an input Event In a dialog box, message box, menu, Or scroll bar. For more information, see the MessageProc hook procedure.
        ''' </summary>
        MessageFilter = -1

        ''' <summary>
        ''' Installs a hook procedure that receives notifications useful To shell applications. For more information, see the ShellProc hook procedure.
        ''' </summary>
        Shell = 10

        ''' <summary>
        ''' Installs a hook procedure that monitors messages generated As a result Of an input Event In a dialog box, message box, menu, Or scroll bar. The hook procedure monitors these messages For all applications In the same desktop As the calling thread. For more information, see the SysMsgProc hook procedure.
        ''' </summary>
        SystemMessageFilter = 6

    End Enum

    Public Enum ShellHookCodes

        ''' <summary>
        ''' The accessibility state has changed.
        ''' </summary>
        AccessibilityState = 11

        ''' <summary>
        ''' The shell should activate its main window.
        ''' </summary>
        ActivateShellWindow = 3

        ''' <summary>
        ''' The user completed an input event (for example, pressed an application command button on the mouse or an application command key on the keyboard), and the application did not handle the WM_APPCOMMAND message generated by that input.
        ''' If the Shell procedure Handles the WM_COMMAND message, it should Not Call CallNextHookEx. See the Return Value section For more information.
        ''' </summary>
        AppCommand = 12

        ''' <summary>
        ''' A window is being minimized or maximized. The system needs the coordinates of the minimized rectangle for the window.
        ''' </summary>
        GetMinRect = 5

        ''' <summary>
        ''' Keyboard language was changed or a new keyboard layout was loaded.
        ''' </summary>
        Language = 8

        ''' <summary>
        ''' The title of a window in the task bar has been redrawn.
        ''' </summary>
        Redraw = 6

        ''' <summary>
        ''' The user has selected the task list. A shell application that provides a task list should return TRUE to prevent Windows from starting its task list.
        ''' </summary>
        TaskMan = 7

        ''' <summary>
        ''' The activation has changed to a different top-level, unowned window.
        ''' </summary>
        WindowActivated = 4

        ''' <summary>
        ''' A top-level, unowned window has been created. The window exists when the system calls this hook.
        ''' </summary>
        WindowCreated = 1

        ''' <summary>
        ''' A top-level, unowned window is about to be destroyed. The window still exists when the system calls this hook.
        ''' </summary>
        WindowDestroyed = 2

        ''' <summary>
        ''' A top-level window is being replaced. The window exists when the system calls this hook.
        ''' </summary>
        WindowReplaced = 13
    End Enum

    Public Delegate Function WinHookCallback(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr

    Public Delegate Function ShellCallback(nCode As ShellHookCodes, wParam As IntPtr, lParam As IntPtr) As IntPtr

    Friend Module WinHooks

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function SetWindowsHookEx(idHook As WindowHookTypes, lpfn As WinHookCallback, hmod As IntPtr, dwThreadId As Integer) As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function SetWindowsHookEx(idHook As WindowHookTypes, lpfn As ShellCallback, hmod As IntPtr, dwThreadId As Integer) As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function CallNextHookEx(hHook As IntPtr, nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Function UnhookWindowsHookEx(hHook As IntPtr) As Boolean
        End Function



        '      HHOOK SetWindowsHookExA(
        'Int       idHook,
        'HOOKPROC  lpfn,
        'HINSTANCE hmod,
        'DWORD     dwThreadId
        ');

    End Module

End Namespace
