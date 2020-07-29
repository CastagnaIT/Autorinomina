Imports System.Runtime.InteropServices
Imports System.Windows.Interop

Module CustomizeWindow
    'Potete utilizzare e distribuire come preferite questo codice, ma non rimuovere quest'intestazione.

    'Tutti i diritti riservati.

    'CustomizeWindow è un progetto Open Source sotto licenza GNU General Public License (GPL) v2
    'http://www.opensource.org/licenses/gpl-license.php

    'CustomizeWindow [Versione 1.2 - 23/06/2011]
    'Autore Gottardo Stefano
    'Web: http://www.autorinomina.it/

    'METODO DI FUNZIONAMENTO
    'Copiare e incollare il seguente codice
    'Protected Overrides Sub OnSourceInitialized(e As EventArgs)
    '    CustomWindow(Me, True)
    'End Sub

#Region "Costanti"
    Private Const GWL_STYLE As Int32 = -16 'Imposta un nuovo stile alla window
    Private Const GWL_EXSTYLE As Int32 = -20 'Stile esteso

    Private Const WS_MAXIMIZEBOX As Int32 = &H10000
    Private Const WS_MINIMIZEBOX As Int32 = &H20000
    Private Const WS_SYSMENU As Int32 = &H80000
    Private Const WS_EX_DLGMODALFRAME As Integer = &H1

    Private Const SWP_NOSIZE As Integer = &H1
    Private Const SWP_NOMOVE As Integer = &H2
    Private Const SWP_NOZORDER As Integer = &H4
    Private Const SWP_FRAMECHANGED As Integer = &H20

    Private Const WM_SETICON As UInteger = &H80

    Private Const ICON_SMALL As Integer = 0
    Private Const ICON_BIG As Integer = 1

    Private Const SC_CLOSE As UInteger = &HF060

    Private Const MF_BYCOMMAND As UInteger = &H0
    Private Const MF_GRAYED As UInteger = &H1
#End Region

#Region "API"
    <DllImport("user32.dll")> _
    Private Function GetWindowLong(ByVal hwnd As IntPtr, ByVal index As Integer) As Integer
    End Function

    <DllImport("user32.dll")> _
    Private Function SetWindowLong(ByVal hwnd As IntPtr, ByVal index As Integer, ByVal newStyle As Integer) As Integer
    End Function

    <DllImport("user32.dll")> _
    Private Function SetWindowPos(ByVal hwnd As IntPtr, ByVal hwndInsertAfter As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal flags As UInteger) As Boolean
    End Function

    <DllImport("User32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Function SendMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Public Function GetSystemMenu(ByVal hwnd As IntPtr, ByVal revert As Boolean) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Public Function DeleteMenu(ByVal hMenu As IntPtr, ByVal position As UInteger, ByVal flags As UInteger) As Boolean
    End Function
#End Region
  

    Public Sub CustomWindow(ByVal window As Window, Optional DisabilitaIcona As Boolean = False, Optional DisabilitaMinimize As Boolean = False, Optional DisabilitaMaximize As Boolean = False, Optional DisabilitaClose As Boolean = False, Optional NascondiClose As Boolean = False)
        Dim hwnd As IntPtr = New Interop.WindowInteropHelper(window).Handle

        If DisabilitaIcona Then
            Dim extendedStyle As Integer = GetWindowLong(hwnd, GWL_EXSTYLE)

            SetWindowLong(hwnd, GWL_EXSTYLE, extendedStyle Or WS_EX_DLGMODALFRAME)

            SendMessage(hwnd, WM_SETICON, ICON_SMALL, IntPtr.Zero)
            SendMessage(hwnd, WM_SETICON, ICON_BIG, IntPtr.Zero)

            SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
        End If

        If DisabilitaClose Then
            Dim NewHwnd As IntPtr = GetSystemMenu(hwnd, False)
            DeleteMenu(NewHwnd, SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED)
        End If

        If DisabilitaMinimize OrElse DisabilitaMaximize OrElse NascondiClose Then
            Dim windowStyle As Int32 = GetWindowLong(hwnd, GWL_STYLE)

            Dim newStyle As Integer = windowStyle
            If DisabilitaMinimize Then newStyle = newStyle And Not WS_MINIMIZEBOX
            If DisabilitaMaximize Then newStyle = newStyle And Not WS_MAXIMIZEBOX
            If NascondiClose Then newStyle = newStyle And Not WS_SYSMENU

            SetWindowLong(hwnd, GWL_STYLE, newStyle)
        End If

    End Sub


End Module

