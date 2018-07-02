Option Strict On

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class NativeMethods

    <StructLayoutAttribute(LayoutKind.Sequential)>
    Private Structure APP_BAR_DATA
        Public cbSize As Integer
        Public hWnd As IntPtr
        Public uCallbackMessage As Integer
        Public uEdge As Integer
        Public rc As NativeMethods.RECT
        Public lParam As IntPtr
    End Structure
    <StructLayoutAttribute(LayoutKind.Sequential)>
    Private Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer

        Public Sub New(left As Integer, top As Integer, right As Integer, bottom As Integer)
            Me.left = left
            Me.top = top
            Me.right = right
            Me.bottom = bottom
        End Sub
        Public Shared Function FromXYWH(x As Integer, y As Integer, width As Integer, height As Integer) As RECT
            Return New RECT(x, y, x + width, y + height)
        End Function
    End Structure
    Private Const ABM_NEW As Integer = 0
    Private Const ABM_REMOVE As Integer = 1
    Private Const ABM_QUERYPOS As Integer = 2
    Private Const ABM_SETPOS As Integer = 3
    Private Const ABM_SETAUTOHIDEBAR As Integer = 8
    Private Const ABM_SETSTATE As Integer = 10
    Private Const ABE_LEFT As Integer = 0
    Private Const ABE_TOP As Integer = 1
    Private Const ABE_RIGHT As Integer = 2
    Private Const ABE_BOTTOM As Integer = 3
    Private Const ABS_AUTOHIDE As Integer = 1
    Private Const ABS_ALWAYSONTOP As Integer = 2
    Friend Const ABN_STATECHANGE As Integer = 0
    Friend Const ABN_POSCHANGED As Integer = 1
    Friend Const ABN_FULLSCREENAPP As Integer = 2
    Friend Const ABN_WINDOWARRANGE As Integer = 3

    Friend Shared Sub DockAppBar(hWnd As IntPtr, edge As Integer, idealSize As Size)
        Dim abd As APP_BAR_DATA = New APP_BAR_DATA()
        abd.cbSize = Marshal.SizeOf(abd)
        abd.hWnd = hWnd
        abd.uEdge = edge
        If edge = 0 OrElse edge = 2 Then
            abd.rc.top = 0
            abd.rc.bottom = SystemInformation.PrimaryMonitorSize.Height
            If edge = 0 Then
                abd.rc.right = idealSize.Width
            Else
                abd.rc.right = SystemInformation.PrimaryMonitorSize.Width
                abd.rc.left = abd.rc.right - idealSize.Width
            End If
        Else
            abd.rc.left = 0
            abd.rc.right = SystemInformation.PrimaryMonitorSize.Width
            If edge = 1 Then
                abd.rc.bottom = idealSize.Height
            Else
                abd.rc.bottom = SystemInformation.PrimaryMonitorSize.Height
                abd.rc.top = abd.rc.bottom - idealSize.Height
            End If
        End If
        SHAppBarMessage(2, abd)
        Select Case edge
            Case 0
                abd.rc.right = abd.rc.left + idealSize.Width
            Case 2
                abd.rc.left = abd.rc.right - idealSize.Width
            Case 1
                abd.rc.bottom = abd.rc.top + idealSize.Height
            Case 3
                abd.rc.top = abd.rc.bottom - idealSize.Height
        End Select
        SHAppBarMessage(3, abd)
        MoveWindow(abd.hWnd, abd.rc.left, abd.rc.top, abd.rc.right - abd.rc.left, abd.rc.bottom - abd.rc.top, True)
    End Sub
    Friend Shared Function RegisterAppBar(hWnd As IntPtr, uCallbackMessage As Integer) As Boolean
        Dim flag As Boolean
        Dim abd As APP_BAR_DATA = New APP_BAR_DATA()
        abd.cbSize = Marshal.SizeOf(abd)
        abd.hWnd = hWnd
        abd.uCallbackMessage = uCallbackMessage
        If SHAppBarMessage(0, abd) = 0 Then
            flag = False
        Else
            flag = True
        End If
        Return flag
    End Function
    Friend Shared Sub UnregisterAppBar(hWnd As IntPtr)
        Dim abd As APP_BAR_DATA = New APP_BAR_DATA()
        abd.cbSize = Marshal.SizeOf(abd)
        abd.hWnd = hWnd
        SHAppBarMessage(1, abd)
    End Sub
    <DllImportAttribute("User32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto)>
    Private Shared Function MoveWindow(hWnd As IntPtr, x As Integer, y As Integer, cx As Integer, cy As Integer, repaint As Boolean) As Boolean
    End Function
    <DllImportAttribute("User32.dll", CharSet:=CharSet.Auto)>
    Friend Shared Function RegisterWindowMessage(msg As String) As Integer
    End Function
    <DllImportAttribute("Shell32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SHAppBarMessage(dwMessage As Integer, ByRef abd As APP_BAR_DATA) As Integer
    End Function

End Class