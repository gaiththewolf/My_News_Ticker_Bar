Option Strict On

Imports Microsoft.VisualBasic
Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Namespace AppBars

    Public Class AppBar

        Inherits Form
        Private Const WS_CAPTION As Integer = 12582912
        Private Const WS_BORDER As Integer = 8388608
        Private components As Container = Nothing
        Private _autoRegisterOnLoad As Boolean
        Private _isAppBarRegistered As Boolean
        Private _appBarDock As AppBarDockStyle
        Private _idealSize As Size
        Private appBarCallback As Integer

        <BrowsableAttribute(False),
        CategoryAttribute("AppBar Behavior"),
        DefaultValueAttribute(AppBarDockStyle.ScreenTop),
        DescriptionAttribute("Designates which side of the screen the AppBar will dock to.")>
        Public Property AppBarDock() As AppBarDockStyle
            Get
                Return _appBarDock
            End Get
            Set(ByVal value As AppBarDockStyle)
                _appBarDock = value
                UpdateDockedAppBarPosition(AppBarDock)
            End Set
        End Property
        <DescriptionAttribute("Registers the AppBar as it loads."),
        DefaultValueAttribute(False),
        CategoryAttribute("AppBar Behavior")>
        Public Property AutoRegisterOnLoad() As Boolean
            Get
                Return _autoRegisterOnLoad
            End Get
            Set(ByVal value As Boolean)
                _autoRegisterOnLoad = value
            End Set
        End Property
        Protected Overrides ReadOnly Property CreateParams() As CreateParams
            Get
                Dim cp As CreateParams = MyBase.CreateParams
                cp.Style = cp.Style And -12582913
                cp.Style = cp.Style And -8388609
                Return cp
            End Get
        End Property
        <CategoryAttribute("AppBar Behavior"),
        DescriptionAttribute("The ideal size of the AppBar. If docked top or bottom, width is ignored. If docked left or right, height is  ignored.")>
        Public Property IdealSize() As Size
            Get
                Return _idealSize
            End Get
            Set(ByVal value As Size)
                _idealSize = value
            End Set
        End Property
        Public Sub New()
            InitializeComponent()
            appBarCallback = NativeMethods.RegisterWindowMessage("Windows Forms AppBar")
            _isAppBarRegistered = False
            AppBarDock = AppBarDockStyle.None
            IdealSize = New Size(100, 100)
            MyBase.FormBorderStyle = FormBorderStyle.FixedToolWindow
        End Sub
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso Not components Is Nothing Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub
        Public Function IsAppBarRegistered() As Boolean
            Return _isAppBarRegistered
        End Function
        Public Sub RefreshPosition()
            UpdateDockedAppBarPosition(AppBarDock)
        End Sub
        Public Function RegisterAppBar() As Boolean
            Dim retVal As Boolean = NativeMethods.RegisterAppBar(MyBase.Handle, appBarCallback)
            _isAppBarRegistered = retVal
            Return retVal
        End Function
        Public Sub UnregisterAppBar()
            NativeMethods.UnregisterAppBar(MyBase.Handle)
            _isAppBarRegistered = False
        End Sub
        Private Sub UpdateDockedAppBarPosition(ByVal dockStyle As AppBarDockStyle)
            Dim edge As Integer = 0
            Select Case dockStyle
                Case AppBarDockStyle.ScreenLeft
                    edge = 0
                Case AppBarDockStyle.ScreenTop
                    edge = 1
                Case AppBarDockStyle.ScreenRight
                    edge = 2
                Case AppBarDockStyle.None
                    Return
                Case Else
                    edge = 3
            End Select
            NativeMethods.DockAppBar(MyBase.Handle, edge, IdealSize)
        End Sub
        Protected Overrides Sub WndProc(ByRef m As Message)
            If m.Msg = appBarCallback Then
                Select Case m.WParam.ToInt32
                    Case NativeMethods.ABN_FULLSCREENAPP
                    Case NativeMethods.ABN_POSCHANGED
                    Case NativeMethods.ABN_STATECHANGE
                    Case NativeMethods.ABN_WINDOWARRANGE
                End Select
            End If
            MyBase.WndProc(m)
        End Sub
        Protected Overrides Sub OnClosing(ByVal e As CancelEventArgs)
            UnregisterAppBar()
            MyBase.OnClosing(e)
        End Sub
        Protected Overrides Sub OnLoad(ByVal e As EventArgs)
            MyBase.OnLoad(e)
            If AutoRegisterOnLoad Then
                RegisterAppBar()
            End If
        End Sub
        Private Sub InitializeComponent()
            Me.SuspendLayout()
            '
            'AppBar
            '
            Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
            Me.ClientSize = New System.Drawing.Size(245, 70)
            Me.Name = "AppBar"
            Me.Text = "AppBar"
            Me.ResumeLayout(False)

        End Sub

        Private Sub AppBar_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        End Sub
    End Class

End Namespace