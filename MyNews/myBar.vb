Option Strict On

Imports MyNews.AppBars

Namespace myBar
    Public Class myBar
        Inherits AppBar
        Friend WithEvents PicBar As PictureBox

        Private Sub InitializeComponent()
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(myBar))
            Me.PicBar = New System.Windows.Forms.PictureBox()
            Me.NTicker1 = New MyNews.NTicker()
            CType(Me.PicBar, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'PicBar
            '
            Me.PicBar.Dock = System.Windows.Forms.DockStyle.Left
            Me.PicBar.Image = CType(resources.GetObject("PicBar.Image"), System.Drawing.Image)
            Me.PicBar.Location = New System.Drawing.Point(0, 0)
            Me.PicBar.Name = "PicBar"
            Me.PicBar.Size = New System.Drawing.Size(60, 60)
            Me.PicBar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
            Me.PicBar.TabIndex = 1
            Me.PicBar.TabStop = False
            '
            'NTicker1
            '
            Me.NTicker1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.NTicker1.IMG = CType(resources.GetObject("NTicker1.IMG"), System.Drawing.Image)
            Me.NTicker1.Location = New System.Drawing.Point(60, 0)
            Me.NTicker1.Name = "NTicker1"
            Me.NTicker1.News = "ssss" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "55454" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "WOLF"
            Me.NTicker1.RTL = True
            Me.NTicker1.Size = New System.Drawing.Size(391, 60)
            Me.NTicker1.Speed = 1.0R
            Me.NTicker1.TabIndex = 2
            Me.NTicker1.TextPosition = MyNews.NTicker.TextPlacement.Middle
            Me.NTicker1.TxTFont = Nothing
            '
            'myBar
            '
            Me.AutoRegisterOnLoad = True
            Me.ClientSize = New System.Drawing.Size(451, 60)
            Me.ControlBox = False
            Me.Controls.Add(Me.NTicker1)
            Me.Controls.Add(Me.PicBar)
            Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
            Me.IdealSize = New System.Drawing.Size(800, 60)
            Me.Name = "myBar"
            Me.ShowIcon = False
            Me.ShowInTaskbar = False
            Me.TopMost = True
            CType(Me.PicBar, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub

        Public Sub New()
            InitializeComponent()
        End Sub

        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso Not NTicker1 Is Nothing Then
                NTicker1.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        <STAThreadAttribute()>
        Public Shared Sub Main()
            Application.Run(New myBar)
        End Sub

        Public Pos As String = "Bottom"
        Public txtcol As Color = Color.Black
        Public pic As Image = My.Resources.Google_News_96px
        Public FormSize As Integer = 60
        Public FName As String = ""
        Public Fsize As Decimal = 0
        Friend WithEvents NTicker1 As NTicker
        Public teext As String = ""

        Private Sub myBar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            If Not RegisterAppBar() Then
                RegisterAppBar()
            End If
            MyBase.Size = New Size(800, FormSize)
            MyBase.IdealSize = New Size(800, FormSize)
            Select Case Pos
                Case "Bottom"
                    MyBase.AppBarDock = CType(4, AppBarDockStyle)
                Case "Top"
                    MyBase.AppBarDock = CType(2, AppBarDockStyle)
                Case Else
                    MyBase.AppBarDock = CType(1, AppBarDockStyle)
            End Select

            'NTicker1.News = teext
            'PicBar.Image = pic
            NTicker1.TxTFont = New Font(FName, Fsize)
            NTicker1.ForeColor = txtcol
        End Sub

    End Class
End Namespace
