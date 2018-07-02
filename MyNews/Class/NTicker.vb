
Public Class NTicker

#Region "Enumerations"
    Enum TextPlacement
        Top
        Middle
        Bottom
    End Enum
#End Region

#Region "Variables"
    Dim X As Single
    Dim Y As Single
    Dim Gr As Graphics
    Dim _textbar As String = "Null"
    Dim RB As RichTextBox = New RichTextBox
    Dim _strSize As SizeF
    Private _tpos As TextPlacement = TextPlacement.Middle
    Private _speed As Double = 1.0
    Private _rtl As Boolean = True
    Private _txtfont As Font
    Private _pic As Image = My.Resources.logo
#End Region

#Region "Events"

    Private Sub NTicker_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetupWindow()
    End Sub
    Private Sub NTicker_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        SetupWindow()
    End Sub
    Private Sub Timer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer.Tick
        UpdateDisplay()
    End Sub
    Private Sub lbl_DoubleClick(sender As Object, e As System.EventArgs) Handles Panel1.DoubleClick
        MyBase.OnDoubleClick(e)
    End Sub
    Private Sub NTicker_hover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.MouseHover, Panel1.MouseHover
        Timer.Stop()
    End Sub
    Private Sub NTicker_leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.MouseLeave, Panel1.MouseLeave
        Timer.Start()
    End Sub
#End Region

#Region "Methods"
    ''' <summary>
    ''' Set up the control for initial use
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetupWindow()
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        If RB Is Nothing Or RB.Text.Equals("") Then
            RB.Text = _textbar
        End If
        makepicbar(RB.Text)
        Panel1.Height = Me.Height - 4
        Panel1.AutoSize = True
        Panel1.Location = New Point(0, Me.Width)
        Me.Dock = DockStyle.Fill
        Me.BorderStyle = Windows.Forms.BorderStyle.None
        If _rtl Then
            X = -Panel1.Width
        Else
            X = Me.Width
        End If
        If (Me.Height < Panel1.Height + 4) Then Me.Height = CInt(Panel1.Height) + 4
        Timer.Interval = 25
        Timer.Start()
    End Sub

    ''' <summary>
    ''' Update the scrolling display by moving the label
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateDisplay()
        Dim tHeight As Double = Panel1.Height
        Select Case Me.TextPosition
            Case TextPlacement.Top : Y = 0
            Case TextPlacement.Middle : Y = CSng((Me.Height / 2) - (tHeight / 2))
            Case TextPlacement.Bottom : Y = CSng(Me.Height - tHeight)
        End Select
        Panel1.Location = New Point(CInt(X), CInt(Y))
        If _rtl Then
            If X >= (Me.Width) Then
                X = -Panel1.Width
            Else
                X = CSng(X + _speed)
            End If
        Else
            If X <= (0 - Panel1.Width) Then
                X = Me.Width
            Else
                X = CSng(X - _speed)
            End If
        End If
    End Sub

    Private Sub makepicbar(ByVal txt As String)
        Dim sstr As String() = Split(txt, "--")
        Dim i As Integer = 0
        Panel1.Controls.Clear()
        Try
            For Each s As String In sstr
                Dim lbl As Label = New Label
                lbl.Text = s
                lbl.Name = "lbl_" & i
                lbl.Dock = DockStyle.Left
                lbl.TextAlign = ContentAlignment.MiddleCenter
                lbl.Font = _txtfont
                lbl.BackColor = Color.Transparent
                'lbl.BorderStyle = BorderStyle.FixedSingle
                lbl.Width = TextRenderer.MeasureText(s, _txtfont).Width
                Panel1.Controls.Add(lbl)
                If i < sstr.Count - 1 Then
                    Dim Pic As PictureBox = New PictureBox
                    Pic.Name = "Pic_" & i
                    Pic.Image = _pic
                    Pic.SizeMode = PictureBoxSizeMode.Zoom
                    Pic.Dock = DockStyle.Left
                    Pic.BackColor = Color.Transparent
                    'Pic.BorderStyle = BorderStyle.FixedSingle
                    Panel1.Controls.Add(Pic)
                End If
                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Properties"

    ''' <summary>
    ''' Set the text to display in the ticker
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property News As String
        Get
            Return _textbar
        End Get
        Set(value As String)
            _textbar = value
            'Need to reset the control with the new text
            SetupWindow()
        End Set
    End Property

    ''' <summary>
    ''' Define the location of the text in the ticker
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TextPosition As TextPlacement
        Get
            Return _tpos
        End Get
        Set(value As TextPlacement)
            _tpos = value
            SetupWindow()
        End Set
    End Property

    ''' <summary>
    ''' Define the Sense of the text in the ticker
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RTL As Boolean
        Get
            Return _rtl
        End Get
        Set(value As Boolean)
            _rtl = value
            SetupWindow()
        End Set
    End Property

    Public Property RBTxt As RichTextBox
        Get
            Return RB
        End Get
        Set(value As RichTextBox)
            RB = value
            'Need to reset the control with the new text
            SetupWindow()
        End Set
    End Property

    Public Property TxTFont As Font
        Get
            Return _txtfont
        End Get
        Set(value As Font)
            _txtfont = value
            SetupWindow()
        End Set
    End Property

    Public Property IMG As Image
        Get
            Return _pic
        End Get
        Set(value As Image)
            _pic = value
            SetupWindow()
        End Set
    End Property

    ''' <summary>
    ''' Set the scrolling speed for the ticker
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Speed As Double
        Get
            Return _speed
        End Get
        Set(value As Double)
            If value < 0.01 Or value > 100 Then
                Throw New ArgumentOutOfRangeException("Speed", "Cannot be less than 0.01 or greater than 100")
            Else
                _speed = value
                SetupWindow()
            End If
        End Set
    End Property

#End Region

End Class
