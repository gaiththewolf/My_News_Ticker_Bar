Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word

Public Class FrmMenu

    Public FrmBar As myBar.myBar


    Public confile As String = System.Windows.Forms.Application.StartupPath & "\conf.ini"

    Private Sub BunifuImageButton1_Click(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
        Try
            System.Windows.Forms.Application.Exit()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub BunifuImageButton2_Click(sender As Object, e As EventArgs) Handles BunifuImageButton2.Click
        WindowState = FormWindowState.Minimized
    End Sub

    Public Sub LoadConfig(Optional ByVal OpenBar As Boolean = False)
        Try
            Dim Backolor As Integer = INIRead(confile, "Config", "Barcolor", "-1")
            selectedColor.BackColor = Color.FromArgb(CInt(Backolor))
            Dim Txtkolor As Integer = INIRead(confile, "Config", "Textcolor", "-16777216")
            selectedtextcolor.BackColor = Color.FromArgb(CInt(Txtkolor))
            Dim Pos As String = INIRead(confile, "Config", "Pos", "Bottom")
            If Pos.Equals("Top") Then
                BarTop.Checked = True
            Else
                BarBottom.Checked = True
            End If
            Dim Siz As Integer = INIRead(confile, "Config", "Size", 60)
            txtsize.Text = Siz
            Dim PicB As String = INIRead(confile, "Config", "Pic", "Default")
            lblselectedPic.Text = PicB
            If IO.File.Exists(lblselectedPic.Text) Then
                selectedPic.Image = Image.FromFile(lblselectedPic.Text)
            Else
                selectedPic.Image = My.Resources.Google_News_96px
            End If

            Dim SplPic As String = INIRead(confile, "Config", "SplitPic", "Default")
            lblselectedsplitpic.Text = SplPic
            If IO.File.Exists(lblselectedsplitpic.Text) Then
                selectedPicsplit.Image = Image.FromFile(lblselectedsplitpic.Text)
            Else
                selectedPicsplit.Image = My.Resources.logo
            End If

            Dim Fname = INIRead(confile, "Font", "Name", "Segoe UI Semibold")
            Dim Fsize = INIRead(confile, "Font", "Size", "11,25")
            lblselectedfont.Text = String.Concat({Fname, "/", Fsize})

            lblselectedfile.Text = INIRead(confile, "Config", "Text", "Default")
            If File.Exists(lblselectedfile.Text) Then
                Dim ext As String = Path.GetExtension(lblselectedfile.Text)
                If ext.Equals(".txt") Then
                    txtnbar.Text = ReadTextFile(lblselectedfile.Text)
                ElseIf ext.Equals(".docx") Then
                    txtnbar.Text = ReadDocFile(lblselectedfile.Text)
                Else
                    If txtnbar.Text.Equals("") Then
                        txtnbar.Text = "Default Text -- 'Hello Wolf' -- Error Reading file"
                    End If
                End If
            Else
                txtnbar.Text = "Default Text -- 'Hello Wolf'"
            End If

            Dim animsens As String = INIRead(confile, "Config", "Anim", "LTR")
            If animsens.Equals("LTR") Then
                animLTR.Checked = True
            Else
                animRTL.Checked = True
            End If

            txtspeed.Text = INIRead(confile, "Config", "Speed", "1")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub FrmMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MyNewsNotif.Icon = Me.Icon

        lblappnamever.Text = "(Build : " & Me.GetType.Assembly.GetName.Version.ToString & ")"

        'Dim strMajorVersion, strMinorVersion, strBuildVersion, strRevisionVersion As String

        'strMajorVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major()
        'strMinorVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor()
        'strBuildVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build()
        'strRevisionVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Revision()
        'lblappnamever.Text = "My News v" & strMajorVersion & "." & strMinorVersion
        'MsgBox("Version - " & strMajorVersion & "." & strMinorVersion & "." & strBuildVersion & "." & strRevisionVersion)

        If Not IO.File.Exists(confile) Then
            IO.File.Create(confile)
            INIWrite(confile, "Config", "Barcolor", "-1")
            INIWrite(confile, "Config", "Textcolor", "-16777216")
            INIWrite(confile, "Config", "Pos", "Bottom")
            INIWrite(confile, "Config", "Size", "60")
            INIWrite(confile, "Config", "Pic", "Default")
            INIWrite(confile, "Font", "Name", "Segoe UI Semibold")
            INIWrite(confile, "Font", "Size", "11,25")
            INIWrite(confile, "Font", "Text", "Default")
            INIWrite(confile, "Config", "Anim", "LTR")
            INIWrite(confile, "Config", "Speed", "1")
            INIWrite(confile, "Config", "SplitPic", "Default")
        Else
            LoadConfig()
        End If
    End Sub

    Private Sub btncolor_Click(sender As Object, e As EventArgs) Handles btncolor.Click
        If ColorDialog1.ShowDialog = DialogResult.OK Then
            INIWrite(confile, "Config", "Barcolor", ColorDialog1.Color.ToArgb)
            Load_Bar()
        End If
    End Sub

    Private Sub btntxtcolorbar_Click(sender As Object, e As EventArgs) Handles btntxtcolorbar.Click
        If ColorDialog1.ShowDialog = DialogResult.OK Then
            INIWrite(confile, "Config", "Textcolor", ColorDialog1.Color.ToArgb)
            Load_Bar()
        End If
    End Sub

    Private Sub btnsetsize_Click(sender As Object, e As EventArgs) Handles btnsetsize.Click
        INIWrite(confile, "Config", "Size", txtsize.Text)
        Load_Bar()
    End Sub

    Private Sub BunifuImageButton3_Click(sender As Object, e As EventArgs) Handles BunifuImageButton3.Click
        If BarBottom.Checked = True Then
            INIWrite(confile, "Config", "Pos", "Bottom")
        Else
            INIWrite(confile, "Config", "Pos", "Top")
        End If
        Load_Bar()
    End Sub

    Public Sub Load_Bar()
        Try
            LoadConfig()
            If FrmBar IsNot Nothing Then
                FrmBar.Close()
            End If
            FrmBar = New myBar.myBar
            FrmBar.FormSize = txtsize.Text
            FrmBar.BackColor = selectedColor.BackColor
            FrmBar.txtcol = selectedtextcolor.BackColor
            Dim pp As String = "Bottom"
            If BarTop.Checked Then
                pp = "Top"
            Else
                pp = "Bottom"
            End If
            If IO.File.Exists(lblselectedPic.Text) Then
                FrmBar.PicBar.Image = Image.FromFile(lblselectedPic.Text)
            Else
                FrmBar.PicBar.Image = My.Resources.Google_News_96px
            End If

            If IO.File.Exists(lblselectedsplitpic.Text) Then
                FrmBar.NTicker1.IMG = Image.FromFile(lblselectedsplitpic.Text)
            Else
                FrmBar.NTicker1.IMG = My.Resources.logo
            End If


            FrmBar.Pos = pp
            Dim gg As String() = lblselectedfont.Text.Split("/")
            FrmBar.FName = gg(0)
            FrmBar.Fsize = CType(gg(1), Decimal)
            FrmBar.AutoRegisterOnLoad = True
            FrmBar.NTicker1.RBTxt.Text = txtnbar.Text
            If animLTR.Checked Then
                FrmBar.PicBar.Dock = DockStyle.Left
                FrmBar.NTicker1.RTL = True
            Else
                FrmBar.PicBar.Dock = DockStyle.Right
                FrmBar.NTicker1.RTL = False
            End If
            FrmBar.NTicker1.Speed = txtspeed.Text
            FrmBar.Show()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub BunifuFlatButton2_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton2.Click
        Load_Bar()
    End Sub

    Private Sub BunifuFlatButton1_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton1.Click
        If FrmBar IsNot Nothing Then
            FrmBar.Close()
        End If
    End Sub

    Private Sub BunifuFlatButton3_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton3.Click
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            selectedPic.Image = Image.FromFile(OpenFileDialog1.FileName)
            INIWrite(confile, "Config", "Pic", OpenFileDialog1.FileName)
        Else
            selectedPic.Image = My.Resources.Google_News_96px
            lblselectedPic.Text = "Default"
            INIWrite(confile, "Config", "Pic", "Default")
        End If
        Load_Bar()
    End Sub

    Private Sub btnFont_Click(sender As Object, e As EventArgs) Handles btnFont.Click
        If FontDialog1.ShowDialog = DialogResult.OK Then
            lblselectedfont.Text = String.Concat({FontDialog1.Font.Name, "/", FontDialog1.Font.Size, "/", FontDialog1.Font.Bold, "/", FontDialog1.Font.Italic})
            INIWrite(confile, "Font", "Name", FontDialog1.Font.Name)
            INIWrite(confile, "Font", "Size", FontDialog1.Font.Size)
        Else
            lblselectedfont.Text = String.Concat({"Segoe UI Semibold", "/", "11,25"})
            INIWrite(confile, "Font", "Name", "Segoe UI Semibold")
            INIWrite(confile, "Font", "Size", "11,25")
        End If
        Load_Bar()
    End Sub

    Private Sub FrmMenu_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If FrmBar IsNot Nothing Then
            FrmBar.Close()
        End If
        Me.Dispose()
    End Sub

    Private Sub BunifuImageButton4_Click(sender As Object, e As EventArgs) Handles BunifuImageButton4.Click
        If File.Exists(lblselectedfile.Text) Then
            Dim flex As String = IO.Path.GetExtension(lblselectedfile.Text)
            If flex.Equals(".txt") Then
                txtnbar.Text = ReadTextFile(lblselectedfile.Text)
            ElseIf flex.Equals(".docx") Then
                txtnbar.Text = ReadDocFile(lblselectedfile.Text)
            Else
                txtnbar.AppendText("Error reading file extention !!" & flex)
            End If
        Else
            txtnbar.AppendText("Error reading file !!")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If OFDText.ShowDialog = DialogResult.OK Then
            lblselectedfile.Text = OFDText.FileName
            INIWrite(confile, "Config", "Text", OFDText.FileName)
        Else
            lblselectedfile.Text = "Default"
            INIWrite(confile, "Config", "Text", "Default")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If OFDDocx.ShowDialog = DialogResult.OK Then
            lblselectedfile.Text = OFDDocx.FileName
            INIWrite(confile, "Config", "Text", OFDDocx.FileName)
        Else
            lblselectedfile.Text = "Default"
            INIWrite(confile, "Config", "Text", "Default")
        End If
    End Sub


    Public Function ReadTextFile(ByVal Pat As String) As String
        Dim RB As RichTextBox = New RichTextBox
        Try
            If IO.File.Exists(Pat) Then
                Dim reader As New StreamReader(Pat, Encoding.UTF8)
                Do While reader.Peek() >= 0
                    RB.AppendText(reader.ReadLine)
                Loop
                reader.Close()
                Return RB.Text
            Else
                RB.AppendText("Default Text 'Hello Wolf' Error Reading file")
                Return RB.Text
            End If
        Catch ex As Exception
            RB.AppendText("Error Reading")
            Return RB.Text
        End Try
    End Function

    Public Function ReadDocFile(ByVal Pat As String) As String
        Dim RB As RichTextBox = New RichTextBox
        Try
            If IO.File.Exists(Pat) Then
                Dim a As String = ""
                Dim App As Application = New Application
                Dim doc As Document
                doc = App.Documents.Open(Pat)
                Dim co As Integer = doc.Words.Count
                For i As Integer = 1 To co
                    RB.AppendText(doc.Words(i).Text)
                Next
                doc.Close()
                Return RB.Text
            Else
                RB.AppendText("Default Text 'Hello Wolf' Error Reading file")
                Return RB.Text
            End If
        Catch ex As Exception
            RB.AppendText("Error Reading")
            Return RB.Text
        End Try
    End Function

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            MyNewsNotif.Visible = True
            MyNewsNotif.Icon = Me.Icon
            MyNewsNotif.BalloonTipIcon = ToolTipIcon.Info
            MyNewsNotif.BalloonTipTitle = "My News " & "(Build : " & Me.GetType.Assembly.GetName.Version.ToString & ")"
            MyNewsNotif.BalloonTipText = "Double click to show menu"
            MyNewsNotif.ShowBalloonTip(50000)
            ShowInTaskbar = False
        End If
    End Sub

    Private Sub MyNewsNotif_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyNewsNotif.DoubleClick
        'Me.Show()
        ShowInTaskbar = True
        Me.WindowState = FormWindowState.Normal
        MyNewsNotif.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        INIWrite(confile, "Config", "Text", "Default")
        Load_Bar()
    End Sub

    Private Sub BunifuImageButton5_Click(sender As Object, e As EventArgs) Handles BunifuImageButton5.Click
        If animLTR.Checked = True Then
            INIWrite(confile, "Config", "Anim", "LTR")
        Else
            INIWrite(confile, "Config", "Anim", "RTL")
        End If
        Load_Bar()
    End Sub

    Private Sub BunifuImageButton6_Click(sender As Object, e As EventArgs) Handles BunifuImageButton6.Click
        Try
            If CDbl(txtspeed.Text) < 100 And CDbl(txtspeed.Text) > 0.01 Then
                INIWrite(confile, "Config", "Speed", txtspeed.Text)
                Load_Bar()
            Else
                txtspeed.Text = "1"
                INIWrite(confile, "Config", "Speed", "1")
                Load_Bar()
            End If
        Catch ex As Exception
            txtspeed.Text = "1"
            INIWrite(confile, "Config", "Speed", "1")
            Load_Bar()
        End Try
    End Sub

    Private Sub btnsplitbar_Click(sender As Object, e As EventArgs) Handles btnsplitbar.Click
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            selectedPicsplit.Image = Image.FromFile(OpenFileDialog1.FileName)
            INIWrite(confile, "Config", "SplitPic", OpenFileDialog1.FileName)
        Else
            selectedPicsplit.Image = My.Resources.logo
            lblselectedsplitpic.Text = "Default"
            INIWrite(confile, "Config", "SplitPic", "Default")
        End If
        Load_Bar()
    End Sub

    Private Sub BunifuImageButton7_Click(sender As Object, e As EventArgs) Handles BunifuImageButton7.Click
        Dim ab As About = New About
        ab.Show()
    End Sub
End Class