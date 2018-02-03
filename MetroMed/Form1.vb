Imports System.Drawing
Imports System.IO
Imports MetroFramework
Public Class Form1
    Dim ButtonAnIndex, RichStep As Integer, Arguments, ParFile As String

    Private Sub Form1_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If File.Exists(MedExtra & "Mini.ini") Then WMetroMed()
    End Sub
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'If File.Exists(Application.StartupPath & "\MedGuiR.exe") = False Then
        'MetroMessageBox.Show(Me, "Unable to start Mednafen", "Process Error", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        'Me.Close()
        'Exit Sub
        'End If
        AniBoxArt(0) = AnimationControl1
        AniBoxArt(1) = AnimationControl2
        AniBoxArt(2) = AnimationControl3
        AniBoxArt(3) = AnimationControl4
        AniBoxArt(4) = AnimationControl5
        AniBoxArt(5) = AnimationControl6
        AniBoxArt(6) = AnimationControl7
        AniBoxArt(7) = AnimationControl8
        AniBoxArt(8) = AnimationControl9
        AniBoxArt(9) = AnimationControl10
        AniBoxArt(10) = AnimationControl11
        AniBoxArt(11) = AnimationControl12
        AniBoxArt(12) = AnimationControl13
        AniBoxArt(13) = AnimationControl14
        AniBoxArt(14) = AnimationControl15
        AniBoxArt(15) = AnimationControl16
        AniBoxArt(16) = AnimationControl17
        AniBoxArt(17) = AnimationControl18
        AniBoxArt(18) = AnimationControl19
        AniBoxArt(19) = AnimationControl20
        AniBoxArt(20) = AnimationControl21
        AniBoxArt(21) = AnimationControl22
        AniBoxArt(22) = AnimationControl23
        AniBoxArt(23) = AnimationControl24

        LoadResource()
        RMetroMed()
        MetroListServer_reload()
        ParseMednafenConfig()
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
    End Sub

    Private Sub MetroListServer_reload()
        mServer.Items.Clear()
        Try
            Dim fullPath As String = MedExtra & "ListServer.txt"
            If File.Exists(fullPath) Then
                mServer.Items.AddRange(IO.File.ReadAllLines(fullPath))
            Else
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub ThemeMod()
        MetroStyleManager1.Theme = cmbTheme.SelectedIndex
        MetroStyleManager1.Style = cmbStyle.SelectedIndex
        MetroToolTip1.Theme = MetroStyleManager1.Theme
        MetroToolTip1.Style = MetroStyleManager1.Style
        MetroTextBox1.Theme = MetroStyleManager1.Theme
        MetroTextBox1.Style = MetroStyleManager1.Style
        MetroTile1.Theme = MetroStyleManager1.Theme
        MetroTile1.Style = MetroStyleManager1.Style
        MetroTile2.Theme = MetroStyleManager1.Theme
        MetroTile2.Style = MetroStyleManager1.Style
        MetroLabel1.Style = MetroStyleManager1.Style
        MetroLabel2.Style = MetroStyleManager1.Style
        MetroLabel3.Theme = MetroStyleManager1.Theme
        MetroLabel3.Style = MetroStyleManager1.Style
        MetroLabel4.Theme = MetroStyleManager1.Theme
        MetroLabel4.Style = MetroStyleManager1.Style
        ColorCheck()
        Me.Style = MetroStyleManager1.Style
        'MetroTile1.ForeColor = System.Drawing.Color.FromName(cmbTheme.Text)
        Me.Theme = MetroStyleManager1.Theme

        Me.Refresh()

    End Sub

    Private Sub ColorCheck()
        Dim BackTheme, BackTheme1 As String

        Select Case cmbTheme.Text()
            Case "Light", "Default"
                BackTheme = "Black"
                BackTheme1 = "White"
            Case "Dark"
                BackTheme = "White"
                BackTheme1 = "Black"
        End Select

        If CheckBox1.Checked = True Then
            CheckBox1.BackColor = Color.FromName(BackTheme)
        Else
            CheckBox1.BackColor = Color.FromName(BackTheme1)
        End If

        If CheckBox2.Checked = True Then
            CheckBox2.BackColor = Color.FromName(BackTheme)
        Else
            CheckBox2.BackColor = Color.FromName(BackTheme1)
        End If

    End Sub

    Private Sub cmbTheme_TextChanged(sender As Object, e As System.EventArgs) Handles cmbTheme.TextChanged
        ThemeMod()
        'WMetroMed()
    End Sub

    Private Sub cmbStyle_TextChanged(sender As Object, e As System.EventArgs) Handles cmbStyle.TextChanged
        ThemeMod()
        'ColorCheck()
    End Sub

    Private Sub mEffect_IndexChanged(sender As Object, e As System.EventArgs) Handles mEffect.SelectedIndexChanged
        Effect = mEffect.Text.Trim
        CoverEffects()
        ReadCSV()
    End Sub

    Private Sub MetroButton16_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton16.Click
        MednafenModule = "def"
        CountRows()
    End Sub

    Private Sub MetroButton17_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton17.Click
        MednafenModule = "fav"
        CountRows()
    End Sub

    Private Sub MetroButton1_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton1.Click
        MednafenModule = "gb"
        CountRows()
    End Sub

    Private Sub MetroButton2_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton2.Click
        MednafenModule = "gba"
        CountRows()
    End Sub

    Private Sub MetroButton3_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton3.Click
        MednafenModule = "gg"
        CountRows()
    End Sub

    Private Sub MetroButton4_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton4.Click
        MednafenModule = "lynx"
        CountRows()
    End Sub

    Private Sub MetroButton5_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton5.Click
        MednafenModule = "md"
        CountRows()
    End Sub

    Private Sub MetroButton6_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton6.Click
        MednafenModule = "nes"
        CountRows()
    End Sub

    Private Sub MetroButton7_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton7.Click
        MednafenModule = "ngp"
        CountRows()
    End Sub

    Private Sub MetroButton8_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton8.Click
        MednafenModule = "pce"
        CountRows()
    End Sub

    Private Sub MetroButton9_Click(sender As System.Object, e As System.EventArgs)
        MednafenModule = "pcfx"
        SelectedImage()
    End Sub

    Private Sub MetroButton10_Click(sender As System.Object, e As System.EventArgs)
        MednafenModule = "psx"
        SelectedImage()
    End Sub

    Private Sub MetroButton11_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton11.Click
        MednafenModule = "sms"
        CountRows()
    End Sub

    Private Sub MetroButton12_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton12.Click
        MednafenModule = "snes"
        CountRows()
    End Sub

    Private Sub MetroButton13_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton13.Click
        MednafenModule = "vi"
        SelectedImage()
    End Sub

    Private Sub MetroButton14_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton14.Click
        MednafenModule = "vb"
        CountRows()
    End Sub

    Private Sub MetroButton15_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton15.Click
        MednafenModule = "wswan"
        CountRows()
    End Sub

    Private Sub SelectedImage()
        Dim vImage As String

        OpenFileDialog1.Title = "Select a virtual image game..."
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            CleanAniboxart()
            vImage = OpenFileDialog1.FileName
            AniTag = Path.GetFileNameWithoutExtension(vImage) & "|_image_console_| |?|" &
                vImage.ToString.Trim & "|Virtual Image|" & MednafenModule & "|." &
            Path.GetExtension(vImage) & "|00000000|"
            AniBoxArt(0).Tag = AniTag & "|0"
            SplitTag()
            MetroToolTip1.SetToolTip(AniBoxArt(0), "     " & cleanpsx(RemoveAmpersand(TagSplit(0))) & " " & TagSplit(2) & "     ")
            'AniCover = MedExtra & "BoxArt\" & TagSplit(5) & "\" & TagSplit(0) & ".png"
            'If File.Exists(AniCover) = False Then
            'SearchScrape()
            'End If
            AniCover = MedExtra & "BoxArt\NoPr.png"
            AniBoxArt(0).AnimatedImage = GamesInfo.Resize(New Bitmap(AniCover), AniBoxArt(0).Width, AniBoxArt(0).Height, True)
            AniBoxArt(0).Animate(10)
        End If
    End Sub

    Private Sub MetroTile2_Click(sender As System.Object, e As System.EventArgs) Handles MetroTile2.Click

        If Val(MetroLabel1.Text) >= Val(MetroLabel2.Text) Then
            CountRows()
            'BoxResult = 0
        Else

            If MetroTextBox1.Text.Trim <> "" Then
                NextSearch()
            Else
                MetroLabel1.Text = Val(MetroLabel1.Text) + 1
                CurrPage = (NextPage + 1)
                NextPage = (NextPage + 24)
            End If
            ReadCSV()
        End If
    End Sub

    Private Sub MetroTile1_Click(sender As System.Object, e As System.EventArgs) Handles MetroTile1.Click
        If Val(MetroLabel1.Text) <= 1 Then Exit Sub
        MetroLabel1.Text = Val(MetroLabel1.Text) - 1
        CurrPage = (CurrPage - 24)
        NextPage = (NextPage - 24)
        ReadCSV()
    End Sub

    Private Sub DetectButton(sender As Object, e As EventArgs) Handles AnimationControl9.MouseEnter, AnimationControl8.MouseEnter, AnimationControl7.MouseEnter, AnimationControl6.MouseEnter, AnimationControl5.MouseEnter, AnimationControl4.MouseEnter, AnimationControl3.MouseEnter, AnimationControl24.MouseEnter, AnimationControl23.MouseEnter, AnimationControl22.MouseEnter, AnimationControl21.MouseEnter, AnimationControl20.MouseEnter, AnimationControl2.MouseEnter, AnimationControl19.MouseEnter, AnimationControl18.MouseEnter, AnimationControl17.MouseEnter, AnimationControl16.MouseEnter, AnimationControl15.MouseEnter, AnimationControl14.MouseEnter, AnimationControl13.MouseEnter, AnimationControl12.MouseEnter, AnimationControl11.MouseEnter, AnimationControl10.MouseEnter, AnimationControl1.MouseEnter
        PictureBox1.Image = Nothing
        PictureBox2.Image = Nothing
        RichTextBox1.Clear()

        Dim fast, faust As String
        Dim b As AnimationControl = CType(sender, AnimationControl)
        AniTag = (b.Tag)
        If AniTag = "" Then Exit Sub
        SplitTag()
        ButtonAnIndex = TagSplit(10)
        ReadXml()
        MetroLabel3.Text = MetroToolTip1.GetToolTip(b).Trim

        Dim title, snap As Bitmap
        If File.Exists(MedExtra & "Snaps\" & TagSplit(5) & "\CRC_Titles\" & Trim(TagSplit(8)) & ".png") Then
            PictureBox1.Tag = MedExtra & "Snaps\" & TagSplit(5) & "\CRC_Titles\" & Trim(TagSplit(8)) & ".png"
            title = New Bitmap(PictureBox1.Tag.ToString)
            PictureBox1.Image = DirectCast(GamesInfo.Resize(title, PictureBox1.Width, PictureBox1.Height, True), Image)
            'Else
            'PictureBox1.Image = Nothing
        End If
        If File.Exists(MedExtra & "Snaps\" & TagSplit(5) & "\CRC_Snaps\" & Trim(TagSplit(8)) & ".png") Then
            PictureBox2.Tag = MedExtra & "Snaps\" & TagSplit(5) & "\CRC_Snaps\" & Trim(TagSplit(8)) & ".png"
            snap = New Bitmap(PictureBox2.Tag.ToString)
            PictureBox2.Image = DirectCast(GamesInfo.Resize(snap, PictureBox2.Width, PictureBox2.Height, True), Image)
            'Else
            'PictureBox2.Image = Nothing
        End If

        If CheckBox1.Checked = True Then
            fast = "-pce.enable 0 "
        Else
            fast = "-pce.enable 1 "
        End If
        If CheckBox2.Checked = True Then
            faust = "-snes.enable 0 "
        Else
            faust = "-snes.enable 1 "
        End If

        ParFile = Chr(34) & TagSplit(4) & Chr(34)
        Arguments = fast & faust & ParFile
        RichTextBox1.Focus()
        RichStep = 0
    End Sub

    Private Sub RichTextBoxScroll(sender As Object, e As MouseEventArgs) Handles AnimationControl9.MouseWheel, AnimationControl8.MouseWheel, AnimationControl7.MouseWheel, AnimationControl6.MouseWheel, AnimationControl5.MouseWheel, AnimationControl4.MouseWheel, AnimationControl3.MouseWheel, AnimationControl24.MouseWheel, AnimationControl23.MouseWheel, AnimationControl22.MouseWheel, AnimationControl21.MouseWheel, AnimationControl20.MouseWheel, AnimationControl2.MouseWheel, AnimationControl19.MouseWheel, AnimationControl18.MouseWheel, AnimationControl17.MouseWheel, AnimationControl16.MouseWheel, AnimationControl15.MouseWheel, AnimationControl14.MouseWheel, AnimationControl13.MouseWheel, AnimationControl12.MouseWheel, AnimationControl11.MouseWheel, AnimationControl10.MouseWheel, AnimationControl1.MouseWheel

        RichTextBox1.HideSelection = True

        If e.Delta < 0 Then
            If RichTextBox1.SelectionStart < RichTextBox1.Text.Length And RichStep >= 0 Then
                RichStep = RichStep + 120
                RichTextBox1.SelectionStart = RichStep
            Else
                RichTextBox1.SelectionStart = RichTextBox1.Text.Length
            End If

        ElseIf e.Delta > 0 Then
            If RichTextBox1.SelectionStart > 120 Then
                RichStep = RichStep - 120
                RichTextBox1.SelectionStart = RichStep
            Else
                RichTextBox1.SelectionStart = 0
            End If
        End If
    End Sub

    Public Sub LaunchGame(sender As Object, e As EventArgs) Handles AnimationControl9.DoubleClick, AnimationControl8.DoubleClick, AnimationControl7.DoubleClick, AnimationControl6.DoubleClick, AnimationControl5.DoubleClick, AnimationControl4.DoubleClick, AnimationControl3.DoubleClick, AnimationControl24.DoubleClick, AnimationControl23.DoubleClick, AnimationControl22.DoubleClick, AnimationControl21.DoubleClick, AnimationControl20.DoubleClick, AnimationControl2.DoubleClick, AnimationControl19.DoubleClick, AnimationControl18.DoubleClick, AnimationControl17.DoubleClick, AnimationControl16.DoubleClick, AnimationControl15.DoubleClick, AnimationControl14.DoubleClick, AnimationControl13.DoubleClick, AnimationControl12.DoubleClick, AnimationControl11.DoubleClick, AnimationControl10.DoubleClick, AnimationControl1.DoubleClick
        StartProcess()
    End Sub

    Private Sub mPlay_Click(sender As Object, e As EventArgs) Handles mPlay.Click
        StartProcess()
    End Sub
    Private Sub OnlinePlayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mOnlinePlay.Click
        Arguments = "-netplay.nick " & mNickname.Text.Trim & " -netplay.host " & mServer.Text.Trim & " -connect " & Arguments
        StartProcess()
    End Sub

    Public Sub StartProcess()
        Dim execute As New Process
        Try
            With execute.StartInfo
                .FileName = "mednafen"
                .Arguments = Arguments
                .WorkingDirectory = MedPath
            End With
            execute.Start()
            Arguments = ""
        Catch ex As Exception
            MetroMessageBox.Show(Me, "Unable to start Mednafen", "Process Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Arguments = ""
        End Try

    End Sub

    Private Sub mGuiMode_Click(sender As System.Object, e As System.EventArgs) Handles mGuiMode.Click
        If File.Exists(Application.StartupPath & "\MedGuiR.exe") Then
            Dim parameters As String = " -file=" & ParFile.Trim
            Process.Start(Application.StartupPath & "\MedGuiR.exe", parameters)
        Else
            MetroMessageBox.Show(Me, "MedGui Reborn not detected!", "MedGuiR not detected...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub About_Click(sender As Object, e As EventArgs) Handles About.Click
        MetroMessageBox.Show(Me, "I'm only a fucking servant of my boss." & vbCrLf &
            "I'm  in a Alpha state if you want a complete GUI call my boss MedGui Reborn", "MetroMed " & Me.Tag, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub ParseMednafenConfig()
        Dim row, splitrow() As String
        Try
            Using reader As New StreamReader(MedPath & "\mednafen-09x.cfg")
                While Not reader.EndOfStream
                    row = reader.ReadLine
                    splitrow = row.Split(" ")
                    Select Case True
                        Case row.Contains("netplay.host")
                            mServer.Text = splitrow(1)
                        Case row.Contains("netplay.port")
                            'port = splitrow(1)
                        Case row.Contains("netplay.nick")
                            mNickname.Text = splitrow(1)
                        Case row.Contains("netplay.password")
                            'If splitrow(1) = "" Then Password = "No" Else Password = "Yes"
                        Case row.Contains("netplay.gamekey")
                            'Gamekey = splitrow(1)
                    End Select
                End While

                reader.Dispose()
                reader.Close()
            End Using

        Catch
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        ColorCheck()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox2.CheckedChanged
        ColorCheck()
    End Sub

    Private Sub MetroTextBox1_KeyUp(sender As Object, e As KeyEventArgs) Handles MetroTextBox1.KeyUp
        Select Case e.KeyCode
            Case Keys.Cancel, Keys.Delete, Keys.Enter
                CountRows()
        End Select
    End Sub

    Private Sub Form1_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        CountRows()
    End Sub

    Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        If File.Exists(MedExtra & "Scanned\" & MednafenModule & ".csv") = False Then Exit Sub
        If Me.WindowState = FormWindowState.Maximized Then
            CountRows()
            PictureBox1.Image = DirectCast(GamesInfo.Resize(New Bitmap(PictureBox1.Tag.ToString), PictureBox1.Width, PictureBox1.Height, True), Image)
            PictureBox2.Image = DirectCast(GamesInfo.Resize(New Bitmap(PictureBox2.Tag.ToString), PictureBox2.Width, PictureBox2.Height, True), Image)
        End If
        If Me.WindowState = FormWindowState.Minimized Then
            PictureBox1.Image = DirectCast(GamesInfo.Resize(New Bitmap(PictureBox1.Tag.ToString), PictureBox1.Width, PictureBox1.Height, True), Image)
            PictureBox2.Image = DirectCast(GamesInfo.Resize(New Bitmap(PictureBox2.Tag.ToString), PictureBox2.Width, PictureBox2.Height, True), Image)
        End If
    End Sub

End Class