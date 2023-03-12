Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports MetroFramework

Public Class MetroMed

    <DllImport("user32")>
    Private Shared Function HideCaret(ByVal hWnd As IntPtr) As Integer
    End Function

    Dim ButtonAnIndex, RichStep, SubStop As Integer, Arguments, FileParameter, vImage, cImage As String
    Dim NProcess, MednafenCore As String

    Private Sub Form1_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If File.Exists(MedExtra & "Mini.ini") Then WMetroMed()
    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        TableLayoutPanel1.Dock = DockStyle.Fill
        Me.Text = "MetroMed " & Me.Tag

        If File.Exists(Application.StartupPath & "\MedGuiR.exe") = False Then
            MetroMessageBox.Show(Me, "Unable to start MedGui Reborn" & vbCrLf &
                                 "You need to put MetroMed in the same path of MedGuiR.exe", "Process Error", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Me.Close()
            Exit Sub
        End If

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
        HideCaret(RichTextBox1.Handle)
        If MednafenModule.Trim <> "" Then CountRows()
    End Sub

    Private Sub MetroListServer_reload()
        mServer.Items.Clear()

        Try
            Dim fullPath As String = MedExtra & "ListServer.txt"
            If IO.File.Exists(fullPath) Then
                Dim oFile As IO.File
                Dim oRead As IO.StreamReader
                Dim Sserver As String
                Dim line() As String

                Try
                    oRead = oFile.OpenText(fullPath)

                    While oRead.Peek <> -1
                        line = Split(oRead.ReadLine(), ":")
                        Sserver = line(0)
                        'port = line(1)
                        mServer.Items.Add(Sserver)
                    End While
                Catch ex As Exception
                Finally
                    oRead.Close()
                End Try

                'cmbServer.Items.AddRange(IO.File.ReadAllLines(fullPath))
                'MedGuiR.ServerToolStripComboBox2.Items.AddRange(IO.File.ReadAllLines(fullPath))
            Else
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub PopulateNetplay()
        mPort.Text = "4046"

        Dim Sport As String
        Dim line() As String

        Using sr As New System.IO.StreamReader(MedExtra & "ListServer.txt")
            Do Until sr.EndOfStream
                Try
                    Dim sBuf As String = sr.ReadLine
                    If sBuf.Contains(mServer.Text) Then
                        If sBuf.Contains(":") Then
                            line = Split(sBuf, ":")
                            Sport = line(1)
                            Dim xvar() As String
                            If line(1).Contains("=") Then
                                xvar = Split(line(1), "=")
                                Sport = xvar(0)
                            End If
                        Else
                            Sport = "4046"
                        End If
                    End If
                    If Sport = "" Then Sport = "4046"
                    mPort.Text = Sport
                Catch ex As Exception

                End Try
            Loop
        End Using
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
        MetroGrid1.Theme = MetroStyleManager1.Theme
        MetroGrid1.Style = MetroStyleManager1.Style
        Me.Style = MetroStyleManager1.Style
        'MetroTile1.ForeColor = System.Drawing.Color.FromName(cmbTheme.Text)
        Me.Theme = MetroStyleManager1.Theme
        ColorCheck()
        Me.Refresh()

    End Sub

    Private Sub ColorCheck()
        MedExtra = Application.StartupPath & "\MedGuiR\"

        Select Case cmbTheme.Text()
            Case "Light", "Default"
                MetroGrid1.GridColor = Color.Black
                CheckBox4.BackgroundImage = (New Bitmap(MedExtra & "Resource\Gui\list_black.png"))
                CheckBox3.BackgroundImage = (New Bitmap(MedExtra & "Resource\Gui\grid_black.png"))
            Case "Dark"
                MetroGrid1.GridColor = Color.White
                CheckBox4.BackgroundImage = (New Bitmap(MedExtra & "Resource\Gui\list_white.png"))
                CheckBox3.BackgroundImage = (New Bitmap(MedExtra & "Resource\Gui\grid_white.png"))
        End Select

        MetroGrid1.DefaultCellStyle.ForeColor = Color.FromName(cmbStyle.Text)

        If CheckBox1.Checked = True Then
            CheckBox1.BackColor = Color.DimGray
        Else
            CheckBox1.BackColor = Me.BackColor
        End If

        If CheckBox2.Checked = True Then
            CheckBox2.BackColor = Color.DimGray
        Else
            CheckBox2.BackColor = Me.BackColor
        End If

        If CheckBox3.Checked = True Then
            CheckBox3.BackColor = Color.DimGray
        Else
            CheckBox3.BackColor = Me.BackColor
        End If

        If CheckBox4.Checked = True Then
            CheckBox4.BackColor = Color.DimGray
        Else
            CheckBox4.BackColor = Me.BackColor
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

    Private Sub MetroButton18_Click(sender As Object, e As EventArgs) Handles MetroButton18.Click
        MednafenModule = "apple2"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton16_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton16.Click
        MednafenModule = "def"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton17_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton17.Click
        MednafenModule = "fav"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton1_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton1.Click
        MednafenModule = "gb"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton2_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton2.Click
        MednafenModule = "gba"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton3_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton3.Click
        MednafenModule = "gg"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton4_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton4.Click
        MednafenModule = "lynx"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton5_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton5.Click
        MednafenModule = "md"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton6_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton6.Click
        MednafenModule = "nes"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton7_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton7.Click
        MednafenModule = "ngp"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton8_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton8.Click
        MednafenModule = "pce"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton10_Click(sender As Object, e As EventArgs) Handles MetroButton10.Click
        MednafenModule = "ss"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton9_Click(sender As Object, e As EventArgs) Handles MetroButton9.Click
        MednafenModule = "psx"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton11_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton11.Click
        MednafenModule = "sms"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton12_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton12.Click
        MednafenModule = "snes"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton13_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton13.Click
        'MednafenModule = ""
        SelectedImage()
    End Sub

    Private Sub MetroButton14_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton14.Click
        MednafenModule = "vb"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub MetroButton15_Click(sender As System.Object, e As System.EventArgs) Handles MetroButton15.Click
        MednafenModule = "wswan"
        If CheckBox4.Checked = True Then
            PopulateGrid()
        Else
            CountRows()
        End If
    End Sub

    Private Sub SelectedImage()
        SubStop = 0
        OpenFileDialog1.Title = "Select a virtual image game..."
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            CleanAniboxart()
            vImage = OpenFileDialog1.FileName
            detectcdtype()
            If SubStop = 1 Then Exit Sub
            AniTag = Path.GetFileNameWithoutExtension(vImage) & "|_image_console_| |?|" &
                vImage.ToString.Trim & "|" & cImage & "|" & MednafenModule & "|." &
            Path.GetExtension(vImage) & "|00000000|"
            AniBoxArt(0).Tag = AniTag & "|0"
            SplitTag()
            MetroToolTip1.SetToolTip(AniBoxArt(0), "     " & cleanpsx(RemoveAmpersand(TagSplit(0))) & " " & TagSplit(2) & "     " &
 vbCrLf & TagSplit(5).Trim)

            'AniCover = MedExtra & "BoxArt\" & TagSplit(5) & "\" & TagSplit(0) & ".png"
            'If File.Exists(AniCover) = False Then
            'SearchScrape()
            'End If
            AniCover = MedExtra & "BoxArt\NoPr.png"
            AniBoxArt(0).AnimatedImage = GamesInfo.Resize(New Bitmap(AniCover), AniBoxArt(0).Width, AniBoxArt(0).Height, True)
            AniBoxArt(0).Animate(performance)
            MetroLabel1.Text = 1
            MetroLabel2.Text = 1
        End If
    End Sub

    Sub detectcdtype()
        Dim offset As Long
        Dim consoletype, rvimage As String

        Select Case LCase(Path.GetExtension(vImage))
            'If LCase(Path.GetExtension(n_psx)) = ".cue" Then
            Case ".cue" ', ".toc"
                Dim righe As String() = File.ReadAllLines(vImage)
                Dim result As String

                For i = 0 To 10
                    If LCase(righe(i)).Contains(" binary") Then
                        result = righe(i)
                        Exit For
                    End If
                Next

                Dim startPosition As Integer
                Dim word2 As String

                If result.Contains(Chr(34)) Then
                    startPosition = result.IndexOf(Chr(34)) + 1
                    word2 = result.Substring(startPosition, result.IndexOf(Chr(34), startPosition) - startPosition)
                Else
                    startPosition = result.IndexOf("E ") + 2
                    word2 = result.Substring(startPosition, result.IndexOf(" B", startPosition) - startPosition).Trim
                End If
                rvimage = Replace(vImage, Path.GetFileName(vImage), "") & word2
            Case ".ccd"
                rvimage = Replace(vImage, ".ccd", "") & ".img"
            Case Else
                Exit Sub
        End Select

        If File.Exists(rvimage) = False Then
            MetroMessageBox.Show(Me, "Bad or corrupted " & Path.GetExtension(vImage) & " file." & vbCrLf &
                   "Open " & vImage & " with a text editor and fix the reference to the binary file",
                   "Bad or copputed " & Path.GetExtension(vImage) & "...", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            SubStop = 1
            Exit Sub
        End If

        Using fs As New FileStream(rvimage, FileMode.Open, FileAccess.Read)

            For offset = 0 To 9500
                consoletype = consoletype & Convert.ToChar(fs.ReadByte())
            Next offset
        End Using

        Select Case True
            Case LCase(consoletype).Contains("sega")
                MednafenModule = "ss"
                cImage = "Sega - Saturn"
            Case LCase(consoletype).Contains("engine")
                MednafenModule = "pce"
                cImage = "NEC - PC Engine - TurboGrafx 16"
            Case LCase(consoletype).Contains("sony")
                MednafenModule = "psx"
                cImage = "Sony PlayStation"
            Case Else
                MednafenModule = "pcfx"
                cImage = "PC-FX"
        End Select
    End Sub

    Private Sub MetroTile2_Click(sender As System.Object, e As System.EventArgs) Handles MetroTile2.Click

        If Val(MetroLabel1.Text) = Val(MetroLabel2.Text) And Val(MetroLabel2.Text) <> 1 Then
            CountRows()
            'BoxResult = 0
        ElseIf Val(MetroLabel1.Text) < Val(MetroLabel2.Text) Then

            If MetroTextBox1.Text.Trim <> "" Then
                NextSearch()
            Else
                MetroLabel1.Text = Val(MetroLabel1.Text) + 1
                CurrPage = (NextPage + 1)
                NextPage = (NextPage + 24)
            End If
            ReadCSV()
        Else
            Exit Sub
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

        Dim b As AnimationControl = CType(sender, AnimationControl)
        AniTag = (b.Tag)

        RetrieveSnap()

        If (TagSplit.Length - 1) = 10 Then
            ButtonAnIndex = TagSplit(10)
            AniBoxArt(ButtonAnIndex).Animate(performance)
        End If

        AddArguments()
    End Sub

    Private Sub RetrieveSnap()
        PictureBox1.Image = Nothing
        PictureBox2.Image = Nothing
        RichTextBox1.Clear()

        If AniTag = "" Then Exit Sub
        SplitTag()
        ReadXml()
        MetroLabel3.Text = cleanpsx(RemoveAmpersand(TagSplit(0))) & " " & TagSplit(2) 'MetroToolTip1.GetToolTip(b).Trim

        Dim title, snap As Bitmap
        If File.Exists(MedExtra & "Snaps\" & TagSplit(5) & "\CRC_Titles\" & Trim(TagSplit(8)) & ".png") Then
            title = New Bitmap(MedExtra & "Snaps\" & TagSplit(5) & "\CRC_Titles\" & Trim(TagSplit(8)) & ".png")
            PictureBox1.Image = DirectCast(GamesInfo.Resize(title, PictureBox1.Width, PictureBox1.Height, True), Image)
            'Else
            'PictureBox1.Image = Nothing
        End If
        If File.Exists(MedExtra & "Snaps\" & TagSplit(5) & "\CRC_Snaps\" & Trim(TagSplit(8)) & ".png") Then
            snap = New Bitmap(MedExtra & "Snaps\" & TagSplit(5) & "\CRC_Snaps\" & Trim(TagSplit(8)) & ".png")
            PictureBox2.Image = DirectCast(GamesInfo.Resize(snap, PictureBox2.Width, PictureBox2.Height, True), Image)
            'Else
            'PictureBox2.Image = Nothing
        End If
    End Sub

    Private Sub AddArguments()
        Dim fast, faust As String
        MednafenCore = TagSplit(6)

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

        FileParameter = TagSplit(4)
        If FileParameter.StartsWith("..\") Then FileParameter = Replace(FileParameter, "..", Application.StartupPath)
        Arguments = fast & faust & Chr(34) & FileParameter & Chr(34)

        If CheckBox3.Checked = True Then
            RichTextBox1.Focus()
            RichStep = 0
            HideCaret(RichTextBox1.Handle)
        End If
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
        HideCaret(RichTextBox1.Handle)
    End Sub

    Public Sub LaunchGame(sender As Object, e As EventArgs) Handles AnimationControl9.DoubleClick, AnimationControl8.DoubleClick, AnimationControl7.DoubleClick, AnimationControl6.DoubleClick, AnimationControl5.DoubleClick, AnimationControl4.DoubleClick, AnimationControl3.DoubleClick, AnimationControl24.DoubleClick, AnimationControl23.DoubleClick, AnimationControl22.DoubleClick, AnimationControl21.DoubleClick, AnimationControl20.DoubleClick, AnimationControl2.DoubleClick, AnimationControl19.DoubleClick, AnimationControl18.DoubleClick, AnimationControl17.DoubleClick, AnimationControl16.DoubleClick, AnimationControl15.DoubleClick, AnimationControl14.DoubleClick, AnimationControl13.DoubleClick, AnimationControl12.DoubleClick, AnimationControl11.DoubleClick, AnimationControl10.DoubleClick, AnimationControl1.DoubleClick
        StartProcess()
    End Sub

    Private Sub mPlay_Click(sender As Object, e As EventArgs) Handles mPlay.Click
        StartProcess()
    End Sub

    Private Sub OnlinePlayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mOnlinePlay.Click
        Dim gk, pw As String
        If mGamekey.Text.Trim <> "" Then gk = mGamekey.Text.Trim Else gk = """"""
        If mPassword.Text.Trim <> "" Then pw = mPassword.Text.Trim Else pw = """"""

        Arguments = "-netplay.nick " & mNickname.Text.Trim & " -netplay.host " & mServer.Text.Trim & " -netplay.port " & mPort.Text.Trim &
            " -netplay.gamekey " & gk & " -netplay.password " & pw & " -connect " & Arguments
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
            If CheckBox3.Checked = True Then Arguments = ""
        Catch ex As Exception
            MetroMessageBox.Show(Me, "Unable to start Mednafen", "Process Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Arguments = ""
        End Try

    End Sub

    Private Sub mGuiMode_Click(sender As System.Object, e As System.EventArgs) Handles mGuiMode.Click
        Dim parameter As String = "-file=" & Chr(34) & FileParameter & Chr(34)
        If File.Exists(Application.StartupPath & "\MedGuiR.exe") Then
            NProcess = "MedGuiR"
            KillProcess()
            Process.Start(Application.StartupPath & "\MedGuiR.exe", parameter)
        Else
            MetroMessageBox.Show(Me, "MedGui Reborn Not detected!", "MedGuiR Not detected...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
        FileParameter = ""
    End Sub

    Private Sub mInput_Click(sender As Object, e As EventArgs) Handles mInput.Click

        Dim portpad As String

        If MednafenCore = "pce" Then
            If CheckBox1.Checked = True Then
                MednafenCore += "_fast"
            End If
        ElseIf MednafenCore = "snes" Then
            If CheckBox2.Checked = True Then
                MednafenCore += "_faust"
            End If
        End If

        Select Case MednafenCore
            Case "apple2", "md", "psx", "snes_faust", "ss"
                portpad = "Virtual Port 1"
            Case "nes", "pce", "pce_fast", "pcfx", "sms", "demo"
                portpad = "Port 1"
            Case "snes"
                portpad = "Port 1/1A"
            Case Else
                portpad = "Built-In"
        End Select

        Dim parameter As String = "-folder=" & Chr(34) & MedPath & Chr(34) & " -console=" & MednafenCore & " -port=" & Chr(34) & portpad & Chr(34) & " -file=" & Chr(34) & Path.GetFileNameWithoutExtension(FileParameter) & Chr(34)
        If File.Exists(MedExtra & "\Plugins\Controller\MedPad.exe") Then
            NProcess = "MedPad"
            KillProcess()
            Process.Start(MedExtra & "\Plugins\Controller\MedPad.exe", parameter)
        Else
            MsgBox("MedPad Not detected!", vbAbort + vbExclamation, "MedPad Not detected...")
        End If
        parameter = ""
    End Sub

    Public Sub SendFolder()
        Dim parameter As String = "-folder=" & Chr(34) & MednafenModule & Chr(34)
        If File.Exists(Application.StartupPath & "\MedGuiR.exe") Then
            NProcess = "MedGuiR"
            KillProcess()
            Process.Start(Application.StartupPath & "\MedGuiR.exe", parameter)
        Else
            MetroMessageBox.Show(Me, "MedGui Reborn Not detected!", "MedGuiR Not detected...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
        MednafenModule = ""
    End Sub

    Public Sub KillProcess()

        Dim myProcesses() As Process
        Dim myProcess As Process
        myProcesses = Process.GetProcessesByName(NProcess)
        For Each myProcess In myProcesses
            Try
                myProcess.Kill()
                myProcess.WaitForExit()
            Catch
            End Try
        Next
        NProcess = ""
    End Sub

    Private Sub About_Click(sender As Object, e As EventArgs) Handles About.Click
        MetroMessageBox.Show(Me, "I'm only a fucking servant of my boss." & vbCrLf &
            "I'm  in a Beta state if you want a complete GUI call my boss MedGui Reborn", "MetroMed " & Me.Tag, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub ParseMednafenConfig()
        Dim MednafenVersion As String

        If File.Exists(MedPath & "\mednafen.cfg") Then
            MednafenVersion = "mednafen"
        ElseIf File.Exists(MedPath & "\mednafen-09x.cfg") Then
            MednafenVersion = "mednafen-09x"
        End If

        Dim row, splitrow() As String
        Try
            Using reader As New StreamReader(MedPath & "\" & MednafenVersion & ".cfg")
                While Not reader.EndOfStream
                    row = reader.ReadLine
                    splitrow = row.Split(" ")
                    Select Case True
                        Case row.Contains("netplay.host")
                            mServer.Text = splitrow(1)
                        Case row.Contains("netplay.port")
                            mPort.Text = splitrow(1)
                        Case row.Contains("netplay.nick")
                            mNickname.Text = splitrow(1)
                        Case row.Contains("netplay.password")
                            mPassword.Text = ""
                        Case row.Contains("netplay.gamekey")
                            mGamekey.Text = splitrow(1)
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

    Private Sub MetroContextMenu1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MetroContextMenu1.Opening
        MetroListServer_reload()
        ParseMednafenConfig()
        MetroToolTip1.Active = False
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            TableLayoutPanel1.Dock = DockStyle.None
            CleanAniboxart()
            TableLayoutPanel1.Visible = False
            MetroGrid1.Dock = DockStyle.Fill
            RichTextBox1.Visible = False
            AnimationControl25.Dock = DockStyle.Fill
            AnimationControl25.Visible = True
            MetroGrid1.Visible = True
            PopulateGrid()
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            MetroGrid1.Dock = DockStyle.None
            allItems.Clear()
            MetroGrid1.Rows.Clear()
            MetroGrid1.Visible = False
            AnimationControl25.Visible = False
            TableLayoutPanel1.Dock = DockStyle.Fill
            RichTextBox1.Visible = True
            RichTextBox1.Dock = DockStyle.Fill
            TableLayoutPanel1.Visible = True
        End If
    End Sub

    Private Sub CheckBox3_Click(sender As Object, e As EventArgs) Handles CheckBox3.Click
        CheckBox3.Checked = True
        CheckBox4.Checked = False
        ColorCheck()
    End Sub

    Private Sub CheckBox4_Click(sender As Object, e As EventArgs) Handles CheckBox4.Click
        CheckBox4.Checked = True
        CheckBox3.Checked = False
        ColorCheck()
    End Sub

    Private Sub MetroTextBox1_KeyUp(sender As Object, e As KeyEventArgs) Handles MetroTextBox1.KeyUp
        If MetroTextBox1.Text.Trim = "" Then Exit Sub
        Select Case e.KeyCode
            Case Keys.Enter
                If CheckBox3.Checked = True Then
                    CountRows()
                ElseIf CheckBox4.Checked = True Then
                    PopulateGrid()
                End If
        End Select
    End Sub

    Private Sub Form1_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        'CountRows()
        'If CheckBox4.Checked = True Then layoutresize()
    End Sub

    Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        ' If File.Exists(MedExtra & "Scanned\" & MednafenModule & ".csv") = False Then Exit Sub
        'If Me.WindowState = FormWindowState.Maximized Then
        'CountRows()
        '      ReadCSV()
        '   End If
        If Me.Width < 862 Then Me.Width = 862
        If Me.Height < 445 Then Me.Height = 445
        If CheckBox4.Checked = True Then layoutresize()
    End Sub

    Private Sub TableLayoutPanel1_SizeChanged(sender As Object, e As EventArgs) Handles TableLayoutPanel1.SizeChanged
        If CheckBox3.Checked = True Then
            If File.Exists(MedExtra & "Scanned\" & MednafenModule & ".csv") = False Then Exit Sub
            If Me.Width < 862 Or Me.Height < 445 Then Exit Sub
            ReadCSV()
        End If
    End Sub

    Dim oldindex As Integer = Nothing

    Private Sub AnimationControl25_SizeChanged(sender As Object, e As EventArgs) Handles AnimationControl25.SizeChanged
        layoutresize()
    End Sub

    Private Sub layoutresize()
        If CheckBox4.Checked = True Then
            If File.Exists(MedExtra & "Scanned\" & MednafenModule & ".csv") = False Then Exit Sub
            If Me.Width < 862 Or Me.Height < 445 Then Exit Sub
            extractAnitag()
            RetrieveSnap()
            If MetroGrid1.Rows.Count > 0 Then MetroGrid1.Rows(oldindex).Selected = True
        End If
    End Sub

    Private Sub mPerformance_IndexChanged(sender As Object, e As EventArgs) Handles mPerformance.SelectedIndexChanged
        Select Case mPerformance.Text
            Case "Performance"
                performance = 60
            Case "Quality"
                performance = 10
            Case "Balanced"
                performance = 30
            Case Else
                performance = 10
        End Select
    End Sub

    Private Sub mServer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles mServer.SelectedIndexChanged
        PopulateNetplay()
    End Sub

    Private Sub MetroContextMenu1_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs) Handles MetroContextMenu1.Closing
        MetroToolTip1.Active = True
    End Sub

    Private allItems As New List(Of String)()
    Private csvList As String

    Private Sub PopulateGrid()
        allItems.Clear()
        MetroGrid1.Rows.Clear()
        If File.Exists(MedExtra & "Scanned\" & MednafenModule & ".csv") Then
            csvList = MedExtra & "Scanned\" & MednafenModule & ".csv"
            ReadCsvList(MetroTextBox1.Text.Trim)
        Else
            MetroTile1.Enabled = False
            MetroTile2.Enabled = False
            MetroLabel1.Text = 0
            MetroLabel2.Text = 0
            MetroLabel3.Text = ""

            Select Case MednafenModule
                Case "fav", "pcfx", ""
                    Exit Sub
            End Select

            Dim risp As String = MetroFramework.MetroMessageBox.Show(Me, "No Prescanned files found!" & vbCrLf & "Do you want to open MedGuiR to do a prescan?",
                                                "No file " & MednafenModule & ".csv found...", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If risp = vbYes Then SendFolder() Else MednafenModule = ""
        End If
    End Sub

    Private Function ReadCsvList(GSplit) As String

        If File.Exists(csvList) Then
            Dim objReader As New StreamReader(csvList)
            Dim sLine As String = ""
            Dim arrText As New ArrayList()
            Dim FullGame() As String
            Do
                sLine = objReader.ReadLine()
                If Not sLine Is Nothing Then
                    arrText.Add(sLine)
                End If
            Loop Until sLine Is Nothing
            objReader.Close()

            For Each sLine In arrText

                FullGame = sLine.Split("|")
                If GSplit = "" Then
                    allItems.Add(sLine)

                    Dim subconsole As String = FullGame(6).ToString
                    Select Case subconsole
                        Case "sasplay"
                            subconsole = "ss"
                    End Select

                    MetroGrid1.Rows.Add(FullGame(0).ToString & " " & FullGame(2).ToString & " *** " & FullGame(8).ToString, GamesInfo.Resize(New Bitmap(MedExtra & "Resource\System\" & subconsole & ".gif"), 32, 32, True))
                Else
                    If UCase(FullGame(0).ToString).Contains(UCase(GSplit)) Then
                        allItems.Add(sLine)
                        MetroGrid1.Rows.Add(FullGame(0).ToString & " " & FullGame(2).ToString & " *** " & FullGame(8).ToString, GamesInfo.Resize(New Bitmap(MedExtra & "Resource\System\" & FullGame(6).ToString & ".gif"), 32, 32, True))
                        'Exit For
                    End If
                End If
            Next
            If MetroGrid1.Rows.Count > 0 Then
                MetroGrid1.Rows(0).Selected = True
                extractAnitag()
                RetrieveSnap()
                AddArguments()
            End If
        End If
    End Function

    Private Sub GetThumb(fullgame5, fullgame0, gif)
        Dim thumb As Bitmap = Nothing
        Dim coverSize As Long = 0

        If File.Exists(MedExtra & "BoxArt\" & fullgame5 & "\" & fullgame0 & ".png") Then
            coverSize = FileLen(MedExtra & "BoxArt\" & fullgame5 & "\" & fullgame0 & ".png")
        End If

        If coverSize > 20000 Then
            AniCover = MedExtra & "BoxArt\" & fullgame5 & "\" & fullgame0 & ".png"
        Else
            SearchScrape(fullgame5, fullgame0, gif)
        End If
    End Sub

    Private Sub extractAnitag()
        Dim Singletag() As String
        For Each sLine In allItems
            Singletag = sLine.Split("|")
            If Singletag(0).ToString & " " & Singletag(2).ToString & " *** " & Singletag(8).ToString = MetroGrid1.SelectedCells(0).Value.ToString Then
                GetThumb(Singletag(5).ToString, Singletag(0).ToString, Singletag(6).ToString())
                AniTag = sLine
                CoverEffects()
                ResizeCover(AnimationControl25)
                AnimationControl25.AnimatedImage = ReBitmap
                AnimationControl25.Animate(performance)
                Exit For
            End If
        Next
    End Sub

    Private Sub MetroGrid1_SelectionChanged(sender As Object, e As EventArgs) Handles MetroGrid1.SelectionChanged
        If CheckBox3.Checked = True Then Exit Sub
        extractAnitag()
        RetrieveSnap()
        AddArguments()
        oldindex = MetroGrid1.CurrentRow.Index
    End Sub

    Private Sub MetroGrid1_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles MetroGrid1.CellMouseDown
        If e.Button = MouseButtons.Right Then
            MetroGrid1.CurrentCell = MetroGrid1(e.ColumnIndex, e.RowIndex)
        End If
    End Sub

    Private Sub MetroGrid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles MetroGrid1.CellDoubleClick
        StartProcess()
    End Sub

    Public Function EmptyBoxart(gif As String)

        Dim subconsole As String = gif
        Select Case subconsole
            Case "sasplay"
                subconsole = "ss"
        End Select

        Dim PathEmptyBox As String = MedExtra & "Resource\Logos\" & subconsole & ".png"
        If IO.File.Exists(PathEmptyBox) Then
            AniCover = PathEmptyBox
        Else
            AniCover = MedExtra & "BoxArt\NoPr.png"
        End If
    End Function

    Private Sub MetroTextBox1_TextChanged(sender As Object, e As EventArgs) Handles MetroTextBox1.TextChanged
        If MetroTextBox1.Text.Trim = "" Then
            If CheckBox3.Checked = True Then
                CountRows()
            ElseIf CheckBox4.Checked = True Then
                PopulateGrid()
            End If
        End If
    End Sub

End Class