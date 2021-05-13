Imports System.Drawing
Imports System.IO

Module Preload

    Public MedPath, MedExtra, MednafenModule, AniTag, AniCover, Effect As String, AniBoxArt(23) As AnimationControl,
        TotalRows, CurrPage, anyindex, NextPage, BoxResult, performance As Integer, TagSplit() As String, ReBitmap As Bitmap

    Public Sub LoadResource()

        MedExtra = Application.StartupPath & "\MedGuiR\"

        MetroMed.MetroButton1.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\gb.gif"))
        MetroMed.MetroButton2.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\gba.gif"))
        MetroMed.MetroButton3.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\gg.gif"))
        MetroMed.MetroButton4.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\lynx.gif"))
        MetroMed.MetroButton5.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\md.gif"))
        MetroMed.MetroButton6.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\nes.gif"))
        MetroMed.MetroButton7.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\ngp.gif"))
        MetroMed.MetroButton8.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\pce.gif"))
        MetroMed.MetroButton9.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\psx.gif"))
        MetroMed.MetroButton10.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\ss.gif"))
        MetroMed.MetroButton11.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\sms.gif"))
        MetroMed.MetroButton12.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\snes.gif"))
        MetroMed.MetroButton13.BackgroundImage = (New Bitmap(MedExtra & "Resource\Gui\dtl.png"))
        MetroMed.MetroButton14.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\vb.gif"))
        MetroMed.MetroButton15.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\wswan.gif"))
        MetroMed.MetroButton16.BackgroundImage = (New Bitmap(MedExtra & "Resource\Gui\folder.png"))
        MetroMed.MetroButton17.BackgroundImage = (New Bitmap(MedExtra & "Resource\Gui\favourite.png"))
        MetroMed.MetroButton18.BackgroundImage = (New Bitmap(MedExtra & "Resource\System\apple2.gif"))
        MetroMed.CheckBox1.BackgroundImage = (New Bitmap(MedExtra & "Resource\Gui\fast.png"))
        MetroMed.CheckBox2.BackgroundImage = (New Bitmap(MedExtra & "Resource\Gui\faust.png"))
        MetroMed.CheckBox4.BackgroundImage = (New Bitmap(MedExtra & "Resource\Gui\list.png"))
        MetroMed.CheckBox3.BackgroundImage = (New Bitmap(MedExtra & "Resource\Gui\grid.png"))
        MetroMed.About.Image = (New Bitmap(MedExtra & "Resource\Gui\info.png"))
        MetroMed.mPlay.Image = (New Bitmap(MedExtra & "Resource\Gui\play.png"))
        MetroMed.mNetPlay.Image = (New Bitmap(MedExtra & "Resource\Gui\netplay.png"))
        MetroMed.mCover.Image = (New Bitmap(MedExtra & "Resource\Gui\cover.png"))
        MetroMed.mGuiMode.Image = (New Bitmap(MedExtra & "Resource\Gui\MedGuiR.png"))
        MetroMed.mServer.Image = (New Bitmap(MedExtra & "Resource\Gui\net.png"))
        MetroMed.mOnlinePlay.Image = (New Bitmap(MedExtra & "Resource\Gui\net.png"))
        MetroMed.mGuiTheme.Image = (New Bitmap(MedExtra & "Resource\Gui\palette.png"))
        MetroMed.mInput.Image = (New Bitmap(MedExtra & "Resource\Gui\medpad.ico"))
        MetroMed.MetroTile1.TileImage = (New Bitmap(MedExtra & "Resource\Gui\play.png"))
        MetroMed.MetroTile1.TileImage.RotateFlip(RotateFlipType.Rotate180FlipNone)
        MetroMed.MetroTile2.TileImage = (New Bitmap(MedExtra & "Resource\Gui\play.png"))
        'SplitContainer2.Panel1.BackgroundImage = (New Bitmap(MedExtra & "Resource\Gui\info.png"))

    End Sub

    Public Sub CountRows()
        Dim crow As Integer = 0
        If File.Exists(MedExtra & "Scanned\" & MednafenModule & ".csv") Then
            Dim readText() As String = File.ReadAllLines(MedExtra & "Scanned\" & MednafenModule & ".csv")
            If MetroMed.MetroTextBox1.Text.Trim <> "" Then
                Dim s As String
                Dim sp() As String

                For Each s In readText
                    sp = Split(s, "|")
                    If LCase(sp(0)).Trim.Contains(LCase(MetroMed.MetroTextBox1.Text)) Then
                        Exit For
                    End If
                    crow = crow + 1
                Next
            End If
            'TotalRows = File.ReadAllLines(MedExtra & "Scanned\" & MednafenModule & ".csv").Length

            TotalRows = readText.Length
            MetroMed.MetroLabel2.Text = (Math.Ceiling(TotalRows / 24))
            MetroMed.MetroLabel1.Text = 1
            If crow = 0 Then
                CurrPage = 0
            Else
                CurrPage = crow
                NextPage = CurrPage + 24
                MetroMed.MetroLabel1.Text = (Math.Ceiling(crow / 24))
            End If
            ReadCSV()
            MetroMed.MetroTile1.Enabled = True
            MetroMed.MetroTile2.Enabled = True
        Else
            MetroMed.MetroTile1.Enabled = False
            MetroMed.MetroTile2.Enabled = False
            MetroMed.MetroLabel1.Text = 0
            MetroMed.MetroLabel2.Text = 0
            MetroMed.MetroLabel3.Text = ""
            CleanAniboxart()

            Select Case MednafenModule
                Case "fav", "pcfx", ""
                    Exit Sub
            End Select

            Dim risp As String = MetroFramework.MetroMessageBox.Show(MetroMed, "No Prescanned files found!" & vbCrLf & "Do you want to open MedGuiR to do a prescan?",
                                                "No file " & MednafenModule & ".csv found...", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If risp = vbYes Then MetroMed.SendFolder() Else MednafenModule = ""
        End If
    End Sub

    Public Sub NextSearch()
        Dim readText() As String = File.ReadAllLines(MedExtra & "Scanned\" & MednafenModule & ".csv")
        Dim sp() As String
        Dim crow As Integer = 0
        For index As Integer = 0 To readText.Length - 1
            sp = Split(readText(index), "|")
            If index > CurrPage + BoxResult Then
                If LCase(sp(0)).Trim.Contains(LCase(MetroMed.MetroTextBox1.Text)) Then
                    Exit For
                End If
                crow = index
            End If
        Next
        BoxResult = 0
        CurrPage = crow
        NextPage = CurrPage + 24
        MetroMed.MetroLabel1.Text = (Math.Ceiling(crow / 24))
    End Sub

    Public Sub CleanAniboxart()

        MetroMed.PictureBox1.Image = Nothing
        MetroMed.PictureBox2.Image = Nothing
        MetroMed.MetroLabel3.Text = ""
        MetroMed.RichTextBox1.Clear()

        For i = 0 To 23
            AniBoxArt(i).Tag = Nothing
            AniBoxArt(i).AnimatedImage = New Bitmap(1, 1)
            AniBoxArt(i).Animate(60)
            MetroMed.MetroToolTip1.SetToolTip(AniBoxArt(i), Nothing)
            AniBoxArt(i).Refresh()
            AniBoxArt(i).AnimationSpeed = 0.7
        Next
    End Sub

    Public Sub ReadCSV()
        If MetroMed.MetroTextBox1.Text.Trim <> "" Then
            MetroMed.MetroTile1.Enabled = False
        Else
            MetroMed.MetroTile1.Enabled = True
        End If

        'LoadCover()
        Dim oRead As String
        Dim oSplit() As String

        CleanAniboxart()
        'Threading.Thread.Sleep(100)

        If CurrPage = 0 Then
            CurrPage = 0
            NextPage = 23
        Else
        End If

        If File.Exists(MedExtra & "Scanned\" & MednafenModule & ".csv") = False Then Exit Sub

        oRead = File.ReadAllText(MedExtra & "Scanned\" & MednafenModule & ".csv")
        oSplit = Split(oRead, (vbCrLf))
        anyindex = 0

        For i = CurrPage To NextPage
            MetroMed.TimerResize.Enabled = False
            Try
Re_Try:
                If i >= TotalRows Then Exit Sub
                AniTag = oSplit(i)
                SplitTag()

                If MetroMed.MetroTextBox1.Text.Trim <> "" Then
                    If cleanpsx(RemoveAmpersand(LCase(TagSplit(0)))).Trim.Contains(LCase(MetroMed.MetroTextBox1.Text.Trim)) = False Then
                        BoxResult = BoxResult + 1
                        Continue For 'i = i + 1 : : anyindex = anyindex + 1 GoTo Re_Try
                    End If

                End If

                AniBoxArt(anyindex).Tag = AniTag & "|" & anyindex

                MetroMed.MetroToolTip1.SetToolTip(AniBoxArt(anyindex), "     " & cleanpsx(RemoveAmpersand(TagSplit(0))) & " " & TagSplit(2) & "     " & vbCrLf &
                        TagSplit(5).Trim)
                AniCover = MedExtra & "BoxArt\" & TagSplit(5) & "\" & TagSplit(0) & ".png"

                If File.Exists(AniCover) = False Then
                    SearchScrape(TagSplit(5), TagSplit(0), TagSplit(6))
                End If
                ResizeCover(AniBoxArt(anyindex))
                'AniBoxArt(anyindex).AnimatedImage = GamesInfo.Resize(New Bitmap(AniCover), AniBoxArt(anyindex).Width, AniBoxArt(anyindex).Height, False)
                AniBoxArt(anyindex).AnimatedImage = ReBitmap
                AniBoxArt(anyindex).Animate(performance)
                anyindex = anyindex + 1
            Catch
                'AniCover = MedExtra & "BoxArt\NoPr.png"
                MetroMed.EmptyBoxart(TagSplit(6))
                ResizeCover(AniBoxArt(anyindex))
                If anyindex = 24 Then anyindex = 23
                AniBoxArt(anyindex).AnimatedImage = ReBitmap
                AniBoxArt(anyindex).Animate(performance)
                If i < NextPage Then i = i + 1 : anyindex = anyindex + 1 : GoTo Re_Try
            End Try
            MetroMed.TimerResize.Enabled = True
        Next (i)

    End Sub

    Public Sub ResizeCover(animation As AnimationControl)
        ReBitmap = New Bitmap(animation.Width, animation.Height)
        Dim Graphics As Graphics = Graphics.FromImage(ReBitmap)
        Dim inputImage As Image = GamesInfo.Resize(New Bitmap(AniCover), animation.Width - 5, animation.Height - 5, True)
        Graphics.DrawImage(inputImage, New Point((animation.Width - inputImage.Width) / 2, (animation.Height - inputImage.Height) / 2))
    End Sub

    Public Sub SearchScrape(console, gname, gif)
        'AniCover = MedExtra & "BoxArt\NoPr.png"
        MetroMed.EmptyBoxart(gif)
        If Directory.Exists(MedExtra & "Scraped\" & console & "\" & gname) Then
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(MedExtra & "Scraped\" & console & "\" & gname)
                If foundFile.Contains("tfront") Then
                    AniCover = foundFile
                End If
            Next
        End If
        'If AniCover = "" Then AniCover = MedExtra & "BoxArt\NoPr.png"
    End Sub

    Public Sub SplitTag()
        TagSplit = AniTag.Split("|")
    End Sub

    Public Sub CoverEffects()
        Dim Anindex As Integer = Nothing

        Select Case Effect
                ' sliding effect
            Case "LeftToRight"
                Anindex = AnimationTypes.LeftToRight
            Case "RighTotLeft"
                Anindex = AnimationTypes.RighTotLeft
            Case "TopToDown"
                Anindex = AnimationTypes.TopToDown
            Case "DownToTop"
                Anindex = AnimationTypes.DownToTop
            Case "TopLeftToBottomRight"
                Anindex = AnimationTypes.TopLeftToBottomRight
            Case "BottomRightToTopLeft"
                Anindex = AnimationTypes.BottomRightToTopLeft
            Case "BottomLeftToTopRight"
                Anindex = AnimationTypes.BottomLeftToTopRight
            Case "TopRightToBottomLeft"
                Anindex = AnimationTypes.TopRightToBottomLeft
                    ' rotating effect
            Case "Maximize"
                Anindex = AnimationTypes.Maximize
            Case "Rotate"
                Anindex = AnimationTypes.Rotate
            Case "Spin"
                Anindex = AnimationTypes.Spin
                    ' shape effect
            Case "Circular"
                Anindex = AnimationTypes.Circular
            Case "Elliptical"
                Anindex = AnimationTypes.Elliptical
            Case "Rectangular"
                Anindex = AnimationTypes.Rectangular
                    ' split effect
            Case "SplitHorizontal"
                Anindex = AnimationTypes.SplitHorizontal
            Case "SplitVertical"
                Anindex = AnimationTypes.SplitVertical
            Case "SplitBoom"
                Anindex = AnimationTypes.SplitBoom
            Case "SplitQuarter"
                Anindex = AnimationTypes.SplitQuarter
                    ' chess effect
            Case "ChessBoard"
                Anindex = AnimationTypes.ChessBoard
            Case "ChessHorizontal"
                Anindex = AnimationTypes.ChessHorizontal
            Case "ChessVertical"
                Anindex = AnimationTypes.ChessVertical
                    ' panorama effect
            Case "Panorama"
                Anindex = AnimationTypes.Panorama
            Case "PanoramaHorizontal"
                Anindex = AnimationTypes.PanoramaHorizontal
            Case "PanoramaVertical"
                Anindex = AnimationTypes.PanoramaVertical
                    ' spiral effect
            Case "Spiral"
                Anindex = AnimationTypes.Spiral
            Case "SpiralBoom"
                Anindex = AnimationTypes.SpiralBoom
                    ' fade effect
            Case "Fade"
                Anindex = AnimationTypes.Fade
            Case "Fade2Images"
                Anindex = AnimationTypes.Fade2Images
        End Select

        If MetroMed.CheckBox3.Checked = True Then
            For i = 0 To 23
                AniBoxArt(i).AnimationType = Anindex
            Next i
        ElseIf MetroMed.CheckBox4.Checked = True Then
            MetroMed.AnimationControl25.AnimationType = Anindex
        End If
    End Sub

End Module