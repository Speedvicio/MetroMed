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
                    SearchScrape()
                End If
                ResizeCover()
                'AniBoxArt(anyindex).AnimatedImage = GamesInfo.Resize(New Bitmap(AniCover), AniBoxArt(anyindex).Width, AniBoxArt(anyindex).Height, False)
                AniBoxArt(anyindex).AnimatedImage = ReBitmap
                AniBoxArt(anyindex).Animate(performance)
                anyindex = anyindex + 1
            Catch
                AniCover = MedExtra & "BoxArt\NoPr.png"
                ResizeCover()
                If anyindex = 24 Then anyindex = 23
                AniBoxArt(anyindex).AnimatedImage = ReBitmap
                AniBoxArt(anyindex).Animate(performance)
                If i < NextPage Then i = i + 1 : anyindex = anyindex + 1 : GoTo Re_Try
            End Try
            MetroMed.TimerResize.Enabled = True
        Next (i)

    End Sub

    Public Sub ResizeCover()
        ReBitmap = New Bitmap(AniBoxArt(anyindex).Width, AniBoxArt(anyindex).Height)
        Dim Graphics As Graphics = Graphics.FromImage(ReBitmap)
        Dim inputImage As Image = GamesInfo.Resize(New Bitmap(AniCover), AniBoxArt(anyindex).Width - 5, AniBoxArt(anyindex).Height - 5, True)
        Graphics.DrawImage(inputImage, New Point((AniBoxArt(anyindex).Width - inputImage.Width) / 2, (AniBoxArt(anyindex).Height - inputImage.Height) / 2))
    End Sub

    Public Sub SearchScrape()
        AniCover = MedExtra & "BoxArt\NoPr.png"
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(MedExtra & "Scraped\" & TagSplit(5) & "\" & TagSplit(0))
            If foundFile.Contains("tfront") Then AniCover = foundFile
        Next
        'If AniCover = "" Then AniCover = MedExtra & "BoxArt\NoPr.png"
    End Sub

    Public Sub SplitTag()
        TagSplit = AniTag.Split("|")
    End Sub

    Public Sub CoverEffects()

        For i = 0 To 23
            Select Case Effect
                ' sliding effect
                Case "LeftToRight"
                    AniBoxArt(i).AnimationType = AnimationTypes.LeftToRight
                Case "RighTotLeft"
                    AniBoxArt(i).AnimationType = AnimationTypes.RighTotLeft
                Case "TopToDown"
                    AniBoxArt(i).AnimationType = AnimationTypes.TopToDown
                Case "DownToTop"
                    AniBoxArt(i).AnimationType = AnimationTypes.DownToTop
                Case "TopLeftToBottomRight"
                    AniBoxArt(i).AnimationType = AnimationTypes.TopLeftToBottomRight
                Case "BottomRightToTopLeft"
                    AniBoxArt(i).AnimationType = AnimationTypes.BottomRightToTopLeft
                Case "BottomLeftToTopRight"
                    AniBoxArt(i).AnimationType = AnimationTypes.BottomLeftToTopRight
                Case "TopRightToBottomLeft"
                    AniBoxArt(i).AnimationType = AnimationTypes.TopRightToBottomLeft
                    ' rotating effect
                Case "Maximize"
                    AniBoxArt(i).AnimationType = AnimationTypes.Maximize
                Case "Rotate"
                    AniBoxArt(i).AnimationType = AnimationTypes.Rotate
                Case "Spin"
                    AniBoxArt(i).AnimationType = AnimationTypes.Spin
                    ' shape effect
                Case "Circular"
                    AniBoxArt(i).AnimationType = AnimationTypes.Circular
                Case "Elliptical"
                    AniBoxArt(i).AnimationType = AnimationTypes.Elliptical
                Case "Rectangular"
                    AniBoxArt(i).AnimationType = AnimationTypes.Rectangular
                    ' split effect
                Case "SplitHorizontal"
                    AniBoxArt(i).AnimationType = AnimationTypes.SplitHorizontal
                Case "SplitVertical"
                    AniBoxArt(i).AnimationType = AnimationTypes.SplitVertical
                Case "SplitBoom"
                    AniBoxArt(i).AnimationType = AnimationTypes.SplitBoom
                Case "SplitQuarter"
                    AniBoxArt(i).AnimationType = AnimationTypes.SplitQuarter
                    ' chess effect
                Case "ChessBoard"
                    AniBoxArt(i).AnimationType = AnimationTypes.ChessBoard
                Case "ChessHorizontal"
                    AniBoxArt(i).AnimationType = AnimationTypes.ChessHorizontal
                Case "ChessVertical"
                    AniBoxArt(i).AnimationType = AnimationTypes.ChessVertical
                    ' panorama effect
                Case "Panorama"
                    AniBoxArt(i).AnimationType = AnimationTypes.Panorama
                Case "PanoramaHorizontal"
                    AniBoxArt(i).AnimationType = AnimationTypes.PanoramaHorizontal
                Case "PanoramaVertical"
                    AniBoxArt(i).AnimationType = AnimationTypes.PanoramaVertical
                    ' spiral effect
                Case "Spiral"
                    AniBoxArt(i).AnimationType = AnimationTypes.Spiral
                Case "SpiralBoom"
                    AniBoxArt(i).AnimationType = AnimationTypes.SpiralBoom
                    ' fade effect
                Case "Fade"
                    AniBoxArt(i).AnimationType = AnimationTypes.Fade
                Case "Fade2Images"
                    AniBoxArt(i).AnimationType = AnimationTypes.Fade2Images
            End Select
        Next i
    End Sub

End Module