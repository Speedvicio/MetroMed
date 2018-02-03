Imports System.Drawing
Imports System.IO

Module GamesInfo

    Public Sub ReadXml()
        Dim TGDBXml, BaseUrl, fBack, fFront As String
        TGDBXml = MedExtra & "Scraped\" & TagSplit(5) & "\" & Trim(TagSplit(0)) & ".xml"

        MetroMed.RichTextBox1.Text = Nothing
        MetroMed.PictureBox1.Image = Nothing
        MetroMed.PictureBox2.Image = Nothing

        If File.Exists(TGDBXml) = False Then Exit Sub

        Dim reader As New System.Xml.XmlTextReader(TGDBXml)
        While reader.Read()
            Dim contents As String
            reader.MoveToContent()

            If reader.NodeType = Xml.XmlNodeType.Element Then
                contents = reader.Name
            End If

            If reader.NodeType = Xml.XmlNodeType.Text Then
                Select Case contents
                    Case "baseImgUrl"
                        BaseUrl = reader.Value
                    Case "GameTitle"
                        'TheGamesDB.Label1.Text = "Game Title: " & Replace(reader.Value, "&", "&&")
                    Case "Platform"
                        'TheGamesDB.Label2.Text = "Platform: " & (reader.Value)
                    Case "ReleaseDate"
                        Dim fdate As String
                        If Len(reader.Value) = 10 Then fdate = reader.Value Else fdate = "0" & reader.Value
                        If Len(reader.Value) = 4 Then
                            'TheGamesDB.Label3.Text = "Release Date: " & (reader.Value)
                        Else
                            Dim expenddt As Date = Date.ParseExact(fdate, "MM/dd/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo).ToString
                            'TheGamesDB.Label3.Text = "Release Date: " & (expenddt)
                        End If
                    Case "Overview"
                        MetroMed.RichTextBox1.Text = (reader.Value).ToString
                    Case "genre"
                        'If Len(TheGamesDB.Label4.Text) <= 7 Then
                        'TheGamesDB.Label4.Text = "Genre: " & (reader.Value)
                        'Else
                        'TheGamesDB.Label4.Text = TheGamesDB.Label4.Text & " - " & (reader.Value)
                        'End If
                    Case "Players"
                        'TheGamesDB.Label11.Text = "Players: " & (reader.Value)
                    Case "Publisher"
                        'TheGamesDB.Label5.Text = "Publisher: " & (reader.Value)
                    Case "Developer"
                        'TheGamesDB.Label6.Text = "Developer: " & (reader.Value)
                    Case "Co-op"
                        'TheGamesDB.Label7.Text = "Co-op: " & (reader.Value)
                    Case "boxart"

                        If reader.Value.Contains("boxart/original/back/") Then
                            fBack = reader.Value
                        ElseIf reader.Value.Contains("boxart/original/front/") Then
                            fFront = reader.Value
                        End If

                End Select

                contents = ""
            End If

            'If File.Exists(MedExtra & "BoxArt\" & MedGuiR.DataGridView1.CurrentRow.Cells(5).Value() & "\" & rn & ".png") = False Then MedGuiR.PictureBox1.Load(SBoxF) : pathimage = SBoxF

        End While
        reader.Close()

    End Sub

    Public Function cleanpsx(ByVal cleanstring As String) As String
        Dim i1, i2 As Integer

        i1 = cleanstring.IndexOf("[")
        i2 = cleanstring.IndexOf("]")
        While i1 >= 0 And i2 >= 0
            cleanstring = cleanstring.Remove(i1, i2 - i1 + 1)
            i1 = cleanstring.IndexOf("[")
            i2 = cleanstring.IndexOf("]")
        End While
        cleanpsx = cleanstring
    End Function

    Public Function RemoveAmpersand(ByVal CleanAmp As String) As String
        RemoveAmpersand = Replace(CleanAmp, " &amp; ", " && ")
        RemoveAmpersand = Replace(RemoveAmpersand, " & ", " && ")
    End Function

    Public Function Resize(ByVal Immagine As Bitmap,
                                        ByVal N_Dim_X As Integer,
                                        ByVal N_Dim_Y As Integer,
                                        Optional ByVal preserveAspectRatio As Boolean = True) As Bitmap
        If preserveAspectRatio Then
            Dim O_Dim_X As Integer = Immagine.Width
            Dim O_Dim_Y As Integer = Immagine.Height
            Dim Percent_X As Single = CSng(N_Dim_X) / CSng(O_Dim_X)
            Dim Percent_Y As Single = CSng(N_Dim_Y) / CSng(O_Dim_Y)
            Dim Percent As Single = If(Percent_Y < Percent_X, Percent_Y, Percent_X)
            N_Dim_X = CInt(O_Dim_X * Percent)
            N_Dim_Y = CInt(O_Dim_Y * Percent)
        End If
        Dim bmp As Image = New Bitmap(N_Dim_X, N_Dim_Y)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.InterpolationMode = Drawing2D.InterpolationMode.Bicubic
            g.DrawImage(Immagine, 0, 0, N_Dim_X, N_Dim_Y)
        End Using
        Immagine.Dispose()
        Return bmp
    End Function
End Module
