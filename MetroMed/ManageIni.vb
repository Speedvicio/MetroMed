Module ManageIni

    Public Sub RMetroMed()

        Try
            Dim RIni As New Ini
            Dim iTheme, iStyle As String
            MedPath = RIni.IniRead(MedExtra & "\Mini.ini", "General", "Mednafen_path")
            'Form1.GuiMode.Text = RIni.IniRead(MedExtra & "\Mini.ini", "General", "GUi_Mode")
            iTheme = Val(RIni.IniRead(MedExtra & "\Mini.ini", "MetroMed", "Theme"))
            iStyle = Val(RIni.IniRead(MedExtra & "\Mini.ini", "MetroMed", "Style"))
            If iTheme = "" Then iTheme = 0
            MetroMed.cmbTheme.SelectedIndex = iTheme
            If iStyle = "" Then iStyle = 0
            MetroMed.cmbStyle.SelectedIndex = iStyle
            MetroMed.mEffect.Text = RIni.IniRead(MedExtra & "\Mini.ini", "MetroMed", "Effect")
            MetroMed.mPerformance.Text = RIni.IniRead(MedExtra & "\Mini.ini", "MetroMed", "Performance")
            MetroMed.CheckBox1.CheckState = RIni.IniRead(MedExtra & "\Mini.ini", "General", "Fast_PCE")
            MetroMed.CheckBox2.CheckState = RIni.IniRead(MedExtra & "\Mini.ini", "General", "SNES_Faust")
            MednafenModule = RIni.IniRead(MedExtra & "\Mini.ini", "General", "Startup_Path")
            'CurrPage = RIni.IniRead(MedExtra & "\Mini.ini", "MetroMed", "Current_Page")
            If MetroMed.mPerformance.Text = "" Then MetroMed.mPerformance.Text = "Balanced" : performance = 30
            Dim iLayout As Integer
            iLayout = Val(RIni.IniRead(MedExtra & "\Mini.ini", "MetroMed", "Layout"))
            If iLayout = 1 Then
                MetroMed.CheckBox3.Checked = True
                MetroMed.CheckBox4.Checked = False
            Else
                MetroMed.CheckBox3.Checked = False
                MetroMed.CheckBox4.Checked = True
            End If
        Catch 'ex As Exception
            'MGRWriteLog("ManageIni - DirectoryRMini: " & ex.Message)
            'Finally
        End Try

    End Sub

    Public Sub WMetroMed()

        Dim WIni As New Ini
        'WIni.IniWrite(MedExtra & "\Mini.ini", "General", "GUi_Mode", Form1.GuiMode.Text)
        WIni.IniWrite(MedExtra & "\Mini.ini", "MetroMed", "Theme", MetroMed.cmbTheme.SelectedIndex)
        WIni.IniWrite(MedExtra & "\Mini.ini", "MetroMed", "Style", MetroMed.cmbStyle.SelectedIndex)
        WIni.IniWrite(MedExtra & "\Mini.ini", "MetroMed", "Effect", MetroMed.mEffect.Text)
        WIni.IniWrite(MedExtra & "\Mini.ini", "MetroMed", "Performance", MetroMed.mPerformance.Text)
        WIni.IniWrite(MedExtra & "\Mini.ini", "MetroMed", "Layout", MetroMed.CheckBox3.CheckState)
        'WIni.IniWrite(MedExtra & "\Mini.ini", "MetroMed", "Current_Page", MetroLabel1.Text)
        WIni.IniWrite(MedExtra & "\Mini.ini", "General", "Fast_PCE", MetroMed.CheckBox1.CheckState)
        WIni.IniWrite(MedExtra & "\Mini.ini", "General", "SNES_Faust", MetroMed.CheckBox2.CheckState)
        WIni.IniWrite(MedExtra & "\Mini.ini", "General", "Startup_Path", MednafenModule)

    End Sub

End Module