Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports ASCOM.Utilities
Imports ASCOM.ES_PMC8

<ComVisible(False)> _
Public Class SetupDialogForm

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click ' OK button event handler
        ' Persist new values of user settings to the ASCOM profile
        Telescope.comPort = ComboBox1.Text ' Update the state variables with results from the dialogue
        Telescope.comSpeed = ComboBox2.Text
        Telescope.traceState = chkTrace.Checked
        Telescope.IPAddress = TextBox1.Text
        Telescope.IPPort = TextBox2.Text
        Telescope.WirelessEnabled = RadioButton3.Checked
        Telescope.WiFiModuleID = cboWiFiModule.Text
        If RadioButton3.Checked Then
            Telescope.WirelessProtocol = "TCP"
        End If

        Telescope.Mount = ComboBox3.Text
        If Telescope.Mount = "Losmandy G-11" Then
            Telescope.MountRACounts = 4608000
            Telescope.MountDECCounts = 4608000
        ElseIf Telescope.Mount = "Losmandy Titan" Then
            Telescope.MountRACounts = 3456000
            Telescope.MountDECCounts = 3456000
        ElseIf Telescope.Mount = "Explore Scientific EXOS II" Then
            Telescope.MountRACounts = 4147200
            Telescope.MountDECCounts = 4147200
        ElseIf Telescope.Mount = "Explore Scientific iEQ100" Then
            Telescope.MountRACounts = 4147200
            Telescope.MountDECCounts = 4147200
        ElseIf Telescope.Mount = "Explore Scientific iEQ300" Then
            Telescope.MountRACounts = 4147200
            Telescope.MountDECCounts = 3456000
        End If

        Telescope.Rate = ComboBox4.Text
        If Telescope.Rate = "Sidereal" Then

        ElseIf Telescope.Rate = "Lunar" Then

        ElseIf Telescope.Rate = "Solar" Then

        ElseIf Telescope.Rate = "King" Then

        End If

        Telescope.SiteLocation = tbSiteLocation.Text
        Telescope.SiteLatitudeValue = tbSiteLatitude.Text
        Telescope.SiteLongitudeValue = tbSiteLongitude.Text
        Telescope.SiteElevationValue = tbSiteElevation.Text
        Telescope.ApertureDiameterValue = tbApertureDiameter.Text
        Telescope.ApertureAreaValue = tbApertureArea.Text
        Telescope.FocalLengthValue = tbFocalLength.Text
        Telescope.RateOffsetValue = tbRateOffset.Text
        Telescope.SiteAmbientTemperatureValue = Convert.ToDouble(tbAmbientTemp.Text)
        Telescope.ApplyRefractionCorrection = cbRefraction.CheckState
        Telescope.RA_SiderealRateFraction = nud_RA.Value
        Telescope.DEC_SiderealRateFraction = nud_DEC.Value
        Telescope.MinimumPulseTime = nud_PulseTime.Value

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click 'Cancel button event handler
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ShowAscomWebPage(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.DoubleClick, PictureBox1.Click
        ' Click on ASCOM logo event handler
        Try
            System.Diagnostics.Process.Start("http://ascom-standards.org/")
        Catch noBrowser As System.ComponentModel.Win32Exception
            If noBrowser.ErrorCode = -2147467259 Then
                MessageBox.Show(noBrowser.Message)
            End If
        Catch other As System.Exception
            MessageBox.Show(other.Message)
        End Try
    End Sub

    Private Sub ShowAscomWebPage2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.DoubleClick, PictureBox2.Click
        ' Click on Explore Scientific logo event handler
        Try
            System.Diagnostics.Process.Start("http://explorescientificusa.com/")
        Catch noBrowser As System.ComponentModel.Win32Exception
            If noBrowser.ErrorCode = -2147467259 Then
                MessageBox.Show(noBrowser.Message)
            End If
        Catch other As System.Exception
            MessageBox.Show(other.Message)
        End Try
    End Sub

    Private Sub SetupDialogForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load ' Form load event handler
        ' Retrieve current values of user settings from the ASCOM Profile 
        Using objSerial As New ASCOM.Utilities.Serial
            For Each item In objSerial.AvailableCOMPorts
                ComboBox1.Items.Add(item)
            Next
        End Using
        chkTrace.Checked = Telescope.traceState
        ComboBox1.Text = Telescope.comPort
        ComboBox2.Text = Telescope.comSpeed
        TextBox1.Text = Telescope.IPAddress
        TextBox2.Text = Telescope.IPPort
        RadioButton1.Checked = Not Telescope.WirelessEnabled
        If Telescope.WirelessProtocol = "TCP" Then
            RadioButton3.Checked = Telescope.WirelessEnabled
        End If
        RadioButton1.Checked = Not Telescope.WirelessEnabled
        ComboBox3.Text = Telescope.Mount
        ComboBox4.Text = Telescope.Rate
        tbSiteLocation.Text = Telescope.SiteLocation
        tbSiteElevation.Text = Telescope.SiteElevationValue.ToString
        tbSiteLatitude.Text = Telescope.SiteLatitudeValue.ToString
        tbSiteLongitude.Text = Telescope.SiteLongitudeValue.ToString
        tbApertureDiameter.Text = Telescope.ApertureDiameterValue.ToString
        tbApertureArea.Text = Telescope.ApertureAreaValue.ToString
        tbFocalLength.Text = Telescope.FocalLengthValue.ToString
        tbRateOffset.Text = Telescope.RateOffsetValue.ToString
        tbAmbientTemp.Text = Telescope.SiteAmbientTemperatureValue.ToString
        cbRefraction.Checked = Telescope.ApplyRefractionCorrection
        nud_RA.Value = Telescope.RA_SiderealRateFraction
        nud_DEC.Value = Telescope.DEC_SiderealRateFraction
        nud_PulseTime.Value = Telescope.MinimumPulseTime
        cboWiFiModule.Text = Telescope.WiFiModuleID

        ' Create the ToolTip and associate with the Form container.
        Dim toolTip1 As New ToolTip()

        ' Set up the delays for the ToolTip.
        toolTip1.AutoPopDelay = 5000
        toolTip1.InitialDelay = 1000
        toolTip1.ReshowDelay = 500
        ' Force the ToolTip text to be displayed whether or not the form is active.
        toolTip1.ShowAlways = True

        ' Set up the ToolTip text for the Button and Checkbox.
        toolTip1.SetToolTip(Me.OK_Button, "Save the changes")
        toolTip1.SetToolTip(Me.Cancel_Button, "Cancel the changes")
        toolTip1.SetToolTip(Me.PictureBox2, "Explore Scientific Website")
        toolTip1.SetToolTip(Me.PictureBox1, "ASCOM Standards Website")
        toolTip1.SetToolTip(Me.ComboBox1, "Select Serial Port Number")
        toolTip1.SetToolTip(Me.ComboBox2, "Select Serial Port Speed")
        toolTip1.SetToolTip(Me.ComboBox3, "Select Mount Type")
        toolTip1.SetToolTip(Me.ComboBox4, "Select Mount Rate")
        toolTip1.SetToolTip(Me.tbSiteLocation, "Set City Name")
        toolTip1.SetToolTip(Me.tbSiteLatitude, "Latitude + for North, - for South")
        toolTip1.SetToolTip(Me.tbSiteLongitude, "Longitude - for West, + for EAST")
        toolTip1.SetToolTip(Me.tbApertureDiameter, "Telescope Aperture Diameter (meters)")
        toolTip1.SetToolTip(Me.tbApertureArea, "Telescope Aperture Area (meter^2)")
        toolTip1.SetToolTip(Me.tbFocalLength, "Telescope Focal Length (meters)")
        toolTip1.SetToolTip(Me.nud_RA, "Pulse Guiding Sidereal Rate Fraction for RA Axis")
        toolTip1.SetToolTip(Me.nud_DEC, "Pulse Guiding Sidereal Rate Fraction for DEC Axis")
        toolTip1.SetToolTip(Me.nud_PulseTime, "Minimum Pulse Guiding Pulse Time Interval milli-seconds")
        toolTip1.SetToolTip(Me.cboWiFiModule, "WiFi Module installed on PMC-Eight")

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Dim frmAbout As New AboutBox1
        frmAbout.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim frmAbout As New AboutBox1
        frmAbout.Show()
    End Sub

    Private Sub btn_ST4Calibration_Click(sender As Object, e As EventArgs) Handles btn_ST4Calibration.Click
        Dim frmST4Calibration As New AboutBox1
    End Sub

End Class
