Public Class Site

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If OverClass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        OverClass.ResetCollection()

        Select Case e.TabPage.Text

            Case "QC"
                StartCombo(Me.ComboBox100)

            Case "ARF Defaults"
                Form1.Specifics(Me.DataGridView101)

            Case "ARF"
                Form1.Specifics(Me.DataGridView102)

            Case "All QC'd Results"
                Form1.Specifics(Me.DataGridView103)


        End Select


        If Not IsNothing(ctl) Then Call Form1.Specifics(ctl)

    End Sub

    Private Sub Site_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Call StartUp(SiteForm)

        Try
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Catch
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside"
        End Try

        Me.Text = SolutionName

    End Sub


    Private Sub DataGridView101_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView101.CellBeginEdit

        If e.ColumnIndex <> sender.columns("DefaultTest").index Then Exit Sub

        If Role <> "Site_Lead" Then
            e.Cancel = True
            MsgBox("Can only be changed by site lead")
        End If


    End Sub

    Private Sub DataGridView102_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView102.CellContentClick

        If e.ColumnIndex = Me.DataGridView102.Columns("Volunteer").Index Then

            AppID = Me.DataGridView102.Item("APP_ID", e.RowIndex).Value
            Dim Volunteer As String = Me.DataGridView102.Item("Volunteer", e.RowIndex).Value
            Dim OK As New PickTests
            OverClass.CreateDataSet("SELECT Result_ID, Virus_ID FROM tblApp_Results " &
                    "WHERE APP_ID=" & AppID, OK.BindingSource1, OK.DataGridView200)
            OK.DataGridView200.Columns("Virus_ID").Visible = False
            OK.DataGridView200.Columns("Result_ID").Visible = False
            Dim cmb As New DataGridViewComboBoxColumn
            cmb.DataSource = OverClass.TempDataTable("SELECT Virus_ID, Description FROM tblVirusStrains")
            cmb.DisplayMember = "Description"
            cmb.ValueMember = "Virus_ID"
            cmb.DataPropertyName = "Virus_ID"
            cmb.HeaderText = "Virus"
            OK.DataGridView200.Columns.Add(cmb)
            OK.Text = Volunteer & " Samples"
            OK.DataGridView200.AllowUserToAddRows = True
            OK.DataGridView200.AllowUserToDeleteRows = True
            Dim cmb2 As New DataGridViewImageColumn
            cmb2.Image = My.Resources.Remove
            cmb2.ImageLayout = DataGridViewImageCellLayout.Stretch
            cmb2.Name = "DeleteSample"
            cmb2.HeaderText = "Delete Sample"
            OK.DataGridView200.Columns.Add(cmb2)
            OK.DataGridView200.RowTemplate.Height = 35
            cmb2.Width = 50
            OK.ShowDialog()
            Me.TabControl1.SelectedIndex = 2
            Me.TabControl1_Selecting(Me.TabControl1, New TabControlCancelEventArgs(TabPage4, 0, False, TabControlAction.Selecting))


        End If

    End Sub

    Private Sub DateTimePicker101_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker101.ValueChanged

        Call Form1.Specifics(DataGridView102)

    End Sub

    Private Sub DateTimePicker102_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker102.ValueChanged

        Call Form1.Specifics(DataGridView102)

    End Sub

    Private Sub DateTimePicker103_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker103.ValueChanged

        Call Form1.Specifics(DataGridView103)

    End Sub

    Private Sub DateTimePicker104_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker104.ValueChanged

        Call Form1.Specifics(DataGridView103)

    End Sub
End Class