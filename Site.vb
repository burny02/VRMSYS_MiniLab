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

            Case "Site QC"
                StartCombo(Me.ComboBox100)

            Case "DefaultTests"
                Form1.Specifics(Me.DataGridView101)

            Case "VolunteerTests"
                Form1.Specifics(Me.DataGridView102)


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

    Private Sub DataGridView102_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView102.CellContentClick

        If e.ColumnIndex = Me.DataGridView102.Columns("Volunteer").Index Then

            Dim AppID As Long = Me.DataGridView102.Item("APP_ID", e.RowIndex).Value
            Dim Volunteer As String = Me.DataGridView102.Item("Volunteer", e.RowIndex).Value
            Dim OK As New PickTests
            OverClass.CreateDataSet("SELECT Virus_ID FROM tblApp_Results " & _
                    "WHERE APP_ID=" & AppID, OK.BindingSource1, OK.DataGridView200)
            OK.DataGridView200.Columns("Virus_ID").Visible = False
            Dim cmb As New DataGridViewComboBoxColumn
            cmb.DataSource = OverClass.TempDataTable("SELECT Virus_ID, Description FROM tblVirusStrains")
            cmb.DisplayMember = "Description"
            cmb.ValueMember = "Virus_ID"
            cmb.DataPropertyName = "Virus_ID"
            OK.DataGridView200.Columns.Add(cmb)
            OK.Text = Volunteer & " Samples"
            OK.ShowDialog()


        End If
    End Sub
End Class