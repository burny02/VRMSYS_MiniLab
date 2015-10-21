Public Class Form1

    Public Sub Specifics(ctl As DataGridView)

        If IsNothing(ctl) Then Exit Sub

        Dim SQLCode As String = vbNullString

        Select Case ctl.name

            Case "DataGridView1"
                If IsNothing(Me.ComboBox1.SelectedValue) Then Exit Sub

                SQLCode = "SELECT Result_ID, Patient_Attendees_ID & ' - ' & Format(Date_Of_Birth,'dd-MMM-yyyy') AS Volunteer, " & _
                    "Start, Result, Batch_No " & _
                    "FROM (tblAppointments a INNER JOIN " & _
                    "tblApp_Results b ON a.ID=b.APP_ID) INNER JOIN tblPatientDemographics c " & _
                    " ON a.Patient_Attendees_ID=c.ID " & _
                    "WHERE Virus_ID=" & Me.ComboBox1.SelectedValue & _
                    " AND Lab_QC=False " & _
                    "ORDER BY START ASC, Patient_Attendees_ID ASC"

                OverClass.CreateDataSet(SQLCode, BindingSource1, DataGridView1)

                ctl.Columns("Result_ID").Visible = False
                ctl.Columns("Result").Visible = False
                ctl.Columns("Start").HeaderText = "Collection Date"
                ctl.Columns("Start").DefaultCellStyle.Format = "dd-MMM-yyyy"
                ctl.Columns("Volunteer").ReadOnly = True
                ctl.Columns("Start").ReadOnly = True

                Dim dt As DataTable
                Dim cmb As New DataGridViewComboBoxColumn
                dt = OverClass.TempDataTable("SELECT Display, ActValue FROM tblResults ORDER BY ACTValue ASC, Display ASC")
                cmb.ValueMember = "ActValue"
                cmb.DisplayMember = "Display"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("Result").ToString
                cmb.HeaderText = "Result"
                cmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
                Dim i As Long = 1
                Do While i <> 13000
                    dt.Rows.Add(i, i)
                    i += 1
                Loop
                cmb.DataSource = dt
                ctl.Columns.Add(cmb)

            Case "DataGridView2"
                If IsNothing(Me.ComboBox2.SelectedValue) Then Exit Sub

                SQLCode = "SELECT Result_ID, Patient_Attendees_ID & ' - ' & Format(Date_Of_Birth,'dd-MMM-yyyy') AS Volunteer, " & _
                    "Start, Result, Batch_No, Lab_QC " & _
                    "FROM (tblAppointments a INNER JOIN " & _
                    "tblApp_Results b ON a.ID=b.APP_ID) INNER JOIN tblPatientDemographics c " & _
                    " ON a.Patient_Attendees_ID=c.ID " & _
                    "WHERE Virus_ID=" & Me.ComboBox2.SelectedValue & _
                    " AND Lab_QC=False " & _
                    "ORDER BY START ASC, Patient_Attendees_ID ASC"

                OverClass.CreateDataSet(SQLCode, BindingSource1, DataGridView2)

                ctl.Columns("Result_ID").Visible = False
                ctl.Columns("Result").Visible = False
                ctl.Columns("Start").HeaderText = "Collection Date"
                ctl.Columns("Start").DefaultCellStyle.Format = "dd-MMM-yyyy"
                ctl.Columns("Lab_QC").HeaderText = "QC Check"
                ctl.Columns("Volunteer").ReadOnly = True
                ctl.Columns("Start").ReadOnly = True
                ctl.Columns("Batch_No").ReadOnly = True
                ctl.Columns("Result").ReadOnly = True

                Dim dt As DataTable
                Dim cmb As New DataGridViewComboBoxColumn
                dt = OverClass.TempDataTable("SELECT Display, ActValue FROM tblResults ORDER BY ACTValue ASC, Display ASC")
                cmb.ValueMember = "ActValue"
                cmb.DisplayMember = "Display"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("Result").ToString
                cmb.HeaderText = "Result"
                cmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                Dim i As Long = 1
                Do While i <> 13000
                    dt.Rows.Add(i, i)
                    i += 1
                Loop
                cmb.DataSource = dt
                cmb.ReadOnly = True
                ctl.Columns.Add(cmb)

                ctl.Columns("Lab_QC").DisplayIndex = 6


            Case "DataGridView3"
                If IsNothing(Me.ComboBox3.SelectedValue) Then Exit Sub

                SQLCode = "SELECT Result_ID, Patient_Attendees_ID & ' - ' & Format(Date_Of_Birth,'dd-MMM-yyyy') AS Volunteer, " & _
                    "Start, Result, Batch_No, Lab_QC_Person, Lab_QC_Date, Released " & _
                    "FROM (tblAppointments a INNER JOIN " & _
                    "tblApp_Results b ON a.ID=b.APP_ID) INNER JOIN tblPatientDemographics c " & _
                    " ON a.Patient_Attendees_ID=c.ID " & _
                    "WHERE Virus_ID=" & Me.ComboBox3.SelectedValue & _
                    " AND Lab_QC=True AND Released=False " & _
                    "ORDER BY START ASC, Patient_Attendees_ID ASC"

                OverClass.CreateDataSet(SQLCode, BindingSource1, DataGridView3)

                ctl.Columns("Result_ID").Visible = False
                ctl.Columns("Result").Visible = False
                ctl.Columns("Start").HeaderText = "Collection Date"
                ctl.Columns("Start").DefaultCellStyle.Format = "dd-MMM-yyyy"
                ctl.Columns("Lab_QC_Date").DefaultCellStyle.Format = "dd-MMM-yyyy"
                ctl.Columns("Lab_QC_Person").HeaderText = "QC'd By"
                ctl.Columns("Lab_QC_Date").HeaderText = "QC'd Date"
                ctl.Columns("Volunteer").ReadOnly = True
                ctl.Columns("Start").ReadOnly = True
                ctl.Columns("Batch_No").ReadOnly = True
                ctl.Columns("Result").ReadOnly = True
                ctl.Columns("Lab_QC_Person").ReadOnly = True
                ctl.Columns("Lab_QC_Date").ReadOnly = True

                Dim dt As DataTable
                Dim cmb As New DataGridViewComboBoxColumn
                dt = OverClass.TempDataTable("SELECT Display, ActValue FROM tblResults ORDER BY ACTValue ASC, Display ASC")
                cmb.ValueMember = "ActValue"
                cmb.DisplayMember = "Display"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("Result").ToString
                cmb.HeaderText = "Result"
                cmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                Dim i As Long = 1
                Do While i <> 13000
                    dt.Rows.Add(i, i)
                    i += 1
                Loop
                cmb.DataSource = dt
                cmb.ReadOnly = True
                ctl.Columns.Add(cmb)

                ctl.Columns("Released").DisplayIndex = 8

            Case "DataGridView4"
                If IsNothing(Me.ComboBox4.SelectedValue) Then Exit Sub

                SQLCode = "SELECT Patient_Attendees_ID & ' - ' & Format(Date_Of_Birth,'dd-MMM-yyyy') AS Volunteer, Start AS Collection_Date, " & _
                    "Result, Lab_QC_Person, Lab_QC_Date, Released_By, Released_Date " & _
                    "FROM (tblAppointments a INNER JOIN " & _
                    "tblApp_Results b ON a.ID=b.APP_ID) INNER JOIN tblPatientDemographics c " & _
                    " ON a.Patient_Attendees_ID=c.ID " & _
                    "WHERE Virus_ID=" & Me.ComboBox4.SelectedValue & _
                    " ORDER BY START ASC, Patient_Attendees_ID ASC"

                OverClass.CreateDataSet(SQLCode, BindingSource1, DataGridView4)

                ctl.ReadOnly = True
                ctl.Columns("Lab_QC_Person").HeaderText = "QC'd By"
                ctl.Columns("Lab_QC_Date").HeaderText = "QC Date"
                ctl.Columns("Released_By").HeaderText = "Released By"
                ctl.Columns("Released_Date").HeaderText = "Released Date"
                ctl.Columns("Collection_Date").HeaderText = "Collection Date"
                ctl.Columns("Result").Visible = False
                ctl.Columns("Lab_QC_Date").DefaultCellStyle.Format = "dd-MMM-yyyy"
                ctl.Columns("Released_Date").DefaultCellStyle.Format = "dd-MMM-yyyy"
                ctl.Columns("Collection_Date").DefaultCellStyle.Format = "dd-MMM-yyyy"

                Dim dt As DataTable
                Dim cmb As New DataGridViewComboBoxColumn
                dt = OverClass.TempDataTable("SELECT Display, ActValue FROM tblResults ORDER BY ACTValue ASC, Display ASC")
                cmb.ValueMember = "ActValue"
                cmb.DisplayMember = "Display"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("Result").ToString
                cmb.HeaderText = "Result"
                cmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                Dim i As Long = 1
                Do While i <> 13000
                    dt.Rows.Add(i, i)
                    i += 1
                Loop
                cmb.DataSource = dt
                cmb.ReadOnly = True
                ctl.Columns.Add(cmb)
                cmb.DisplayIndex = 3

            Case "DataGridView100"

                If IsNothing(SiteForm.ComboBox100.SelectedValue) Then Exit Sub

                SQLCode = "SELECT Patient_Attendees_ID & ' - ' & Format(Date_Of_Birth,'dd-MMM-yyyy') AS Volunteer, Start AS Collection_Date, " & _
                    "Result_ID, Result, Released_By, Released_Date, Site_QC " & _
                    "FROM (tblAppointments a INNER JOIN " & _
                    "tblApp_Results b ON a.ID=b.APP_ID) INNER JOIN tblPatientDemographics c " & _
                    " ON a.Patient_Attendees_ID=c.ID " & _
                    "WHERE Virus_ID=" & SiteForm.ComboBox100.SelectedValue & _
                    " AND Released=True AND Site_QC=False " & _
                    "ORDER BY START ASC, Patient_Attendees_ID ASC"

                OverClass.CreateDataSet(SQLCode, SiteForm.BindingSource1, SiteForm.DataGridView100)


                ctl.Columns("Site_QC").HeaderText = "QC Check"
                ctl.Columns("Released_By").HeaderText = "Released By"
                ctl.Columns("Released_Date").HeaderText = "Released Date"
                ctl.Columns("Collection_Date").HeaderText = "Collection Date"
                ctl.Columns("Result").Visible = False
                ctl.Columns("Result_ID").Visible = False
                ctl.Columns("Released_Date").DefaultCellStyle.Format = "dd-MMM-yyyy"
                ctl.Columns("Collection_Date").DefaultCellStyle.Format = "dd-MMM-yyyy"
                ctl.Columns("Volunteer").ReadOnly = True
                ctl.Columns("Collection_Date").ReadOnly = True
                ctl.Columns("Result").ReadOnly = True
                ctl.Columns("Released_By").ReadOnly = True
                ctl.Columns("Released_Date").ReadOnly = True

                Dim dt As DataTable
                Dim cmb As New DataGridViewComboBoxColumn
                dt = OverClass.TempDataTable("SELECT Display, ActValue FROM tblResults ORDER BY ACTValue ASC, Display ASC")
                cmb.ValueMember = "ActValue"
                cmb.DisplayMember = "Display"
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("Result").ToString
                cmb.HeaderText = "Result"
                cmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
                Dim i As Long = 1
                Do While i <> 13000
                    dt.Rows.Add(i, i)
                    i += 1
                Loop
                cmb.DataSource = dt
                cmb.ReadOnly = True
                ctl.Columns.Add(cmb)
                cmb.DisplayIndex = 3

            Case "DataGridView101"

                SQLCode = "SELECT Virus_ID, Description, DefaultTest FROM tblVirusStrains"

                OverClass.CreateDataSet(SQLCode, SiteForm.BindingSource1, SiteForm.DataGridView101)


                ctl.Columns("Virus_ID").Visible = False
                ctl.Columns("Description").HeaderText = "Virus"
                ctl.Columns("DefaultTest").HeaderText = "Test as default"
                ctl.Columns("Description").ReadOnly = True

            Case "DataGridView102"

                SQLCode = "SELECT * FROM WhichTests"

                OverClass.CreateDataSet(SQLCode, SiteForm.BindingSource1, SiteForm.DataGridView102)

                Dim fonter As Font

                fonter = New Font("Arial", 10, FontStyle.Underline)

                ctl.Columns("Volunteer").DefaultCellStyle.Font = fonter
                ctl.Columns("Volunteer").DefaultCellStyle.ForeColor = Color.Blue
                ctl.Columns("App_ID").Visible = False
                ctl.ReadOnly = True
                ctl.Columns("Collection_Date").DefaultCellStyle.Format = "dd-MMM-yyyy"


        End Select

    End Sub


    Private Sub Form1_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load


        Me.WindowState = FormWindowState.Maximized

        Call StartUp(Me)

        Try
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Catch
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside"
        End Try

        Me.Text = SolutionName


    End Sub

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

            Case "InputLab"
                StartCombo(Me.ComboBox1)

            Case "QCLab"
                StartCombo(Me.ComboBox2)

            Case "ReleaseLab"
                StartCombo(Me.ComboBox3)

            Case "All Results"
                StartCombo(Me.ComboBox4)

        End Select


        If Not IsNothing(ctl) Then Call Specifics(ctl)


    End Sub
End Class

