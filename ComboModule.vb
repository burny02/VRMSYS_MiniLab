Module ComboModule

    Public Sub GenericCombo(sender As Object, e As EventArgs)

        If OverClass.UnloadData() = True Then Exit Sub
        OverClass.ResetCollection()
        Call SubCombo(sender)


    End Sub

    Public Sub SubCombo(sender As ComboBox)

        Select Case sender.Name.ToString

            Case Else
                ComboRefreshData(sender)

        End Select

    End Sub

    Public Sub StartCombo(ctl As ComboBox)

        Select Case ctl.Name.ToString()

            Case "ComboBox1", "ComboBox2", "ComboBox3", "ComboBox100"
                ctl.DataSource = OverClass.TempDataTable("SELECT Virus_ID, Description FROM tblVirusStrains " & _
                                                         "WHERE Redundant=FALSE")
                ctl.ValueMember = "Virus_ID"
                ctl.DisplayMember = "Description"


        End Select

        ComboRefreshData(ctl)

    End Sub

    Public Sub ComboRefreshData(sender As ComboBox)

        Dim Grid As DataGridView = Nothing

        Select Case sender.Name.ToString()

            Case "ComboBox1"
                Grid = Form1.DataGridView1

            Case "ComboBox2"
                Grid = Form1.DataGridView2

            Case "ComboBox3"
                Grid = Form1.DataGridView3

            Case "ComboBox100"
                Grid = SiteForm.DataGridView100

        End Select


        If Not IsNothing(Grid) Then Call Form1.Specifics(Grid)

    End Sub

End Module
