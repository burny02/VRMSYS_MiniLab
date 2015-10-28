
Module ButtonModule

    Public Sub ButtonSpecifics(sender As Object, e As EventArgs)

        Dim ctl As Object = Nothing

        Select Case sender.name.ToString

            Case "Button1"
                Call Saver(Form1.DataGridView1)

            Case "Button2"
                Call Saver(Form1.DataGridView2)

            Case "Button3"
                Call Saver(Form1.DataGridView3)

            Case "Button4"
                Call ExportExcel("SELECT * FROM LabExport")

            Case "Button100"
                Call Saver(SiteForm.DataGridView100)

            Case "Button101"
                Call Saver(SiteForm.DataGridView101)

            Case "Button102"
                Call ExportExcel("SELECT * FROM LabExport")

            Case "Button103"
                Call ExportExcel(OverClass.CurrentDataAdapter.SelectCommand.CommandText)

        End Select



    End Sub


End Module
