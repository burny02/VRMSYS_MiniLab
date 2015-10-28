Public Class MainForm

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Hide()

        Call StartUp(Me)


        Select Case Role

            Case "Site"
                SiteForm = New Site
                SiteForm.ShowDialog()

            Case "Lab"
                LabForm = New Form1
                Form1.ShowDialog()

        End Select


    End Sub
End Class