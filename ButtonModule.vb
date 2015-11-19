Imports Microsoft.Reporting.WinForms

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
                Dim OK As New ReportDisplay

                OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "VRMSYS_MiniLab.ResultExport.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                        OverClass.TempDataTable(
                                                        OverClass.CurrentDataAdapter.SelectCommand.CommandText)))

                OK.ReportViewer1.RefreshReport()

                Dim RandNo As String = Format(DateAndTime.Now, "ddssmmhhyymm")

                Dim pdfContent As Byte() = OK.ReportViewer1.LocalReport.Render("PDF", Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                Dim pdfPath As String = ReportPath & RandNo & ".pdf"
                Dim pdfFile As New System.IO.FileStream(pdfPath, System.IO.FileMode.Create)
                pdfFile.Write(pdfContent, 0, pdfContent.Length)
                pdfFile.Close()


                OK.Close()

                Try

                    Process.Start("explorer.exe", pdfPath)

                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Case "Button100"
                Call Saver(SiteForm.DataGridView100)

            Case "Button101"
                Call Saver(SiteForm.DataGridView101)

            Case "Button102"
                Dim OK As New ReportDisplay

                OK.ReportViewer1.ProcessingMode = ProcessingMode.Local
                OK.ReportViewer1.LocalReport.ReportEmbeddedResource = "VRMSYS_MiniLab.ResultExport.rdlc"
                OK.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                        OverClass.TempDataTable(
                                                        OverClass.CurrentDataAdapter.SelectCommand.CommandText)))

                OK.ReportViewer1.RefreshReport()

                Dim RandNo As String = Format(DateAndTime.Now, "ddssmmhhyymm")

                Dim pdfContent As Byte() = OK.ReportViewer1.LocalReport.Render("PDF", Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                Dim pdfPath As String = ReportPath & RandNo & ".pdf"
                Dim pdfFile As New System.IO.FileStream(pdfPath, System.IO.FileMode.Create)
                pdfFile.Write(pdfContent, 0, pdfContent.Length)
                pdfFile.Close()


                OK.Close()

                Try

                    Process.Start("explorer.exe", pdfPath)

                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

        End Select



    End Sub


End Module
