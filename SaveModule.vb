Module SaveModule
    Public Sub Saver(ctl As Object)

        Dim DisplayMessage As Boolean = True

        'Get a generic command list first - Ignore errors (Multi table)
        Dim cb As New OleDb.OleDbCommandBuilder(OverClass.CurrentDataAdapter)

        Try
            OverClass.CurrentDataAdapter.UpdateCommand = cb.GetUpdateCommand()
        Catch
        End Try
        Try
            OverClass.CurrentDataAdapter.InsertCommand = cb.GetInsertCommand()
        Catch
        End Try
        Try
            OverClass.CurrentDataAdapter.DeleteCommand = cb.GetDeleteCommand()
        Catch
        End Try


        'Create and overwrite a custom one if needed (More than 1 table) ...OLEDB Parameters must be added in the order they are used
        Select Case ctl.name


            Case "DataGridView1"

                Dim Person As String = "'" & WhichUser & "'"
                Dim ThisDate As String = OverClass.SQLDate(DateTime.Now)

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE tblApp_Results " &
                                                                    "Set Result=@P1, Batch_No=@P2, Entered_Person=" & Person &
                                                                    ", Entered_Date=" & ThisDate & "WHERE Result_ID=@P3")


                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double, 255, "Result")
                    .Add("@P2", OleDb.OleDbType.VarChar, 255, "Batch_No")
                    .Add("@P3", OleDb.OleDbType.Double, 255, "Result_ID")
                End With

            Case "DataGridView2"

                Dim Person As String = "'" & WhichUser & "'"
                Dim ThisDate As String = OverClass.SQLDate(DateTime.Now)

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE tblApp_Results " &
                                                                    "Set Lab_QC=@P1, Lab_QC_Date=" & ThisDate & ", Lab_QC_Person=" & Person _
                                                                          & " WHERE Result_ID=@P2 AND @P3=true")


                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Boolean, 255, "Lab_QC")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "Result_ID")
                    .Add("@P3", OleDb.OleDbType.Boolean, 255, "Lab_QC")
                End With

            Case "DataGridView3"

                Dim Person As String = "'" & WhichUser & "'"
                Dim ThisDate As String = OverClass.SQLDate(DateTime.Now)

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE tblApp_Results " &
                                                                    "Set Released=@P1, Released_Date=" & ThisDate & ", Released_By=" & Person _
                                                                          & " WHERE Result_ID=@P2 AND @P3=true")


                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Boolean, 255, "Released")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "Result_ID")
                    .Add("@P3", OleDb.OleDbType.Boolean, 255, "Released")
                End With

            Case "DataGridView100"

                Dim Person As String = "'" & WhichUser & "'"
                Dim ThisDate As String = OverClass.SQLDate(DateTime.Now)

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE tblApp_Results " &
                                                                    "Set Site_QC=@P1, Site_QC_Date=" & ThisDate & ", Site_QC_Person=" & Person _
                                                                          & " WHERE Result_ID=@P2 AND @P3=true")


                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Boolean, 255, "Site_QC")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "Result_ID")
                    .Add("@P3", OleDb.OleDbType.Boolean, 255, "Site_QC")
                End With

            Case "DataGridView101"

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE tblVirusStrains " & _
                                                                    "Set DefaultTest=@P1 WHERE Virus_ID=@P2")


                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Boolean, 255, "DefaultTest")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "Virus_ID")
                End With

            Case "DataGridView200"

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE tblApp_Results " & _
                                                                    "Set Virus_ID=@P1 WHERE Result_ID=@P2")


                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double, 255, "Virus_ID")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "Result_ID")
                End With

                OverClass.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO tblApp_Results " & _
                                                                    "(App_ID, Virus_ID) VALUES (" & AppID & ", @P1)")


                With OverClass.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double, 255, "Virus_ID")
                End With

                OverClass.CurrentDataAdapter.DeleteCommand = New OleDb.OleDbCommand("DELETE FROM tblApp_Results " & _
                                                                    "WHERE Result_ID=@P1")


                With OverClass.CurrentDataAdapter.DeleteCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double, 255, "Result_ID")
                End With

        End Select


        Call OverClass.SetCommandConnection()
        Call OverClass.UpdateBackend(ctl, DisplayMessage)
        If DisplayMessage = False Then MsgBox("Table Updated")


    End Sub


End Module
