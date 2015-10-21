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

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE tblApp_Results " & _
                                                                    "Set Result=@P1, Batch_No=@P2 " & _
                                                                          "WHERE Result_ID=@P3")


                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Double, 255, "Result")
                    .Add("@P2", OleDb.OleDbType.VarChar, 255, "Batch_No")
                    .Add("@P3", OleDb.OleDbType.Double, 255, "Result_ID")
                End With

            Case "DataGridView2"

                Dim Person As String = "'" & OverClass.GetUserName & "'"
                Dim ThisDate As String = OverClass.SQLDate(DateTime.Now)

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE tblApp_Results " & _
                                                                    "Set Lab_QC=@P1, Lab_QC_Date=" & ThisDate & ", Lab_QC_Person=" & Person _
                                                                          & " WHERE Result_ID=@P2")


                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Boolean, 255, "Lab_QC")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "Result_ID")
                End With

            Case "DataGridView3"

                Dim Person As String = "'" & OverClass.GetUserName & "'"
                Dim ThisDate As String = OverClass.SQLDate(DateTime.Now)

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE tblApp_Results " & _
                                                                    "Set Released=@P1, Released_Date=" & ThisDate & ", Released_By=" & Person _
                                                                          & " WHERE Result_ID=@P2")


                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Boolean, 255, "Released")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "Result_ID")
                End With

            Case "DataGridView100"

                Dim Person As String = "'" & OverClass.GetUserName & "'"
                Dim ThisDate As String = OverClass.SQLDate(DateTime.Now)

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE tblApp_Results " & _
                                                                    "Set Site_QC=@P1, Site_QC_Date=" & ThisDate & ", Site_QC_Person=" & Person _
                                                                          & " WHERE Result_ID=@P2")


                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Boolean, 255, "Site_QC")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "Result_ID")
                End With

            Case "DataGridView101"

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE tblVirusStrains " & _
                                                                    "Set DefaultTest=@P1 WHERE Virus_ID=@P2")


                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.Boolean, 255, "DefaultTest")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "Virus_ID")
                End With
        End Select


        Call OverClass.SetCommandConnection()
        Call OverClass.UpdateBackend(ctl, DisplayMessage)
        If DisplayMessage = False Then MsgBox("Table Updated")


    End Sub


End Module
