Public Class PickTests

    Private Sub DataGridView200_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView200.CellContentClick

        If e.ColumnIndex = sender.columns("DeleteSample").index Then
            If IsDBNull(Me.DataGridView200.Item("Virus_ID", e.RowIndex).Value) Then Exit Sub
            If MsgBox("Are you sure you want to delete?" & vbNewLine & "Table must be saved to commit delete", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Dim row As DataGridViewRow
                row = sender.rows(e.RowIndex)
                sender.rows.remove(row)
            End If
        End If

    End Sub

    Private Sub Button200_Click(sender As Object, e As EventArgs) Handles Button200.Click

        Call Saver(Me.DataGridView200)

    End Sub
End Class