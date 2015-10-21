Module ExportModule

    Public Sub ExportExcel(SQLCode As String)

        Dim dt As New DataTable
        Dim da As OleDb.OleDbDataAdapter = OverClass.NewDataAdapter(SQLCode)
        da.Fill(dt)

        da = Nothing

        Dim i As Integer
        Dim j As Integer

        Dim xlApp As Object
        xlApp = CreateObject("Excel.Application")
        With xlApp
            .Visible = False
            .Workbooks.Add()
            .Sheets("Sheet1").Select()

            'Add column heading
            For i = 1 To dt.Columns.Count
                xlApp.activesheet.Cells(1, i).Value = dt.Columns(i - 1).ColumnName
            Next i

            'Add Rows
            For i = 0 To dt.Rows.Count - 1
                For j = 0 To dt.Columns.Count - 1
                    xlApp.activesheet.Cells(i + 2, j + 1) = dt.Rows(i).Item(j)
                Next j
            Next i

            xlApp.Cells.EntireColumn.AutoFit()
            .activesheet.Range("$A$1:$Z$1").AutoFilter()

        End With

        Dim numrow As Long
        numrow = dt.Rows.Count + 1
        dt = Nothing
        da = Nothing

        xlApp.Visible = True

    End Sub

End Module
