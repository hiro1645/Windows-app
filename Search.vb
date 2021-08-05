Public Class Search

    Private dt As New DataTable
    Private flug As Boolean
    Private dic As New Dictionary(Of String, Boolean)

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles 検索.Click

        Using db As New dbConnection

            Dim sql As String = ""

            sql = ""
            sql &= "select * from mst_inserts3 "

            dt = db.getDtSql(sql)

        End Using


        DataGridView1.DataSource = dt

        TWO()

    End Sub

    Private Sub TWO()

        Dim dt2 As New DataTable
        dt2.Columns.Add("タイトル")

        For i = 0 To dt.Columns.Count - 1
            dt2.Rows.Add(dt.Columns(i).ColumnName)
        Next

        DataGridView2.DataSource = dt2

        Dim column As New DataGridViewCheckBoxColumn
        DataGridView2.Columns.Add(column)

        DataGridView2.AllowUserToAddRows = False


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim dtr As DataTable = dt.Copy

        Dim high As New List(Of String)

        For w = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(w).Cells(1).Value = False Then
                high.Add(DataGridView2.Rows(w).Cells(0).Value)
            End If
        Next

        For Each t In high
            Dim o As Integer = dtr.Columns(t).Ordinal
            dtr.Columns.RemoveAt(o)
        Next

        DataGridView1.DataSource = dtr

        Dim i As New List(Of Integer)

        Dim value As New List(Of String)


        If flug = False Then
            For aaa = 0 To DataGridView2.Rows.Count - 1
                dic.Add(DataGridView2.Rows(aaa).Cells(0).Value.ToString, CType(DataGridView2.Rows(aaa).Cells(1).Value, Boolean))
            Next
            flug = True
        End If


        For a = 0 To DataGridView2.Rows.Count - 1
            If CType(DataGridView2.Rows(a).Cells(1).Value, Boolean) = False Then
                i.Add(a)
                value.Add(DataGridView2.Rows(a).Cells(0).Value.ToString)
            End If
        Next

        Dim data As DataTable = CType(DataGridView2.DataSource, DataTable)
        Dim rows As DataRow()

        rows = data.Select("タイトル = 'JANCODE'")


        For Each t As String In value

            rows = data.Select("タイトル =  '" & t & "' ")

            For Each Row As DataRow In rows
                data.Rows.Remove(Row)
            Next

        Next

        DataGridView2.DataSource = Nothing
        DataGridView2.Columns.Clear()

        DataGridView2.DataSource = CType(data, DataTable)

        Dim column As New DataGridViewCheckBoxColumn
        DataGridView2.Columns.Add(column)

        For check = 0 To DataGridView2.Rows.Count - 1
            DataGridView2.Rows(check).Cells(1).Value = True
        Next

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dic2 As New Dictionary(Of String, Boolean)

        For aaa = 0 To DataGridView2.Rows.Count - 1
            dic2.Add(DataGridView2.Rows(aaa).Cells(0).Value.ToString, CType(DataGridView2.Rows(aaa).Cells(1).Value, Boolean))
        Next

        DataGridView2.DataSource = Nothing

        DataGridView2.Columns.Clear()

        Dim data As New DataTable

        data.Columns.Add("タイトル")

        For Each a In dic
            data.Rows.Add(a.Key)
        Next

        DataGridView2.DataSource = CType(data, DataTable)

        Dim column As New DataGridViewCheckBoxColumn

        DataGridView2.Columns.Add(CType(column, DataGridViewCheckBoxColumn))

        For k = 0 To DataGridView2.Rows.Count - 1
            For Each t In dic2
                If DataGridView2.Rows(k).Cells(0).Value.ToString = t.Key Then
                    DataGridView2.Rows(k).Cells(1).Value = t.Value
                End If
            Next
        Next

        DataGridView1.DataSource = dt
    End Sub
End Class