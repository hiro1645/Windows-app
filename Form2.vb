Public Class Form2
    Private dt As New DataTable
    Private dt2 As New DataTable
    Private count As Integer

    Sub New(a As Integer, b As String, c As Integer, dtt As DataTable)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        DataGridView1.AllowUserToAddRows = False

        DataGridView1.ReadOnly = True
        If a > 0 Then
            DataGridView1.DataSource = dtt
            dt2 = dtt
            Dim colBtn As New DataGridViewButtonColumn()
            colBtn.Name = "削除"
            colBtn.UseColumnTextForButtonValue = True
            colBtn.Text = "削除"

            DataGridView1.Columns.Add(colBtn)

            DataGridView1.Rows(a - 1).Cells(c).Value = b

            count = 1
        End If

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        makecombo()

        dt.Columns.Add("日付")
        dt.Columns.Add("項目")
        dt.Columns.Add("内容")

        DataGridView2.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GetGrid()

        If count = 0 Then
            Dim colBtn As New DataGridViewButtonColumn()
            colBtn.Name = "削除"
            colBtn.UseColumnTextForButtonValue = True
            colBtn.Text = "削除"

            DataGridView1.Columns.Add(colBtn)

        End If


        count = 1

    End Sub

    Private Sub GetGrid()

        Dim unused = MonthCalendar1.SelectionStart
        Dim data As String = MaskedTextBox1.Text
        Dim combo As String = ComboBox1.Text

        dt.Rows.Add(unused, data, combo)

        DataGridView1.DataSource = dt



    End Sub

    Private Sub makecombo()
        Dim dt As New DataTable

        dt.Columns.Add("a")
        dt.Columns.Add("b")

        dt.Rows.Add("1", "給料")

        dt.Rows.Add("2", "ガス代")

        ComboBox1.DataSource = dt
        ComboBox1.DisplayMember = "b"
        ComboBox1.ValueMember = "a"
    End Sub


    Private index As Integer
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        index = e.RowIndex + 1

        Dim a As String = (index.ToString() + "行目を削除してしまってよろしいですか？")


        If DataGridView1.Columns(e.ColumnIndex).Name = "削除" Then
            Dim result As DialogResult = MessageBox.Show(a,
                                             "質問",
                                             MessageBoxButtons.YesNo,
                                             MessageBoxIcon.Exclamation,
                                             MessageBoxDefaultButton.Button2)
            '何が選択されたか調べる 
            If result = DialogResult.Yes Then
                '「はい」が選択された時 
                DataGridView1.Rows.RemoveAt(DataGridView1.CurrentCell.RowIndex)
            Else
                Exit Sub
            End If

        End If
    End Sub

    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        Dim row As Integer = e.RowIndex
        Dim row2 As Integer = e.ColumnIndex
        Dim value As String = DataGridView1.Rows(row).Cells(row2).Value

        Dim f As New Form8(value, row, row2, dt)
        f.ShowDialog()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim money As Integer = 0
        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(2).Value = "給料" Then
                money += CInt(DataGridView1.Rows(i).Cells(1).Value)
            End If
        Next
        dt.Rows.Add(MonthCalendar1.SelectionStart, money & " 集計 ", "給料")
        DataGridView1.DataSource = dt
        MessageBox.Show(money)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dtSelect As DataTable = dt.Copy
        'dtSelect.Columns.Add("日付")
        'dtSelect.Columns.Add("項目")
        'dtSelect.Columns.Add("内容")

        'Dim ret As DataRow()

        'ret = dtSelect.Select("内容 = '給料'")

        'dt.Rows.Clear()

        Dim con As New List(Of Integer)
        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(2).Value = "給料" Then
                con.Add(i)
            End If
        Next

        For aa = 0 To con.Count - 1
            If aa >= DataGridView1.Rows.Count Then
                Exit For
            End If
            If DataGridView1.Rows(aa).Cells(2).Value = "給料" Then
                DataGridView1.Rows.RemoveAt(aa)
            End If
        Next

        'DataGridView1.Visible = False
        ''DataGridView2.DataSource = dtSelect
        'DataGridView2.Visible = True

    End Sub
End Class