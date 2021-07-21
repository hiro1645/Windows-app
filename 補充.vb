Imports System.IO
Imports System.Text
Public Class Form7
    Private loginID As Integer
    'Private Item As Integer


    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loginID = LoginForm1.ID
        Label3.Text = LoginForm1.Namae
        Label2.Text = "JANコード(14桁)で検索して下さい"
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        'If TextBox1.Text = "" Then
        '    MessageBox.Show("JANコードを入力してください")
        'Else
        If e.KeyCode = Keys.Enter Then
            If TextBox1.Text = String.Empty Then
                MessageBox.Show("JANコード(14桁)を入力してください")
                Exit Sub
            End If
            Rep()
            Me.Label2.Visible = False
        End If


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, MyBase.Enter

        Try
            Dim sql As String

            sql = ""
            sql &= "SELECT jan_code, "
            sql &= "item_name, "
            sql &= "standard, "
            sql &= "stock_count, "
            sql &= "mst_item_id, "
            sql &= "place_name, "
            sql &= "order_num "
            sql &= "FROM trn_stocks as a "
            sql &= "Left join mst_items as b "
            sql &= "on a.mst_item_id = b.id "
            sql &= " Left join mst_areas as c "
            sql &= " on a.place_id = c.order_num "
            sql &= "where mst_dealer_id = " & loginID
            sql &= " and stock_count < 20 "
            sql &= " and "
            sql &= "jan_code = '" & TextBox1.Text & "' "
            sql &= " order by order_num "

            Dim dt As DataTable = New DataTable

            'getDtSqlの場合
            Using db As New dbConnection()

                dt = db.getDtSql(sql)

            End Using

            If dt.Rows.Count = 0 Then
                MessageBox.Show("完了しています", "正常", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'Throw New Exception
                Exit Sub
            End If

            Dim count As Integer = 0
            Dim count2 As Integer = 0

            For Each row As DataRow In dt.Rows
                For i As Integer = 1 To 6
                    If row("stock_count") < 20 Then
                        sql = ""
                        sql &= " update trn_stocks"
                        sql &= " set stock_count = stock_count + 1"
                        sql &= " where mst_item_id =" & Rep()
                        sql &= " and place_id =" & row("order_num")
                        Using db As New dbConnection()

                            Try
                                db.trnStart()
                                db.executeSql(sql)
                                db.commit()
                            Catch ex As Exception
                                db.rollback()
                                Throw
                            End Try

                        End Using

                        'row("stock_count") += 1
                        count += 1
                        count2 += 1
                    End If

                    If count = 6 Then
                        Exit For
                    End If
                Next

                If count2 > 0 Then
                    MessageBox.Show(row("place_name") & count2 & "補充")
                End If

                count2 = 0

                If count = 6 Then
                    Exit For
                End If
            Next
            Rep()
            main()

        Catch ex As Exception
            MessageBox.Show("エラー" & vbCrLf & "[" & System.Reflection.MethodBase.GetCurrentMethod.Name & "]" & ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

    Private Function Rep()
        'データベース接続

        Dim sql As String
        sql = ""
        sql &= " select"
        sql &= " c.mst_item_id,"
        sql &= " a.jan_code,"
        sql &= " a.item_name,"
        sql &= " a.standard,"
        sql &= " sum(20 - stock_count) as rep_count"
        sql &= " from mst_items as a"
        sql &= " left join trn_stocks as c"
        sql &= " on a.id = c.mst_item_id"
        sql &= " where a.mst_dealer_id = " & loginID


        sql &= " and "
        sql &= " jan_code = '" & TextBox1.Text & "' "

        sql &= " And "
        sql &= " stock_count < 20"
        sql &= " group by a.id, c.mst_item_id"
        sql &= " order by c.mst_item_id asc"

        ' SQL実行
        Dim dt As DataTable = New DataTable

        'getDtSqlの場合
        Using db As New dbConnection()
            dt = db.getDtSql(sql)
        End Using

        If dt.Rows.Count = 0 Then
            MessageBox.Show("補充するものがありません。またはJANコードが存在しません")
            Exit Function
        End If

        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        DataGridView1.Columns.Add("jan_code", "JANコード")
        DataGridView1.Columns.Add("item_name", "商品名")
        DataGridView1.Columns.Add("standard", "規格")
        DataGridView1.Columns.Add("rep_count", "補充数")

        If (dt.Rows).Count > 0 Then
            Dim row As DataRow = dt.Rows(0)
            Dim janc As String = row("jan_code")
            Dim item As String = row("item_name")
            Dim standard As String = row("standard")
            Dim rep_count As Integer = row("rep_count")
            Dim mst_item_id As Integer = row("mst_item_id")

            DataGridView1.Rows.Add(janc, item, standard, rep_count)

            Return mst_item_id
        End If


    End Function

    Private Sub main()

        Me.DataGridView2.Visible = False

        Dim sql As String
        sql = ""
        sql &= " select"
        sql &= " c.mst_item_id,"
        sql &= " a.jan_code,"
        sql &= " a.item_name,"
        sql &= " a.standard,"
        sql &= "stock_count, "
        sql &= "20 -stock_count as rep_count "
        sql &= " from mst_items as a"
        sql &= " left join trn_stocks as c"
        sql &= " on a.id = c.mst_item_id"
        sql &= " where a.mst_dealer_id = " & LoginForm1.ID
        sql &= " And a.jan_code ='" & TextBox1.Text & "' "
        sql &= " order by c.mst_item_id asc"

        Dim dt As DataTable = New DataTable

        'getDtSqlの場合
        Using db As New dbConnection()

            dt = db.getDtSql(sql)

        End Using

        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()

        DataGridView2.Columns.Add("jan_code", "JANコード")
        DataGridView2.Columns.Add("item_name", "商品名")
        DataGridView2.Columns.Add("standard", "規格")
        DataGridView2.Columns.Add("stock_count", "在庫数")
        DataGridView2.Columns.Add("rep_count", "欠品数")

        For Each row As DataRow In dt.Rows
            Dim janc As String = row("jan_code")
            Dim item As String = row("item_name")
            Dim standard As String = row("standard")
            Dim stock As Integer = row("stock_count")
            Dim rep As Integer = row("rep_count")
            DataGridView2.Rows.Add(janc, item, standard, stock, rep)
        Next

        Using sfd As SaveFileDialog = New SaveFileDialog
            'デフォルトのファイル名を指定します
            sfd.FileName = "output.csv"

            If sfd.ShowDialog() = DialogResult.OK Then
                Using writer As New StreamWriter(sfd.FileName, False, Encoding.GetEncoding("shift_jis"))

                    Dim rowCount As Integer = DataGridView2.Rows.Count
                    Dim ColumnCount As Integer = DataGridView2.Columns.Count
                    ' ユーザによる行追加が許可されている場合は、最後の新規入力用の
                    ' 1行分を差し引く
                    If (DataGridView2.AllowUserToAddRows = True) Then
                        rowCount = rowCount - 1
                    End If

                    Dim strList1 As New List(Of String)
                    For i = 0 To DataGridView2.ColumnCount - 1
                        strList1.Add(DataGridView2.Columns(i).HeaderText.ToString)
                    Next
                    Dim strary1 As String() = strList1.ToArray
                    Dim strCsvData1 As String = String.Join(",", strary1)
                    writer.WriteLine(strCsvData1)

                    ' 行
                    For i As Integer = 0 To rowCount - 1
                        ' リストの初期化
                        Dim strList As New List(Of String)
                        ' 列
                        For j As Integer = 0 To ColumnCount - 1
                            strList.Add(DataGridView2(j, i).Value.ToString())
                        Next
                        Dim strArray As String() = strList.ToArray() ' 配列へ変換

                        ' CSV 形式に変換
                        Dim strCsvData As String = String.Join(",", strArray)

                        writer.WriteLine(strCsvData)
                    Next
                    MessageBox.Show("CSV ファイルを出力しました")
                End Using
            End If
        End Using

    End Sub

End Class