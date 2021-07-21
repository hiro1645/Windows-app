
Public Class 購入画面
    Private area As String
    Private Buyitem As Integer
    Private Buyitem2 As Integer
    Private Buyitem3 As Integer
    Private Buyitem4 As Integer
    Private Buyitem5 As Integer
    Private Buyitem6 As Integer
    Private Buyitem7 As Integer
    Private Buyitem8 As Integer
    Private Buyitem9 As Integer

    Private Buyprice As Integer
    Private Buyprice2 As Integer
    Private Buyprice3 As Integer
    Private Buyprice4 As Integer
    Private Buyprice5 As Integer
    Private Buyprice6 As Integer
    Private Buyprice7 As Integer
    Private Buyprice8 As Integer
    Private Buyprice9 As Integer

    Private oturi As Integer

    Private htInput As New Hashtable

    Private After As Integer = 1
    Private back As Integer = 0

    Private Sub First()
        Place()

        Dim sql As String
        sql = ""
        sql &= "SELECT "
        sql &= "a.id, "
        sql &= "a.item_name, "
        sql &= "b.price "
        sql &= "FROM mst_items as a "
        sql &= "Left join mst_prices as b "
        sql &= "on a.id = b.mst_item_id "
        sql &= "Left Join trn_stocks as c "
        sql &= "On a.id = c.id "
        sql &= "WHERE start_date <= CURRENT_DATE "
        sql &= " And CURRENT_DATE <= end_date "
        sql &= "and place_id = " & After

        sql &= "order by id"

        Dim dt As DataTable = New DataTable()

        'getDtSqlの場合
        Using db As New dbConnection()

            dt = db.getDtSql(sql)

        End Using
        ' ボタン表示用変数
        Dim row As DataRow

        row = dt.Rows(0)
        Buyprice = row("price")
        Button1.Text = row("item_name")
        Buyitem = row("id")

        row = dt.Rows(1)
        Buyprice2 = row("price")
        Button2.Text = row("item_name")
        Buyitem2 = row("id")

        row = dt.Rows(2)
        Buyprice3 = row("price")
        Button3.Text = row("item_name")
        Buyitem3 = row("id")

        row = dt.Rows(3)
        Buyprice4 = row("price")
        Button4.Text = row("item_name")
        Buyitem4 = row("id")

        row = dt.Rows(4)
        Buyprice5 = row("price")
        Button5.Text = row("item_name")
        Buyitem5 = row("id")

        row = dt.Rows(5)
        Buyprice6 = row("price")
        Button6.Text = row("item_name")
        Buyitem6 = row("id")

        row = dt.Rows(6)
        Buyprice7 = row("price")
        Button7.Text = row("item_name")
        Buyitem7 = row("id")

        row = dt.Rows(7)
        Buyprice8 = row("price")
        Button8.Text = row("item_name")
        Buyitem8 = row("id")

        row = dt.Rows(8)
        Buyprice9 = row("price")
        Button9.Text = row("item_name")
        Buyitem9 = row("id")


    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        First()

        Label1.Text = 0
        Label2.Text = 0
        Button10.Text = "10円"
        Button11.Text = "50円"
        Button12.Text = "100円"
        Button13.Text = "500円"
        Button14.Text = "1000円"

    End Sub

    Public Sub Buysystem(ByVal id As Integer, ByVal Buyprice As Integer)

        Try
            'id,商品名,料金,在庫数を取得する処理
            Dim sql As String
            sql = ""
            sql &= "select "
            sql &= "a.id, "
            sql &= "a.item_name, "
            sql &= "b.price, "
            sql &= "c.stock_count "
            sql &= "from mst_items as a "
            sql &= "left join mst_prices as b "
            sql &= "on a.id = b.mst_item_id "
            sql &= "left join trn_stocks as c "
            sql &= "on a.id = c.mst_item_id "
            sql &= "where c.place_id = " & After
            sql &= " and "
            sql &= " a.id = " & id
            sql &= " order by a.id asc "

            Dim dt As DataTable = New DataTable

            'getDtSqlの場合
            Using db As New dbConnection()

                dt = db.getDtSql(sql)

            End Using

            Dim row As DataRow

            ' ボタン１
            row = dt.Rows(0)

            Dim Name As String = row("item_name")
            Dim price As Integer = row("price")
            Dim stock As Integer = row("stock_count")
            '料金が投入金額より大きいとき
            If price > Label1.Text Then
                MessageBox.Show("買えません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            '在庫が0の時
            If stock = 0 Then
                MessageBox.Show("売り切れです", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If


            '在庫があり料金が投入金額より少ないとき
            If stock > 0 And Label1.Text >= price Then
                oturi = CInt(Label2.Text)
            End If

            Using db As New dbConnection()
                Try

                    db.trnStart()


                    Dim cmd As New Npgsql.NpgsqlCommand
                    'データベースのストックを１減らす
                    sql = ""
                    sql &= "UPDATE "
                    sql &= "trn_stocks "
                    sql &= " SET "
                    sql &= "stock_count "
                    sql &= "= "
                    sql &= "stock_count-1 "
                    sql &= "where mst_item_id=" & id
                    sql &= "and "
                    sql &= "place_id =" & After

                    db.executeSql(sql)


                    ''データベースに商品購入時に実績として追加する処理

                    sql = ""
                    sql &= "  insert into trn_claims "
                    sql &= "(place_id, mst_item_id, buy_count, buy_price, buy_date) "
                    sql &= " values(" & After & ", " & id & ", 1 , " & price & ", current_timestamp );"



                    db.executeSql(sql)




                    For Each ht As DictionaryEntry In htInput

                        sql = ""
                        sql &= "Update "
                        sql &= "money_prices "
                        sql &= "set "
                        sql &= "money_stock = "
                        sql &= " money_stock + "
                        sql &= " " & ht.Value
                        sql &= " where "
                        sql &= "money_price = "
                        sql &= " " & ht.Key

                        db.executeSql(sql)
                    Next


                    db.commit()
                Catch ex As Npgsql.PostgresException
                    db.rollback()
                    Throw
                End Try
            End Using
            If stock < 10 Then
                MessageBox.Show("在庫が少ないです", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try

        Label1.Text -= Buyprice

    End Sub

    Private Sub Place()

        Dim sql As String
        sql = ""
        sql &= " select place_name from mst_areas "

        Dim dt As DataTable = New DataTable

        'getDtSqlの場合
        Using db As New dbConnection()

            dt = db.getDtSql(sql)

        End Using

        back = (dt.Rows).Count
        Dim row As DataRow

        row = dt.Rows(After - 1)
        area = row("place_name")
        Label5.Text = area

    End Sub

    '次へ
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Try
            If After >= 0 And After < back Then
                After += 1
            ElseIf After >= back Then
                Throw New Exception
            End If
            Label5.Text = area
            First()
        Catch
            MessageBox.Show("ありません")
        End Try

    End Sub
    '前へ
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Try
            If After > 1 Then
                After -= 1
            ElseIf After = 1 Then
                Throw New Exception
            End If
            Label5.Text = area
            First()
        Catch
            MessageBox.Show("戻れません")
        End Try
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Label2.Text = Label1.Text
        Label1.Text = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Buysystem(Buyitem, Buyprice)
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Buysystem(Buyitem2, Buyprice2)
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Buysystem(Buyitem3, Buyprice3)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Buysystem(Buyitem4, Buyprice4)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Buysystem(Buyitem5, Buyprice5)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Buysystem(Buyitem6, Buyprice6)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Buysystem(Buyitem7, Buyprice7)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Buysystem(Buyitem8, Buyprice8)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Buysystem(Buyitem9, Buyprice9)
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Label1.Text += 10
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Label1.Text += 50
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Label1.Text += 100
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Label1.Text += 500
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Label1.Text += 1000
    End Sub



End Class