Imports Npgsql
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

'Imports Excel.XlHAlign

Public Enum csv商品2
    janコード = 0
    商品名
    規格
    販売元名
End Enum


Public Class エクセル
    Private DataTable1 As DataTable = New DataTable
    Private DataTable2 As DataTable = New DataTable
    Private datatable3 As DataTable
    'Dim dict As New Dictionary(Of String, String)
    Dim dict As New Hashtable
    Private Forda As String
    Private total As Integer

    Private Sub combo()

        Dim dtCombo As New DataTable
        dtCombo.Columns.Add("id")
        dtCombo.Columns.Add("name")

        Dim dtRowCombo As DataRow
        dtRowCombo = dtCombo.NewRow
        dtRowCombo("id") = ""
        dtRowCombo("name") = ""
        dtCombo.Rows.Add(dtRowCombo)

        Using db As New dbConnection()

            Dim sql As String
            sql = ""
            sql &= "select distinct temprate_id, temprate_name from save_template "
            sql &= " order by temprate_id "


            Dim dt As DataTable = New DataTable

            dt = db.getDtSql(sql)



            For Each row As DataRow In dt.Rows
                dtRowCombo = dtCombo.NewRow
                dtRowCombo("id") = row("temprate_id")
                dtRowCombo("name") = row("temprate_name")
                dtCombo.Rows.Add(dtRowCombo)
            Next

            ComboBox1.DataSource = dtCombo
            ComboBox1.DisplayMember = "name"
            ComboBox1.ValueMember = "id"

        End Using

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim FilePath As String = String.Empty
        Dim SelectFile As String = String.Empty
        Dim Ret As DialogResult

        DataTable1.Columns.Clear()
        DataTable1.Rows.Clear()

        Try

            Using Dialog As New OpenFileDialog()

                With Dialog
                    .Title = "ダイアログボックスのサンプル"
                    .CheckFileExists = True
                    .Filter = "テキストファイル|*.txt;*.csv|すべてのファイル|*.*"
                End With

                Ret = Dialog.ShowDialog()

                If Ret = DialogResult.OK Then
                    Forda = Dialog.FileName
                    TextBox1.Text = Dialog.SafeFileName

                    If TextBox1.Text.EndsWith(".csv") Or TextBox1.Text.EndsWith(".tsv") Or TextBox1.Text.EndsWith(".txt") Then
                    Else
                        TextBox1.Text = ""
                        DataGridView1.Columns.Clear()
                        DataGridView1.Rows.Clear()
                        DataGridView2.Rows.Clear()
                        MessageBox.Show("エラー")
                        Exit Sub
                    End If
                    Me.Label1.Visible = True
                End If

            End Using

        Catch ex As Exception
            Label1.Text = "[" & System.Reflection.MethodBase.GetCurrentMethod.Name & "]" & ex.Message
        End Try

        '読み取り
        If Forda = Nothing Then
            MessageBox.Show("ファイルが選択されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ComboBox1.Text = String.Empty


        'If Not DataGridView1.Rows.Count = Nothing Then
        DataGridView1.DataSource = Nothing
            DataGridView1.Columns.Clear()
            'End If

            If Not DataGridView2.Rows.Count = Nothing Then
            DataGridView2.Rows.Clear()
        End If

        If Shadow > 0 Then
            DataGridView1.DataSource = Nothing
        End If
        Dim SR As New StreamReader(Forda, System.Text.Encoding.GetEncoding("shift_jis")) 'StreamReader文字化け防止

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        ' 重い処理　
        'System.Threading.Thread.Sleep(10000)

        Dim line As String

        Using db As New dbConnection()

            Dim i As Integer
            Dim s As Integer = 0

            If i = 0 Then
                line = SR.ReadLine()

                If line = Nothing Then
                    Exit Sub
                End If

                Dim Item() As String = filecheck(line)



                DataGridView2.Visible = True

                For Each v In Item
                    DataGridView2.Rows.Add(Item(s))
                    DataGridView2.Rows(s).Cells(1).Value = "文字列"
                    DataGridView2.Rows(s).Cells(2).Value = True
                    DataGridView1.Columns.Add("row_count", Item(s))
                    DataTable1.Columns.Add(Item(s))
                    s += 1
                Next

                i += 1

            End If

            Do

                If i > 0 Then

                    line = SR.ReadLine()
                    If line = Nothing Then
                        Exit Do
                    End If

                    Dim Item() As String = filecheck(line)

                    DataGridView1.Rows.Add(Item)

                    DataTable1.Rows.Add(Item)


                    i += 1

                End If

            Loop

        End Using


        DataGridView1.AllowUserToAddRows = False
        DataGridView2.AllowUserToAddRows = False


        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        total = DataGridView1.Rows.Count

        For center As Integer = 0 To DataGridView2.Columns.Count - 1
            DataGridView2.Columns(center).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next

        For center As Integer = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(center).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next


        DataGridView3.DataSource = DataTable1
        For center As Integer = 0 To DataGridView3.Columns.Count - 1
            DataGridView3.Columns(center).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
    End Sub

    Private Sub エクセル_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView3.Visible = False

        DataGridView2.Columns.Add("row_count", "行数")


        Dim colCombo As New DataGridViewComboBoxColumn
        colCombo.Name = "形式"
        colCombo.HeaderText = "形式"
        colCombo.Width = "100"
        colCombo.Items.Add("文字列")
        colCombo.Items.Add("数値")
        colCombo.Items.Add("数値(小数点二桁)")
        colCombo.Items.Add("日付")
        DataGridView2.Columns.Add(colCombo)


        Dim colCheck As New DataGridViewCheckBoxColumn

        colCheck.Name = "select"
        colCheck.HeaderText = "選択"
        colCheck.Width = 40
        DataGridView2.Columns.Add(colCheck)

        Dim colCheck2 As New DataGridViewCheckBoxColumn
        colCheck2.Name = "select"
        colCheck2.HeaderText = "選択"
        colCheck2.Width = 40
        DataGridView2.Columns.Add(colCheck2)


        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill


        combo()

    End Sub




    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ComboBox1.Text = String.Empty Then
            MessageBox.Show("テンプレート名を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim sql As String

        Dim ItemName1 As String
        Dim itemName As String
        Dim format As String
        Dim outPut As Boolean
        Dim sum As Boolean



        Using db As New dbConnection()
            sql = ""
            sql &= " select item_name,temprate_id , temprate_name from save_template "
            sql &= " where temprate_name = '" & ComboBox1.Text & "' "
            sql &= " order by temprate_id  desc "


            Dim dt2 As DataTable = New DataTable

            dt2 = db.getDtSql(sql)

            If dt2.Rows.Count = 0 Then

            End If


            Dim templateId As Integer = GetTemplate()


            If dt2.Rows.Count = 0 Then
                'テンプレ登録
                For i As Integer = 0 To DataGridView2.Rows.Count - 1

                    ItemName1 = DataGridView1.Columns(i).HeaderCell.Value
                    itemName = DataGridView2.Rows(i).Cells(0).Value
                    format = DataGridView2.Rows(i).Cells(1).Value
                    outPut = DataGridView2.Rows(i).Cells(2).Value
                    sum = DataGridView2.Rows(i).Cells(3).Value

                    Dim item_id As Integer = ID(i)


                    If item_id = 0 Then
                        item_id = No()
                    End If

                    sql = ""
                    sql &= " insert"
                    sql &= " into save_template"
                    sql &= " (temprate_id, temprate_name, item_name, format, output_target, total_target,item_id)"
                    sql &= " values"
                    sql &= " (" & templateId & ", '" & ComboBox1.Text & "', '" & itemName & "', '" & format & "', " & outPut & "," & sum & "," & item_id & ")"

                    'SQL実行

                    dt2 = db.getDtSql(sql)


                Next


            Else
                'テンプレ更新
                For i As Integer = 0 To DataGridView2.Rows.Count - 1

                    ItemName1 = DataGridView1.Columns(i).HeaderCell.Value
                    itemName = DataGridView2.Rows(i).Cells(0).Value
                    format = DataGridView2.Rows(i).Cells(1).Value
                    outPut = DataGridView2.Rows(i).Cells(2).Value
                    sum = DataGridView2.Rows(i).Cells(3).Value

                    Dim item_id As Integer = ID2(i)


                    'Dim sql As String
                    sql = ""
                    sql &= " update save_template"
                    sql &= " set item_name ='" & itemName & "',"
                    sql &= " format = '" & format & "',"
                    sql &= " output_target = " & outPut & ","
                    sql &= " total_target = " & sum
                    sql &= " where temprate_name = '" & ComboBox1.Text & "'"
                    sql &= " and item_id = '" & item_id & "'"

                    'SQL実行

                    dt2 = db.getDtSql(sql)


                Next
            End If


            MessageBox.Show("完了しました")
            combo()

        End Using


    End Sub


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If Forda = Nothing Then
            MessageBox.Show("ファイルが選択されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If




        If CheckBox1.Checked = False Then
            DataGridView2.Columns(2).ReadOnly = False
            DataGridView1.Visible = False
            DataGridView3.Visible = True
            DataGridView3.DataSource = DataTable1
            DataGridView3.AllowUserToAddRows = False
        End If
        Try


            If CheckBox1.Checked = True Then


                DataGridView2.Columns(2).ReadOnly = True


                DataTable2.Columns.Clear()
                DataTable2.Rows.Clear()


                DataGridView3.Visible = False
                DataGridView1.Visible = True

                Dim SR As New StreamReader(Forda, System.Text.Encoding.GetEncoding("shift_jis")) 'StreamReader文字化け防止


                Dim count1 As Integer = total

                'If DataGridView1.Rows.Count = 0 Then
                '    count1 = total
                'End If

                'If DataGridView1.Rows.Count > 0 Then
                '    count1 = DataGridView1.Rows.Count
                'End If


                Dim count2 As Integer = DataGridView2.Rows.Count - 1



                Dim line(count1) As String

                For aa As Integer = 0 To count1
                    line(aa) = SR.ReadLine()
                Next


                Dim a As Integer = line.Count

                'Dim split() As String = line(0).Split(vbTab)
                'Dim split() As String = line(0).Split(",")

                Dim Item() As String

                If TextBox1.Text.EndsWith(".csv") Then
                    Item = line(0).Split(",")
                ElseIf TextBox1.Text.EndsWith(".tsv") Then
                    Item = line(0).Split(vbTab)
                ElseIf TextBox1.Text.EndsWith(".txt") Then
                    Item = line(0).Split(vbTab)
                End If

                'If DataGridView1.Rows.Count > 0 Then
                'DataGridView1.Rows.Clear()
                'DataGridView1.Columns.Clear()
                '    DataGridView1.Rows.Clear()
                'End If


                'If Shadow > 0 Then
                DataGridView1.DataSource = Nothing
                    DataGridView1.Columns.Clear()
                    'End If

                    'DataGridView1.Rows.Clear()

                    Dim k As Integer = 0

                For i = 0 To DataGridView2.Rows.Count - 1
                    If DataGridView2.Rows(i).Cells(2).Value = True Then
                        DataGridView1.Columns.Add("row_count", Item(i))
                        DataTable2.Columns.Add(Item(i))
                        For t = 1 To count1
                            Dim Split() As String = line(t).Split(",")

                            If TextBox1.Text.EndsWith(".csv") Then
                                Split = line(t).Split(",")
                            ElseIf TextBox1.Text.EndsWith(".tsv") Then
                                Split = line(t).Split(vbTab)
                            ElseIf TextBox1.Text.EndsWith(".txt") Then
                                Split = line(t).Split(vbTab)
                            End If

                            If k = 0 Then
                                DataGridView1.Rows.Add(Split(i))
                                DataTable2.Rows.Add(Split(i))
                            End If

                            DataGridView1.Rows(t - 1).Cells(k).Value = Split(i)
                            DataTable2.Rows(t - 1)(k) = Split(i)
                        Next
                        k += 1
                    End If
                Next

            Else
            End If

            For center As Integer = 0 To DataGridView1.Columns.Count - 1
                DataGridView1.Columns(center).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Next

            DataTable1.EndInit()
        'DataGridView3.DataSource = DataTable2

        'DataTable2 = DataGridView1.DataSource

        Catch ex As Exception
        Dim a As Integer = DataGridView1.Rows.Count
        Dim b As Integer = DataGridView1.Columns.Count

        MessageBox.Show("例外なエラー")
        Exit Sub
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        dict.Clear()

        If Forda = Nothing Then
            MessageBox.Show("ファイルが選択されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If ComboBox1.Text = String.Empty Then
            MessageBox.Show("テンプレート名を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim total2 As Integer = DataGridView2.Rows.Count


        Dim sql As String


        Using db As New dbConnection()
            For s As Integer = 0 To total2 - 1


                Dim h As String = GetItemname(s)


                If Not h = Nothing Then
                    sql = ""
                    sql &= "select * from save_template "
                    sql &= "where item_name = '" & h & "' "
                    sql &= " and "
                    sql &= " temprate_id =  " & ComboBox1.SelectedIndex
                Else
                    sql = ""
                    sql &= "select * from save_template "
                    sql &= "where item_name = '" & DataGridView2.Rows(s).Cells(0).Value & "' "
                    sql &= " and "
                    sql &= " temprate_id =  " & ComboBox1.SelectedIndex
                End If


                Dim dt As DataTable = New DataTable

                dt = db.getDtSql(sql)


                If dt.Rows.Count = 0 Then
                    DataGridView2.Rows(s).Cells(0).Value = DataGridView2.Rows(s).Cells(0).Value
                    DataGridView2.Rows(s).Cells(1).Value = String.Empty
                    DataGridView2.Rows(s).Cells(2).Value = False
                    DataGridView2.Rows(s).Cells(3).Value = False
                End If

                If dt.Rows.Count > 0 Then

                    Dim row As DataRow = dt.Rows(0)
                    Dim item_name As String = row("item_name")
                    Dim format As String = row("format")
                    Dim output As Boolean = row("output_target")
                    Dim sum As Boolean = row("total_target")


                    DataGridView2.Rows(s).Cells(0).Value = row("item_name")
                    DataGridView2.Rows(s).Cells(1).Value = row("format")
                    DataGridView2.Rows(s).Cells(2).Value = row("output_target")
                    DataGridView2.Rows(s).Cells(3).Value = row("total_target")


                End If

            Next

        End Using
    End Sub

    Private Shadow As Integer
    Private dark As Integer


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        'If CheckBox1.Checked = False Then
        '    MessageBox.Show("出力対象にチェックしてください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    Exit Sub
        'End If

        DataGridView3.Visible = True
        DataGridView1.Visible = False




        If Shadow = 0 Then
            datatable3 = DataTable1.Copy
        End If


        If CheckBox1.Checked = True Then
            DataTable1 = DataTable2
            DataGridView3.DataSource = DataTable2
        End If

        DataGridView3.AllowUserToAddRows = False

        If Shadow > 0 Then
            DataGridView1.DataSource = Nothing
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = DataTable1
        End If



        Shadow += 1

        For f As Integer = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Rows(f).Cells(1).Value = Nothing Then
                MessageBox.Show("形式を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        Next

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If Forda = String.Empty Then
            MessageBox.Show("ファイルが選択されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        ' EXCEL関連オブジェクトの定義
        Dim objExcel As Excel.Application = New Excel.Application
        Dim objWorkBook As Excel.Workbook = objExcel.Workbooks.Add
        Dim objSheet As Excel.Worksheet = Nothing

        '現在日時を取得
        Dim timestanpText As String = Format(Now, "yyyyMMddHHmmss")

        '保存ディレクトリとファイル名を設定
        Dim saveFileName As String
        saveFileName = objExcel.GetSaveAsFilename(
            InitialFilename:="ファイル名_" & timestanpText,
            FileFilter:="Excel File (*.xlsx),*.xlsx")

        '保存先ディレクトリの設定が有効の場合はブックを保存
        If saveFileName <> "False" Then
            objWorkBook.SaveAs(Filename:=saveFileName)
        End If

        Dim columnMaxNum As Integer = DataGridView1.Columns.Count - 1

        Dim rowMaxNum As Integer = DataGridView1.Rows.Count - 1



        Dim x As Integer = 0
        Dim y As Integer = 1
        Dim total As Decimal
        Dim total2 As Decimal
        Dim Value As Integer
        Dim value2 As String
        Dim value3 As Decimal

        Dim value5 As Date

        'Value = CInt(DataGridView1.Rows(0).Cells(5).Value)
        Try


            For b As Integer = 0 To DataGridView2.Rows.Count - 1
                For a As Integer = 2 To DataGridView1.Rows.Count + 1
                    If DataGridView2.Rows(x).Cells(1).Value = "文字列" Then
                        objWorkBook.Sheets(1).Cells(a, y).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    End If
                    If DataGridView2.Rows(x).Cells(1).Value = "数値" Then
                        If DataGridView2.Rows(x).Cells(3).Value = True Then
                            total = CInt(DataGridView1.Rows(a - 2).Cells(y - 1).Value)
                            total2 += total
                        End If
                        If CheckBox1.Checked = True Then
                            If DataGridView2.Rows(x).Cells(2).Value = False Then
                                Exit For
                            End If
                        End If
                        Value = CInt(DataGridView1.Rows(a - 2).Cells(y - 1).Value)
                        DataGridView1.Rows(a - 2).Cells(y - 1).Value = String.Format("{0:#,0}", Value)
                        objWorkBook.Sheets(1).Cells(a, y).Value = Value
                        objWorkBook.Sheets(1).Cells(a, y).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    End If
                    If DataGridView2.Rows(x).Cells(1).Value = "数値(小数点二桁)" Then
                        If DataGridView2.Rows(x).Cells(3).Value = True Then
                            total = Decimal.Parse(DataGridView1.Rows(a - 2).Cells(y - 1).Value)
                            total2 += total
                        End If
                        If CheckBox1.Checked = True Then
                            If DataGridView2.Rows(x).Cells(2).Value = False Then
                                Exit For
                            End If
                        End If
                        value3 = Decimal.Parse(DataGridView1.Rows(a - 2).Cells(y - 1).Value)
                        DataGridView1.Rows(a - 2).Cells(y - 1).Value = value3.ToString("0.00")
                        objWorkBook.Sheets(1).Cells(a, y).Value = value3
                        objWorkBook.Sheets(1).Cells(a, y).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    End If
                    If DataGridView2.Rows(x).Cells(1).Value = "日付" Then
                        'value2 = DataGridView1.Rows(a - 2).Cells(y - 1).Value
                        If CheckBox1.Checked = True Then
                            If DataGridView2.Rows(x).Cells(2).Value = False Then
                                Exit For
                            End If
                            If DataGridView2.Rows(x).Cells(2).Value = True Then
                                If DataGridView1.Rows(a - 2).Cells(y - 1).Value.length = 8 Then
                                    value2 = DataGridView1.Rows(a - 2).Cells(y - 1).Value.Insert(6, "/").Insert(4, "/")
                                    DataGridView1.Rows(a - 2).Cells(y - 1).Value = value2
                                    objWorkBook.Sheets(1).Cells(a, y).Value = value2
                                End If
                            End If
                        End If
                        'If DataGridView1.Rows(a - 2).Cells(y - 1).Value.length = 8 Then
                        'value5 = DataGridView1.Rows(a - 2).Cells(y - 1).Value.Insert(6, "/").Insert(4, "/")
                        value5 = DataGridView1.Rows(a - 2).Cells(y - 1).Value
                        value2 = Strings.Format(value5, "yyyy/MM/dd")

                        DataGridView1.Rows(a - 2).Cells(y - 1).Value = value2
                        objWorkBook.Sheets(1).Cells(a, y).Value = value2
                        'End If
                        objWorkBook.Sheets(1).Cells(a, y).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    End If
                Next

                If DataGridView2.Rows(x).Cells(3).Value = True Then
                    objWorkBook.Sheets(1).Cells(rowMaxNum + 3, y).Borders.LineStyle = True
                    ' 項目の表示行に背景色を設定
                    objWorkBook.Sheets(1).Cells(rowMaxNum + 3, y).Interior.Color = RGB(140, 140, 140)
                    ' 文字のフォントを設定
                    objWorkBook.Sheets(1).Cells(rowMaxNum + 3, y).Font.Color = RGB(255, 255, 255)
                    objWorkBook.Sheets(1).Cells(rowMaxNum + 3, y).Font.Bold = True
                    objWorkBook.Sheets(1).Cells(rowMaxNum + 3, y).value = "合計"

                    objWorkBook.Sheets(1).Cells(rowMaxNum + 4, y).value = total2
                End If

                x += 1
                If CheckBox1.Checked = False Then
                    y += 1
                End If
                If CheckBox1.Checked = True Then
                    'If DataGridView2.Rows(x - 1).Cells(3).Value = True Then
                    y += 1
                        'End If
                    End If
            Next



        Catch ex As Exception
            MessageBox.Show("変換できません")
        End Try

        'シートの最大表示列項目数

        Dim columnList As New List(Of String)
        '項目名を取得
        For i As Integer = 0 To DataGridView2.Rows.Count - 1
            If CheckBox1.Checked = True Then
                If DataGridView2.Rows(i).Cells(2).Value = True Then
                    columnList.Add(DataGridView2.Rows(i).Cells(0).Value)
                End If
            End If

            If CheckBox1.Checked = False Then
                columnList.Add(DataGridView2.Rows(i).Cells(0).Value)
            End If
        Next

        'セルのデータ取得用二次元配列を宣言
        Dim v As String(,) = New String(rowMaxNum, columnMaxNum) {}


        For row As Integer = 0 To rowMaxNum
            For col As Integer = 0 To columnMaxNum
                'If DataGridView1.Rows(row).Cells(col).Value Is Nothing = False Then
                ' セルに値が入っている場合、二次元配列に格納
                v(row, col) = DataGridView1.Rows(row).Cells(col).Value.ToString()
                'End If
            Next
        Next


        For i As Integer = 1 To DataGridView1.Columns.Count
            ' シートの一行目に項目を挿入
            objWorkBook.Sheets(1).Cells(1, i) = columnList(i - 1)


            ' 罫線を設定
            objWorkBook.Sheets(1).Cells(1, i).Borders.LineStyle = True
            ' 項目の表示行に背景色を設定
            objWorkBook.Sheets(1).Cells(1, i).Interior.Color = RGB(140, 140, 140)
            ' 文字のフォントを設定
            objWorkBook.Sheets(1).Cells(1, i).Font.Color = RGB(255, 255, 255)
            objWorkBook.Sheets(1).Cells(1, i).Font.Bold = True

        Next


        Dim data As String = ""

        If columnMaxNum <= 25 Then
            data = "A2:" & Chr(Asc("A") + columnMaxNum) & DataGridView1.Rows.Count + 1
        End If
        If columnMaxNum > 25 Then
            data = "A2:" & ConvertToLetter(columnMaxNum) & DataGridView1.Rows.Count + 1
        End If


        objWorkBook.Sheets(1).Range(data) = v

        ' データの表示範囲に罫線を設定
        objWorkBook.Sheets(1).Range(data).Borders.LineStyle = True

        ' エクセル表示
        objExcel.Visible = True

        DataTable1 = datatable3.Copy

        'If Shadow > 0 Then
        '    DataGridView1.DataBindings. = False
        'End If
        'CheckBox1.Checked = False
        'For s As Integer = 0 To DataGridView2.Rows.Count - 1
        '    DataGridView2.Rows(s).Cells(2).Value = True
        'Next

        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub

    Function ConvertToLetter(iCol As Long) As String
        Dim a As Long
        Dim b As Long
        a = iCol
        ConvertToLetter = ""
        Do While iCol > 0
            a = Int((iCol - 1) / 25)
            b = (iCol - 1) Mod 25
            ConvertToLetter = Chr(b + 65) & ConvertToLetter
            iCol = a
        Loop
    End Function

    Private Function filecheck(line)
        Dim Item() As String

        If TextBox1.Text.EndsWith(".csv") Then
            Item = line.Split(",")
            Return Item
        ElseIf TextBox1.Text.EndsWith(".tsv") Then
            Item = line.Split(vbTab)
            Return Item
        ElseIf TextBox1.Text.EndsWith(".txt") Then
            Item = line.Split(vbTab)
            Return Item
        Else
            MessageBox.Show("その拡張子は読み取れません")
            Exit Function
        End If

    End Function


    Private Function GetItemname(s As Integer)
        Using db As New dbConnection()


            Dim sql As String
            sql = ""
            sql &= " select item_name, item_id  from save_template "
            sql &= " where temprate_name = '" & ComboBox1.Text & "'  "
            sql &= " order by id"

            Dim dt As DataTable = New DataTable

            dt = db.getDtSql(sql)


            'Dim row As DataRow
            Dim item_id As Integer

            Dim Item2 As Integer = Item(s)

            Dim i As Integer

            For Each row As DataRow In dt.Rows
                row = dt.Rows(i)
                item_id = row("item_id")
                If DataGridView1.Columns(s).HeaderCell.Value = DataGridView2.Rows(s).Cells(0).Value Then
                    If Not DataGridView1.Columns(s).HeaderCell.Value = row("item_name") Then
                        If item_id = Item2 Then
                            If Not dict(DataGridView1.Columns(s).HeaderCell.Value) = row("item_name") Then
                                dict.Add(DataGridView1.Columns(s).HeaderCell.Value, row("item_name"))
                                Return row("item_name")
                            End If
                        End If
                    End If
                End If
                i += 1
            Next
        End Using
    End Function

    Function Item(s As Integer)
        Using db As New dbConnection()
            Dim sql As String

            sql = ""
            sql &= " select distinct item_id  from save_template "
            sql &= " where item_name = '" & DataGridView2.Rows(s).Cells(0).Value & "' "
            sql &= " and "
            sql &= " temprate_name = '" & ComboBox1.Text & "' "

            Dim dt As DataTable = New DataTable

            dt = db.getDtSql(sql)

            Dim row As DataRow

            Dim item_id As Integer
            If dt.Rows.Count > 0 Then
                row = dt.Rows(0)
                item_id = row("item_id")
                Return item_id
            End If


        End Using
    End Function

    Private Function ID(s As Integer)

        Using db As New dbConnection()
            Dim sql As String

            sql = ""
            sql &= " select item_id from save_template "
            sql &= " where item_name = '" & DataGridView1.Columns(s).HeaderCell.Value & "' "
            sql &= " and "
            sql &= " temprate_name = '" & ComboBox1.Text & "' "

            Dim dt As DataTable = New DataTable

            dt = db.getDtSql(sql)

            Dim row As DataRow

            If dt.Rows.Count > 0 Then
                row = dt.Rows(0)
                Dim item_id As Integer = row("item_id")
                Return item_id
            End If


        End Using

    End Function

    Private Clear As Integer
    Private Function ID2(s As Integer)

        Using db As New dbConnection()
            Dim sql As String

            sql = ""
            sql &= " select item_id from save_template "
            sql &= " where item_name = '" & DataGridView2.Rows(s).Cells(0).Value & "' "
            sql &= " and "
            sql &= " temprate_name = '" & ComboBox1.Text & "' "

            Dim dt As DataTable = New DataTable

            dt = db.getDtSql(sql)

            Dim row As DataRow

            If dt.Rows.Count > 0 Then
                row = dt.Rows(0)
                Dim item_id As Integer = row("item_id")
                Return item_id
            End If


            Clear += 1

            If dt.Rows.Count = 0 Then
                Return Clear
            End If

        End Using

    End Function

    Private Function No()

        Using db As New dbConnection()
            Dim sql As String

            sql = ""
            sql &= " select item_id from save_template "
            sql &= " where temprate_name = '" & ComboBox1.Text & "' "
            sql &= " order by item_id desc "


            Dim dt As DataTable = New DataTable

            dt = db.getDtSql(sql)

            Dim row As DataRow
            Dim item_id As Integer = 0

            If dt.Rows.Count > 0 Then
                row = dt.Rows(0)
                item_id = row("item_id")
                Return item_id + 1
            End If

            If dt.Rows.Count = 0 Then
                Dim i As Integer = 1
                Return i
            End If

        End Using
    End Function

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If CheckBox1.Checked = True Then
            MessageBox.Show("編集できません。出力対象チェックを外してから再度お試しください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub



    Function GetTemplate()

        Dim sql As String
        sql = ""
        sql &= " select distinct"
        sql &= " temprate_id,"
        sql &= " temprate_name"
        sql &= " from save_template"
        sql &= " order by temprate_id desc"

        'SQL実行
        Dim dt As DataTable = New DataTable()

        Using db As New dbConnection()
            dt = db.getDtSql(sql)
        End Using

        Dim rowCount As Integer = dt.Rows.Count
        Dim temprateId As Integer

        If rowCount = 0 Then
            temprateId = 1
        Else
            temprateId = dt.Rows(0).Item("temprate_id") + 1
        End If

        Return temprateId

    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox1.Text = String.Empty Then
            MessageBox.Show("テンプレート名を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim sql As String

        Dim ItemName1 As String
        Dim itemName As String
        Dim format As String
        Dim outPut As Boolean
        Dim sum As Boolean
        Dim temprateId As Integer
        Using db As New dbConnection()

            sql = ""
            sql &= " select item_name,temprate_id , temprate_name from save_template "
            sql &= " where temprate_name = '" & ComboBox1.Text & "' "
            sql &= " order by temprate_id  desc "


            Dim dt2 As DataTable = New DataTable

            dt2 = db.getDtSql(sql)




            If dt2.Rows.Count = 0 Then
                sql = ""
                sql &= " select item_name,temprate_id from save_template order by temprate_id desc  "

                Dim dt As DataTable = New DataTable

                dt = db.getDtSql(sql)

                If dt.Rows.Count = 0 Then
                    temprateId = 1
                End If

                If dt.Rows.Count > 0 Then
                    Dim row As DataRow = dt.Rows(0)
                    'item_name = row("item_name")
                    temprateId = row("temprate_id") + 1
                End If


            End If

            If dt2.Rows.Count > 0 Then
                Dim row As DataRow = dt2.Rows(0)
                temprateId = row("temprate_id")
            End If


            For s As Integer = 0 To DataGridView2.Rows.Count - 1

                itemName = DataGridView2.Rows(s).Cells(0).Value
                format = DataGridView2.Rows(s).Cells(1).Value
                outPut = DataGridView2.Rows(s).Cells(2).Value
                sum = DataGridView2.Rows(s).Cells(3).Value

                ItemName1 = DataGridView1.Columns(s).HeaderCell.Value

                Dim item_id As Integer = ID(s)


                If item_id = 0 Then
                    item_id = No()
                End If


                If dt2.Rows.Count = 0 Then

                    sql = ""
                    sql &= " insert into save_template "
                    sql &= " (item_name, format, output_target, total_target, temprate_id, temprate_name, item_id ) "
                    sql &= " values('" & itemName & "', '" & format & "', " & outPut & ", " & sum & "," & temprateId & ", '" & ComboBox1.Text & "'," & item_id & ") "

                    db.executeSql(sql)

                End If

                If dt2.Rows.Count > 0 Then


                    sql = ""
                    sql &= " update save_template set (item_name, format, output_target, total_target, temprate_id, temprate_name,item_id) "
                    sql &= " = "
                    sql &= "('" & itemName & "', '" & format & "', " & outPut & ", " & sum & "," & temprateId & ", '" & ComboBox1.Text & "'," & item_id & ")"
                    sql &= "where item_name = '" & itemName & "' "
                    sql &= " and "
                    sql &= " temprate_id = " & temprateId


                    db.executeSql(sql)


                    If Not itemName = ItemName1 Then

                        sql = ""
                        sql &= " select item_id from save_template where item_name = '" & itemName & "' "
                        sql &= " and "
                        sql &= " temprate_name = '" & ComboBox1.Text & "' "

                        Dim dt As DataTable = New DataTable

                        dt = db.getDtSql(sql)

                        If dt.Rows.Count = 0 Then
                            sql = ""
                            sql &= "insert into save_template "
                            sql &= " (item_name, format, output_target, total_target, temprate_id,temprate_name, item_id) "
                            sql &= " values('" & itemName & "', '" & format & "', " & outPut & ", " & sum & "," & temprateId & ", '" & ComboBox1.Text & "'," & item_id & ") "

                            db.executeSql(sql)


                        End If

                        If dt.Rows.Count > 0 Then
                            sql = ""
                            sql &= " update save_template set (item_name, format, output_target, total_target, temprate_id, temprate_name,item_id) "
                            sql &= " = "
                            sql &= "('" & itemName & "', '" & format & "', " & outPut & ", " & sum & "," & temprateId & ", '" & ComboBox1.Text & "'," & item_id & ")"
                            sql &= "where item_name = '" & itemName & "' "
                            sql &= " and "
                            sql &= " temprate_id = " & temprateId

                            db.executeSql(sql)


                        End If
                    End If
                End If

            Next

            MessageBox.Show("完了しました")
            combo()

        End Using
    End Sub
End Class