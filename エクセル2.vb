Imports Npgsql
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class エクセル2
    Private forda As String
    Private dt As DataTable
    Private dt2 As DataTable = New DataTable



    'Private datatable2 As DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        GetGrid()

    End Sub

    Private Function GetFile() As StreamReader
        Dim ofd As New OpenFileDialog()


        '2番目の「すべてのファイル」が選択されているようにする
        ofd.FilterIndex = 2
        'タイトルを設定する
        ofd.Title = "開くファイルを選択してください"
        'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
        ofd.RestoreDirectory = True



        'ダイアログを表示する
        If ofd.ShowDialog() = DialogResult.OK Then
            'OKボタンがクリックされたとき、選択されたファイル名を表示する
            Console.WriteLine(ofd.FileName)
        End If

        Dim SR As New StreamReader(ofd.FileName, System.Text.Encoding.GetEncoding("shift_jis")) 'StreamReader文字化け防止
        If ofd.FileName.EndsWith(".csv") Or ofd.FileName.EndsWith(".tsv") Or ofd.FileName.EndsWith(".txt") Then
        Else
            DataGridView1.DataSource = Nothing
            MessageBox.Show("エラー")
            Return Nothing
        End If
        Return SR
    End Function

    Private Sub GetGrid()
        Dim SR As StreamReader = GetFile()
        If SR Is Nothing Then
            Exit Sub
        End If
        dt = New DataTable

        Dim line As String = String.Empty
        Dim i As Integer = 0
        Dim dr As DataRow

        If i = 0 Then
            line = SR.ReadLine

            Dim Item() As String = line.Split(vbTab)

            Dim s As Integer = 0

            For Each v In Item
                dt.Columns.Add(Item(s))
                dt2.Rows.Add(Item(s))
                s += 1
            Next
            i += 1
        End If
        Do
            If i > 0 Then
                line = SR.ReadLine
                If line = Nothing Then
                    Exit Do
                End If
                Dim Item() As String = line.Split(vbTab)

                dr = dt.NewRow
                dr.ItemArray = Item
                dt.Rows.Add(dr)
            End If
        Loop
        DataGridView1.DataSource = dt
        DataGridView2.DataSource = dt2

        For s = 0 To DataGridView2.Rows.Count - 2
            DataGridView2.Rows(s).Cells(1).Value = "数値"
            DataGridView2.Rows(s).Cells(2).Value = "a"
            DataGridView2.Rows(s).Cells(3).Value = "a"
        Next

    End Sub
    Private Sub エクセル2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dt2.Columns.Add("内容")
        dt2.Columns.Add("中身")
        dt2.Columns.Add("使用")
        dt2.Columns.Add("行数")
        Me.Label1.Text = Now.ToString("hh:mm:ss")

        Me.Timer1.Interval = 1000

        Me.Timer1.Enabled = True



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' EXCEL関連オブジェクトの定義
        Dim objExcel As Excel.Application = New Excel.Application
        Dim objWorkBook As Excel.Workbook = objExcel.Workbooks.Add
        Dim objSheet As Excel.Worksheet = Nothing

        '現在日時を取得
        Dim timestanpText As String = Format(Now, "yyyyMMddHHmmss")

        Dim aa As Integer = dt.Rows.Count



        '保存ディレクトリとファイル名を設定
        Dim saveFileName As String
        saveFileName = objExcel.GetSaveAsFilename(
            InitialFilename:="ファイル名_" & timestanpText,
            FileFilter:="Excel File (*.xlsx),*.xlsx")

        '保存先ディレクトリの設定が有効の場合はブックを保存
        If saveFileName <> "False" Then
            objWorkBook.SaveAs(Filename:=saveFileName)
        End If

        'シートの最大表示列項目数
        Dim columnMaxNum As Integer = dt.Columns.Count - 1
        'シートの最大表示行項目数
        Dim rowMaxNum As Integer = dt.Rows.Count - 1

        Dim x As Integer
        Dim y As Integer


        Dim Last As String = dt.Rows(0)(0)

        Dim value As Object
        Dim total As Integer

        For a As Integer = 0 To DataGridView2.Rows.Count - 1
            x = 0
            For b As Integer = 0 To dt.Rows.Count - 1

                If DataGridView2.Rows(a).Cells(1).Value Is Nothing Then
                    Exit For
                End If

                If DataGridView2.Rows(a).Cells(1).Value = "数値" Then
                    If IsNumeric(dt.Rows(b)(y)) Then
                        If DataGridView2.Rows(a).Cells(2).Value = "合計" Then
                            total += GetSub(b, y)
                        End If
                        value = CInt(dt.Rows(b)(y))
                        Last = String.Format("{0:#,0}", value)
                        objWorkBook.Sheets(1).Cells(x + 2, y + 1).Value = Last
                    Else
                        objWorkBook.Sheets(1).Cells(x + 2, y + 1).Value = dt.Rows(b)(y)
                    End If
                End If

                If DataGridView2.Rows(a).Cells(1).Value = "小数" Then
                    If TypeOf dt.Rows(b)(y) Is Decimal Then
                        value = Decimal.Parse(dt.Rows(b)(y))
                        Last = String.Format("{0:#,0.00}", value)
                        objWorkBook.Sheets(1).Cells(x + 2, y + 1).Value = Last
                    Else
                        objWorkBook.Sheets(1).Cells(x + 2, y + 1).Value = dt.Rows(b)(y)
                    End If
                End If

                If DataGridView2.Rows(a).Cells(1).Value = "日付" Then
                    If TypeOf dt.Rows(b)(y) Is DateTime Then
                        objWorkBook.Sheets(1).Cells(x + 2, y + 1).numberformat = "yyyy/MM/dd"
                        objWorkBook.Sheets(1).Cells(x + 2, y + 1).Value = dt.Rows(b)(y)
                    Else
                        objWorkBook.Sheets(1).Cells(x + 2, y + 1).Value = dt.Rows(b)(y)
                    End If
                End If
                x += 1
            Next
            y += 1
            If total > 0 Then
                objWorkBook.Sheets(1).Cells(rowMaxNum + 3, y).value = total
                total = 0
            End If
        Next




        '項目名格納用リストを宣言
        Dim columnList As New List(Of String)
        '項目名を取得
        For i As Integer = 0 To (columnMaxNum)
            columnList.Add(dt.Columns(i).Caption)
        Next

        ''セルのデータ取得用二次元配列を宣言
        'Dim v As Object(,) = New Object(rowMaxNum, columnMaxNum) {}

        'For row As Integer = 0 To rowMaxNum
        '    For col As Integer = 0 To columnMaxNum
        '        'If dt.Rows(row)(col).Value Is Nothing = False Then
        '        ' セルに値が入っている場合、二次元配列に格納
        '        v(row, col) = dt.Rows(row)(col)
        '        'End If
        '    Next
        'Next

        ' EXCELに項目名を転送
        For i As Integer = 1 To dt.Columns.Count
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



        '' EXCELにデータを範囲指定で転送
        'Dim data As String = "A2:" & Chr(Asc("A") + columnMaxNum) & dt.Rows.Count + 1
        'objWorkBook.Sheets(1).Range(data) = v

        '' データの表示範囲に罫線を設定
        'objWorkBook.Sheets(1).Range(data).Borders.LineStyle = True

        ' エクセル表示
        objExcel.Visible = True

        ' EXCEL解放
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim dv As DataView
        dv = New DataView(dt)
        dv.Sort = "JANCODE2"
        dt = dv.ToTable
        Dim row As DataRow = dt.NewRow '追加行を宣言

        '値をセット
        row("JANCODE") = "1"
        row("JANCODE2") = "一郎"
        row("JANCODE3") = "山田"
        row("JANCODE4") = "1"
        row("日付") = "170"

        dt.Rows.InsertAt(row, 0)

        Using sfd As SaveFileDialog = New SaveFileDialog
            'デフォルトのファイル名を指定します
            sfd.FileName = "output.csv"

            If sfd.ShowDialog() = DialogResult.OK Then
                Using writer As New StreamWriter(sfd.FileName, False, Encoding.GetEncoding("shift_jis"))

                    Dim rowCount As Integer = dt.Rows.Count
                    Dim ColumnCount As Integer = dt.Columns.Count

                    Dim strList1 As New List(Of String)
                    For i = 0 To ColumnCount - 1
                        strList1.Add(dt.Columns(i).Caption)
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
                            strList.Add(dt(i)(j))
                            'strList.Add(DataGridView1.Columns(j).HeaderCell.Value)
                        Next
                        'Dim strArray As String() = strList.ToArray() ' 配列へ変換


                        ' CSV 形式に変換
                        Dim strCsvData As String = String.Join(",", strList)

                        writer.WriteLine(strCsvData)
                    Next
                    MessageBox.Show("CSV ファイルを出力しました")
                End Using
            End If


        End Using
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim total2 As Integer = DataGridView2.Rows.Count


        Dim sql As String = String.Empty


        Using db As New dbConnection()


            sql &= "select * from save_template "

                Dim dt As DataTable = New DataTable

                dt = db.getDtSql(sql)
            For s As Integer = 0 To total2 - 2

                'If dt.Rows.Count = 0 Then
                dt2.Rows(s).Item(0) = DataGridView1.Columns(0).HeaderText
                dt2.Rows(s).Item(1) = String.Empty
                dt2.Rows(s).Item(2) = String.Empty
                dt2.Rows(s).Item(3) = String.Empty
                'End If

                If dt.Rows.Count > 0 Then

                    Dim row As DataRow = dt.Rows(s)
                    Dim item_name As String = row("item_name")
                    Dim format As String = row("format")
                    Dim output As Boolean = row("output_target")
                    Dim sum As Boolean = row("total_target")


                    dt2.Rows(s).Item(0) = item_name
                    dt2.Rows(s).Item(1) = format
                    dt2.Rows(s).Item(2) = output
                    dt2.Rows(s).Item(3) = sum

                    'If IsNumeric(dt2.Rows(1).Item(0)) = False Then '数値チェック Not つかってもOK　Microsoft.VisualBasic　DLL
                    '    MessageBox.Show("完了しました")
                    'End If

                End If

            Next



        End Using
    End Sub

    Private Function GetSub(a As Integer, b As Integer)

        Dim total As Integer = 0
        If dt.Rows(a)(b) = String.Empty Then
            total = 0
        Else
            total += dt.Rows(a)(b)
        End If


        Return total


    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        'If forda = "" Then
        '    MessageBox.Show("ファイルを選択して下さい。")
        '    Exit Sub
        'End If

        Dim xlApp As Object = Nothing
        Dim xlBooks As Object = Nothing
        Dim xlBook As Object = Nothing
        Dim xlSheet As Object = Nothing
        Dim xlCells As Object = Nothing
        Dim xlRange As Object = Nothing
        Dim xlCellStart As Object = Nothing
        Dim xlCellEnd As Object = Nothing

        'エクセルシート２
        Dim xlSheet2 As Object = Nothing
        Dim xlRange22 As Object = Nothing
        Dim xlCells22 As Object = Nothing

        Try
            xlApp = CreateObject("Excel.Application")
            xlBooks = xlApp.Workbooks
            xlBook = xlApp.Workbooks.Add
            xlSheet = xlBook.WorkSheets(1)
            'xlSheet2 = xlBook.WorkSheets.add
            xlCells = xlSheet.Cells



            Dim dc As DataColumn
            Dim columnData(dt.Rows.Count, 1) As Object
            Dim row As Integer = 1
            Dim col As Integer = 1
            Dim col2 As Integer = 1
            Dim changeExcelRow As String
            Dim sumExcel As Object
            Dim xlRange2 As Object
            Dim xlRange3 As Object
            Dim exist As Integer = 0

            Dim dtc As Integer = dt.Columns.Count

            For col = 1 To dt.Columns.Count

                Dim CheckFlg As Boolean = False

                If DataGridView2.Rows(col - 1).Cells(2).Value = "a" Then

                    row = 1
                    dc = dt.Columns(col - 1)
                    'ヘッダー行の出力
                    xlCells(row, col2).value = dc.ColumnName
                    row = row + 1

                    ' 列データを配列に格納
                    For i As Integer = 0 To dt.Rows.Count - 1
                        columnData(i, 0) = String.Format(dt.Rows(i)(col - 1))
                    Next
                    xlCellStart = xlCells(row, col2)
                    xlCellEnd = xlCells(row + dt.Rows.Count - 1, col2)
                    xlRange = xlSheet.Range(xlCellStart, xlCellEnd)


                    Dim count As Integer = DataGridView2.Rows.Count
                    Dim check As Object = xlCells(row, col).value
                    Dim check2 As Object = dc.DataType
                    Dim DtCount As Integer = dt.Rows.Count


                    Dim dtt As DataTable = dt
                    Select Case DataGridView2.Rows(col - 1).Cells(1).Value
                        Case "文字列"
                            xlRange.NumberFormatLocal = "@"
                        Case "数値"
                            For b = 1 To DtCount - 1
                                If IsNumeric(dt.Rows(b)(col - 1)) Then
                                    xlRange.NumberFormatLocal = "###,##0"
                                    CheckFlg = True
                                End If
                            Next
                        Case "日付"
                            For b = 1 To DtCount - 1
                                If TypeOf dt.Rows(b)(col - 1) Is Date Then
                                    xlRange.NumberFormatLocal = "yyyy/mm/dd"
                                End If
                            Next
                        'End ig
                        Case "小数"
                            For b = 1 To DtCount - 1
                                If TypeOf dt.Rows(b)(col - 1) Is Decimal Then
                                    xlRange.NumberFormatLocal = "#,##.00"
                                End If
                            Next
                            CheckFlg = True
                    End Select

                    xlRange.value = columnData

                    changeExcelRow = ConvertToLetter(col2)
                    '合計出力
                    If CheckFlg = True Then
                        If DataGridView2.Rows(col - 1).Cells(3).Value = "a" Then
                            sumExcel = "=SUM(" & changeExcelRow & "2:" & changeExcelRow & dt.Rows.Count + 1 & ")"
                            xlRange2 = xlSheet.Cells(dt.Rows.Count + 3, changeExcelRow)
                            xlRange2.value = sumExcel
                            exist += 1
                        End If
                    End If

                    If exist > 0 Then
                        xlRange3 = xlSheet.Cells(dt.Rows.Count + 2, changeExcelRow)
                        xlRange3.value = "合計"
                    End If


                    col2 += 1

                End If

            Next
            If col2 = 1 Then
                MessageBox.Show("出力対象がありません")
                Exit Sub
            End If

            'xlSheet.Copy(Before:=xlSheet)
            'xlSheet.Previous.Name = "TEST"
            'xlSheet2 = xlSheet.copysheet

            xlCells.EntireColumn.AutoFit()
            xlRange = xlSheet.UsedRange
            xlRange.Borders.LineStyle = 1   'xlContinuous
            xlApp.Visible = True

            'xlBook.WorkSheets(3) = xlBook.WorkSheets("kk").copy



        Catch
            xlApp.DisplayAlerts = False
            xlApp.Quit()
            Throw
        Finally
            If xlCellStart IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCellStart)
            If xlCellEnd IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCellEnd)
            If xlRange IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange)
            If xlCells IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCells)
            'If xlSheet IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet)
            If xlBooks IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks)
            If xlBook IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook)
            If xlApp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)

            GC.Collect()

        End Try

    End Sub

#Region "メソッド"　' 

#End Region
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

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Me.Label1.Text = Now.ToString("hh:mm:ss")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim excelApp As New Excel.Application()
        Dim excelBooks As Excel.Workbooks
        excelBooks = excelApp.Workbooks
        excelBooks.Open("C:\Users\aokihiro\Desktop\(CUSTOM2021_B001)機能定義書(EXCEL変換アプリ).xlsx")
        excelApp.Visible = True

    End Sub


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim Wdt As DataTable = New DataTable
        Wdt.Columns.Add("ああ")
        Wdt.Columns.Add("いい")

        Wdt.Rows.Add("aa", "ii")
        Wdt.Rows.Add("aa", "ii")
        Wdt.Rows.Add("aa", "ii")
        Wdt.Rows.Add("aa", "ii")


        Dim xlApp As Object = Nothing
        Dim xlBooks As Object = Nothing
        Dim xlBook As Object = Nothing
        Dim xlSheet As Object = Nothing
        Dim xlCells As Object = Nothing
        Dim xlRange As Object = Nothing
        Dim xlCellStart As Object = Nothing
        Dim xlCellEnd As Object = Nothing

        'エクセルシート２
        Dim xlSheet2 As Object = Nothing
        Dim xlRange22 As Object = Nothing
        Dim xlCells22 As Object = Nothing
        Dim xlCellStart22 As Object = Nothing
        Dim xlCellEnd22 As Object = Nothing

        Try
            xlApp = CreateObject("Excel.Application")
            xlBooks = xlApp.Workbooks
            xlBook = xlApp.Workbooks.Add
            xlSheet = xlBook.WorkSheets(1)
            xlSheet2 = xlBook.WorkSheets.add
            xlCells = xlSheet.Cells

            Dim dc As DataColumn
            Dim columnData(dt.Rows.Count, 1) As Object
            Dim row As Integer = 1
            Dim col As Integer = 1
            Dim col2 As Integer = 1
            Dim changeExcelRow As String
            Dim sumExcel As Object
            Dim xlRange2 As Object
            Dim xlRange3 As Object
            Dim exist As Integer = 0

            Dim dtc As Integer = dt.Columns.Count

            For col = 1 To dt.Columns.Count

                Dim CheckFlg As Boolean = False

                If DataGridView2.Rows(col - 1).Cells(2).Value = "a" Then

                    row = 1
                    dc = dt.Columns(col - 1)
                    'ヘッダー行の出力
                    xlCells(row, col2).value = dc.ColumnName
                    row = row + 1

                    ' 列データを配列に格納
                    For i As Integer = 0 To dt.Rows.Count - 1
                        columnData(i, 0) = String.Format(dt.Rows(i)(col - 1))
                    Next
                    xlCellStart = xlCells(row, col2)
                    xlCellEnd = xlCells(row + dt.Rows.Count - 1, col2)
                    xlRange = xlSheet.Range(xlCellStart, xlCellEnd)


                    Dim count As Integer = DataGridView2.Rows.Count
                    Dim check As Object = xlCells(row, col).value
                    Dim check2 As Object = dc.DataType
                    Dim DtCount As Integer = dt.Rows.Count


                    Dim dtt As DataTable = dt
                    Select Case DataGridView2.Rows(col - 1).Cells(1).Value
                        Case "文字列"
                            xlRange.NumberFormatLocal = "@"
                        Case "数値"
                            For b = 1 To DtCount - 1
                                If IsNumeric(dt.Rows(b)(col - 1)) Then
                                    xlRange.NumberFormatLocal = "###,##0"
                                    CheckFlg = True
                                End If
                            Next
                        Case "日付"
                            For b = 1 To DtCount - 1
                                If TypeOf dt.Rows(b)(col - 1) Is Date Then
                                    xlRange.NumberFormatLocal = "yyyy/mm/dd"
                                End If
                            Next
                        'End ig
                        Case "小数"
                            For b = 1 To DtCount - 1
                                If TypeOf dt.Rows(b)(col - 1) Is Decimal Then
                                    xlRange.NumberFormatLocal = "#,##.00"
                                End If
                            Next
                            CheckFlg = True
                    End Select

                    xlRange.value = columnData

                    changeExcelRow = ConvertToLetter(col2)
                    '合計出力
                    If CheckFlg = True Then
                        If DataGridView2.Rows(col - 1).Cells(3).Value = "a" Then
                            sumExcel = "=SUM(" & changeExcelRow & "2:" & changeExcelRow & dt.Rows.Count + 1 & ")"
                            xlRange2 = xlSheet.Cells(dt.Rows.Count + 3, changeExcelRow)
                            xlRange2.value = sumExcel
                            exist += 1
                        End If
                    End If

                    If exist > 0 Then
                        xlRange3 = xlSheet.Cells(dt.Rows.Count + 2, changeExcelRow)
                        xlRange3.value = "合計"
                    End If


                    col2 += 1

                End If

            Next

            Dim col22 As Integer = 1
            Dim dc2 As DataColumn
            xlCells22 = xlSheet2.Cells
            Dim columnData2(Wdt.Rows.Count, 1) As Object
            Dim couneeet As Integer = Wdt.Columns.Count
            For Scol = 1 To Wdt.Columns.Count
                row = 1
                dc2 = Wdt.Columns(Scol - 1)
                'ヘッダー行の出力
                xlCells22(row, col22).value = dc2.ColumnName
                row = row + 1

                ' 列データを配列に格納
                For i As Integer = 0 To Wdt.Rows.Count - 1
                    columnData2(i, 0) = String.Format(Wdt.Rows(i)(Scol - 1))
                Next
                xlCellStart22 = xlCells22(row, col22)
                xlCellEnd22 = xlCells22(row + Wdt.Rows.Count - 1, col22)
                xlRange22 = xlSheet2.Range(xlCellStart22, xlCellEnd22)

                Dim check As Object = xlCells22(row, Scol).value
                Dim check2 As Object = dc2.DataType
                Dim DtCount As Integer = Wdt.Rows.Count

                xlRange22.value = columnData2

                col22 += 1
            Next


            If col2 = 1 Then
                MessageBox.Show("出力対象がありません")
                Exit Sub
            End If

            'xlSheet.Copy(Before:=xlSheet)
            'xlSheet.Previous.Name = "TEST"
            'xlSheet2 = xlSheet.copysheet

            xlCells.EntireColumn.AutoFit()
            xlRange = xlSheet.UsedRange
            xlRange.Borders.LineStyle = 1

            xlCells22.EntireColumn.AutoFit()
            xlRange22 = xlSheet2.UsedRange
            xlRange22.Borders.LineStyle = 1   'xlContinuous

            xlApp.Visible = True

            'xlBook.WorkSheets(3) = xlBook.WorkSheets("kk").copy



        Catch
            xlApp.DisplayAlerts = False
            xlApp.Quit()
            Throw
        Finally
            If xlCellStart IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCellStart)
            If xlCellEnd IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCellEnd)
            If xlRange IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange)
            If xlCells IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCells)
            'If xlSheet IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet)
            If xlBooks IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks)
            If xlBook IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook)
            If xlApp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)

            GC.Collect()

        End Try
    End Sub






End Class