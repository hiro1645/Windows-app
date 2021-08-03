Imports Npgsql
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Xml
Imports System.Windows.Input
Imports Newtonsoft.Json

Public Class insert
    Private Property _dt As DataTable
    Private dt As DataTable
    Private Fdt As DataTable
    Private dic As New Dictionary(Of String, Boolean)
    Private flug As Boolean = False
    Private flug2 As Boolean = False
    Private flug3 As Boolean = False
    Private dealerflug As Boolean = False

    Sub New()


        'Dim productInfo As New Dictionary(Of String, Object)
        'Dim image As New Dictionary(Of String, Object)
        'Dim thumbnail As New Dictionary(Of String, Object)

        'productInfo.Add("Image", image)
        'image.Add("Width", 800)
        'image.Add("Height", 600)
        'image.Add("Title", "15階からの眺望")
        'image.Add("Thumbnail", thumbnail)
        'thumbnail.Add("Url", "http://www.example.com/image/481989943")
        'thumbnail.Add("Height", 125)
        'thumbnail.Add("Width", 100)
        'image.Add("Animated", False)
        'image.Add("IDs", {116, 943, 234, 38793})

        'Dim jsonText As String = JsonConvert.SerializeObject(productInfo, Xml.Formatting.Indented)

        'Debug.WriteLine(jsonText)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        'write()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Columns()

        DataGridView2.AllowUserToAddRows = False

        _dt = CType(dt.Copy, DataTable)

        Dim person1 = New With {.Name = "徳川家康", .Age = 20}
        Dim person2 = New With {.Name = "豊臣秀吉", .Age = 25}
        Dim person3 = New With {.Name = "織田信長", .Age = 30}

        Dim person = {person1, person2, person3}

        combo()

        TextBox1.Text = "KIRIN モンスター"


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

        If ofd.FileName = Nothing Then
            Exit Function
        End If


        Dim SR As New StreamReader(ofd.FileName, System.Text.Encoding.GetEncoding("shift_jis")) 'StreamReader文字化け防止
        If ofd.FileName.EndsWith(".csv") Or ofd.FileName.EndsWith(".tsv") Or ofd.FileName.EndsWith(".txt") Then
        Else
            'DataGridView1.DataSource = Nothing
            MessageBox.Show("エラー")
            Return Nothing
        End If
        Return SR
#Disable Warning BC42105 ' 関数が、すべてのコード パスで値を返しません
    End Function
#Enable Warning BC42105 ' 関数が、すべてのコード パスで値を返しません

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
            line = SR.ReadLine.ToString

            Dim Item() As String = line.Split(vbTab.ToCharArray)

            Dim s As Integer = 0

            'For Each v In Item
            '    dt.Columns.Add(Item(s))
            '    s += 1
            'Next
            For Each v In Item
                dt.Columns.Add(v)
            Next
            i += 1
        End If

        Do
            If i > 0 Then
                line = SR.ReadLine
                If line = Nothing Then
                    Exit Do
                End If
                Dim Item() As String = line.Split(vbTab.ToCharArray)

                dr = CType(dt.NewRow, DataRow)
                dr.ItemArray = Item
                dt.Rows.Add(CType(dr, DataRow))
            End If
        Loop


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        If dt Is Nothing Then
            MessageBox.Show("ファイルを読み込んでください")
            Exit Sub
        End If




        Dim a As String = "登録します。よろしいですか？"
        Dim result As DialogResult = MessageBox.Show(a,
                                             "質問",
                                             MessageBoxButtons.YesNo,
                                             MessageBoxIcon.Exclamation,
                                             MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If result = DialogResult.No Then
            Exit Sub
        End If

        dt = CType(DataGridView1.DataSource, DataTable)

        Dim dt2 As DataTable

        Dim i As Integer = 10
        Using db As New dbConnection()

            Try
                db.trnStart()

                Dim sql As String = ""

                Dim sql2 As String = ""
                Dim count As Integer = CInt(dt.Rows.Count)
                Dim check As Integer
                Dim aa As Integer

                sql = ""
                sql &= " insert into mst_inserts3 "
                sql &= " (a, b, c, d, f,g,h,i,j,k,l) values "
                For Each row As DataRow In CType(dt.Rows, DataRowCollection)
                    sql2 = ""
                    sql2 &= "select a from mst_inserts3 "
                    sql2 &= " where a = '" & row(0).ToString & "' "

                    dt2 = db.getDtSql(sql2)
                    check += 1

                    If dt2.Rows.Count = 0 Then

                        aa += 1
                        If check = count Then
                            sql &= " ('" & row(0).ToString & "','" & row(1).ToString & "', '" & row(2).ToString & "','" & row(3).ToString & "','" & row(4).ToString & "',
                    '" & row(5).ToString & "','" & row(6).ToString & "', '" & row(7).ToString & "','" & row(8).ToString & "','" & row(9).ToString & "','" & row(10).ToString & "')"
                        Else
                            sql &= " ('" & row(0).ToString & "','" & row(1).ToString & "', '" & row(2).ToString & "','" & row(3).ToString & "','" & row(4).ToString & "',
                                '" & row(5).ToString & "','" & row(6).ToString & "', '" & row(7).ToString & "','" & row(8).ToString & "','" & row(9).ToString & "','" & row(10).ToString & "'),"
                        End If
                    End If

                Next
                If aa = 0 Then
                    MessageBox.Show("すべて登録済みです")
                    Exit Sub
                End If


                db.executeSql(sql)
                db.commit()  'コミットする


            Catch ex As Exception
                db.rollback()
                MessageBox.Show("例外的なエラーです")
            Finally

            End Try

            MessageBox.Show("完了しました")


        End Using
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim Lst As New List(Of String)

        For s = 0 To DataGridView2.Rows.Count - 1
            'If DataGridView2.Rows(s).Visible = True Then
            Lst.Add(DataGridView2.Rows(s).Cells(0).Value.ToString)
            'End If
        Next



        Dim dt2 As New DataTable
        Dim v As Integer = 0

        Using db As New dbConnection()
            Dim sql As String = ""

            sql = ""
            sql &= " select "
            For Each s In Lst
                v += 1
                If v = Lst.Count Then
                    sql &= s
                    Exit For
                End If
                sql &= s & ","
            Next
            'sql &= " as 商品名,"
            'sql &= " h as 適用開始日"
            sql &= " from mst_inserts3 "
            'sql &= " where id <= 10000"
            'sql &= " And"
            'sql &= " ( h >= '2021/5/01' and h <= '2021/5/31')"

            dt2 = db.getDtSql(sql)

        End Using

        DataGridView1.DataSource = Nothing
        DataGridView1.DataSource = CType(dt2, DataTable)

        If CheckBox1.Checked = False Then
            Csv(CType(dt2, DataTable))
        End If

        If CheckBox1.Checked = True Then
            Excel(CType(dt2, DataTable))
        End If

    End Sub

    Private Sub insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Private Sub Columns()

        dt = Nothing

        GetGrid()

        DataGridView1.DataSource = CType(dt, DataTable)


        DataGridView2.DataSource = Nothing
        DataGridView2.Columns.Clear()

        Dim DtColumns As Integer = CInt(dt.Columns.Count)
        Dim dt3 As New DataTable

        dt3.Columns.Add("タイトル")

        Dim dr As DataRow

        For a = 0 To DtColumns - 1
            dr = CType(dt3.NewRow, DataRow)
            dr("タイトル") = dt.Columns(a)
            dt3.Rows.Add(dr)
        Next
        DataGridView2.DataSource = CType(dt3, DataTable)
        Dim column As New DataGridViewCheckBoxColumn
        DataGridView2.Columns.Add(column)

    End Sub

    Private Sub Csv(dt2 As DataTable)
        Using sfd As SaveFileDialog = New SaveFileDialog
            'デフォルトのファイル名を指定します
            sfd.FileName = "output.csv"

            If sfd.ShowDialog() = DialogResult.OK Then
                Using writer As New StreamWriter(sfd.FileName, False, Encoding.GetEncoding("shift_jis"))

                    Dim rowCount As Integer = CInt(dt2.Rows.Count)
                    Dim ColumnCount As Integer = CInt(dt2.Columns.Count)

                    Dim strList1 As New List(Of String)
                    For i = 0 To ColumnCount - 1
                        strList1.Add(dt2.Columns(i).Caption)
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
                            strList.Add(dt2(i)(j).ToString)
                        Next
                        Dim strArray As String() = strList.ToArray() ' 配列へ変換


                        ' CSV 形式に変換
                        Dim strCsvData As String = String.Join(",", strList)

                        writer.WriteLine(strCsvData.ToString)
                    Next
                    MessageBox.Show("CSV ファイルを出力しました")
                End Using
            End If

        End Using
    End Sub

    Private Sub Excel(dt2 As DataTable)
        Dim objExcel As Excel.Application = New Excel.Application
        Dim objWorkBook As Excel.Workbook = objExcel.Workbooks.Add
        Dim objSheet As Excel.Worksheet = Nothing

        '現在日時を取得
        Dim timestanpText As String = Format(Now, "yyyyMMdd")

        Dim aa As Integer = dt2.Rows.Count


        '保存ディレクトリとファイル名を設定
        Dim saveFileName As String
        saveFileName = objExcel.GetSaveAsFilename(
            InitialFilename:="ファイル名_" & timestanpText,
            FileFilter:="Excel File (*.xlsx),*.xlsx").ToString

        '保存先ディレクトリの設定が有効の場合はブックを保存
        If saveFileName <> "False" Then
            objWorkBook.SaveAs(Filename:=saveFileName)
        End If

        'シートの最大表示列項目数
        Dim columnMaxNum As Integer = dt2.Columns.Count - 1
        'シートの最大表示行項目数
        Dim rowMaxNum As Integer = dt2.Rows.Count - 1

        Dim Last As String = dt2.Rows(0)(0).ToString


        '項目名格納用リストを宣言
        Dim columnList As New List(Of String)
        '項目名を取得
        For i As Integer = 0 To (columnMaxNum)
            columnList.Add(dt2.Columns(i).Caption)
        Next

        ''セルのデータ取得用二次元配列を宣言
        Dim v As Object(,) = New Object(rowMaxNum, columnMaxNum) {}

        For row As Integer = 0 To rowMaxNum
            For col As Integer = 0 To columnMaxNum
                'If dt.Rows(row)(col).Value Is Nothing = False Then
                ' セルに値が入っている場合、二次元配列に格納
                v(row, col) = dt2.Rows(row)(col)
                'End If
            Next
        Next


        ' EXCELに項目名を転送
        For i As Integer = 1 To dt2.Columns.Count
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

        Dim data As String = "A2:" & Chr(Asc("A") + columnMaxNum) & dt2.Rows.Count + 1
        objWorkBook.Sheets(1).Range(data) = v

        ' エクセル表示
        objExcel.Visible = True

        ' EXCEL解放
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub

    'Private Sub write()

    '    Dim xmlDoc As New XmlDocument()

    '    Const fileName As String = "XMLFile1.xml"

    '    xmlDoc.Load(fileName)

    '    Dim node As XmlNode = xmlDoc.DocumentElement.SelectSingleNode(Me.Name)

    '    '更新かける
    '    If Not node Is Nothing Then
    '        node.Item("Width").InnerText = Me.ClientSize.Width.ToString
    '        node.Item("Height").InnerText = Me.ClientSize.Height.ToString
    '        node.Item("X").InnerText = Me.Location.X.ToString
    '        node.Item("Y").InnerText = Me.Location.Y.ToString

    '        xmlDoc.Save(fileName)

    '        Return
    '    End If

    '    'インサートする
    '    Dim person = New XElement(New XElement(Me.Name, New XElement("Width", Me.ClientSize.Width),
    '                                       New XElement("Height", Me.ClientSize.Height), New XElement("X", Me.Location.X),
    '                                         New XElement("Y", Me.Location.Y)))


    '    Dim xmlFile = XElement.Load(fileName)

    '    xmlFile.Add(person)

    '    xmlFile.Save(fileName)

    'End Sub

    'Private Sub Read()


    '    Const fileName As String = "XMLFile1.xml"

    '    Dim xmlDoc As New XmlDocument()
    '    xmlDoc.Load(fileName)

    '    Dim node As XmlNode = xmlDoc.DocumentElement.SelectSingleNode(Me.Name)

    '    If node Is Nothing Then
    '        Return
    '    End If

    '    Dim size As Size = New Size(CInt(node.Item("Width").InnerText), CInt(node.Item("Height").InnerText))

    '    Dim point As Point = New Point(CInt(node.Item("X").InnerText), CInt(node.Item("Y").InnerText))

    '    Me.ClientSize = size

    '    Me.Location = point

    'End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click


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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click


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


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        If dealerflug = True Then
            dt = Fdt
        End If

        Dim datacolumns As DataColumnCollection = CType(dt.Columns, DataColumnCollection)

        Dim datarows As New List(Of Object)

        For f = 0 To dt.Rows.Count - 1
            datarows.Add(dt.Rows(f)(1))
        Next

        Dim kskk3 = datarows.ToArray

        Dim k As DataColumn = datacolumns(0)

        Dim check = From key In dt.Columns
                    Order By key.ToString.Length

        Dim Scheck = From key In dt.Columns
                     Order By key.ToString.Length Descending

#Disable Warning BC42017
        Dim check2 = From key In dt.Rows Order By key.item(1).ToString.Length Select key.item(1)

        Dim kskk() As Object

        If flug = False Then
            kskk = check.ToArray
            flug = True
        Else
            kskk = Scheck.ToArray
            flug = False
        End If

        Dim kskk2 = check2.ToArray

        Dim aa = kskk2.ToString

        'For i = 0 To kskk2.Length - 1
        '    If kskk2(i) Is kskk3(i) Then
        '        Continue For
        '    End If
        '    Dim s = kskk2(i)
        '    Dim count As Integer = 0
        '    For v = 0 To kskk3.Length - 1
        '        If kskk3(v) Is s Then
        '            count = v
        '            Exit For
        '        End If
        '    Next
        'Next

        For l = 0 To kskk.Length - 1
            If datacolumns(l) Is kskk(l) Then
                Continue For
            End If
            Dim s = kskk(l)
            Dim count As Integer = 0
            For v = 0 To datacolumns.Count - 1
                If datacolumns(v) Is s Then
                    count = v
                    Exit For
                End If
            Next
            dt.Columns(count).SetOrdinal(l)
        Next

        DataGridView1.DataSource = Nothing

        DataGridView1.Columns.Clear()

        DataGridView1.DataSource = CType(dt, DataTable)

        'Dim min As String = a.Min
        Debug.Print("a")
        'Dim aa As New List(Of String)
        'Dim count As Integer = 0

        'a.Sort()

        'Dim ss As New List(Of Integer)

        'For p = 0 To a.Count - 1
        '    ss.Add(a(p).Length)
        'Next

        'ss.Sort()

        'Dim l As New List(Of String)

        'For k = 0 To a.Count - 1
        '    For Each aaaa In a
        '        If ss(k) = aaaa.Length Then
        '            l.Add(aaaa)
        '            Exit For
        '        End If
        '    Next
        'Next

    End Sub



    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Dim datarows As New List(Of Object)

        Dim 商品名 As Integer = dt.Columns("商品名").Ordinal

        For f = 0 To dt.Rows.Count - 1
            datarows.Add(dt.Rows(f)(商品名))
        Next

        Dim sort = From key In datarows Order By key.ToString.Length

        Dim k() As Object
        If flug3 = False Then
            k = sort.ToArray
            flug3 = True
            Button9.Text = "JANCODE並び替え"
        Else
            k = datarows.ToArray
            flug3 = False
            Button9.Text = "商品名並び替え"
        End If


        Dim rows As DataRow()

        Dim dtt As DataTable = dt.Copy

        Dim name As String = ""
        For row = 0 To k.Length - 1

            If name = k(row).ToString Then
                Continue For
            End If

            rows = dtt.Select("商品名 =  '" & k(row).ToString & "' ")

            If rows.Length > 1 Then
                name = k(row).ToString
            End If

            Dim select2 As DataRow = Nothing

            For g = 0 To rows.Length - 1
                select2 = rows(g)
                dtt.Rows.Add(select2.ItemArray)
            Next

            For Each Row2 As DataRow In rows
                dtt.Rows.Remove(Row2)
            Next

        Next

        Debug.Print("s")
        DataGridView1.DataSource = Nothing
        DataGridView1.Columns.Clear()

        DataGridView1.DataSource = dtt


        'DataGridView3.DataSource = dtt
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim datarows As New List(Of Object)

        Dim 日付 As Integer = dt.Columns("日付").Ordinal

        For f = 0 To dt.Rows.Count - 1
            datarows.Add(dt.Rows(f)(日付))
        Next

        Dim sort = From key In datarows Order By key

        Dim time = sort.ToArray

        Dim rows As DataRow()

        Dim dtt As DataTable = dt.Copy

        Dim name As String = ""
        For row = 0 To time.Length - 1
            '行を取得
            rows = dtt.Select("日付 =  '" & time(row).ToString & "' ")

            '同じ日付あった場合
            If name = time(row).ToString Then
                Continue For
            End If
            If rows.Length > 1 Then
                name = time(row).ToString
            End If

            Dim select2 As DataRow = Nothing

            For g = 0 To rows.Length - 1
                select2 = rows(g)
                dtt.Rows.Add(select2.ItemArray)
            Next

            For Each Row2 As DataRow In rows
                dtt.Rows.Remove(Row2)
            Next

        Next

        DataGridView1.DataSource = dtt

        Debug.Print("a")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Dim datarows As New List(Of Object)

        Dim 販売元名 As Integer = dt.Columns("販売元名").Ordinal

        For f = 0 To dt.Rows.Count - 1
            datarows.Add(dt.Rows(f)(販売元名))
        Next

        Dim sort = From key In datarows Order By key

        Dim dealer = sort.ToArray

        Dim rows As DataRow()

        Dim dtt As DataTable = dt.Copy

        Dim name As String = ""
        For row = 0 To dealer.Length - 1
            '行を取得
            rows = dtt.Select("販売元名 =  '" & dealer(row).ToString & "' ")

            '同じ日付あった場合
            If name = dealer(row).ToString Then
                Continue For
            End If
            If rows.Length > 1 Then
                name = dealer(row).ToString
            End If

            Dim select2 As DataRow = Nothing

            For g = 0 To rows.Length - 1
                select2 = rows(g)
                dtt.Rows.Add(select2.ItemArray)
            Next

            For Each Row2 As DataRow In rows
                dtt.Rows.Remove(Row2)
            Next

        Next

        DataGridView1.DataSource = dtt

        Debug.Print("a")
    End Sub



    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        If ComboBox1.SelectedIndex < 0 Then
            MessageBox.Show("選択してください")
            Return
        End If

        If dealerflug = False Then
            Fdt = dt.Copy
        End If

        Dim datarows As New List(Of Object)

        Dim 販売元名 As Integer = dt.Columns("販売元名").Ordinal

        For f = 0 To dt.Rows.Count - 1
            datarows.Add(dt.Rows(f)(販売元名))
        Next

        Dim sort = From key In datarows Order By key Where key.ToString = ComboBox1.SelectedItem.ToString

        Dim dealer = sort.ToArray

        Dim newdt As DataTable = dt.Copy

        newdt.Rows.Clear()

        For ff = 0 To datarows.Count - 1
            If dt(ff)(販売元名).ToString = dealer(0).ToString Then
                Dim item() = dt(ff).ItemArray
                newdt.Rows.Add(item)
            End If
        Next

        DataGridView1.DataSource = newdt

        Debug.Print("a")

        dt = newdt

        dealerflug = True

    End Sub

    Private Sub combo()

        ComboBox1.Items.Clear()

        Dim datarows As New List(Of String)

        Dim 販売元名 As Integer = dt.Columns("販売元名").Ordinal

        For f = 0 To dt.Rows.Count - 1
            datarows.Add(dt.Rows(f)(販売元名).ToString)
        Next

        Dim check = From key In datarows Select key Distinct

        Dim result As Integer = Aggregate number In datarows Into Count(CType(number.Length, Boolean))

        Dim sort = check.ToArray

        For a = 0 To sort.Length - 1
            ComboBox1.Items.Add(sort(a))
        Next

        'sum()
    End Sub

    Private Sub sum()

        Dim datarows As New List(Of String)

        Dim 販売元名 As Integer = dt.Columns("商品名").Ordinal

        For f = 0 To dt.Rows.Count - 1
            datarows.Add(dt.Rows(f)(販売元名).ToString)
        Next

        Dim check As Integer = Aggregate number In datarows Into Sum(CInt(number))

        Dim check2 = Aggregate number In datarows Into Average(CInt(number))

        MessageBox.Show(check.ToString)

        MessageBox.Show(check2.ToString)
    End Sub

    Private Sub search()

        'Try
        If TextBox1.Text = String.Empty Then
            DataGridView1.DataSource = Nothing
            DataGridView1.DataSource = dt
            Exit Sub
        End If
        Dim datarow As New List(Of String)

        Dim datacolumn As New List(Of String)

        Dim 販売元名 As Integer = dt.Columns("販売元名").Ordinal

        Dim 商品名 As Integer = dt.Columns("商品名").Ordinal

        For f = 0 To dt.Rows.Count - 1
            datarow.Add(dt.Rows(f)(販売元名) & "," & dt.Rows(f)(商品名))
        Next

        'For f = 0 To dt.Columns.Count - 1
        '    datacolumn.Add(dt.Columns(f).ColumnName)
        'Next

        Dim newdt As DataTable = dt.Copy

        newdt.Rows.Clear()



        If TextBox1.Text = String.Empty Then
            DataGridView1.DataSource = newdt
            Return
        End If


        Dim item = TextBox1.Text.Split(" ")

        For v = 0 To item.Length - 1
            If item(v) = "" Then
                Array.Clear(item, v, 1)

            End If
        Next

        Dim folder As New List(Of Object)

        Dim e As Integer = 0

        For e = 0 To item.Length - 1

            If item(e) = Nothing Then
                Exit For
            End If
            Dim check = From key In datarow Where key.Split(",").First.Contains(item(e)) Select key.Split(",").First

            folder.Add(check.ToArray)
            Debug.Print("a")
        Next

        For e = 0 To item.Length - 1
            If item(e) = Nothing Then
                Exit For
            End If

            Dim check = From key In datarow Where key.Split(",").Last.Contains(item(e)) Select key.Split(",").Last

            folder.Add(check.ToArray)

        Next

        For o = 0 To folder.Count - 1
            Dim g = folder(o)
        Next


        Try
            Dim dealer As Integer = 100


            Dim gift As Integer = 100

            For f = 0 To folder.Count - 1
                If dealer = 100 Then
                    Try
                        For u = 0 To dt.Rows.Count - 1
                            If dt(u)(販売元名).ToString = folder(f)(0).ToString Then
                                dealer = f
                                Exit For
                            End If
                        Next
                    Catch ex As Exception
                        Exit Try
                    End Try
                End If

                If gift = 100 Then
                    For u = 0 To dt.Rows.Count - 1
                        Try
                            If dt(u)(商品名).ToString = folder(f)(0).ToString Then
                                gift = f
                                Exit For
                            End If
                        Catch ex As Exception
                            Exit For
                        End Try
                    Next
                End If

            Next


            If Not dealer = 100 Then
                For ff = 0 To datarow.Count - 1
                    If dt(ff)(販売元名).ToString = folder(dealer)(0).ToString Then
                        Dim item2() = dt(ff).ItemArray
                        newdt.Rows.Add(item2)
                    End If
                Next
            End If

            If gift = 100 Then
                DataGridView1.DataSource = newdt

                Exit Sub
            End If


            Dim newdt2 As New DataTable


            If Not gift = 100 Then

                If dealer = 100 Then
                    newdt2 = dt.Copy
                Else
                    newdt2 = newdt.Copy
                End If

                newdt2.Rows.Clear()

                For ff = 0 To datarow.Count - 1
                    If dt(ff)(商品名).ToString = folder(gift)(0).ToString Then
                        Dim item2() = dt(ff).ItemArray
                        newdt2.Rows.Add(item2)
                    End If
                Next
            End If

            DataGridView1.DataSource = newdt2

        Catch ex As Exception

        Finally
            If DataGridView1.Rows.Count <= 1 Then
                MessageBox.Show("検索結果がありません")
            End If
        End Try

    End Sub


    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        search()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        DataGridView1.DataSource = dt
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim yu As New DataTable
        yu = DataGridView1.DataSource
        Csv(yu)
    End Sub

End Class