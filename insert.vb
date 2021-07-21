Imports Npgsql
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Xml
Public Class insert
    Private dt As DataTable

    Sub New()
        aa()
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        write()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        GetGrid()

        DataGridView1.DataSource = dt

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


        Label1.Text = "ファイルを読み込みました"
        Label1.ForeColor = Color.Blue
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click



        If dt Is Nothing Then
            MessageBox.Show("ファイルを読み込んでください")
            Exit Sub
        End If

        Label1.Text = "実行中"


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

        dt = DataGridView1.DataSource
        Dim dt2 As DataTable



        Dim i As Integer = 10
        Using db As New dbConnection()

            Try
                db.trnStart()

                Dim sql As String = ""

                Dim sql2 As String = ""
                Dim count As Integer = dt.Rows.Count
                Dim check As Integer
                Dim aa As Integer
                sql = ""
                sql &= " insert into mst_inserts3 "
                sql &= " (a, b, c, d, f,g,h,i,j,k,l) values "
                For Each row As DataRow In dt.Rows


                    sql2 = ""
                    sql2 &= "select a from mst_inserts3 "
                    sql2 &= " where a = '" & row(0) & "' "

                    dt2 = db.getDtSql(sql2)
                    check += 1

                    If dt2.Rows.Count = 0 Then

                        aa += 1
                        If check = count Then
                            sql &= " ('" & row(0) & "','" & row(1) & "', '" & row(2) & "','" & row(3) & "','" & row(4) & "',
                    '" & row(5) & "','" & row(6) & "', '" & row(7) & "','" & row(8) & "','" & row(9) & "','" & row(10) & "')"
                        Else
                            sql &= " ('" & row(0) & "','" & row(1) & "', '" & row(2) & "','" & row(3) & "','" & row(4) & "',
                                '" & row(5) & "','" & row(6) & "', '" & row(7) & "','" & row(8) & "','" & row(9) & "','" & row(10) & "'),"
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
                Label1.Text = "実行完了しています"
            End Try

            MessageBox.Show("完了しました")


        End Using
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Label1.Text = "実行中"

        Dim dt2 As New DataTable

        Using db As New dbConnection()
            Dim sql As String = ""

            sql = ""
            sql &= " select d as 商品名,"
            sql &= " h as 適用開始日"
            sql &= " from mst_inserts3 "
            sql &= " where id <= 10000"
            sql &= " and"
            sql &= " ( h >= '2021/5/01' and h <= '2021/5/31')"

            dt2 = db.getDtSql(sql)

        End Using

        DataGridView1.DataSource = Nothing
        DataGridView1.DataSource = dt2

        If CheckBox1.Checked = False Then
            Csv(dt2)
        End If

        If CheckBox1.Checked = True Then
            Excel(dt2)
        End If

        Label1.Text = "実行完了しています"

    End Sub

    Private Sub insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "ファイルを読み込んでください"
        Label1.ForeColor = Color.Red
    End Sub

    Private Sub Csv(dt2 As DataTable)
        Using sfd As SaveFileDialog = New SaveFileDialog
            'デフォルトのファイル名を指定します
            sfd.FileName = "output.csv"

            If sfd.ShowDialog() = DialogResult.OK Then
                Using writer As New StreamWriter(sfd.FileName, False, Encoding.GetEncoding("shift_jis"))

                    Dim rowCount As Integer = dt2.Rows.Count
                    Dim ColumnCount As Integer = dt2.Columns.Count

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
                            strList.Add(dt2(i)(j))
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
            FileFilter:="Excel File (*.xlsx),*.xlsx")

        '保存先ディレクトリの設定が有効の場合はブックを保存
        If saveFileName <> "False" Then
            objWorkBook.SaveAs(Filename:=saveFileName)
        End If

        'シートの最大表示列項目数
        Dim columnMaxNum As Integer = dt2.Columns.Count - 1
        'シートの最大表示行項目数
        Dim rowMaxNum As Integer = dt2.Rows.Count - 1

        Dim Last As String = dt2.Rows(0)(0)


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

    Private Sub write()

        Dim xmlDoc As New XmlDocument()

        Const fileName As String = "XMLFile1.xml"

        xmlDoc.Load(fileName)

        Dim node As XmlNode = xmlDoc.DocumentElement.SelectSingleNode(Me.Name)

        '更新かける
        If Not node Is Nothing Then
            node.Item("Width").InnerText = Me.ClientSize.Width.ToString
            node.Item("Height").InnerText = Me.ClientSize.Height.ToString
            node.Item("X").InnerText = Me.Location.X.ToString
            node.Item("Y").InnerText = Me.Location.Y.ToString

            xmlDoc.Save(fileName)

            Return
        End If

        'インサートする
        Dim person = New XElement(New XElement(Me.Name, New XElement("Width", Me.ClientSize.Width),
                                           New XElement("Height", Me.ClientSize.Height), New XElement("X", Me.Location.X),
                                             New XElement("Y", Me.Location.Y)))

        Dim xmlFile = XElement.Load(fileName)

        xmlFile.Add(person)

        xmlFile.Save(fileName)

    End Sub

    Private Sub Read()


        Const fileName As String = "XMLFile1.xml"

        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(fileName)

        Dim node As XmlNode = xmlDoc.DocumentElement.SelectSingleNode(Me.Name)

        If node Is Nothing Then
            Return
        End If

        Dim size As Size = New Size(CInt(node.Item("Width").InnerText), CInt(node.Item("Height").InnerText))

        Dim point As Point = New Point(CInt(node.Item("X").InnerText), CInt(node.Item("Y").InnerText))

        Me.ClientSize = size

        Me.Location = point

    End Sub

    Private Sub aa()
        Dim check As Integer() = {1, 2, 3, 4, 5}
        Dim aa As Integer = Array.IndexOf(check, 4)
        Dim i As Integer = 0
        For Each a In check
            i += 1
            If a = 4 Then
                MessageBox.Show(i)
            End If
        Next
    End Sub
End Class