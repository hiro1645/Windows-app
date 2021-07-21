Imports System.Xml
Public Class BaseForm
    Private fileName As String = IO.Path.Combine(Application.StartupPath, "XMLFile1.xml")

    ''' <summary>
    ''' ボタン連打対応用
    ''' </summary>
    Public buttonProcessing As Boolean

    ''' <summary>
    ''' 説明設定用プロパティ
    ''' </summary>
    Public WriteOnly Property InfoLabel As String
        Set(value As String)
            Controls("InfoLabel").Text = value
        End Set
    End Property

    Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

        'ボタン連打対応用
        AddHandler System.Windows.Forms.Application.Idle, AddressOf Application_Idle

        'Form設定
        Font = New Font("Yu Gothic UI", 12)
        Width = 1600
        Height = 800
        BackColor = Color.AliceBlue

    End Sub

    Private Sub Application_Idle(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.buttonProcessing = False
    End Sub

    Private Sub BaseForm_Closed(sender As Object, e As EventArgs) Handles Me.Closed

    End Sub


    Private Sub BaseForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Write()
    End Sub

    Private Sub BaseForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If DesignMode = False Then
            Read()
        End If
    End Sub

    Private Sub Read()
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

    Private Sub Write()
        Dim xmlDoc As New XmlDocument()

        xmlDoc.Load(fileName)

        Dim node As XmlNode = xmlDoc.DocumentElement.SelectSingleNode(Me.Name)

        '既に存在していたら更新かける
        If Not node Is Nothing Then
            node.Item("Width").InnerText = Me.ClientSize.Width.ToString
            node.Item("Height").InnerText = Me.ClientSize.Height.ToString
            node.Item("X").InnerText = Me.Location.X.ToString
            node.Item("Y").InnerText = Me.Location.Y.ToString

            xmlDoc.Save(fileName)

            Return
        End If

        '存在していなかったら登録する
        Dim person = New XElement(New XElement(Me.Name, New XElement("Width", Me.ClientSize.Width),
                                           New XElement("Height", Me.ClientSize.Height), New XElement("X", Me.Location.X),
                                             New XElement("Y", Me.Location.Y)))

        Dim xmlFile = XElement.Load(fileName)

        xmlFile.Add(person)

        xmlFile.Save(fileName)
    End Sub
End Class