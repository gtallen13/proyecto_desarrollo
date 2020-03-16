Imports System.Xml
Public Class XmlVersiculo


    ' Private Sub BtnCargarXml_Click(sender As Object, e As EventArgs) Handles btnCargarXml.Click
    ' Dim Doc As New XmlDocument()
    'Dim xmlnode As XmlNodeList
    'Dim i As Integer = 0
    '   Doc.Load(Application.StartupPath & "\Versiculo.xml")
    '   xmlnode = Doc.GetElementsByTagName("Versiculo")
    '  txtBiblia.Text = xmlnode(i).ChildNodes.Item(0).InnerText.Trim()
    '  btnCrearXml.Text = "GuardarXML"
    'End Sub



    '  Private Sub BtnAceptar_Click_1(sender As Object, e As EventArgs) Handles btnAceptar.Click
    '      txtBiblia.Text = ""
    '  Me.Close()
    '    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Me.Dispose()
    End Sub



    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Me.Dispose()
    End Sub

    Private Sub TxtBiblia_TextChanged(sender As Object, e As EventArgs) Handles txtBiblia.TextChanged

    End Sub

    Private Sub XmlVersiculo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Doc As New XmlDocument()
        Dim xmlnode As XmlNodeList
        Dim i As Integer = 0
        Doc.Load(Application.StartupPath & "\Versiculo.xml")
        xmlnode = Doc.GetElementsByTagName("Versiculo")
        txtBiblia.Text = xmlnode(i).ChildNodes.Item(0).InnerText.Trim()

    End Sub

    Private Sub BtnCrearXml_Click_1(sender As Object, e As EventArgs) Handles btnCrearXml.Click
        If txtBiblia.Text = "" Then
            MessageBox.Show("No ha ingresado el versiculo")
        Else

            Dim Doc As New XmlDocument, biblia As XmlNode
            biblia = Doc.CreateElement("Versiculo")

            Doc.AppendChild(biblia)
            biblia = Doc.CreateElement("Biblia")
            biblia.InnerText = txtBiblia.Text
            Doc.DocumentElement.AppendChild(biblia)
            Doc.Save(Application.StartupPath & "\Versiculo.xml")
            btnCrearXml.Text = "GuardarXML"
            MessageBox.Show("Se ha guardado exitosamente")

        End If
        Me.Dispose()
    End Sub
End Class