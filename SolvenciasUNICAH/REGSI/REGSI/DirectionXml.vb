Imports System.Xml

Public Class DirectionXml
    Private Sub BtnCrearXml_Click(sender As Object, e As EventArgs) Handles btnCrearXml.Click
        If DireccionSolvencias.Text = "" And DireccionPlantilla.Text = "" Then
            MessageBox.Show("No ha ingresado la dirreccion del archivo y la direccion del pdf")
        Else
            Dim Doc As New XmlDocument, direct, direct1 As XmlNode
            direct = Doc.CreateElement("Direccion")

            Doc.AppendChild(direct)
            direct = Doc.CreateElement("Direccion")
            direct.InnerText = DireccionSolvencias.Text
            Doc.DocumentElement.AppendChild(direct)

            direct1 = Doc.CreateElement("Direccion")
            direct1.InnerText = DireccionPlantilla.Text
            Doc.DocumentElement.AppendChild(direct1)

            Doc.Save(Application.StartupPath & "\Direccion.xml")

            Me.Dispose()
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        Me.Dispose()
    End Sub
    Private Sub DirectionXml_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Doc As New XmlDocument()
        Dim xmlnode As XmlNodeList
        Dim i As Integer = 0
        Doc.Load(Application.StartupPath & "\Direccion.xml")
        xmlnode = Doc.GetElementsByTagName("Direccion")
        DireccionPlantilla.Text = xmlnode(i).ChildNodes.Item(0).InnerText.Trim()
        DireccionSolvencias.Text = xmlnode(i).ChildNodes.Item(1).InnerText.Trim()
    End Sub

    Private Sub BtnOpenFile_Click(sender As Object, e As EventArgs) Handles btnOpenFile.Click
        Dim openFileDialog1 As New OpenFileDialog
        If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            DireccionPlantilla.Text = openFileDialog1.FileName
            Exit Sub
        End If
    End Sub

    Private Sub BtnFolderOpen_Click(sender As Object, e As EventArgs) Handles btnFolderOpen.Click
        Dim folder As New FolderBrowserDialog
        If folder.ShowDialog() = Windows.Forms.DialogResult.OK Then
            DireccionSolvencias.Text = folder.SelectedPath & "\"
            Exit Sub
        End If
    End Sub
End Class