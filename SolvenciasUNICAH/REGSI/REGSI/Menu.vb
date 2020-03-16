Imports WindowsApplication1
Public Class Menu
    Dim Solvencia As New WindowsApplication1.Form1
    Private Sub Menu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
     
    End Sub

    Private Sub Cerrar_Click(sender As Object, e As EventArgs) Handles Cerrar.Click
        Me.Dispose()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles BtnAbrirSolvencia.Click
        Solvencia.ShowDialog()
    End Sub

    Private Sub Minimizar_Click(sender As Object, e As EventArgs) Handles Minimizar.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles Xml.Click
        'mostrara el login 
        'si ingresa la contraseña correcta le mostrara el formulario de DirectionXml
        Decision.valor = 1
        Login.ShowDialog()
    End Sub

    Private Sub AboutBox_Click(sender As Object, e As EventArgs) Handles AboutBox.Click
        AcercaDe.ShowDialog()
    End Sub

    Private Sub PictureBox3_Click_1(sender As Object, e As EventArgs) Handles PictureBox3.Click
        'mostrara el login 
        'si ingresa la contraseña correcta le mostrara el formulario de XmlVersiculo
        Decision.valor = 0
        Login.ShowDialog()
    End Sub
End Class