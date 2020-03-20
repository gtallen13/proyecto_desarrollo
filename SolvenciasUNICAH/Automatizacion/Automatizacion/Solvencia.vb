Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports iTextSharp
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.xml
Imports System.IO
Imports sigplusnet_vbnet_lcd15_demo
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Text.RegularExpressions




Public Class Form1
    Dim firmar As New sigplusnet_vbnet_lcd15_demo.Firma
    Dim Correlativo As String = "PRES"
    Dim blnGuardado As Boolean = False
    Dim intCorrelativo As Integer = 0

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        txtSolicitante.Text = ""
        txtLugar.Text = ""
        txtFechaReservacion.Text = ""
        txtHoraInicial.Text = ""
        txtHoraFinal.Text = ""
        rtxtOtros.Text = ""
        rtxtEvento.Text = ""
        chkMicrofono.Checked = False
        chkSonido.Checked = False
        chkDataShow.Checked = False
        chkParlante.Checked = False
        chkOtros.Checked = False
        chkControlAC.Checked = False
        chkControlPantalla.Checked = False
        chkControlDataShow.Checked = False
        chkGrabadora.Checked = False
        chkDocente.Checked = False
        chkPersonalAdmin.Checked = False
        chkEstudiante.Checked = False



        TabControl1.SelectedTab = TabPage1


        Dim fecha As Date
        fecha = Now
        lblFecha.Text = Format(fecha, "dd/MM/yyyy")
        Label20.Text = "Favor Firmar el Documento"
        Label20.ForeColor = Color.Orange
        btnGuardar.Enabled = True
        rtxtOtros.Enabled = False

        Dim Doc As New XmlDocument()
        Dim xmlnode As XmlNodeList
        Dim i As Integer = 0
        Doc.Load(Application.StartupPath & "\Versiculo.xml")
        xmlnode = Doc.GetElementsByTagName("Versiculo")
        lbBiblia.Text = xmlnode(0).ChildNodes.Item(0).InnerText.Trim()

    End Sub


    Private Sub Regresar_Click(sender As Object, e As EventArgs) Handles Regresar.Click
        TabControl1.SelectedTab = TabPage1
    End Sub

    Private Sub BotonFirmar_Click(sender As Object, e As EventArgs) Handles BotonFirmar.Click
        Dim firmar As New Firma
        Try
            firmar.ShowDialog()
            If firmar.firmado = True Then

                Label20.Text = "El documento ha sido Firmado"
                Label20.ForeColor = Color.LimeGreen
                btnGuardar.Enabled = True
                blnGuardado = True

            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        If blnGuardado = False Then
            MessageBox.Show("Por favor firme el documento", "Informacion Incompleta", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim direccionPlantilla As String
        Dim direccionCarpeta As String
        Dim direccionFirmas As String = Application.StartupPath()

        Dim doc As New XmlDocument
        Dim xmlnode As XmlNodeList
        Dim i As Integer = 0
        doc.Load(Application.StartupPath & "\Direccion.xml")
        xmlnode = doc.GetElementsByTagName("Direccion")
        direccionCarpeta = xmlnode(0).ChildNodes.Item(0).InnerText.Trim()
        direccionPlantilla = xmlnode(0).ChildNodes.Item(1).InnerText.Trim()


        Dim pdfTemplate As String = direccionPlantilla
        Dim newFile As String


        'en caso que el mismo usuario elija varios documentos del mismo tipo
        If intCorrelativo > 1 Then
            newFile = direccionCarpeta & txtSolicitante.Text & "SOL" & intCorrelativo & ".pdf"
            'el siguiente codigo iterara sobre los archivo creados para asi asignar el numero seguido del Correlativo
            Dim di As New DirectoryInfo(direccionCarpeta)
            Dim firArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo
            For Each fri In firArr
                If fri.FullName = newFile Then
                    intCorrelativo += 1
                    newFile = direccionCarpeta & txtSolicitante.Text & "SOL" & intCorrelativo & ".pdf"
                End If
            Next
        End If




        Dim pdfReader As New PdfReader(pdfTemplate)
        Dim pdfStamper As New PdfStamper(pdfReader, New FileStream(newFile, FileMode.Create))
        Dim pdfFormFields As AcroFields = pdfStamper.AcroFields
        Dim pcbContent As PdfContentByte = Nothing
        Dim img As System.Drawing.Image = System.Drawing.Image.FromFile(direccionFirmas & "\firma.bmp") 'aqui ira la direccion de las firmas
        Dim sap As PdfSignatureAppearance = pdfStamper.SignatureAppearance
        Dim rect As iTextSharp.text.Rectangle = Nothing
        Dim imagen As iTextSharp.text.Image
        Dim loc As String


        loc = direccionFirmas & "\firma.bmp" 'aqui ira la direccion de las firmas

        imagen = iTextSharp.text.Image.GetInstance(loc)
        imagen.SetAbsolutePosition(427, 43)
        imagen.ScaleToFit(130, 130)
        pcbContent = pdfStamper.GetUnderContent(1)
        pcbContent.AddImage(imagen)

        ' set form pdfFormFields
        pdfFormFields.SetField("Solicitante", txtSolicitante.Text)
        pdfFormFields.SetField("Lugar", txtLugar.Text)
        pdfFormFields.SetField("Fecha", txtFechaReservacion.Text)
        pdfFormFields.SetField("HoraInicial", txtHoraInicial.Text)
        pdfFormFields.SetField("HoraFinal", txtHoraFinal.Text)
        pdfFormFields.SetField("Fecha_5", lblFecha.Text)
        pdfFormFields.SetField("Otros", rtxtOtros.Text)
        pdfFormFields.SetField("Evento", rtxtEvento.Text)


        'pdfFormFields.SetField("signature5", TextBox1.Text)

        ' The form's checkboxes
        If chkMicrofono.Checked = True Then
            pdfFormFields.SetField("Microfono", "On")
        End If

        If chkSonido.Checked = True Then
            pdfFormFields.SetField("Sonido", "On")
        End If

        If chkDataShow.Checked = True Then
            pdfFormFields.SetField("DataShow", "On")
        End If

        If chkParlante.Checked = True Then
            pdfFormFields.SetField("Parlante", "On")
        End If

        If chkControlAC.Checked = True Then
            pdfFormFields.SetField("ControlAC", "On")
        End If

        If chkControlPantalla.Checked = True Then
            pdfFormFields.SetField("ControlPantalla", "On")
        End If

        If chkControlDataShow.Checked = True Then
            pdfFormFields.SetField("ControlDataShow", "On")
        End If

        If chkGrabadora.Checked = True Then
            pdfFormFields.SetField("Grabadora", "On")
        End If
        If chkDocente.Checked = True Then
            pdfFormFields.SetField("Docente", "On")
        End If
        If chkPersonalAdmin.Checked = True Then
            pdfFormFields.SetField("PersonalAdmin", "On")
        End If
        If chkEstudiante.Checked = True Then
            pdfFormFields.SetField("Estudiante", "On")
        End If
        If chkOtros.Checked = True Then
            pdfFormFields.SetField("Otros", "On")
        End If






        MessageBox.Show("Datos Guardados Satisfactoriamente")

        ' flatten the form to remove editting options, set it to false
        ' to leave the form open to subsequent manual edits
        pdfStamper.FormFlattening = False

        ' close the pdf
        pdfStamper.Close()


        txtSolicitante.Text = ""
        txtFechaReservacion.Text = ""
        txtLugar.Text = ""
        txtHoraInicial.Text = ""
        txtHoraFinal.Text = ""
        rtxtEvento.Text = ""
        rtxtOtros.Text = ""
        chkDocente.CheckState = 0
        chkPersonalAdmin.CheckState = 0
        chkEstudiante.CheckState = 0
        chkMicrofono.CheckState = 0
        chkSonido.CheckState = 0
        chkDataShow.CheckState = 0
        chkParlante.CheckState = 0
        chkOtros.CheckState = 0
        chkControlAC.CheckState = 0
        chkControlPantalla.CheckState = 0
        chkControlDataShow.CheckState = 0
        chkGrabadora.CheckState = 0

        Me.Close()

    End Sub



    Private Sub CheckOtros_Razones_CheckedChanged(sender As Object, e As EventArgs) Handles chkOtros.CheckedChanged
        If chkOtros.Checked = True Then
            rtxtOtros.Enabled = True
        Else
            rtxtOtros.Enabled = False
            rtxtOtros.Text = ""
        End If
    End Sub



    ''' <summary>
    ''' Revisa que una dirreccion de correo sea valida
    ''' </summary>
    ''' <param name="correo1"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function IsValidEmailFormat(ByVal correo1 As String) As Boolean
        'Return Regex.IsMatch(correo1, "^([0-9a-zA-Z]([-\.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$")
        Return Regex.IsMatch(correo1, "\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9]*[a-z0-9])?)\Z", RegexOptions.IgnoreCase)
    End Function

    ''' <summary>
    ''' Segunda funcion que revisa que una dirreccion de correo sea valida
    ''' </summary>
    ''' <param name="correo1"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function IsValidEmailFormatB(ByVal correo1 As String) As Boolean
        Return Regex.IsMatch(correo1, "^([0-9a-zA-Z]([-\.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$")
    End Function

    Private Sub pbSiguiente_Click(sender As Object, e As EventArgs) Handles pbSiguiente.Click
        TabControl1.SelectedTab = TabPage2
    End Sub








    'ING. ANIBAL LO AMAMOS <3
End Class
