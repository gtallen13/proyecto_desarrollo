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
    Dim Correlativo As String
    Dim contadorAdmnistrativo As Integer = 0
    Dim contadorAcademico As Integer = 0
    Dim ContadorRazones As Integer = 0
    Dim intcorrelativoSOL As Integer = 0
    Dim intcorrelativoCA As Integer = 0
    Dim intCorrelativoPRA As Integer = 0
    Dim intCorrelativoTL As Integer = 0
    Dim intCorrelativoCE As Integer = 0
    Dim intCorrelativoGRA As Integer = 0
    Dim intCorrelativoRTEM As Integer = 0
    Dim intCorrelativoCAN As Integer = 0
    Dim intCorrelativoRTOT As Integer = 0
    Dim intCorrelativoRI As Integer = 0
    Dim intCorrelativoGeneral As Integer = 0
    Dim blnGuardado As Boolean = False

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        CmboBoxCarrera.Text = ""
        cmbCampus.Text = "San Isidro - La Ceiba"
        TxtBoxCuenta.Text = ""
        TxtTelefono.Text = ""
        TxtCorreo.Text = ""
        TxtBoxNombreAlumno.Text = ""
        checkinternado.Checked = False
        CheckInicioPractica.Checked = False
        CheckGraduacion.Checked = False
        CheckOtros_Razones.Checked = False
        CheckOtros_TramitesAcademicos.Checked = False
        CheckRetiroDeLaUniversidad.Checked = False
        CheckRetiroPorNormasQueRigenLaVidaEstudiantil.Checked = False
        CheckSalud.Checked = False
        CheckServicioSocial.Checked = False
        CheckSyllabusAsignatura.Checked = False
        CheckTrabajo.Checked = False
        CheckTraslado.Checked = False
        CheckAbandonoDeLaUniversidad.Checked = False
        CheckActivarCuentaParaReintegro.Checked = False

        TabControl1.SelectedTab = TabPage1
        Becado.Checked = True

        Dim fecha As Date
        fecha = Now
        lblFecha.Text = Format(fecha, "dd/MM/yyyy")
        Label20.Text = "Favor Firmar el Documento"
        Label20.ForeColor = Color.Orange
        btnGuardar.Enabled = True
        RichTextBoxEspecificarRazones.Enabled = False

        Dim Doc As New XmlDocument()
        Dim xmlnode As XmlNodeList
        Dim i As Integer = 0
        Doc.Load(Application.StartupPath & "\Versiculo.xml")
        xmlnode = Doc.GetElementsByTagName("Versiculo")
        lbBiblia.Text = xmlnode(0).ChildNodes.Item(0).InnerText.Trim()

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtBoxNombreAlumno.TextChanged

        Dim theText As String = TxtBoxNombreAlumno.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TxtBoxNombreAlumno.SelectionStart
        Dim Change As Integer
        Dim charactersDisallowed As String = "1234567890"


        For x As Integer = 0 To TxtBoxNombreAlumno.Text.Length - 1
            Letter = TxtBoxNombreAlumno.Text.Substring(x, 1)
            If charactersDisallowed.Contains(Letter) Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TxtBoxNombreAlumno.Text = theText
        TxtBoxNombreAlumno.Select(SelectionIndex - Change, 0)
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckTraslado.CheckedChanged
        If CheckTraslado.Checked = True Then
            CcmbCampusDestino.Enabled = True
            Correlativo = "TL"
            intCorrelativoTL += 1
        Else
            CcmbCampusDestino.Text = ""
            CcmbCampusDestino.Enabled = False
            Correlativo = ""
            intCorrelativoTL -= 1
        End If
    End Sub

    Private Sub TabPage2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckSyllabusAsignatura.CheckedChanged
        If CheckSyllabusAsignatura.CheckState = CheckState.Checked Then
            Correlativo = "SOL"
            intcorrelativoSOL += 1
            intCorrelativoGeneral = 1
        Else
            Correlativo = ""
            intcorrelativoSOL -= 1
            intCorrelativoGeneral -= 1
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckCertificadoDeEstudio.CheckedChanged
        If CheckCertificadoDeEstudio.CheckState = CheckState.Checked Then
            Correlativo = "CE"
            intCorrelativoCE += 1
            intCorrelativoGeneral += 1

        Else
            Correlativo = ""
            intCorrelativoCE -= 1
            intCorrelativoGeneral -= 1
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckConstanciaAcademica.CheckedChanged
        If CheckConstanciaAcademica.CheckState = CheckState.Checked Then
            Correlativo = "CA"
            intcorrelativoCA += 1
            intCorrelativoGeneral += 1

        Else
            Correlativo = ""
            intcorrelativoCA -= 1
            intCorrelativoGeneral -= 1

        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckGraduacion.CheckedChanged
        If CheckGraduacion.CheckState = CheckState.Checked Then
            Correlativo = "GRA"
            intCorrelativoGRA += 1
            intCorrelativoGeneral += 1

        Else
            Correlativo = ""
            intCorrelativoGRA -= 1
            intCorrelativoGeneral -= 1

        End If
    End Sub

    Private Sub CheckInicioPractica_CheckedChanged(sender As Object, e As EventArgs) Handles CheckInicioPractica.CheckedChanged
        If CheckInicioPractica.CheckState = CheckState.Checked Then
            Correlativo = "PRA"
            intCorrelativoPRA += 1
            intCorrelativoGeneral += 1

        Else
            Correlativo = ""
            intCorrelativoPRA -= 1
            intCorrelativoGeneral -= 1

        End If
    End Sub


    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckAbandonoDeLaUniversidad.CheckedChanged
        If CheckAbandonoDeLaUniversidad.CheckState = CheckState.Checked Then
            Correlativo = "RTEM"
            intCorrelativoRTEM += 1
            intCorrelativoGeneral += 1

        Else
            Correlativo = ""
            intCorrelativoRTEM -= 1
            intCorrelativoGeneral -= 1

        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckCancelacionDelPeriodo.CheckedChanged
        If CheckCancelacionDelPeriodo.CheckState = CheckState.Checked Then
            Correlativo = "CAN"
            intCorrelativoCAN += 1
            intCorrelativoGeneral += 1

        Else
            Correlativo = ""
            intCorrelativoCAN -= 1
            intCorrelativoGeneral -= 1

        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckRetiroDeLaUniversidad.CheckedChanged
        If CheckSyllabusAsignatura.CheckState = CheckState.Checked Then
            Correlativo = "RTOT"
            intCorrelativoRTOT += 1
            intCorrelativoGeneral += 1

        Else
            Correlativo = ""
            intCorrelativoRTOT -= 1
            intCorrelativoGeneral -= 1

        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckRetiroPorNormasQueRigenLaVidaEstudiantil.CheckedChanged
        If CheckSyllabusAsignatura.CheckState = CheckState.Checked Then
            Correlativo = "RTOT"
            intCorrelativoRTOT += 1
            intCorrelativoGeneral += 1

        Else
            Correlativo = ""
            intCorrelativoRTOT -= 1
            intCorrelativoGeneral -= 1

        End If
    End Sub

    Private Sub CheckBox16_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckActivarCuentaParaReintegro.CheckedChanged
        If CheckActivarCuentaParaReintegro.CheckState = CheckState.Checked Then
            Correlativo = "RI"
            intCorrelativoRI += 1
            intCorrelativoGeneral += 1

        Else
            Correlativo = ""
            intCorrelativoRI -= 1
            intCorrelativoGeneral -= 1

        End If
    End Sub

    Private Sub CheckBox11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckOtros_TramitesAcademicos.CheckedChanged
        If CheckSyllabusAsignatura.CheckState = CheckState.Checked Then
            Correlativo = "SOL"
            intcorrelativoSOL += 1
            intCorrelativoGeneral += 1

        Else
            Correlativo = ""
            intCorrelativoRI -= 1
            intCorrelativoGeneral -= 1

        End If
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckSalud.CheckedChanged
        If CheckSalud.Checked = True Then
            RichTextBoxEspecificarRazones.Enabled = False
        End If
    End Sub
    Private Sub Checkinternado_CheckedChanged(sender As Object, e As EventArgs) Handles checkinternado.CheckedChanged
        If checkinternado.CheckState = CheckState.Checked Then
            Correlativo = "SOL1"
            intcorrelativoSOL += 1
            intCorrelativoGeneral += 1

        Else
            Correlativo = ""
            intcorrelativoSOL -= 1
            intCorrelativoGeneral -= 1
        End If
    End Sub
    Private Sub CheckServicioSocial_CheckedChanged(sender As Object, e As EventArgs) Handles CheckServicioSocial.CheckedChanged
        If CheckServicioSocial.CheckState = CheckState.Checked Then
            Correlativo = "SOL"
            intcorrelativoSOL += 1
            intCorrelativoGeneral += 1
        Else
            Correlativo = ""
            intcorrelativoSOL -= 1
            intCorrelativoGeneral -= 1
        End If
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
        If intCorrelativoGeneral > 1 Then
            newFile = direccionCarpeta & TxtBoxCuenta.Text & "SOL" & intCorrelativoGeneral & ".pdf"
            'el siguiente codigo iterara sobre los archivo creados para asi asignar el numero seguido del Correlativo
            Dim di As New DirectoryInfo(direccionCarpeta)
            Dim firArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo
            For Each fri In firArr
                If fri.FullName = newFile Then
                    intCorrelativoGeneral += 1
                    newFile = direccionCarpeta & TxtBoxCuenta.Text & "SOL" & intcorrelativoSOL & ".pdf"
                End If
            Next
        ElseIf intcorrelativoSOL > 0 Then
            newFile = direccionCarpeta & TxtBoxCuenta.Text & "SOL" & intcorrelativoSOL & ".pdf"
            Dim di As New DirectoryInfo(direccionCarpeta)
            Dim firArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo
            For Each fri In firArr
                If fri.FullName = newFile Then
                    intcorrelativoSOL += 1
                    newFile = direccionCarpeta & TxtBoxCuenta.Text & "SOL" & intcorrelativoSOL & ".pdf"
                End If
            Next
        ElseIf intCorrelativoRTOT > 0 Then
            newFile = direccionCarpeta & TxtBoxCuenta.Text & "RTOT" & intCorrelativoRTOT & ".pdf"
            Dim di As New DirectoryInfo(direccionCarpeta)
            Dim firArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo
            For Each fri In firArr
                If fri.FullName = newFile Then
                    intcorrelativoSOL += 1
                    newFile = direccionCarpeta & TxtBoxCuenta.Text & "SOL" & intcorrelativoSOL & ".pdf"
                End If
            Next
        ElseIf intcorrelativoCA > 0 Then
            newFile = direccionCarpeta & TxtBoxCuenta.Text & "CA" & intcorrelativoCA & ".pdf"
            Dim di As New DirectoryInfo(direccionCarpeta)
            Dim firArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo
            For Each fri In firArr
                If fri.FullName = newFile Then
                    intcorrelativoCA += 1
                    newFile = direccionCarpeta & TxtBoxCuenta.Text & "CA" & intcorrelativoCA & ".pdf"
                End If
            Next
        ElseIf intCorrelativoPRA > 0 Then
            newFile = direccionCarpeta & TxtBoxCuenta.Text & "PRA" & intCorrelativoPRA & ".pdf"
            Dim di As New DirectoryInfo(direccionCarpeta)
            Dim firArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo
            For Each fri In firArr
                If fri.FullName = newFile Then
                    intCorrelativoPRA += 1
                    newFile = direccionCarpeta & TxtBoxCuenta.Text & "PRA" & intCorrelativoPRA & ".pdf"
                End If
            Next
        ElseIf intCorrelativoTL > 0 Then
            newFile = direccionCarpeta & TxtBoxCuenta.Text & "TL" & intCorrelativoTL & ".pdf"
            Dim di As New DirectoryInfo(direccionCarpeta)
            Dim firArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo
            For Each fri In firArr
                If fri.FullName = newFile Then
                    intCorrelativoTL += 1
                    newFile = direccionCarpeta & TxtBoxCuenta.Text & "TL" & intcorrelativoSOL & ".pdf"
                End If
            Next
        ElseIf intCorrelativoCE > 0 Then
            newFile = direccionCarpeta & TxtBoxCuenta.Text & "CE" & intCorrelativoCE & ".pdf"
            Dim di As New DirectoryInfo(direccionCarpeta)
            Dim firArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo
            For Each fri In firArr
                If fri.FullName = newFile Then
                    intCorrelativoCE += 1
                    newFile = direccionCarpeta & TxtBoxCuenta.Text & "CE" & intcorrelativoSOL & ".pdf"
                End If
            Next
        ElseIf intCorrelativoGRA > 0 Then
            newFile = direccionCarpeta & TxtBoxCuenta.Text & "GRA" & intCorrelativoGRA & ".pdf"
            Dim di As New DirectoryInfo(direccionCarpeta)
            Dim firArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo
            For Each fri In firArr
                If fri.FullName = newFile Then
                    intCorrelativoGRA += 1
                    newFile = direccionCarpeta & TxtBoxCuenta.Text & "GRA" & intCorrelativoGRA & ".pdf"
                End If
            Next
        ElseIf intCorrelativoRTEM > 0 Then
            newFile = direccionCarpeta & TxtBoxCuenta.Text & "RTEM" & intCorrelativoRTEM & ".pdf"
            Dim di As New DirectoryInfo(direccionCarpeta)
            Dim firArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo
            For Each fri In firArr
                If fri.FullName = newFile Then
                    intCorrelativoRTEM += 1
                    newFile = direccionCarpeta & TxtBoxCuenta.Text & "RTEM" & intCorrelativoRTEM & ".pdf"
                End If
            Next
        ElseIf intCorrelativoCAN > 0 Then
            newFile = direccionCarpeta & TxtBoxCuenta.Text & "CAN" & intCorrelativoCAN & ".pdf"
            Dim di As New DirectoryInfo(direccionCarpeta)
            Dim firArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo
            For Each fri In firArr
                If fri.FullName = newFile Then
                    intCorrelativoCAN += 1
                    newFile = direccionCarpeta & TxtBoxCuenta.Text & "CAN" & intCorrelativoCAN & ".pdf"
                End If
            Next
        ElseIf intCorrelativoRI > 0 Then
            newFile = direccionCarpeta & TxtBoxCuenta.Text & "RI" & intCorrelativoRI & ".pdf"
            Dim di As New DirectoryInfo(direccionCarpeta)
            Dim firArr As FileInfo() = di.GetFiles()
            Dim fri As FileInfo
            For Each fri In firArr
                If fri.FullName = newFile Then
                    intCorrelativoRI += 1
                    newFile = direccionCarpeta & TxtBoxCuenta.Text & "RI" & intCorrelativoRI & ".pdf"
                End If
            Next
        Else
            newFile = direccionCarpeta & TxtBoxCuenta.Text & Correlativo & ".pdf"
        End If



        Dim pdfReader As New PdfReader(pdfTemplate)
        Dim pdfStamper As New PdfStamper(pdfReader, New FileStream(
            newFile, FileMode.Create))
        Dim pdfFormFields As AcroFields = pdfStamper.AcroFields
        Dim pcbContent As PdfContentByte = Nothing
        Dim img As System.Drawing.Image = System.Drawing.Image.FromFile(direccionFirmas & "\firma.bmp") 'aqui ira la direccion de las firmas
        Dim sap As PdfSignatureAppearance = pdfStamper.SignatureAppearance
        Dim rect As iTextSharp.text.Rectangle = Nothing
        Dim imagen As iTextSharp.text.Image
        Dim loc As String


        loc = direccionFirmas & "\firma.bmp" 'aqui ira la direccion de las firmas

        imagen = iTextSharp.text.Image.GetInstance(Loc)
        imagen.SetAbsolutePosition(427, 43)
        imagen.ScaleToFit(130, 130)
        pcbContent = pdfStamper.GetUnderContent(1)
        pcbContent.AddImage(imagen)

        ' set form pdfFormFields
        pdfFormFields.SetField("Cuenta", TxtBoxCuenta.Text)
        pdfFormFields.SetField("Carrera", CmboBoxCarrera.Text)
        pdfFormFields.SetField("Nombre", TxtBoxNombreAlumno.Text)
        pdfFormFields.SetField("RazonOtro", RichTextBoxEspecificarRazones.Text)
        pdfFormFields.SetField("Fecha_5", lblFecha.Text)
        pdfFormFields.SetField("Tel", TxtTelefono.Text)
        pdfFormFields.SetField("Email", TxtCorreo.Text)
        pdfFormFields.SetField("Campus", cmbCampus.Text)
        pdfFormFields.SetField("CampusDestino", CcmbCampusDestino.Text)
        pdfFormFields.SetField("Recibo", TxtRecibo.Text)
        pdfFormFields.SetField("Internado", checkinternado.Text)
        pdfFormFields.SetField("Social", CheckServicioSocial.Text)
        'pdfFormFields.SetField("signature5", TextBox1.Text)

        ' The form's checkboxes
        If CheckSyllabusAsignatura.Checked = True Then
            pdfFormFields.SetField("Contenidos", "On")
        End If

        If CheckCertificadoDeEstudio.Checked = True Then
            pdfFormFields.SetField("Certificado", "On")
        End If

        If CheckConstanciaAcademica.Checked = True Then
            pdfFormFields.SetField("Constancia", "On")
        End If

        If CheckGraduacion.Checked = True Then
            pdfFormFields.SetField("Graduacion", "On")
        End If

        If CheckInicioPractica.Checked = True Then
            pdfFormFields.SetField("Practica", "On")
        End If

        If CheckTraslado.Checked = True Then
            pdfFormFields.SetField("Traslado", "On")
        End If

        If CheckAbandonoDeLaUniversidad.Checked = True Then
            pdfFormFields.SetField("Abandono", "On")
        End If

        If CheckCancelacionDelPeriodo.Checked = True Then
            pdfFormFields.SetField("Cancelacion", "On")
        End If

        If CheckRetiroDeLaUniversidad.Checked = True Then
            pdfFormFields.SetField("Retiro", "On")
        End If

        If CheckRetiroPorNormasQueRigenLaVidaEstudiantil.Checked = True Then
            pdfFormFields.SetField("Normas", "On")
        End If

        If CheckOtros_TramitesAcademicos.Checked = True Then
            pdfFormFields.SetField("Otros", "On")
        End If

        If CheckSalud.Checked = True Then
            pdfFormFields.SetField("Salud", "On")
        End If

        If CheckTrabajo.Checked = True Then
            pdfFormFields.SetField("Trabajo", "On")
        End If

        If CheckEconomicos.Checked = True Then
            pdfFormFields.SetField("Economicos", "On")
        End If

        If CheckOtros_Razones.Checked = True Then
            pdfFormFields.SetField("Otros_2", "On")
        End If

        If CheckActivarCuentaParaReintegro.Checked = True Then
            pdfFormFields.SetField("Reintegro", "On")
        End If

        If Becado.Checked = True Then
            pdfFormFields.SetField("Beca Si", "On")
        End If

        If NoBecado.Checked = True Then
            pdfFormFields.SetField("Beca No", "Off")
        End If
        If NoBecado.Checked = True Then
            pdfFormFields.SetField("BecasTexto", "N/A")
        End If
        If checkinternado.Checked = True Then
            pdfFormFields.SetField("Internado", "On")
        End If
        If CheckServicioSocial.Checked = True Then
            pdfFormFields.SetField("Social", "On")
        End If

        MessageBox.Show("Datos Guardados Satisfactoriamente")

        ' flatten the form to remove editting options, set it to false
        ' to leave the form open to subsequent manual edits
        pdfStamper.FormFlattening = False

        ' close the pdf
        pdfStamper.Close()


        TxtBoxCuenta.Text = ""
        CmboBoxCarrera.SelectedIndex = -1
        CmboBoxCarrera.SelectedIndex = -1
        TxtBoxNombreAlumno.Text = ""
        TxtTelefono.Text = ""
        TxtCorreo.Text = ""
        TxtRecibo.Text = ""
        CheckSyllabusAsignatura.CheckState = 0
        CheckCertificadoDeEstudio.CheckState = 0
        CheckConstanciaAcademica.CheckState = 0
        CheckGraduacion.CheckState = 0
        CheckInicioPractica.CheckState = 0
        CheckTraslado.CheckState = 0
        CheckAbandonoDeLaUniversidad.CheckState = 0
        CheckCancelacionDelPeriodo.CheckState = 0
        CheckRetiroDeLaUniversidad.CheckState = 0
        CheckRetiroPorNormasQueRigenLaVidaEstudiantil.CheckState = 0
        CheckOtros_TramitesAcademicos.CheckState = 0
        CheckSalud.CheckState = 0
        CheckTrabajo.CheckState = 0
        CheckEconomicos.CheckState = 0
        CheckOtros_Razones.CheckState = 0
        RichTextBoxEspecificarRazones.Text = ""
        RichTextBoxEspecificarRazones.Enabled = False




        Me.Close()

        ' End If
    End Sub

    ''' <summary>
    ''' Aqui se validan todos los tramites academicos
    ''' </summary>
    Sub ValidarAdministrativos()
        contadorAdmnistrativo = 0
        If CheckSyllabusAsignatura.Checked Then
            contadorAdmnistrativo += 1
        ElseIf CheckCertificadoDeEstudio.Checked Then
        contadorAdmnistrativo += 1
        ElseIf CheckConstanciaAcademica.Checked Then
            contadorAdmnistrativo += 1
        ElseIf CheckGraduacion.Checked Then
            contadorAdmnistrativo += 1
        ElseIf CheckInicioPractica.Checked Then
            contadorAdmnistrativo += 1
        ElseIf CheckAbandonoDeLaUniversidad.Checked Then
            contadorAdmnistrativo += 1
        ElseIf CheckCancelacionDelPeriodo.Checked Then
            contadorAdmnistrativo += 1
        ElseIf CheckRetiroDeLaUniversidad.Checked Then
            contadorAdmnistrativo += 1
        ElseIf CheckRetiroPorNormasQueRigenLaVidaEstudiantil.Checked Then
            contadorAdmnistrativo += 1
        ElseIf CheckActivarCuentaParaReintegro.Checked Then
            contadorAdmnistrativo += 1
        ElseIf checkinternado.Checked Then
            contadorAdmnistrativo += 1
        ElseIf CheckServicioSocial.Checked Then
            contadorAdmnistrativo += 1
        ElseIf CheckTraslado.Checked Then
            contadorAdmnistrativo += 1
        ElseIf CheckOtros_TramitesAcademicos.Checked Then
            contadorAdmnistrativo += 1
        Else
            contadorAdmnistrativo = 0
        End If

    End Sub


    ''' <summary>
    ''' Aqui se validan los tramites Academicos
    ''' </summary>
    Sub ValidarAcademicos()
        ContadorRazones = 0 'Utilizado para pedir una razon al usuario si él no la ingresa
        contadorAcademico = 0
        If CheckAbandonoDeLaUniversidad.Checked Then
            contadorAcademico += 1
        End If
        If CheckCancelacionDelPeriodo.Checked Then
            contadorAcademico += 1
        End If
        If CheckRetiroDeLaUniversidad.Checked Then
            contadorAcademico += 1
        End If
        If CheckRetiroPorNormasQueRigenLaVidaEstudiantil.Checked Then
            contadorAcademico += 1
        End If
        If CheckActivarCuentaParaReintegro.Checked Then
            contadorAcademico += 1
        End If
        If CheckEconomicos.Checked Then
            ContadorRazones += 1
        End If
        If CheckSalud.Checked Then
            ContadorRazones += 1
        End If
        If CheckTrabajo.Checked Then
            ContadorRazones += 1
        End If
        If CheckOtros_Razones.Checked And RichTextBoxEspecificarRazones.Text = "" Then
            ContadorRazones += 1
        ElseIf CheckOtros_Razones.Checked And RichTextBoxEspecificarRazones.Text > "" Then
            ContadorRazones += 1
        End If
    End Sub
    Private Sub PictureBox2_Click_1(sender As Object, e As EventArgs) Handles btnSiguiente.Click
        If CheckOtros_Razones.Checked = True And RichTextBoxEspecificarRazones.Text = "" Then
            MessageBox.Show("Especificaciones vacias", "Campo vacio", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If IsNothing(CcmbCampusDestino.Text) = True Then
            MessageBox.Show("Por favor seleccione un Campus", "Informacion Incompleta", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If Me.TxtBoxCuenta.Text.Length = 0 Then
            MessageBox.Show("Ingrese su Numero de Cuenta", "Campo vacio", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        If Me.TxtBoxNombreAlumno.Text.Length = 0 Then
            MessageBox.Show("Ingrese su Nombre", "Campo vacio", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        If CmboBoxCarrera.Text = Nothing Then
            MessageBox.Show("No ha seleccionado una carrera", "Campo Vacio", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If TxtCorreo.Text = Nothing Then
            MessageBox.Show("No ha ingresado un correo", "Campo Vacio", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If IsValidEmailFormat(TxtCorreo.Text) = False And IsValidEmailFormatB(TxtCorreo.Text) = False Then
            MessageBox.Show("Ingrese un correo valido", "Informacion Incompleta", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ValidarAcademicos()
        If contadorAcademico <> 0 Then
            If ContadorRazones = 0 Then
                MessageBox.Show("Seleccione al menos una Razon.", "Informacion Incompleta", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        End If

        ValidarAdministrativos()
        If contadorAdmnistrativo = 0 Then
            MessageBox.Show("Seleccione al menos un tramite", "Campo vacio", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        'preguntara si es becado y guardara la respuesta
        Dim respuesta As Integer = MessageBox.Show("¿Usted es becado?", "Confirmacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If respuesta = DialogResult.No Then
                NoBecado.Checked = True
            ElseIf respuesta = DialogResult.Yes Then
                Becado.Checked = True
            End If
        TabControl1.SelectedTab = TabPage2
    End Sub

    Private Sub TxtTelefono_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtTelefono.KeyPress
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) Or Asc(e.KeyChar) = 8 Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub CheckOtros_Razones_CheckedChanged(sender As Object, e As EventArgs) Handles CheckOtros_Razones.CheckedChanged
        If CheckOtros_Razones.Checked = True Then
            RichTextBoxEspecificarRazones.Enabled = True
        Else
            RichTextBoxEspecificarRazones.Enabled = False
            RichTextBoxEspecificarRazones.Text = ""
        End If
    End Sub

    Private Sub CheckEconomicos_CheckedChanged(sender As Object, e As EventArgs) Handles CheckEconomicos.CheckedChanged
        If CheckEconomicos.Checked = True Then
            RichTextBoxEspecificarRazones.Enabled = False
        End If
    End Sub

    Private Sub CheckTrabajo_CheckedChanged(sender As Object, e As EventArgs) Handles CheckTrabajo.CheckedChanged
        If CheckTrabajo.Checked = True Then
            RichTextBoxEspecificarRazones.Enabled = False
        End If
    End Sub

    Private Sub TxtBoxCuenta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtBoxCuenta.KeyPress
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) Or Asc(e.KeyChar) = 8 Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub CmboBoxCarrera_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmboBoxCarrera.SelectedIndexChanged
        CmboBoxCarrera.Text = ""
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





    'ING. ANIBAL LO AMAMOS <3
End Class
