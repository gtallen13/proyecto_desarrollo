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


Public Class Form1
    Dim firmar As New sigplusnet_vbnet_lcd15_demo.Firma
    Dim Correlativo As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        If TextBox1.Text = "" Then
            MessageBox.Show("Favor Completar El Formulario")
        Else


            Dim pdfTemplate As String = "C:\Users\Estudiante\Dropbox\REG-PS.505 Solicitud de Solvencias.pdf"
            Dim newFile As String = "C:\Users\Estudiante\Dropbox\Solvencias\Pendientes\" & TextBox1.Text & Correlativo & ".pdf"
            Dim pdfReader As New PdfReader(pdfTemplate)
            Dim pdfStamper As New PdfStamper(pdfReader, New FileStream( _
                newFile, FileMode.Create))
            Dim pdfFormFields As AcroFields = pdfStamper.AcroFields
            Dim pcbContent As PdfContentByte = Nothing
            Dim img As System.Drawing.Image = System.Drawing.Image.FromFile("C:\Firmas\Firma.bmp")
            Dim sap As PdfSignatureAppearance = pdfStamper.SignatureAppearance
            Dim rect As iTextSharp.text.Rectangle = Nothing
            Dim imagen As iTextSharp.text.Image
            Dim loc As String

            loc = "C:\Firmas\Firma.bmp"
            imagen = iTextSharp.text.Image.GetInstance(loc)
            imagen.SetAbsolutePosition(389, 92)
            imagen.ScaleToFit(130, 130)
            pcbContent = pdfStamper.GetUnderContent(1)
            pcbContent.AddImage(imagen)

            ' set form pdfFormFields
            pdfFormFields.SetField("Cuenta", TextBox1.Text)
            pdfFormFields.SetField("Carrera", ComboBox1.Text)
            pdfFormFields.SetField("Nombre", TextBox3.Text)
            pdfFormFields.SetField("RazonOtro", RichTextBox1.Text)
            pdfFormFields.SetField("Fecha_5", Label18.Text)
            pdfFormFields.SetField("Tel", txtTelefono.Text)
            pdfFormFields.SetField("Email", TxtCorreo.Text)
            pdfFormFields.SetField("Campus", "San Isidro")
            pdfFormFields.SetField("CampusDestino", ComboBox2.Text)
            pdfFormFields.SetField("Recibo", txtrecibo.Text)
            'pdfFormFields.SetField("signature5", TextBox1.Text)

            ' The form's checkboxes
            If CheckBox1.Checked = True Then
                pdfFormFields.SetField("Contenidos", "On")
            End If

            If CheckBox2.Checked = True Then
                pdfFormFields.SetField("Certificado", "On")
            End If

            If CheckBox3.Checked = True Then
                pdfFormFields.SetField("Constancia", "On")
            End If

            If CheckBox4.Checked = True Then
                pdfFormFields.SetField("Graduacion", "On")
            End If

            If CheckBox5.Checked = True Then
                pdfFormFields.SetField("Practica", "On")
            End If

            If CheckBox6.Checked = True Then
                pdfFormFields.SetField("Traslado", "On")
            End If

            If CheckBox7.Checked = True Then
                pdfFormFields.SetField("Abandono", "On")
            End If

            If CheckBox8.Checked = True Then
                pdfFormFields.SetField("Cancelacion", "On")
            End If

            If CheckBox9.Checked = True Then
                pdfFormFields.SetField("Retiro", "On")
            End If

            If CheckBox10.Checked = True Then
                pdfFormFields.SetField("Normas", "On")
            End If

            If CheckBox11.Checked = True Then
                pdfFormFields.SetField("Otros", "On")
            End If

            If CheckBox12.Checked = True Then
                pdfFormFields.SetField("Salud", "On")
            End If

            If CheckBox13.Checked = True Then
                pdfFormFields.SetField("Trabajo", "On")
            End If

            If CheckBox14.Checked = True Then
                pdfFormFields.SetField("Economicos", "On")
            End If

            If CheckBox15.Checked = True Then
                pdfFormFields.SetField("Otros_2", "On")
            End If

            If CheckBox16.Checked = True Then
                pdfFormFields.SetField("Reintegro", "On")
            End If


            MessageBox.Show("Datos Guardados Satisfactoriamente")

            ' flatten the form to remove editting options, set it to false
            ' to leave the form open to subsequent manual edits
            pdfStamper.FormFlattening = False

            ' close the pdf
            pdfStamper.Close()


            TextBox1.Text = ""
            ComboBox1.SelectedIndex = -1
            ComboBox1.SelectedIndex = -1
            TextBox3.Text = ""
            txtTelefono.Text = ""
            TxtCorreo.Text = ""
            txtrecibo.Text = ""
            CheckBox1.CheckState = 0
            CheckBox2.CheckState = 0
            CheckBox3.CheckState = 0
            CheckBox4.CheckState = 0
            CheckBox5.CheckState = 0
            CheckBox6.CheckState = 0
            CheckBox7.CheckState = 0
            CheckBox8.CheckState = 0
            CheckBox9.CheckState = 0
            CheckBox10.CheckState = 0
            CheckBox11.CheckState = 0
            CheckBox12.CheckState = 0
            CheckBox13.CheckState = 0
            CheckBox14.CheckState = 0
            CheckBox15.CheckState = 0
            RichTextBox1.Text = ""
            RichTextBox1.Enabled = False



            Me.Close()

        End If
    End Sub




    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.Text = "Arquitectura"
        TabControl1.SelectedTab = TabPage1

        Dim fecha As Date
        fecha = Now
        Label18.Text = Format(fecha, "MM/dd/yyyy")
        Label20.Text = "Favor Firmar el Documento"
        Label20.ForeColor = Color.Orange
        Button1.Enabled = False
        RichTextBox1.Enabled = False


    End Sub

    

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            firmar.ShowDialog()
            If firmar.firmado = True Then
                Label20.Text = "El documento ha sido Firmado"
                Label20.ForeColor = Color.LimeGreen
                Button1.Enabled = True
            End If
        Catch ex As Exception

        End Try

    End Sub


    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

        Dim theText As String = TextBox3.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = TextBox3.SelectionStart
        Dim Change As Integer
        Dim charactersDisallowed As String = "1234567890"


        For x As Integer = 0 To TextBox3.Text.Length - 1
            Letter = TextBox3.Text.Substring(x, 1)
            If charactersDisallowed.Contains(Letter) Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        TextBox3.Text = theText
        TextBox3.Select(SelectionIndex - Change, 0)
    End Sub




    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub CheckBox15_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox15.CheckedChanged


        If CheckBox15.Checked = True Then
            RichTextBox1.Enabled = True
        Else
            RichTextBox1.Enabled = False
            RichTextBox1.Text = ""
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TabControl1.SelectedTab = TabPage2
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TabControl1.SelectedTab = TabPage1
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TabControl1.SelectedTab = TabPage1
    End Sub

    Private Sub TxtCorreo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCorreo.TextChanged

    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            ComboBox2.Enabled = True
            Correlativo = "TL1"
        Else
            ComboBox2.Text = ""
            ComboBox2.Enabled = False
            Correlativo = ""
        End If
    End Sub

    Private Sub TabPage2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = CheckState.Checked Then
            Correlativo = "SOL1"
        Else
            Correlativo = ""
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.CheckState = CheckState.Checked Then
            Correlativo = "SOL1"
        Else
            Correlativo = ""
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.CheckState = CheckState.Checked Then
            Correlativo = "SOL1"
        Else
            Correlativo = ""
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.CheckState = CheckState.Checked Then
            Correlativo = "SOL2"
        Else
            Correlativo = ""
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.CheckState = CheckState.Checked Then
            Correlativo = "SOL1"
        Else
            Correlativo = ""
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.CheckState = CheckState.Checked Then
            Correlativo = "RTOT"
        Else
            Correlativo = ""
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.CheckState = CheckState.Checked Then
            Correlativo = "RTEM1"
        Else
            Correlativo = ""
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.CheckState = CheckState.Checked Then
            Correlativo = "RTOT"
        Else
            Correlativo = ""
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.CheckState = CheckState.Checked Then
            Correlativo = "RTOT"
        Else
            Correlativo = ""
        End If
    End Sub

    Private Sub CheckBox16_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox16.CheckedChanged
        If CheckBox16.CheckState = CheckState.Checked Then
            Correlativo = "SOL1"
        Else
            Correlativo = ""
        End If
    End Sub

    Private Sub CheckBox11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.CheckState = CheckState.Checked Then
            Correlativo = "SOL1"
        Else
            Correlativo = ""
        End If
    End Sub
End Class
