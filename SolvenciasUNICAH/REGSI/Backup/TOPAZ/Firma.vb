Imports System.Data
Imports System
Imports Topaz
Imports System.IO
Imports Topaz.SigPlusNET


Public Class Firma
    Dim sign As Bitmap
    Dim ok As Bitmap
    Dim clear As Bitmap
    Dim please As Bitmap
    Dim lcdX As Integer
    Dim lcdY As Integer
    Dim lcdSize As Integer
    Dim screen As Integer
    Dim data As String
    Dim data2 As String
    Public firmado As Boolean

    Private Sub cmdStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStart.Click
        'The following code will write BMP images out to the LCD 1X5 screen


    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        'reset hardware
        SigPlusNET1.LCDRefresh(0, 0, 0, 240, 64)
        SigPlusNET1.LCDSetWindow(0, 0, 240, 64)
        SigPlusNET1.SetSigWindow(1, 0, 0, 240, 64)
        SigPlusNET1.KeyPadClearHotSpotList()
        SigPlusNET1.SetLCDCaptureMode(1)
        SigPlusNET1.SetTabletState(0)
    End Sub

    Private Sub SigPlusNET1_PenDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SigPlusNET1.PenDown

    End Sub

    Private Sub SigPlusNET1_PenUp(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SigPlusNET1.PenUp
        Dim strSig As String

        If SigPlusNET1.KeyPadQueryHotSpot(0) > 0 Then 'If the Continue hotspot is tapped, then...

            If screen = 1 Then

                SigPlusNET1.ClearSigWindow(1)
                SigPlusNET1.LCDRefresh(1, 16, 45, 50, 15) 'Refresh LCD at 'Continue' to indicate to user that this option has been sucessfully chosen
                SigPlusNET1.ClearTablet()
                SigPlusNET1.LCDRefresh(0, 0, 0, 240, 64)

                'Demo text
                Dim f As Font
                f = New System.Drawing.Font("Arial", 9.0F, System.Drawing.FontStyle.Regular)
                data2 = "We'll bind the signature to all the displayed text. Please press Continue."
                Dim words() As String = data2.Split(New Char() {" "c})
                Dim writeData As String
                Dim tempData As String

                writeData = ""
                tempData = ""

                Dim xSize As Integer
                Dim ySize As Integer
                Dim i As Integer
                Dim yPos As Integer

                yPos = 0

                For i = 0 To (words.Length - 1)

                    tempData = tempData & words(i)

                    xSize = SigPlusNET1.LCDStringWidth(f, tempData)

                    If xSize < lcdX Then

                        writeData = tempData
                        tempData = tempData & " "

                        xSize = SigPlusNET1.LCDStringWidth(f, tempData)

                        If xSize < lcdX Then

                            writeData = tempData

                        End If

                    Else

                        ySize = SigPlusNET1.LCDStringHeight(f, tempData)

                        SigPlusNET1.LCDWriteString(0, 2, 0, yPos, f, writeData)

                        tempData = ""
                        writeData = ""
                        yPos += ySize
                        i = (i - 1)

                    End If
                Next i

                If writeData <> "" Then

                    SigPlusNET1.LCDWriteString(0, 2, 0, yPos, f, writeData)

                End If


                'Hotspot text
                SigPlusNET1.LCDWriteString(0, 2, 15, 45, f, "Continuar")
                SigPlusNET1.LCDWriteString(0, 2, 200, 45, f, "Back")

                screen = 2

            ElseIf screen = 2 Then

                SigPlusNET1.ClearSigWindow(1)
                SigPlusNET1.LCDRefresh(1, 16, 45, 50, 15) 'Refresh LCD at 'Continue' to indicate to user that this option has been sucessfully chosen
                SigPlusNET1.LCDRefresh(2, 0, 0, 240, 64) 'Brings the background image already loaded into foreground
                SigPlusNET1.ClearTablet()
                SigPlusNET1.KeyPadClearHotSpotList()
                SigPlusNET1.KeyPadAddHotSpot(2, 1, 10, 5, 53, 17) 'For CLEAR button
                SigPlusNET1.KeyPadAddHotSpot(3, 1, 197, 5, 19, 17) 'For OK button
                SigPlusNET1.LCDSetWindow(0, 22, 238, 40)
                SigPlusNET1.SetSigWindow(1, 0, 22, 240, 40) 'Sets the area where ink is permitted in the SigPlus object

            End If

            SigPlusNET1.SetLCDCaptureMode(2)

        End If

        If SigPlusNET1.KeyPadQueryHotSpot(1) > 0 Then ' If the Exit hotspot is tapped, then...

            If screen = 1 Then

                SigPlusNET1.ClearSigWindow(1)
                SigPlusNET1.LCDRefresh(1, 200, 45, 20, 15) 'Refresh (invert) LCD at 'EXIT' to indicate to user that this option has been sucessfully chosen
                SigPlusNET1.SetLCDCaptureMode(1)
                SigPlusNET1.LCDRefresh(0, 0, 0, 240, 64)
                Me.Close()
                'reset hardware
                SigPlusNET1.SetTabletState(0)
                'Application.Exit()

            ElseIf screen = 2 Then

                SigPlusNET1.ClearSigWindow(1)
                SigPlusNET1.LCDRefresh(1, 200, 45, 25, 15) 'Refresh (invert) LCD at 'Back' to indicate to user that this option has been sucessfully chosen
                SigPlusNET1.ClearTablet()
                SigPlusNET1.LCDRefresh(0, 0, 0, 240, 64)

                'Demo text
                Dim f As Font
                f = New System.Drawing.Font("Arial", 9.0F, System.Drawing.FontStyle.Regular)
                data = "Esta Seguro que desea salir."
                Dim words() As String = data.Split(New Char() {" "c})
                Dim writeData As String
                Dim tempData As String

                writeData = ""
                tempData = ""

                Dim xSize As Integer
                Dim ySize As Integer
                Dim i As Integer
                Dim yPos As Integer

                yPos = 0

                For i = 0 To (words.Length - 1)

                    tempData = tempData & words(i)

                    xSize = SigPlusNET1.LCDStringWidth(f, tempData)

                    If xSize < lcdX Then

                        writeData = tempData
                        tempData = tempData & " "

                        xSize = SigPlusNET1.LCDStringWidth(f, tempData)

                        If xSize < lcdX Then

                            writeData = tempData

                        End If

                    Else

                        ySize = SigPlusNET1.LCDStringHeight(f, tempData)

                        SigPlusNET1.LCDWriteString(0, 2, 0, yPos, f, writeData)

                        tempData = ""
                        writeData = ""
                        yPos += ySize
                        i = (i - 1)

                    End If
                Next i

                If writeData <> "" Then

                    SigPlusNET1.LCDWriteString(0, 2, 0, yPos, f, writeData)

                End If


                'Hotspot text
                SigPlusNET1.LCDWriteString(0, 2, 15, 45, f, "Continuar")
                SigPlusNET1.LCDWriteString(0, 2, 200, 45, f, "Salir")

                'Create the hot spots for the Continue and Exit buttons
                SigPlusNET1.KeyPadAddHotSpot(0, 1, 12, 40, 40, 15) 'For Continue button
                SigPlusNET1.KeyPadAddHotSpot(1, 1, 195, 40, 20, 15) 'For Exit button

                SigPlusNET1.ClearTablet()

                screen = 2

            End If

            SigPlusNET1.SetLCDCaptureMode(2)

        End If

        If SigPlusNET1.KeyPadQueryHotSpot(2) > 0 Then 'if the CLEAR hotspot is tapped, then...
            SigPlusNET1.ClearSigWindow(1)
            SigPlusNET1.LCDRefresh(1, 10, 0, 53, 17) 'Refresh LCD at 'CLEAR' to indicate to user that this option has been sucessfully chosen
            SigPlusNET1.LCDRefresh(2, 0, 0, 240, 64) 'Brings the background image already loaded into foreground
            SigPlusNET1.ClearTablet()
        End If

        If SigPlusNET1.KeyPadQueryHotSpot(3) > 0 Then 'if the OK hotspot is tapped, then...
            SigPlusNET1.ClearSigWindow(1)

            'the following code is used to cryptographically bind the
            'signature to some specific data, passed in
            'using the AutoKeyData property
            'the signature will not be decrypted without this data
            SigPlusNET1.SetSigCompressionMode(1)
            SigPlusNET1.AutoKeyStart()
            SigPlusNET1.SetAutoKeyData(data)
            'SigPlusNET1.SetAutoKeyData(data2)
            SigPlusNET1.AutoKeyFinish()
            SigPlusNET1.SetEncryptionMode(2)

            '*********************Two ways to save the signature*
            '*************************************'
            'Method 1--storing as an ASCII string value
            strSig = SigPlusNET1.GetSigString()
            'the strSig String variable now holds the signature as a long ASCII string.
            'this can be stored as desired, in a database, etc.
            'Method 2--storing as a SIG file on the hard drive

            SigPlusNET1.ExportSigFile("C:\Firmas\Firma.sig")
            'The commented-out function above will export the signature to the SIG file
            'specified (in this case C:\SigFile1.sig, saving the signature as a file on your hardrive
            '*****************************************************************************************'
            SigPlusNET1.SetImageXSize(500)
            SigPlusNET1.SetImageYSize(150)
            SigPlusNET1.SetJustifyMode(5)

            Dim myimage As Image

            myimage = SigPlusNET1.GetSigImage()
            myimage.Save("C:\Firmas\Firma.bmp", System.Drawing.Imaging.ImageFormat.Bmp)

            SigPlusNET1.SetJustifyMode(0)







            SigPlusNET1.LCDRefresh(1, 210, 3, 14, 14) 'Refresh LCD at 'OK' to indicate to user that this option has been sucessfully chosen

            If SigPlusNET1.NumberOfTabletPoints > 0 Then
                SigPlusNET1.LCDRefresh(0, 0, 0, 240, 64)
                Dim f As Font = New System.Drawing.Font("Arial", 9.0F, System.Drawing.FontStyle.Regular)
                SigPlusNET1.LCDWriteString(0, 2, 35, 25, f, "Firma Capturada Exitosamente")
                System.Threading.Thread.Sleep(2000)
                firmado = True
                'Application.Exit()

                Me.Close()
            Else
                SigPlusNET1.LCDRefresh(0, 0, 0, 240, 64)
                SigPlusNET1.LCDSendGraphic(0, 2, 4, 20, please)
                System.Threading.Thread.Sleep(750)
                SigPlusNET1.ClearTablet()
                SigPlusNET1.LCDRefresh(2, 0, 0, 240, 64)
                SigPlusNET1.SetLCDCaptureMode(2)   'Sets mode so ink will not disappear after a few seconds
            End If
        End If

        SigPlusNET1.ClearSigWindow(1)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        firmado = False

        sign = New System.Drawing.Bitmap(Application.StartupPath & "\Sign.bmp")
        ok = New System.Drawing.Bitmap(Application.StartupPath & "\OK.bmp")
        clear = New System.Drawing.Bitmap(Application.StartupPath & "\CLEAR.bmp")
        please = New System.Drawing.Bitmap(Application.StartupPath & "\please.bmp")

        SigPlusNET1.SetTabletState(1)  'Turns tablet on to collect signature
        SigPlusNET1.LCDRefresh(0, 0, 0, 240, 64)
        SigPlusNET1.SetTranslateBitmapEnable(False)

        'Images sent to the background
        SigPlusNET1.LCDSendGraphic(1, 2, 0, 20, sign)
        SigPlusNET1.LCDSendGraphic(1, 2, 207, 4, ok)
        SigPlusNET1.LCDSendGraphic(1, 2, 15, 4, clear)

        'Check for LCD size in pixels.
        'lcdSize = SigPlusNET1.GetLCDSize()
        'lcdX = lcdSize And &HFFFF
        'lcdY = (lcdSize >> 16) And &HFFFF
        lcdX = 240
        lcdY = 64

        'Demo text
        Dim f As Font
        f = New System.Drawing.Font("Arial", 7.0F, System.Drawing.FontStyle.Regular)
        'data = "These are sample terms and conditions. Please press Continue."
        data = "UNIVERSIDAD CATOLICA DE HONDURAS CAMPUS SAN ISIDRO."
        Dim words() As String = data.Split(New Char() {" "c})
        Dim writeData As String
        Dim tempData As String

        writeData = ""
        tempData = ""

        Dim xSize As Integer
        Dim ySize As Integer
        Dim i As Integer
        Dim yPos As Integer

        yPos = 0

        For i = 0 To (words.Length - 1)

            tempData = tempData & words(i)

            xSize = SigPlusNET1.LCDStringWidth(f, tempData)

            If xSize < lcdX Then

                writeData = tempData
                tempData = tempData & " "

                xSize = SigPlusNET1.LCDStringWidth(f, tempData)

                If xSize < lcdX Then

                    writeData = tempData

                End If

            Else

                ySize = SigPlusNET1.LCDStringHeight(f, tempData)

                SigPlusNET1.LCDWriteString(0, 2, 0, yPos, f, writeData)

                tempData = ""
                writeData = ""
                yPos += ySize
                i = (i - 1)

            End If
        Next i

        If writeData <> "" Then

            SigPlusNET1.LCDWriteString(0, 2, 0, yPos, f, writeData)

        End If

        'Hotspot text
        SigPlusNET1.LCDWriteString(0, 2, 15, 45, f, "Continuar")
        SigPlusNET1.LCDWriteString(0, 2, 200, 45, f, "Salir")

        'Create the hot spots for the Continue and Exit buttons
        SigPlusNET1.KeyPadAddHotSpot(0, 1, 12, 40, 40, 15) 'For Continue button
        SigPlusNET1.KeyPadAddHotSpot(1, 1, 195, 40, 20, 15) 'For Exit button

        SigPlusNET1.ClearTablet()

        SigPlusNET1.LCDSetWindow(0, 0, 1, 1)
        SigPlusNET1.SetSigWindow(1, 0, 0, 1, 1) 'Sets the area where ink is permitted in the SigPlus object
        SigPlusNET1.SetLCDCaptureMode(2)   'Sets mode so ink will not disappear after a few seconds
        screen = 2
    End Sub

    Private Sub SigPlusNET1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SigPlusNET1.Click

    End Sub
End Class
