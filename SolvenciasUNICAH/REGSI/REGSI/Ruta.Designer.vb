<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Ruta
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtDir = New System.Windows.Forms.TextBox()
        Me.txtDir1 = New System.Windows.Forms.TextBox()
        Me.btnAceptar = New System.Windows.Forms.PictureBox()
        Me.btnCargar = New System.Windows.Forms.PictureBox()
        Me.btnCrear = New System.Windows.Forms.PictureBox()
        CType(Me.btnAceptar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnCargar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnCrear, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtDir
        '
        Me.txtDir.Location = New System.Drawing.Point(307, 128)
        Me.txtDir.Multiline = True
        Me.txtDir.Name = "txtDir"
        Me.txtDir.Size = New System.Drawing.Size(466, 46)
        Me.txtDir.TabIndex = 0
        '
        'txtDir1
        '
        Me.txtDir1.Location = New System.Drawing.Point(307, 216)
        Me.txtDir1.Multiline = True
        Me.txtDir1.Name = "txtDir1"
        Me.txtDir1.Size = New System.Drawing.Size(466, 46)
        Me.txtDir1.TabIndex = 1
        '
        'btnAceptar
        '
        Me.btnAceptar.Location = New System.Drawing.Point(663, 375)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(100, 50)
        Me.btnAceptar.TabIndex = 2
        Me.btnAceptar.TabStop = False
        '
        'btnCargar
        '
        Me.btnCargar.Location = New System.Drawing.Point(144, 375)
        Me.btnCargar.Name = "btnCargar"
        Me.btnCargar.Size = New System.Drawing.Size(100, 50)
        Me.btnCargar.TabIndex = 3
        Me.btnCargar.TabStop = False
        '
        'btnCrear
        '
        Me.btnCrear.Location = New System.Drawing.Point(22, 375)
        Me.btnCrear.Name = "btnCrear"
        Me.btnCrear.Size = New System.Drawing.Size(100, 50)
        Me.btnCrear.TabIndex = 4
        Me.btnCrear.TabStop = False
        '
        'Ruta
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.btnCrear)
        Me.Controls.Add(Me.btnCargar)
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me.txtDir1)
        Me.Controls.Add(Me.txtDir)
        Me.Name = "Ruta"
        Me.Text = "Ruta"
        CType(Me.btnAceptar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnCargar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnCrear, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtDir As TextBox
    Friend WithEvents txtDir1 As TextBox
    Friend WithEvents btnAceptar As PictureBox
    Friend WithEvents btnCargar As PictureBox
    Friend WithEvents btnCrear As PictureBox

    Private Sub TxtDir_TextChanged(sender As Object, e As EventArgs) Handles txtDir.TextChanged

    End Sub
End Class
