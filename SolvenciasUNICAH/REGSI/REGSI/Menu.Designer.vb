<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Menu
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Menu))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.AboutBox = New System.Windows.Forms.PictureBox()
        Me.Xml = New System.Windows.Forms.PictureBox()
        Me.Cerrar = New System.Windows.Forms.PictureBox()
        Me.Minimizar = New System.Windows.Forms.PictureBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.BtnAbrirSolvencia = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AboutBox, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Xml, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Cerrar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Minimizar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BtnAbrirSolvencia, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(30, Byte), Integer), CType(CType(38, Byte), Integer), CType(CType(70, Byte), Integer))
        Me.Panel1.Controls.Add(Me.PictureBox3)
        Me.Panel1.Controls.Add(Me.AboutBox)
        Me.Panel1.Controls.Add(Me.Xml)
        Me.Panel1.Controls.Add(Me.Cerrar)
        Me.Panel1.Controls.Add(Me.Minimizar)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(631, 42)
        Me.Panel1.TabIndex = 13
        '
        'PictureBox3
        '
        Me.PictureBox3.BackgroundImage = Global.REGSI.My.Resources.Resources.Imagen17
        Me.PictureBox3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox3.Location = New System.Drawing.Point(85, 6)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(33, 31)
        Me.PictureBox3.TabIndex = 4
        Me.PictureBox3.TabStop = False
        '
        'AboutBox
        '
        Me.AboutBox.BackgroundImage = Global.REGSI.My.Resources.Resources.Imagen1
        Me.AboutBox.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.AboutBox.Location = New System.Drawing.Point(46, 7)
        Me.AboutBox.Name = "AboutBox"
        Me.AboutBox.Size = New System.Drawing.Size(33, 31)
        Me.AboutBox.TabIndex = 3
        Me.AboutBox.TabStop = False
        '
        'Xml
        '
        Me.Xml.BackgroundImage = Global.REGSI.My.Resources.Resources.Cargarxml1
        Me.Xml.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Xml.Location = New System.Drawing.Point(7, 6)
        Me.Xml.Name = "Xml"
        Me.Xml.Size = New System.Drawing.Size(33, 31)
        Me.Xml.TabIndex = 2
        Me.Xml.TabStop = False
        '
        'Cerrar
        '
        Me.Cerrar.BackgroundImage = Global.REGSI.My.Resources.Resources.Close_67_
        Me.Cerrar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Cerrar.Location = New System.Drawing.Point(579, 9)
        Me.Cerrar.Name = "Cerrar"
        Me.Cerrar.Size = New System.Drawing.Size(23, 23)
        Me.Cerrar.TabIndex = 1
        Me.Cerrar.TabStop = False
        '
        'Minimizar
        '
        Me.Minimizar.BackgroundImage = Global.REGSI.My.Resources.Resources.Minimize
        Me.Minimizar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Minimizar.Location = New System.Drawing.Point(528, 9)
        Me.Minimizar.Name = "Minimizar"
        Me.Minimizar.Size = New System.Drawing.Size(23, 23)
        Me.Minimizar.TabIndex = 0
        Me.Minimizar.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(30, Byte), Integer), CType(CType(38, Byte), Integer), CType(CType(70, Byte), Integer))
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 306)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(631, 22)
        Me.Panel2.TabIndex = 14
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox2.BackgroundImage = Global.REGSI.My.Resources.Resources.TITULO
        Me.PictureBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox2.Location = New System.Drawing.Point(240, 48)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(362, 121)
        Me.PictureBox2.TabIndex = 16
        Me.PictureBox2.TabStop = False
        '
        'BtnAbrirSolvencia
        '
        Me.BtnAbrirSolvencia.BackColor = System.Drawing.Color.Transparent
        Me.BtnAbrirSolvencia.BackgroundImage = Global.REGSI.My.Resources.Resources.Boton1
        Me.BtnAbrirSolvencia.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnAbrirSolvencia.Location = New System.Drawing.Point(309, 187)
        Me.BtnAbrirSolvencia.Name = "BtnAbrirSolvencia"
        Me.BtnAbrirSolvencia.Size = New System.Drawing.Size(242, 91)
        Me.BtnAbrirSolvencia.TabIndex = 15
        Me.BtnAbrirSolvencia.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.BackgroundImage = Global.REGSI.My.Resources.Resources.LOGO_2B1_68_
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox1.Location = New System.Drawing.Point(0, 43)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(206, 257)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 12
        Me.PictureBox1.TabStop = False
        '
        'Menu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(59, Byte), Integer), CType(CType(104, Byte), Integer))
        Me.BackgroundImage = Global.REGSI.My.Resources.Resources.Imagen2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(631, 328)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.BtnAbrirSolvencia)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PictureBox1)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Menu"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Menu"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AboutBox, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Xml, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Cerrar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Minimizar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BtnAbrirSolvencia, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents BtnAbrirSolvencia As PictureBox
    Friend WithEvents Minimizar As PictureBox
    Friend WithEvents Cerrar As PictureBox
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents Xml As PictureBox
    Friend WithEvents AboutBox As PictureBox
    Friend WithEvents PictureBox3 As PictureBox
End Class
