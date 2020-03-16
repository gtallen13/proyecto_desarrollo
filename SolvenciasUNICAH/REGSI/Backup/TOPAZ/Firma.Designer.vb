<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Firma
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Firma))
        Me.cmdStart = New System.Windows.Forms.Button
        Me.SigPlusNET1 = New Topaz.SigPlusNET
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdStart
        '
        Me.cmdStart.Location = New System.Drawing.Point(371, 121)
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.Size = New System.Drawing.Size(75, 23)
        Me.cmdStart.TabIndex = 0
        Me.cmdStart.Text = "Iniciar"
        Me.cmdStart.UseVisualStyleBackColor = True
        Me.cmdStart.Visible = False
        '
        'SigPlusNET1
        '
        Me.SigPlusNET1.BackColor = System.Drawing.Color.White
        Me.SigPlusNET1.ForeColor = System.Drawing.Color.Black
        Me.SigPlusNET1.Location = New System.Drawing.Point(12, 12)
        Me.SigPlusNET1.Name = "SigPlusNET1"
        Me.SigPlusNET1.Size = New System.Drawing.Size(353, 132)
        Me.SigPlusNET1.TabIndex = 1
        Me.SigPlusNET1.Text = "SigPlusNET1"
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(371, 29)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(100, 50)
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'Firma
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(458, 156)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.SigPlusNET1)
        Me.Controls.Add(Me.cmdStart)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Firma"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Firmar Documento"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdStart As System.Windows.Forms.Button
    Friend WithEvents SigPlusNET1 As Topaz.SigPlusNET
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox

End Class
