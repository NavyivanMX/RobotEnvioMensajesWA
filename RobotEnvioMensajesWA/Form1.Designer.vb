<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrincipal
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
        Me.components = New System.ComponentModel.Container()
        Me.LBLOGO = New System.Windows.Forms.ListBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SB = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LBLEJE = New System.Windows.Forms.Label()
        Me.SB.SuspendLayout()
        Me.SuspendLayout()
        '
        'LBLOGO
        '
        Me.LBLOGO.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LBLOGO.BackColor = System.Drawing.Color.Black
        Me.LBLOGO.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBLOGO.ForeColor = System.Drawing.Color.White
        Me.LBLOGO.HorizontalScrollbar = True
        Me.LBLOGO.ItemHeight = 20
        Me.LBLOGO.Location = New System.Drawing.Point(13, 109)
        Me.LBLOGO.Margin = New System.Windows.Forms.Padding(4)
        Me.LBLOGO.Name = "LBLOGO"
        Me.LBLOGO.ScrollAlwaysVisible = True
        Me.LBLOGO.Size = New System.Drawing.Size(949, 504)
        Me.LBLOGO.TabIndex = 1226
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'SB
        '
        Me.SB.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.SB.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2, Me.ToolStripStatusLabel3})
        Me.SB.Location = New System.Drawing.Point(0, 617)
        Me.SB.Name = "SB"
        Me.SB.Size = New System.Drawing.Size(975, 25)
        Me.SB.TabIndex = 1227
        Me.SB.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(138, 20)
        Me.ToolStripStatusLabel1.Text = "Los Mochis, Sinaloa"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(225, 20)
        Me.ToolStripStatusLabel2.Text = "Envio de mensajes de WhatsApp"
        '
        'ToolStripStatusLabel3
        '
        Me.ToolStripStatusLabel3.Name = "ToolStripStatusLabel3"
        Me.ToolStripStatusLabel3.Size = New System.Drawing.Size(95, 20)
        Me.ToolStripStatusLabel3.Text = "Fecha y Hora"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Black
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(13, 41)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 20)
        Me.Label1.TabIndex = 1248
        Me.Label1.Text = "Impuestos"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Black
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(13, 9)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(168, 29)
        Me.Label2.TabIndex = 1247
        Me.Label2.Text = "Timbrado 4.0"
        '
        'LBLEJE
        '
        Me.LBLEJE.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LBLEJE.BackColor = System.Drawing.Color.Black
        Me.LBLEJE.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBLEJE.ForeColor = System.Drawing.Color.Cyan
        Me.LBLEJE.Location = New System.Drawing.Point(16, 67)
        Me.LBLEJE.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LBLEJE.Name = "LBLEJE"
        Me.LBLEJE.Size = New System.Drawing.Size(853, 33)
        Me.LBLEJE.TabIndex = 1246
        Me.LBLEJE.Text = "Ejecutando desde: "
        '
        'frmPrincipal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(975, 642)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.LBLEJE)
        Me.Controls.Add(Me.SB)
        Me.Controls.Add(Me.LBLOGO)
        Me.Name = "frmPrincipal"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Robot de envio de mensajes de whatsapp"
        Me.SB.ResumeLayout(False)
        Me.SB.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LBLOGO As ListBox
    Friend WithEvents Timer1 As Timer
    Friend WithEvents SB As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel3 As ToolStripStatusLabel
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents LBLEJE As Label
End Class
