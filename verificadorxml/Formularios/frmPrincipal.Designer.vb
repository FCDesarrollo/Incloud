<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmprincipal
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmprincipal))
        Me.niClose = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.tMenu = New System.Windows.Forms.ToolStrip()
        Me.MConfig = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.tMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'niClose
        '
        Me.niClose.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.niClose.BalloonTipText = "Aplicación Corriendo"
        Me.niClose.BalloonTipTitle = "App is Running"
        Me.niClose.Icon = CType(resources.GetObject("niClose.Icon"), System.Drawing.Icon)
        Me.niClose.Text = "Incloud"
        Me.niClose.Visible = True
        '
        'tMenu
        '
        Me.tMenu.AutoSize = False
        Me.tMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MConfig, Me.ToolStripButton1})
        Me.tMenu.Location = New System.Drawing.Point(0, 0)
        Me.tMenu.Name = "tMenu"
        Me.tMenu.Size = New System.Drawing.Size(210, 62)
        Me.tMenu.TabIndex = 5
        Me.tMenu.Text = "tMenu"
        '
        'MConfig
        '
        Me.MConfig.AutoSize = False
        Me.MConfig.Image = CType(resources.GetObject("MConfig.Image"), System.Drawing.Image)
        Me.MConfig.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.MConfig.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.MConfig.Name = "MConfig"
        Me.MConfig.Size = New System.Drawing.Size(85, 60)
        Me.MConfig.Tag = "-1"
        Me.MConfig.Text = "Configuración"
        Me.MConfig.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.AutoSize = False
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(85, 55)
        Me.ToolStripButton1.Text = "Procesar"
        Me.ToolStripButton1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'frmprincipal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(210, 76)
        Me.Controls.Add(Me.tMenu)
        Me.MaximizeBox = False
        Me.Name = "frmprincipal"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inicio"
        Me.tMenu.ResumeLayout(False)
        Me.tMenu.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents niClose As NotifyIcon
    Friend WithEvents tMenu As ToolStrip
    Friend WithEvents MConfig As ToolStripButton
    Friend WithEvents ToolStripButton1 As ToolStripButton
End Class
