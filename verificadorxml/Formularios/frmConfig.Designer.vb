<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConfig
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
        Me.cbEmpresas = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DTInic = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtPlantillaFac = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnselec = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ckFactura = New System.Windows.Forms.CheckBox()
        Me.CKPoliza = New System.Windows.Forms.CheckBox()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnbuspol = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtPlantillaPol = New System.Windows.Forms.TextBox()
        Me.Lmen = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cbEmpresas
        '
        Me.cbEmpresas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbEmpresas.FormattingEnabled = True
        Me.cbEmpresas.Location = New System.Drawing.Point(12, 29)
        Me.cbEmpresas.Name = "cbEmpresas"
        Me.cbEmpresas.Size = New System.Drawing.Size(319, 21)
        Me.cbEmpresas.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Empresas:"
        '
        'DTInic
        '
        Me.DTInic.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTInic.Location = New System.Drawing.Point(15, 74)
        Me.DTInic.Name = "DTInic"
        Me.DTInic.Size = New System.Drawing.Size(101, 20)
        Me.DTInic.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Fecha Inicio"
        '
        'txtPlantillaFac
        '
        Me.txtPlantillaFac.Location = New System.Drawing.Point(80, 117)
        Me.txtPlantillaFac.Name = "txtPlantillaFac"
        Me.txtPlantillaFac.Size = New System.Drawing.Size(234, 20)
        Me.txtPlantillaFac.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(77, 101)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(85, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Plantilla Factura:"
        '
        'btnselec
        '
        Me.btnselec.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnselec.Location = New System.Drawing.Point(320, 117)
        Me.btnselec.Name = "btnselec"
        Me.btnselec.Size = New System.Drawing.Size(28, 20)
        Me.btnselec.TabIndex = 6
        Me.btnselec.Text = "..."
        Me.btnselec.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnGuardar.Location = New System.Drawing.Point(73, 192)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(75, 31)
        Me.btnGuardar.TabIndex = 7
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(19, 229)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(312, 32)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Solo se consideran las empresas que tienen Fecha de Inicio" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "para los procesos"
        '
        'ckFactura
        '
        Me.ckFactura.AutoSize = True
        Me.ckFactura.Location = New System.Drawing.Point(7, 119)
        Me.ckFactura.Name = "ckFactura"
        Me.ckFactura.Size = New System.Drawing.Size(67, 17)
        Me.ckFactura.TabIndex = 9
        Me.ckFactura.Text = "Facturas"
        Me.ckFactura.UseVisualStyleBackColor = True
        '
        'CKPoliza
        '
        Me.CKPoliza.AutoSize = True
        Me.CKPoliza.Location = New System.Drawing.Point(7, 166)
        Me.CKPoliza.Name = "CKPoliza"
        Me.CKPoliza.Size = New System.Drawing.Size(59, 17)
        Me.CKPoliza.TabIndex = 10
        Me.CKPoliza.Text = "Polizas"
        Me.CKPoliza.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.BackColor = System.Drawing.Color.Red
        Me.btnEliminar.Enabled = False
        Me.btnEliminar.Location = New System.Drawing.Point(165, 192)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(75, 31)
        Me.btnEliminar.TabIndex = 11
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = False
        '
        'btnbuspol
        '
        Me.btnbuspol.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnbuspol.Location = New System.Drawing.Point(320, 166)
        Me.btnbuspol.Name = "btnbuspol"
        Me.btnbuspol.Size = New System.Drawing.Size(28, 20)
        Me.btnbuspol.TabIndex = 14
        Me.btnbuspol.Text = "..."
        Me.btnbuspol.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(77, 150)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(77, 13)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Plantilla Poliza:"
        '
        'txtPlantillaPol
        '
        Me.txtPlantillaPol.Location = New System.Drawing.Point(80, 166)
        Me.txtPlantillaPol.Name = "txtPlantillaPol"
        Me.txtPlantillaPol.Size = New System.Drawing.Size(234, 20)
        Me.txtPlantillaPol.TabIndex = 12
        '
        'Lmen
        '
        Me.Lmen.AutoSize = True
        Me.Lmen.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lmen.ForeColor = System.Drawing.Color.Red
        Me.Lmen.Location = New System.Drawing.Point(126, 74)
        Me.Lmen.Name = "Lmen"
        Me.Lmen.Size = New System.Drawing.Size(212, 13)
        Me.Lmen.TabIndex = 15
        Me.Lmen.Text = "La Empresa no tiene fecha de Inicio"
        Me.Lmen.Visible = False
        '
        'frmConfig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(350, 270)
        Me.Controls.Add(Me.Lmen)
        Me.Controls.Add(Me.btnbuspol)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtPlantillaPol)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.CKPoliza)
        Me.Controls.Add(Me.ckFactura)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.btnselec)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtPlantillaFac)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DTInic)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbEmpresas)
        Me.MaximizeBox = False
        Me.Name = "frmConfig"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Configuración"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cbEmpresas As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents DTInic As DateTimePicker
    Friend WithEvents Label2 As Label
    Friend WithEvents txtPlantillaFac As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents btnselec As Button
    Friend WithEvents btnGuardar As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents ckFactura As CheckBox
    Friend WithEvents CKPoliza As CheckBox
    Friend WithEvents btnEliminar As Button
    Friend WithEvents btnbuspol As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents txtPlantillaPol As TextBox
    Friend WithEvents Lmen As Label
End Class
