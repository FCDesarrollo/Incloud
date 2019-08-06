<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProcesar
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbEmpresas = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DTInic = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DTFin = New System.Windows.Forms.DateTimePicker()
        Me.CKPoliza = New System.Windows.Forms.CheckBox()
        Me.ckFactura = New System.Windows.Forms.CheckBox()
        Me.btnProcesar = New System.Windows.Forms.Button()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Empresas:"
        '
        'cbEmpresas
        '
        Me.cbEmpresas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbEmpresas.FormattingEnabled = True
        Me.cbEmpresas.Location = New System.Drawing.Point(12, 28)
        Me.cbEmpresas.Name = "cbEmpresas"
        Me.cbEmpresas.Size = New System.Drawing.Size(283, 21)
        Me.cbEmpresas.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Fecha Inicio"
        '
        'DTInic
        '
        Me.DTInic.CustomFormat = "dd/MM/yyyy"
        Me.DTInic.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTInic.Location = New System.Drawing.Point(15, 80)
        Me.DTInic.Name = "DTInic"
        Me.DTInic.Size = New System.Drawing.Size(101, 20)
        Me.DTInic.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(146, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Fecha Final"
        '
        'DTFin
        '
        Me.DTFin.CustomFormat = "dd/MM/yyyy"
        Me.DTFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTFin.Location = New System.Drawing.Point(149, 80)
        Me.DTFin.Name = "DTFin"
        Me.DTFin.Size = New System.Drawing.Size(101, 20)
        Me.DTFin.TabIndex = 6
        '
        'CKPoliza
        '
        Me.CKPoliza.AutoSize = True
        Me.CKPoliza.Location = New System.Drawing.Point(122, 119)
        Me.CKPoliza.Name = "CKPoliza"
        Me.CKPoliza.Size = New System.Drawing.Size(59, 17)
        Me.CKPoliza.TabIndex = 14
        Me.CKPoliza.Text = "Polizas"
        Me.CKPoliza.UseVisualStyleBackColor = True
        '
        'ckFactura
        '
        Me.ckFactura.AutoSize = True
        Me.ckFactura.Location = New System.Drawing.Point(15, 119)
        Me.ckFactura.Name = "ckFactura"
        Me.ckFactura.Size = New System.Drawing.Size(67, 17)
        Me.ckFactura.TabIndex = 13
        Me.ckFactura.Text = "Facturas"
        Me.ckFactura.UseVisualStyleBackColor = True
        '
        'btnProcesar
        '
        Me.btnProcesar.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnProcesar.Location = New System.Drawing.Point(59, 156)
        Me.btnProcesar.Name = "btnProcesar"
        Me.btnProcesar.Size = New System.Drawing.Size(75, 31)
        Me.btnProcesar.TabIndex = 15
        Me.btnProcesar.Text = "Procesar"
        Me.btnProcesar.UseVisualStyleBackColor = False
        '
        'btnSalir
        '
        Me.btnSalir.Location = New System.Drawing.Point(149, 156)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(75, 31)
        Me.btnSalir.TabIndex = 16
        Me.btnSalir.Text = "Salir"
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'frmProcesar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(307, 199)
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.btnProcesar)
        Me.Controls.Add(Me.CKPoliza)
        Me.Controls.Add(Me.ckFactura)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.DTFin)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DTInic)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbEmpresas)
        Me.MaximizeBox = False
        Me.Name = "frmProcesar"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Procesos"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents cbEmpresas As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents DTInic As DateTimePicker
    Friend WithEvents Label3 As Label
    Friend WithEvents DTFin As DateTimePicker
    Friend WithEvents CKPoliza As CheckBox
    Friend WithEvents ckFactura As CheckBox
    Friend WithEvents btnProcesar As Button
    Friend WithEvents btnSalir As Button
End Class
