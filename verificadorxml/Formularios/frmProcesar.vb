Imports System.Data.SqlClient

Public Class frmProcesar
    Private Sub FrmProcesar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadEmpresas()
    End Sub

    Private Sub loadEmpresas()
        Dim lQue As String
        ferror = FC_Conexion()
        If ferror <> 0 Then Exit Sub

        Try
            lQue = "SELECT id, NomEmpresa FROM EEFEmpresas WHERE FechaAutomatic IS NOT NULL"
            Dim da As New SqlDataAdapter(lQue, FC_Con)
            Dim ds As New DataSet
            da.Fill(ds)
            With Me.cbEmpresas
                .DataSource = ds.Tables(0)
                .ValueMember = "id"
                .DisplayMember = "NomEmpresa"
                .SelectedItem = Nothing
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub BtnProcesar_Click(sender As Object, e As EventArgs) Handles btnProcesar.Click
        If cbEmpresas.Text <> "" Then
            esBoton = True
            If (ckFactura.Checked = True Or CKPoliza.Checked = True) Then
                If ckFactura.Checked = True Then
                    CreaXML(cbEmpresas.Text, False, DTInic.Value.Date, DTFin.Value.Date)
                    CreaExpediente(cbEmpresas.Text, False, DTInic.Value.Date, DTFin.Value.Date)
                End If
                If CKPoliza.Checked = True Then
                    CreaPoliza(cbEmpresas.Text, "POLIZA", "Polizas", False, DTInic.Value.Date, DTFin.Value.Date)
                End If
                esBoton = False
                'KillAllExcels()
                MsgBox("Se ha Procesado Correctamente.", vbInformation, "Validación")
            Else
                MsgBox("No ha seleccionado ningun proceso.", vbInformation, "Validación")
            End If
        Else
                MsgBox("Seleccione la Empresa para procesar.", vbInformation, "Validación")
        End If
    End Sub
End Class