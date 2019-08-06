Imports System.Data.SqlClient
Imports System.IO

Public Class frmConfig
    Private sBandLoad As Boolean
    Private Sub FrmConfig_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sBandLoad = True
        loadEmpresas()
        sBandLoad = False
    End Sub

    Private Sub loadEmpresas()
        Dim lQue As String
        ferror = FC_Conexion()
        If ferror <> 0 Then Exit Sub

        Try
            lQue = "SELECT id, NomEmpresa FROM EEFEmpresas"
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

    Private Sub CbEmpresas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbEmpresas.SelectedIndexChanged
        Dim idEmp As Integer, cQue As String, cAux As String

        If sBandLoad = False Then
            idEmp = CInt(cbEmpresas.SelectedValue)
            cQue = "SELECT FechaAutomatic FROM EEFEmpresas WHERE id=@id"
            Using cCom = New SqlCommand(cQue, FC_Con)
                cCom.Parameters.AddWithValue("@id", idEmp)
                Using dRs = cCom.ExecuteReader()
                    dRs.Read()
                    If dRs.HasRows Then
                        If dRs("FechaAutomatic") IsNot DBNull.Value Then
                            Lmen.Visible = False
                            DTInic.Value = CDate(dRs("FechaAutomatic"))
                            btnEliminar.Enabled = True
                        Else
                            Lmen.Visible = True
                            btnEliminar.Enabled = False
                        End If
                        cAux = "SELECT activo,tipo,plantilla FROM EEFPlantillaDoc WHERE idempresa=@idem"
                        Using cPlan = New SqlCommand(cAux, FC_Con)
                            cPlan.Parameters.AddWithValue("@idem", idEmp)
                            Using rAux = cPlan.ExecuteReader()
                                Do While rAux.Read()
                                    If rAux("tipo") = tFactura Then
                                        ckFactura.Checked = rAux("activo")
                                        txtPlantillaFac.Tag = rAux("plantilla")
                                        txtPlantillaFac.Text = Path.GetFileName(rAux("plantilla"))
                                    ElseIf rAux("tipo") = tPoliza Then
                                        CKPoliza.Checked = rAux("activo")
                                        txtPlantillaPol.Tag = rAux("plantilla")
                                        txtPlantillaPol.Text = Path.GetFileName(rAux("plantilla"))
                                    End If
                                Loop
                            End Using
                        End Using
                    End If
                End Using
            End Using
        End If
    End Sub

    Private Sub GuardarEmpresa()
        Dim gQue As String, idEmp As Integer
        Dim sEmpresa As String

        idEmp = CInt(cbEmpresas.SelectedValue)
        sEmpresa = cbEmpresas.Text

        If Not Directory.Exists(FC_RutaModulos & "\ARCHIVOSXML\" & sEmpresa) Then
            My.Computer.FileSystem.CreateDirectory(FC_RutaModulos & "\ARCHIVOSXML\" & sEmpresa)
        End If

        gQue = "SELECT BDDCon FROM EEFEmpresas WHERE id=@id"
        Using cCom = New SqlCommand(gQue, FC_Con)
            cCom.Parameters.AddWithValue("@id", idEmp)
            Using cr = cCom.ExecuteReader()
                cr.Read()
                If cr.HasRows Then
                    CreaTablas(cr("BDDCon"))
                Else
                    MsgBox("No se encontro la empresa.", vbInformation, "Validación")
                    Exit Sub
                End If
            End Using
        End Using

        gQue = "UPDATE EEFEmpresas SET FechaAutomatic=@fecha WHERE id=@idemp"
        Using cCom = New SqlCommand(gQue, FC_Con)
            cCom.Parameters.AddWithValue("@fecha", Format(DTInic.Value, "yyyy-MM-dd"))
            cCom.Parameters.AddWithValue("@idemp", idEmp)
            cCom.ExecuteNonQuery()
        End Using

        If ckFactura.Checked = True Then
            FileSystem.FileCopy(txtPlantillaFac.Tag, FC_RutaModulos & "\ARCHIVOSXML\" & sEmpresa & "\" & txtPlantillaFac.Text)
            gQue = "INSERT INTO EEFPlantillaDoc(idempresa,activo,tipo,plantilla)
                    VALUES(@idemp, @activo, @tipo, @plantilla)"
            Using cCom = New SqlCommand(gQue, FC_Con)
                cCom.Parameters.AddWithValue("@idemp", idEmp)
                cCom.Parameters.AddWithValue("@activo", 1)
                cCom.Parameters.AddWithValue("@tipo", tFactura)
                cCom.Parameters.AddWithValue("@plantilla", FC_RutaModulos & "\ARCHIVOSXML\" & sEmpresa & "\" & txtPlantillaFac.Text)
                cCom.ExecuteNonQuery()
            End Using
        End If

        If CKPoliza.Checked = True Then
            FileSystem.FileCopy(txtPlantillaPol.Tag, FC_RutaModulos & "\ARCHIVOSXML\" & sEmpresa & "\" & txtPlantillaPol.Text)
            gQue = "INSERT INTO EEFPlantillaDoc(idempresa,activo,tipo,plantilla)
                    VALUES(@idemp, @activo, @tipo, @plantilla)"
            Using cCom = New SqlCommand(gQue, FC_Con)
                cCom.Parameters.AddWithValue("@idemp", idEmp)
                cCom.Parameters.AddWithValue("@activo", 1)
                cCom.Parameters.AddWithValue("@tipo", tPoliza)
                cCom.Parameters.AddWithValue("@plantilla", FC_RutaModulos & "\ARCHIVOSXML\" & sEmpresa & "\" & txtPlantillaPol.Text)
                cCom.ExecuteNonQuery()
            End Using
        End If
    End Sub

    Private Sub Btnselec_Click(sender As Object, e As EventArgs) Handles btnselec.Click
        If ckFactura.Checked = False Then Exit Sub
        Dim OpenFile As New OpenFileDialog()
        OpenFile.Filter = "Excel Worksheets|*.xlsx"
        If OpenFile.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtPlantillaFac.Tag = OpenFile.FileName
            txtPlantillaFac.Text = Path.GetFileName(OpenFile.FileName)
        End If
    End Sub

    Private Sub Btnbuspol_Click(sender As Object, e As EventArgs) Handles btnbuspol.Click
        If CKPoliza.Checked = False Then Exit Sub
        Dim OpenFile As New OpenFileDialog()
        OpenFile.Filter = "Excel Worksheets|*.xlsx"
        If OpenFile.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtPlantillaPol.Tag = OpenFile.FileName
            txtPlantillaPol.Text = Path.GetFileName(OpenFile.FileName)
        End If
    End Sub

    Private Sub BtnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        If cbEmpresas.Text <> "" Then
            If Lmen.Visible = True Then
                GuardarEmpresa()
            Else
                UpdateEmpresa()
            End If
            Limpiar()
            sBandLoad = True
            loadEmpresas()
            sBandLoad = False
        End If
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        If cbEmpresas.Text <> "" Then
            EliminarEmpresa()
            Limpiar()
            sBandLoad = True
            loadEmpresas()
            sBandLoad = False
        End If
    End Sub

    Private Sub Limpiar()
        ckFactura.Checked = False
        CKPoliza.Checked = False
        txtPlantillaFac.Tag = ""
        txtPlantillaFac.Text = ""
        txtPlantillaPol.Tag = ""
        txtPlantillaPol.Text = ""
        Lmen.Visible = False
    End Sub
    Private Sub EliminarEmpresa()
        Dim eQue As String, idEmp As Integer
        idEmp = CInt(cbEmpresas.SelectedValue)
        eQue = "UPDATE EEFEmpresas SET FechaAutomatic=@fech WHERE id=@idemp"
        Using cCom = New SqlCommand(eQue, FC_Con)
            cCom.Parameters.AddWithValue("@fech", DBNull.Value)
            cCom.Parameters.AddWithValue("@idemp", idEmp)
            cCom.ExecuteNonQuery()
        End Using

        eQue = "DELETE FROM EEFPlantillaDoc WHERE idempresa=@idemp"
        Using cCom = New SqlCommand(eQue, FC_Con)
            cCom.Parameters.AddWithValue("@idemp", idEmp)
            cCom.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub UpdateEmpresa()
        Dim eQue As String, idEmp As Integer, sEmpresa As String
        idEmp = CInt(cbEmpresas.SelectedValue)
        sEmpresa = cbEmpresas.Text
        eQue = "UPDATE EEFEmpresas SET FechaAutomatic=@fech WHERE id=@idemp"
        Using cCom = New SqlCommand(eQue, FC_Con)
            cCom.Parameters.AddWithValue("@fech", DBNull.Value)
            cCom.Parameters.AddWithValue("@idemp", idEmp)
            cCom.ExecuteNonQuery()
        End Using

        If txtPlantillaFac.Tag <> "" Then
            FileSystem.FileCopy(txtPlantillaFac.Tag, FC_RutaModulos & "\ARCHIVOSXML\" & sEmpresa & "\" & txtPlantillaFac.Text)
        End If

        eQue = "UPDATE EEFPlantillaDoc SET activo=@acti, plantilla=@plant
                    WHERE idempresa=@idemp AND tipo=@tip"
        Using cCom = New SqlCommand(eQue, FC_Con)
            cCom.Parameters.AddWithValue("@acti", IIf(ckFactura.Checked, 1, 0))
            cCom.Parameters.AddWithValue("@plant", FC_RutaModulos & "\ARCHIVOSXML\" & sEmpresa & "\" & txtPlantillaFac.Text)
            cCom.Parameters.AddWithValue("@idemp", idEmp)
            cCom.Parameters.AddWithValue("@tip", tFactura)
            cCom.ExecuteNonQuery()
        End Using

        If txtPlantillaPol.Tag <> "" Then
            FileSystem.FileCopy(txtPlantillaPol.Tag, FC_RutaModulos & "\ARCHIVOSXML\" & sEmpresa & "\" & txtPlantillaPol.Text)
        End If

        eQue = "UPDATE EEFPlantillaDoc SET activo=@acti, plantilla=@plant
                    WHERE idempresa=@idemp AND tipo=@tip"
        Using cCom = New SqlCommand(eQue, FC_Con)
            cCom.Parameters.AddWithValue("@acti", IIf(CKPoliza.Checked, 1, 0))
            cCom.Parameters.AddWithValue("@plant", FC_RutaModulos & "\ARCHIVOSXML\" & sEmpresa & "\" & txtPlantillaPol.Text)
            cCom.Parameters.AddWithValue("@idemp", idEmp)
            cCom.Parameters.AddWithValue("@tip", tPoliza)
            cCom.ExecuteNonQuery()
        End Using
    End Sub
End Class