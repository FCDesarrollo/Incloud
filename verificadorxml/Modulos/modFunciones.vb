Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop

Module modFuncione
    Public cgXML As New Collection
    Public Sub CreaXML(ByVal cNombreEmpresa As String,
                       Optional allEmp As Boolean = True,
                       Optional Fechai As Date = Nothing,
                       Optional FechaF As Date = Nothing)

        Dim cQuery As String, cVersion As Integer, cVersionAnt As Integer
        Dim nomCon As String, fechaini As Date, fechafin As Date
        Dim cVersionGuarda As Integer, idEmpresa As Integer, plantilla As String
        ferror = FC_Conexion()
        If ferror <> 0 Then Exit Sub
        For t = 0 To DConexionesXML.Count - 1
            nomCon = DConexionesXML.Keys(t)
            plantilla = ""
            If cNombreEmpresa = nomCon Or allEmp = True Then
                If IsNothing(DConexiones) Then FC_GetCons()
                DConexiones("CON").ChangeDatabase(DConexionesCFDI(nomCon).Database)

                cVersion = 0

                cQuery = "SELECT TOP (1) lastVersion FROM zEEFControlVersion WHERE Tipo='" & "XML" & "' ORDER BY lastVersion DESC"
                cVersionAnt = GetDatoInt(cQuery, "lastVersion", DConexiones("CON"))

                cQuery = "SELECT id, FechaAutomatic FROM EEFEmpresas WHERE NomEmpresa='" & nomCon & "'"
                idEmpresa = GetDatoInt(cQuery, "id", FC_Con)

                If allEmp = True Then
                    fechaini = Format(GetDatoFecha(cQuery, "FechaAutomatic", FC_Con), "yyyy-MM-dd")
                    fechafin = Format(Date.Now, "yyyy-MM-dd")
                Else
                    fechaini = Format(Fechai, "yyyy-MM-dd")
                    fechafin = Format(FechaF, "yyyy-MM-dd")
                End If

                cQuery = "SELECT plantilla FROM EEFPlantillaDoc
                            WHERE idempresa=@idemp AND tipo=@tip"
                Using cCom = New SqlCommand(cQuery, FC_Con)
                    cCom.Parameters.AddWithValue("@idemp", idEmpresa)
                    cCom.Parameters.AddWithValue("@tip", tFactura)
                    Using cR = cCom.ExecuteReader()
                        cR.Read()
                        If cR.HasRows Then
                            If System.IO.File.Exists(cR("plantilla")) Then
                                plantilla = Path.GetFileName(cR("plantilla"))
                            Else
                                If allEmp = True Then
                                    My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ARCHIVOSXML\errores.log", Format(Now, "01/MM/yyy HH:mm") & " - la plantilla no se encontro de la empresa." & nomCon & "" & vbCrLf, True)
                                Else
                                    MsgBox("La plantilla de la empresa. " & nomCon & " no se encontro", vbInformation, "Validación")
                                End If
                                GoTo Otraempresa
                            End If
                        End If
                    End Using
                End Using

                If cVersionAnt > 0 Then
                    cQuery = "DECLARE @last_synchronization_version bigint
                        SET @last_synchronization_version = CHANGE_TRACKING_MIN_VALID_VERSION(OBJECT_ID('dbo.Comprobante'))
                            SELECT TOP 1 Ct.* FROM CHANGETABLE(CHANGES Comprobante, @last_synchronization_version) as CT  ORDER  BY CT.SYS_CHANGE_VERSION DESC"
                    cVersion = GetDatoInt(cQuery, "SYS_CHANGE_VERSION", DConexionesXML(nomCon))
                End If

                If esBoton = True Then
                    cVersionGuarda = cVersion
                    cVersion = 0
                End If

                If cVersion = 0 Then
                    ConsultaXML(cVersion, DConexionesXML(nomCon), DConexiones("CON"), fechaini, fechafin, nomCon, plantilla)
                ElseIf cVersion <> cVersionAnt Then
                    cQuery = "DECLARE @last_synchronization_version bigint
                        SET @last_synchronization_version = CHANGE_TRACKING_MIN_VALID_VERSION(OBJECT_ID('dbo.Comprobante'))
                            SELECT  Ct.* FROM CHANGETABLE(CHANGES Comprobante, @last_synchronization_version) as CT 
                            WHERE SYS_CHANGE_VERSION > @uversion ORDER  BY CT.SYS_CHANGE_VERSION DESC"
                    Using mcom = New SqlCommand(cQuery, DConexionesXML(nomCon))
                        mcom.Parameters.AddWithValue("@uversion", cVersionAnt)
                        Using mRs = mcom.ExecuteReader()
                            Do While mRs.Read()
                                ConsultaXML(cVersion, DConexionesXML(nomCon),
                                            DConexiones("CON"), fechaini, fechafin, nomCon, plantilla, GetDatoUUID(mRs("GuidDocument"), DConexionesXML(nomCon)))
                                cVersion = mRs("SYS_CHANGE_VERSION")
                            Loop
                        End Using
                    End Using

                End If

                If esBoton = True Then cVersion = cVersionGuarda
                cVersion = IIf(cVersion = 0, 1, cVersion)
                cQuery = "DELETE FROM zEEFControlVersion WHERE Tipo=@tip"
                Using dleCom = New SqlCommand(cQuery, DConexiones("CON"))
                    dleCom.Parameters.AddWithValue("@tip", "XML")
                    dleCom.ExecuteNonQuery()
                End Using
                cQuery = "INSERT INTO zEEFControlVersion(lastVersion, fecha_version, Tipo)
                                            VALUES(@last, @fechave, @tip)"
                Using gCom = New SqlCommand(cQuery, DConexiones("CON"))
                    gCom.Parameters.AddWithValue("@last", cVersion)
                    gCom.Parameters.AddWithValue("@fechave", Date.Now.Date)
                    gCom.Parameters.AddWithValue("@tip", "XML")
                    gCom.ExecuteNonQuery()

                End Using
            End If
Otraempresa:
        Next
    End Sub


    Public Sub CreaExpediente(ByVal cNombreEmpresa As String,
                       Optional allEmp As Boolean = True,
                       Optional Fechai As Date = Nothing,
                       Optional FechaF As Date = Nothing)
        Dim cQuery As String, cVersion As Integer, cVersionAnt As Integer
        Dim nomCon As String, fechaini As Date, fechafin As Date
        Dim cVersionGuarda As Integer, idEmpresa As Integer, cQueryAsoc As String

        For t = 0 To DConexionesCFDI.Count - 1
            nomCon = DConexionesCFDI.Keys(t)
            If cNombreEmpresa = nomCon Or allEmp = True Then
                If IsNothing(DConexiones) Then FC_GetCons()
                DConexiones("CON").ChangeDatabase(DConexionesCFDI(nomCon).Database)

                cVersion = 0

                cQuery = "SELECT TOP (1) lastVersion FROM zEEFControlVersion WHERE Tipo='" & "CFDI" & "' ORDER BY lastVersion DESC"
                cVersionAnt = GetDatoInt(cQuery, "lastVersion", DConexiones("CON"))
                cQuery = "SELECT id, FechaAutomatic FROM EEFEmpresas WHERE NomEmpresa='" & nomCon & "'"
                idEmpresa = GetDatoInt(cQuery, "id", FC_Con)

                If allEmp = True Then
                    fechaini = Format(GetDatoFecha(cQuery, "FechaAutomatic", FC_Con), "yyyy-MM-dd")
                    fechafin = Format(Date.Now, "yyyy-MM-dd")
                Else
                    fechaini = Format(Fechai, "yyyy-MM-dd")
                    fechafin = Format(FechaF, "yyyy-MM-dd")
                End If

                If cVersionAnt > 0 Then
                    cQuery = "DECLARE @last_synchronization_version bigint
                        SET @last_synchronization_version = CHANGE_TRACKING_MIN_VALID_VERSION(OBJECT_ID('dbo.AsocCFDIs'))
                            SELECT TOP 1 Ct.* FROM CHANGETABLE(CHANGES AsocCFDIs, @last_synchronization_version) as CT  ORDER  BY CT.SYS_CHANGE_VERSION DESC"
                    cVersion = GetDatoInt(cQuery, "SYS_CHANGE_VERSION", DConexionesCFDI(nomCon))
                End If

                If esBoton = True Then
                    cVersionGuarda = cVersion
                    cVersion = 0
                End If

                If cVersion = 0 Then
                    cQuery = "SELECT DISTINCT a.UUID FROM AsocCFDIs a 
                            LEFT JOIN MovimientosPoliza AS m ON a.GuidRef = m.Guid 
                            WHERE  Cast(m.Fecha As Date)>=@fech "
                    Using mcom = New SqlCommand(cQuery, DConexionesCFDI(nomCon))
                        mcom.Parameters.AddWithValue("@fech", Format(fechaini, "yyyy-MM-dd"))
                        'mcom.Parameters.AddWithValue("@fechFin", Format(fechafin, "yyyy-MM-dd"))
                        Using mRs = mcom.ExecuteReader()
                            Do While mRs.Read()
                                ImprimeExpediente(mRs("UUID"), DConexionesCFDI(nomCon), nomCon)
                            Loop
                        End Using
                    End Using
                ElseIf cVersion <> cVersionAnt Then
                    cQuery = "DECLARE @last_synchronization_version bigint
                        SET @last_synchronization_version = CHANGE_TRACKING_MIN_VALID_VERSION(OBJECT_ID('dbo.AsocCFDIs'))
                            SELECT  Ct.* FROM CHANGETABLE(CHANGES AsocCFDIs, @last_synchronization_version) as CT 
                            WHERE SYS_CHANGE_VERSION > @uversion ORDER  BY CT.SYS_CHANGE_VERSION ASC"
                    Using mcom = New SqlCommand(cQuery, DConexionesCFDI(nomCon))
                        mcom.Parameters.AddWithValue("@uversion", cVersionAnt)
                        Using mRs = mcom.ExecuteReader()
                            Do While mRs.Read()
                                If mRs("SYS_CHANGE_OPERATION") = "I" Then
                                    cQuery = "SELECT a.UUID FROM AsocCFDIs a 
                                            LEFT JOIN MovimientosPoliza AS m ON a.GuidRef = m.Guid 
                                            WHERE Cast(m.Fecha As Date)>=@fech  AND a.id=@id"
                                    Using mcomA = New SqlCommand(cQuery, DConexionesCFDI(nomCon))
                                        mcomA.Parameters.AddWithValue("@fech", Format(fechaini, "yyyy-MM-dd"))
                                        ' mcomA.Parameters.AddWithValue("@fechFin", Format(fechafin, "yyyy-MM-dd"))
                                        mcomA.Parameters.AddWithValue("@id", mRs("id"))
                                        Using mRsA = mcomA.ExecuteReader()
                                            Do While mRsA.Read()
                                                ImprimeExpediente(mRsA("UUID"), DConexionesCFDI(nomCon), nomCon)
                                                cQueryAsoc = "INSERT INTO zEEFControlUUID(idAsocCFDI, UUID)VALUES(@idasoc, @uuid)"
                                                Using cCom = New SqlCommand(cQueryAsoc, DConexionesCFDI(nomCon))
                                                    cCom.Parameters.AddWithValue("@idasoc", mRs("id"))
                                                    cCom.Parameters.AddWithValue("@uuid", mRsA("UUID"))
                                                    cCom.ExecuteNonQuery()
                                                End Using
                                            Loop
                                        End Using
                                    End Using
                                ElseIf mRs("SYS_CHANGE_OPERATION") = "D" Then
                                    cQueryAsoc = "SELECT UUID FROM zEEFControlUUID WHERE idAsocCFDI=@idasoc"
                                    Using cCom = New SqlCommand(cQueryAsoc, DConexionesCFDI(nomCon))
                                        cCom.Parameters.AddWithValue("@idasoc", mRs("id"))
                                        Using cr = cCom.ExecuteReader()
                                            cr.Read()
                                            If cr.HasRows Then
                                                ImprimeExpediente(cr("UUID"), DConexionesCFDI(nomCon), nomCon)
                                                cQueryAsoc = "DELETE FROM zEEFControlUUID WHERE idAsocCFDI=@idasoc"
                                                Using cComD = New SqlCommand(cQueryAsoc, DConexionesCFDI(nomCon))
                                                    cComD.Parameters.AddWithValue("@idasoc", mRs("id"))
                                                    cComD.ExecuteNonQuery()
                                                End Using
                                            End If
                                        End Using
                                    End Using
                                End If
                                cVersion = mRs("SYS_CHANGE_VERSION")
                            Loop
                        End Using
                    End Using
                End If

                If esBoton = True Then cVersion = cVersionGuarda
                cVersion = IIf(cVersion = 0, 1, cVersion)
                cQuery = "DELETE FROM zEEFControlVersion WHERE Tipo=@tip"
                Using dleCom = New SqlCommand(cQuery, DConexiones("CON"))
                    dleCom.Parameters.AddWithValue("@tip", "CFDI")
                    dleCom.ExecuteNonQuery()
                End Using
                cQuery = "INSERT INTO zEEFControlVersion(lastVersion, fecha_version, Tipo)
                                            VALUES(@last, @fechave, @tip)"
                Using gCom = New SqlCommand(cQuery, DConexiones("CON"))
                    gCom.Parameters.AddWithValue("@last", cVersion)
                    gCom.Parameters.AddWithValue("@fechave", Date.Now.Date)
                    gCom.Parameters.AddWithValue("@tip", "CFDI")
                    gCom.ExecuteNonQuery()

                End Using
            End If
Otraempresa:
        Next
    End Sub

    Public Function BaseConSeguimiento(ByVal bddEmp As String) As Boolean
        Dim bQue As String
        BaseConSeguimiento = False
        bQue = "SELECT db_name(database_id) as 'Object Name'
                from sys.change_tracking_databases 
                WHERE db_name(database_id) = '" & bddEmp & "'"
        Using cCom = New SqlCommand(bQue, DConexiones("CON"))
            Using cCr = cCom.ExecuteReader()
                cCr.Read()
                If cCr.HasRows Then
                    BaseConSeguimiento = True
                End If
            End Using
        End Using
    End Function

    Public Function TablaConSeguimiento(ByVal tTabla As String, ByVal tCone As SqlConnection) As Boolean
        Dim bQue As String
        TablaConSeguimiento = False
        bQue = "SELECT OBJECT_NAME(OBJECT_ID) as 'Object Name',*
                FROM sys.change_tracking_tables
                WHERE OBJECT_NAME(OBJECT_ID)='" & tTabla & "'"
        Using cCom = New SqlCommand(bQue, tCone)
            Using cCr = cCom.ExecuteReader()
                cCr.Read()
                If cCr.HasRows Then
                    TablaConSeguimiento = True
                End If
            End Using
        End Using
    End Function

    Public Sub CreaTablas(ByVal nombase As String)
        Dim cpCom As SqlCommand
        Dim cQue As String
        If IsNothing(DConexiones) Then FC_GetCons()
        DConexiones("CON").ChangeDatabase(nombase)

        cQue = "IF OBJECT_ID('dbo.zEEFControlVersion') IS NULL " &
                    "CREATE TABLE [dbo].[zEEFControlVersion](
	                [lastVersion] [bigint] NOT NULL,
	                [fecha_version] [date] NULL,
	                [fecha_inicio] [date] NULL,
	                [Tipo] [nvarchar](50) NULL) ON [PRIMARY]"
        cpCom = New SqlCommand(cQue, DConexiones("CON"))
        cpCom.ExecuteNonQuery()
        cpCom.Dispose()

        cQue = "IF OBJECT_ID('dbo.zEEFControlUUID') IS NULL " &
                    "CREATE TABLE [dbo].[zEEFControlUUID](
	                    [idAsocCFDI] [int] NULL,
	                    [UUID] [varchar](36) NULL) ON [PRIMARY] "
        cpCom = New SqlCommand(cQue, DConexiones("CON"))
        cpCom.ExecuteNonQuery()
        cpCom.Dispose()
    End Sub

    Public Sub CreaPoliza(ByVal cNombreEmpresa As String, cTipo As String, cTabla As String,
                       Optional allEmp As Boolean = True,
                       Optional Fechai As Date = Nothing,
                       Optional FechaF As Date = Nothing)
        Dim cQuery As String, cVersion As Integer, cVersionAnt As Integer
        Dim nomCon As String, fechaini As Date, fechafin As Date
        Dim cVersionGuarda As Integer, idEmpresa As Integer, cQueryAsoc As String
        Dim plantilla As String, guid As String, clasEmpresa As CLEmpresa

        For t = 0 To PConexionesPol.Count - 1
            nomCon = PConexionesPol.Keys(t)
            If cNombreEmpresa = nomCon Or allEmp = True Then
                If IsNothing(DConexiones) Then FC_GetCons()
                DConexiones("CON").ChangeDatabase(PConexionesPol(nomCon).Database)

                cVersion = 0

                cQuery = "SELECT TOP (1) lastVersion FROM zEEFControlVersion WHERE Tipo='" & cTipo & "' ORDER BY lastVersion DESC"
                cVersionAnt = GetDatoInt(cQuery, "lastVersion", DConexiones("CON"))
                cQuery = "SELECT id, FechaAutomatic FROM EEFEmpresas WHERE NomEmpresa='" & nomCon & "'"
                idEmpresa = GetDatoInt(cQuery, "id", FC_Con)

                If allEmp = True Then
                    fechaini = Format(GetDatoFecha(cQuery, "FechaAutomatic", FC_Con), "yyyy-MM-dd")
                    fechafin = Format(Date.Now, "yyyy-MM-dd")
                Else
                    fechaini = Format(Fechai, "yyyy-MM-dd")
                    fechafin = Format(FechaF, "yyyy-MM-dd")
                End If

                cQuery = "SELECT IdEmpresa,RFC,GuidDSL,Direccion,CodPostal,RegCamara,RegEstatal
                          FROM Parametros"
                Using cCom = New SqlCommand(cQuery, DConexiones("CON"))
                    Using cCr = cCom.ExecuteReader()
                        cCr.Read()
                        If cCr.HasRows Then
                            clasEmpresa = New CLEmpresa
                            clasEmpresa.CIDEmpresa = cCr("IdEmpresa")
                            clasEmpresa.CRFCEmpresa = cCr("RFC")
                            clasEmpresa.CGuidDSL = cCr("GuidDSL")
                            clasEmpresa.CCodigoPostal = cCr("CodPostal")
                            clasEmpresa.CDireccion = Replace(cCr("Direccion"), Chr(10), "")
                            clasEmpresa.CRegCamara = cCr("RegCamara")
                            clasEmpresa.CRegEstatal = cCr("RegEstatal")
                            clasEmpresa.CNomEmpresa = clasEmpresa.ObtenerNombreEmpresa(cCr("IdEmpresa"))
                            FC_ConexionGUID("document_" & cCr("GuidDSL") & "_metadata")
                        Else
                            Exit Sub
                        End If
                    End Using
                End Using

                DConexiones("CON").ChangeDatabase(PConexionesPol(nomCon).Database)

                cQuery = "SELECT plantilla FROM EEFPlantillaDoc
                            WHERE idempresa=@idemp AND tipo=@tip"
                Using cCom = New SqlCommand(cQuery, FC_Con)
                    cCom.Parameters.AddWithValue("@idemp", idEmpresa)
                    cCom.Parameters.AddWithValue("@tip", tPoliza)
                    Using cR = cCom.ExecuteReader()
                        cR.Read()
                        If cR.HasRows Then
                            If System.IO.File.Exists(cR("plantilla")) Then
                                plantilla = Path.GetFileName(cR("plantilla"))
                            Else
                                If allEmp = True Then
                                    My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ARCHIVOSXML\errores.log", Format(Now, "01/MM/yyy HH:mm") & " - la plantilla no se encontro de la empresa." & nomCon & "" & vbCrLf, True)
                                Else
                                    MsgBox("La plantilla de la empresa. " & nomCon & " no se encontro", vbInformation, "Validación")
                                End If
                                GoTo Otraempresa
                            End If
                        End If
                    End Using
                End Using

                If cVersionAnt > 0 Then
                    cQuery = "DECLARE @last_synchronization_version bigint
                        SET @last_synchronization_version = CHANGE_TRACKING_MIN_VALID_VERSION(OBJECT_ID('dbo." & cTabla & "'))
                            SELECT TOP 1 Ct.* FROM CHANGETABLE(CHANGES " & cTabla & ", @last_synchronization_version) as CT  ORDER  BY CT.SYS_CHANGE_VERSION DESC"
                    cVersion = GetDatoInt(cQuery, "SYS_CHANGE_VERSION", PConexionesPol(nomCon))
                End If

                If esBoton = True Then
                    cVersionGuarda = cVersion
                    cVersion = 0
                End If

                If cVersion = 0 Then
                    cQuery = "SELECT Id FROM Polizas 
                            WHERE Cast(Fecha As Date)>=@fech AND Cast(Fecha As Date)<=@fechFin"
                    Using mcom = New SqlCommand(cQuery, PConexionesPol(nomCon))
                        mcom.Parameters.AddWithValue("@fech", Format(fechaini, "yyyy-MM-dd"))
                        mcom.Parameters.AddWithValue("@fechFin", Format(fechafin, "yyyy-MM-dd"))
                        Using mRs = mcom.ExecuteReader()
                            Do While mRs.Read()
                                guid = ImprimePoliza(nomCon, mRs("id"), plantilla, fechaini, fechafin, clasEmpresa)
                                If guid <> "" Then
                                    cQueryAsoc = "INSERT INTO zEEFControlPoliza(idPoliza, Guid)VALUES(@idpol, @guid)"
                                    Using cCom = New SqlCommand(cQueryAsoc, DConexionesCFDI(nomCon))
                                        cCom.Parameters.AddWithValue("@idpol", mRs("id"))
                                        cCom.Parameters.AddWithValue("@guid", guid)
                                        cCom.ExecuteNonQuery()
                                    End Using
                                End If
                            Loop
                        End Using
                    End Using
                ElseIf cVersion <> cVersionAnt Then
                    If cTabla = "Polizas" Then
                        cQuery = "DECLARE @last_synchronization_version bigint
                        SET @last_synchronization_version = CHANGE_TRACKING_MIN_VALID_VERSION(OBJECT_ID('dbo." & cTabla & "'))
                            SELECT  Ct.* FROM CHANGETABLE(CHANGES " & cTabla & ", @last_synchronization_version) as CT 
                            WHERE SYS_CHANGE_VERSION > @uversion ORDER  BY CT.SYS_CHANGE_VERSION ASC"
                    Else
                        cQuery = "DECLARE @last_synchronization_version bigint
                        SET @last_synchronization_version = CHANGE_TRACKING_MIN_VALID_VERSION(OBJECT_ID('dbo." & cTabla & "'))
                            SELECT  Ct.*,t.IdPoliza FROM CHANGETABLE(CHANGES '" & cTabla & "', @last_synchronization_version) as CT 
                            INNER JOIN " & cTabla & " t ON  Ct.id=t.Id
                            WHERE SYS_CHANGE_VERSION > @uversion ORDER  BY CT.SYS_CHANGE_VERSION ASC"
                    End If
                    Using mcom = New SqlCommand(cQuery, PConexionesPol(nomCon))
                        mcom.Parameters.AddWithValue("@uversion", cVersionAnt)
                        Using mRs = mcom.ExecuteReader()
                            Do While mRs.Read()
                                If mRs("SYS_CHANGE_OPERATION") <> "D" Then

                                    guid = ImprimePoliza(nomCon, mRs("id"), plantilla, fechaini, fechafin, clasEmpresa)
                                    If guid <> "" Then
                                        cQueryAsoc = "INSERT INTO zEEFControlPoliza(idPoliza, Guid)VALUES(@idpol, @guid)"
                                        Using cCom = New SqlCommand(cQueryAsoc, DConexionesCFDI(nomCon))
                                            cCom.Parameters.AddWithValue("@idpol", mRs("id"))
                                            cCom.Parameters.AddWithValue("@guid", guid)
                                            cCom.ExecuteNonQuery()
                                        End Using
                                    End If
                                Else
                                    'cQueryAsoc = "SELECT UUID FROM zEEFControlUUID WHERE idAsocCFDI=@idasoc"
                                    'Using cCom = New SqlCommand(cQueryAsoc, DConexionesCFDI(nomCon))
                                    '    cCom.Parameters.AddWithValue("@idasoc", mRs("id"))
                                    '    Using cr = cCom.ExecuteReader()
                                    '        cr.Read()
                                    '        If cr.HasRows Then
                                    '            ImprimeExpediente(cr("UUID"), DConexionesCFDI(nomCon), nomCon)
                                    '            cQueryAsoc = "DELETE FROM zEEFControlUUID WHERE idAsocCFDI=@idasoc"
                                    '            Using cComD = New SqlCommand(cQueryAsoc, DConexionesCFDI(nomCon))
                                    '                cComD.Parameters.AddWithValue("@idasoc", mRs("id"))
                                    '                cComD.ExecuteNonQuery()
                                    '            End Using
                                    '        End If
                                    '    End Using
                                    'End Using
                                End If
                                cVersion = mRs("SYS_CHANGE_VERSION")
                            Loop
                        End Using
                    End Using
                End If

                If esBoton = False Then
                    cVersion = IIf(cVersion = 0, 1, cVersion)
                    cQuery = "DELETE FROM zEEFControlVersion WHERE Tipo=@tip"
                    Using dleCom = New SqlCommand(cQuery, DConexiones("CON"))
                        dleCom.Parameters.AddWithValue("@tip", cTipo)
                        dleCom.ExecuteNonQuery()
                    End Using
                    cQuery = "INSERT INTO zEEFControlVersion(lastVersion, fecha_version, Tipo)
                                            VALUES(@last, @fechave, @tip)"
                    Using gCom = New SqlCommand(cQuery, DConexiones("CON"))
                        gCom.Parameters.AddWithValue("@last", cVersion)
                        gCom.Parameters.AddWithValue("@fechave", Date.Now.Date)
                        gCom.Parameters.AddWithValue("@tip", cTipo)
                        gCom.ExecuteNonQuery()

                    End Using
                End If
            End If
Otraempresa:
        Next
    End Sub

    Public Function getLastRow(ByRef sht As Excel.Worksheet) As Long
        On Error GoTo Err
        getLastRow = sht.Cells.Find("*", SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlPrevious).Row
        Exit Function
Err:
        If Err.Number = 91 Then getLastRow = 0
    End Function
End Module
