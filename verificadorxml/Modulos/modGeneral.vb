Imports Microsoft.Win32
Imports System.Data.SqlClient
Module modGeneral
    Public Const FC_REGKEY As String = "HKEY_LOCAL_MACHINE\SOFTWARE\FCModulos\"
    Public Const FC_REGKEYWRITE As String = "SOFTWARE\FCModulos"

    ''VARIABLES DE CONEXION
    Public DConexiones As Dictionary(Of String, SqlConnection)
    Public DConexionesXML As Dictionary(Of String, SqlConnection)
    Public DConexionesCFDI As Dictionary(Of String, SqlConnection)

    Public DConexionesConten As Dictionary(Of String, SqlConnection)

    ''VARIABLES DE CONEXION PARA LOS CAMBIOS DE POLIZAS
    Public PConexionesPol As Dictionary(Of String, SqlConnection)
    'Public PConexionesMovPol As Dictionary(Of String, SqlConnection)
    'Public PConexionesDevo As Dictionary(Of String, SqlConnection)
    'Public PConexionesCausa As Dictionary(Of String, SqlConnection)
    'Public PConexionesAsoc As Dictionary(Of String, SqlConnection)


    Public letr As NumaLet
    Public esBoton As Boolean = False
    Public ferror As Long
    Public FC_Con As New SqlConnection
    Public FC_ConGuid As New SqlConnection

    Private Rs As SqlDataReader
    Private ComRs As SqlCommand

    Public Function FC_GetDatos() As String()
        On Error Resume Next
        Dim sArr(0 To 2) As String
        sArr(0) = My.Computer.Registry.GetValue(FC_REGKEY, "Instancia", Nothing)
        sArr(1) = My.Computer.Registry.GetValue(FC_REGKEY, "Uid", Nothing)
        sArr(2) = My.Computer.Registry.GetValue(FC_REGKEY, "Password", Nothing)
        FC_GetDatos = sArr
    End Function

    Public Sub FC_SetDatos(ByVal Inst As String, ByVal Uid As String, ByVal Pwd As String)
        WriteToRegistry("Instancia", Inst)
        WriteToRegistry("Password", Pwd)
        WriteToRegistry("Uid", Uid)
    End Sub

    Public Sub WriteToRegistry(ByVal Key As String, ByVal Value As String)
        Dim runK As RegistryKey = Registry.LocalMachine.OpenSubKey(FC_REGKEYWRITE, True)
        runK.SetValue(Key, Value)
    End Sub

    Public Property FC_DATABASE() As String
        Get
            On Error Resume Next
            FC_DATABASE = My.Computer.Registry.GetValue(FC_REGKEY, "BDDGen", Nothing)
        End Get
        Set(ByVal val As String)
            On Error Resume Next
            WriteToRegistry("BDDGen", val)
        End Set
    End Property

    Public Property FC_RutaModulos() As String
        Get
            On Error Resume Next
            FC_RutaModulos = My.Computer.Registry.GetValue(FC_REGKEY, "RutaModulos", Nothing)
        End Get
        Set(ByVal val As String)
            On Error Resume Next
            WriteToRegistry("RutaModulos", val)
        End Set
    End Property
    Public Function FC_Conexion() As Long
        Dim conData() As String
        On Error GoTo ERR_CON

        If FC_Con.State = ConnectionState.Closed Then
            conData = FC_GetDatos()
            FC_Con = New SqlConnection()
            FC_Con.ConnectionString = "Data Source=" + conData(0) + " ;" &
                         "Initial Catalog=" + FC_DATABASE + ";" &
                         "User Id=" + conData(1) + ";Password=" + conData(2) + ";MultipleActiveResultSets=True"
            FC_Con.Open()
        End If
        FC_Conexion = 0
        Exit Function
ERR_CON:
        MsgBox(Err.Description)
        FC_Conexion = Err.Number
    End Function

    Public Function FC_ConexionGUID(ByVal BaseGuid As String) As Long
        Dim conData() As String
        On Error GoTo ERR_CON

        If FC_ConGuid.State = ConnectionState.Open Then FC_ConGuid.Close()
        conData = FC_GetDatos()
        FC_ConGuid = New SqlConnection()
        FC_ConGuid.ConnectionString = "Data Source=" + conData(0) + " ;" &
                         "Initial Catalog=" + BaseGuid + ";" &
                         "User Id=" + conData(1) + ";Password=" + conData(2) + ";MultipleActiveResultSets=True"
        FC_ConGuid.Open()
        FC_ConexionGUID = 0
        Exit Function
ERR_CON:
        MsgBox(Err.Description)
        FC_ConexionGUID = Err.Number
    End Function

    Public Function GetDate(ByVal dateString As String) As String

        Try
            Dim anio As String = dateString.Substring(0, 4)
            Dim mes As String = dateString.Substring(4, 2)
            Dim dia As String = dateString.Substring(6, 2)

            Dim fecha As String = (dia + "/" + mes + "/" + anio)
            Return fecha

        Catch ex As Exception
            Return dateString
        End Try
    End Function

    Public Function GetTime(ByVal dateString As String) As String

        Try
            Dim hora As String = dateString.Substring(8, 2)
            Dim min As String = dateString.Substring(10, 2)
            Dim seg As String = dateString.Substring(12, 2)
            Dim fecha As String = (hora + ":" + min + ":" + seg)

            Return fecha

        Catch ex As Exception
            Return dateString
        End Try
    End Function

    Public Function UnixToDateTime(ByVal strUnixTime As String) As DateTime

        Dim nTimestamp As Double = strUnixTime
        Dim nDateTime As System.DateTime = New System.DateTime(1970, 1, 1, 0, 0, 0, 0)
        nDateTime = nDateTime.AddSeconds(nTimestamp)

        Return nDateTime

    End Function

    Public Sub InsertSQLGen(Cadena As String)
        Dim comando As SqlCommand
        FC_Conexion()
        Try
            comando = New SqlCommand(Cadena, FC_Con)
            comando.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function FC_GetCons() As Long
        Dim cRs As SqlDataReader
        Dim comCrs As SqlCommand
        Dim Exists As Boolean
        Dim prevKey As String

        On Error GoTo ERR_GETCONS
        FC_Conexion()
        comCrs = New SqlCommand("SELECT * FROM Instancias", FC_Con)
        cRs = comCrs.ExecuteReader()
        If cRs.HasRows = False Then MsgBox("No hay instancias registradas.", vbExclamation, "Error de instancias FC") : FC_GetCons = 1 : Exit Function

        DConexiones = New Dictionary(Of String, SqlConnection)

        Do While cRs.Read()
            Exists = False
            If Not Exists Then
                DConexiones.Add(CStr(cRs("nombre")), New SqlConnection)
                DConexiones(CStr(cRs("nombre"))).ConnectionString = "Data Source=" + cRs("server") + " ;" &
                         "User Id=" + cRs("uid") + ";Password=" + cRs("pwd") + ";MultipleActiveResultSets=True"
                DConexiones(CStr(cRs("nombre"))).Open()
            End If
        Loop
        FC_GetCons = 0
        cRs.Close()

        Exit Function
ERR_GETCONS:
        FC_GetCons = Err.Number
    End Function

    Public Function GetDatoInt(ByVal gQuery As String, gCampo As String, ByVal gCon As SqlConnection) As Integer
        Dim RsGet As SqlDataReader

        GetDatoInt = 0
        Using ComRsGet As New SqlCommand(gQuery, gCon)
            RsGet = ComRsGet.ExecuteReader()
            RsGet.Read()
            If RsGet.HasRows = True Then
                If RsGet(gCampo) IsNot DBNull.Value Then GetDatoInt = RsGet(gCampo)
            End If
            RsGet.Close()
        End Using
    End Function

    Public Function GetDatoFecha(ByVal gQuery As String, gCampo As String, ByVal gCon As SqlConnection) As Date
        Dim RsGet As SqlDataReader

        GetDatoFecha = Date.Now.Date
        Using ComRsGet As New SqlCommand(gQuery, gCon)
            RsGet = ComRsGet.ExecuteReader()
            RsGet.Read()
            If RsGet.HasRows = True Then
                If RsGet(gCampo) IsNot DBNull.Value Then GetDatoFecha = RsGet(gCampo)
            End If
            RsGet.Close()
        End Using
    End Function

    Public Function GetDatoUUID(ByVal gGuid As Guid, ByVal gCon As SqlConnection) As String
        Dim RsGet As SqlDataReader, gQuery As String

        GetDatoUUID = ""
        gQuery = "SELECT UUID FROM Comprobante WHERE GuidDocument=@guid"
        Using ComRsGet As New SqlCommand(gQuery, gCon)
            ComRsGet.Parameters.AddWithValue("@guid", gGuid)
            RsGet = ComRsGet.ExecuteReader()
            RsGet.Read()
            If RsGet.HasRows = True Then
                If RsGet("UUID") IsNot DBNull.Value Then GetDatoUUID = RsGet("UUID").ToString
            End If
            RsGet.Close()
        End Using
    End Function

    Public Function GetRutaFile(ByVal gRuta As String, ByVal gNombre As String) As String
        Dim files() As String = System.IO.Directory.GetFiles(
            gRuta, gNombre, IO.SearchOption.AllDirectories)
        GetRutaFile = ""
        For Each file In files
            GetRutaFile = file
            Exit For
        Next file
    End Function
End Module
