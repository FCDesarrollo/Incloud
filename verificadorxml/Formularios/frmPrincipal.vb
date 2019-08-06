Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Public Class frmprincipal
    ' Public m_connect As String = "Data Source=10.10.10.15\COMPAC;Initial Catalog=FCModulos;User Id=sa;Password=compac$1;MultipleActiveResultSets=True"
    Private con As SqlConnection = Nothing
    Public Delegate Sub NewHome()
    Public Event OnNewHome As NewHome
    Public Event OnNewHome2 As NewHome

    Private sload As Boolean = True
    Private DSQLDepende As Dictionary(Of String, SqlDependency)
    Private DSQLDepende2 As Dictionary(Of String, SqlDependency)
    Public Sub CargaConexiones()

        ' Esta llamada es exigida por el diseñador.
        'InitializeComponent()
        Try
            Dim ss As SqlClientPermission
            ss = New SqlClientPermission(System.Security.Permissions.PermissionState.Unrestricted)
            ss.Demand()

        Catch ex As Exception
            Throw
        End Try

        ferror = FC_Conexion()
        If ferror <> 0 Then Exit Sub
        If IsNothing(DConexiones) Then FC_GetCons()
        DConexiones("CON").ChangeDatabase("master")
        Dim m_connect As String, cQue As String, Quer As String
        Dim baseXml As String, baseConten As String
        Dim conData() As String
        conData = FC_GetDatos()
        DConexionesXML = New Dictionary(Of String, SqlConnection)
        DConexionesCFDI = New Dictionary(Of String, SqlConnection)
        DConexionesConten = New Dictionary(Of String, SqlConnection)

        cQue = "SELECT  id,NomEmpresa,BDDCon FROM EEFEmpresas WHERE FechaAutomatic IS NOT NULL"
        Using eCom = New SqlCommand(cQue, FC_Con)
            Using rsTras = eCom.ExecuteReader()
                Do While rsTras.Read()
                    DConexiones("CON").ChangeDatabase(rsTras("BDDCon"))

                    If BaseConSeguimiento(rsTras("BDDCon")) = False Then
                        cQue = "ALTER DATABASE " & rsTras("BDDCon") &
                           " SET CHANGE_TRACKING = ON
                            (CHANGE_RETENTION = 7 DAYS, AUTO_CLEANUP = ON)"
                        Using cCom = New SqlCommand(cQue, DConexiones("CON"))
                            cCom.ExecuteNonQuery()
                        End Using
                    End If

                    'DConexiones("CON").ChangeDatabase(rsTras("BDDCon"))
                    If TablaConSeguimiento("AsocCFDIs", DConexiones("CON")) = False Then
                        cQue = "ALTER TABLE AsocCFDIs
                            ENABLE CHANGE_TRACKING
                            WITH (TRACK_COLUMNS_UPDATED = ON)"
                        Using cCom = New SqlCommand(cQue, DConexiones("CON"))
                            cCom.ExecuteNonQuery()
                        End Using
                    End If
                    m_connect = "Data Source=" & conData(0) & ";Initial Catalog=" & rsTras("BDDCon") & ";User Id=" & conData(1) & ";Password=" & conData(2) & ";MultipleActiveResultSets=True"
                    DConexionesCFDI.Add(rsTras("NomEmpresa"), New SqlConnection)
                    DConexionesCFDI(rsTras("NomEmpresa")).ConnectionString = m_connect
                    DConexionesCFDI(rsTras("NomEmpresa")).Open()
                    SqlDependency.Stop(m_connect)
                    SqlDependency.Start(m_connect)

                    Quer = "SELECT GuidDSL FROM Parametros"
                    Using dCom = New SqlCommand(Quer, DConexionesCFDI(rsTras("NomEmpresa")))
                        Using dRs = dCom.ExecuteReader
                            dRs.Read()
                            If dRs.HasRows Then
                                baseXml = "document_" & dRs("GuidDSL") & "_metadata"
                                m_connect = "Data Source=" & conData(0) & ";Initial Catalog=" & baseXml & ";User Id=" & conData(1) & ";Password=" & conData(2) & ";MultipleActiveResultSets=True"
                                DConexionesXML.Add(rsTras("NomEmpresa"), New SqlConnection)
                                DConexionesXML(rsTras("NomEmpresa")).ConnectionString = m_connect
                                DConexionesXML(rsTras("NomEmpresa")).Open()
                                SqlDependency.Stop(m_connect)
                                SqlDependency.Start(m_connect)

                                If BaseConSeguimiento(baseXml) = False Then
                                    cQue = "ALTER DATABASE [" & baseXml &
                                        "] SET CHANGE_TRACKING = ON
                                        (CHANGE_RETENTION = 7 DAYS, AUTO_CLEANUP = ON)"
                                    Using cCom = New SqlCommand(cQue, DConexiones("CON"))
                                        cCom.ExecuteNonQuery()
                                    End Using
                                End If

                                If TablaConSeguimiento("Comprobante", DConexionesXML(rsTras("NomEmpresa"))) = False Then
                                    cQue = "ALTER TABLE Comprobante
                                        ENABLE CHANGE_TRACKING
                                        WITH (TRACK_COLUMNS_UPDATED = ON)"
                                    Using cCom = New SqlCommand(cQue, DConexionesXML(rsTras("NomEmpresa")))
                                        cCom.ExecuteNonQuery()
                                    End Using
                                End If

                                baseConten = "document_" & dRs("GuidDSL") & "_content"
                                    m_connect = "Data Source=" & conData(0) & ";Initial Catalog=" & baseConten & ";User Id=" & conData(1) & ";Password=" & conData(2) & ";MultipleActiveResultSets=True"
                                    DConexionesConten.Add(rsTras("NomEmpresa"), New SqlConnection)
                                    DConexionesConten(rsTras("NomEmpresa")).ConnectionString = m_connect
                                    DConexionesConten(rsTras("NomEmpresa")).Open()
                                    SqlDependency.Stop(m_connect)
                                    SqlDependency.Start(m_connect)
                                End If
                        End Using
                    End Using
                Loop
            End Using
        End Using
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        letr = New NumaLet
        letr.ApocoparUnoParteEntera = True
        CargaConexiones()
        AddHandler OnNewHome, New NewHome(AddressOf Form1_OnNewHome)
        AddHandler OnNewHome2, New NewHome(AddressOf Form1_OnNewHome2)
        sload = True
        LoadData()
        LoadCFDI()
        sload = False
    End Sub

    Public Sub Form1_OnNewHome()
        Dim i As ISynchronizeInvoke = CType(Me, ISynchronizeInvoke)

        If i.InvokeRequired Then
            Dim dd As NewHome = New NewHome(AddressOf Form1_OnNewHome)
            i.BeginInvoke(dd, Nothing)
            Return
        End If
        LoadData()
    End Sub

    Public Sub Form1_OnNewHome2()
        Dim i As ISynchronizeInvoke = CType(Me, ISynchronizeInvoke)

        If i.InvokeRequired Then
            Dim dd As NewHome = New NewHome(AddressOf Form1_OnNewHome2)
            i.BeginInvoke(dd, Nothing)
            Return
        End If
        LoadCFDI()
    End Sub

    Public Sub LoadData()
        Dim dtshow As DataTable = New DataTable()
        Dim dtshow2 As DataTable = New DataTable()
        Dim dt As DataTable = New DataTable()
        Dim DDtable As Dictionary(Of String, DataTable)
        Dim t As Integer, nomCon As String, sQuery As String

        Dim DSQLCommand As Dictionary(Of String, SqlCommand)
        Dim DSQLCommandNoti As Dictionary(Of String, SqlCommand)


        DSQLCommand = New Dictionary(Of String, SqlCommand)
        DSQLCommandNoti = New Dictionary(Of String, SqlCommand)
        DSQLDepende = New Dictionary(Of String, SqlDependency)

        DDtable = New Dictionary(Of String, DataTable)

        For t = 0 To DConexionesXML.Count - 1
            nomCon = DConexionesXML.Keys(t)

            If DConexionesXML(nomCon).State = ConnectionState.Closed Then
                DConexionesXML(nomCon).Open()
            End If

            ' DSQLCommand(nomCon) = New SqlCommand("SELECT  ID, NombreEmpresa, RutaCONTPAQ, Indexado FROM PCHEmpresas", DConexiones(nomCon))

            sQuery = "SELECT UUID FROM dbo.Comprobante"
            DSQLCommandNoti(nomCon) = New SqlCommand(sQuery, DConexionesXML(nomCon))

            DSQLCommandNoti(nomCon).Notification = Nothing

            DSQLDepende(nomCon) = New SqlDependency(DSQLCommandNoti(nomCon))

            DDtable(nomCon) = New DataTable


            DDtable(nomCon).Load(DSQLCommandNoti(nomCon).ExecuteReader(CommandBehavior.CloseConnection))


            AddHandler DSQLDepende(nomCon).OnChange, AddressOf de_OnChange

        Next


        For t = 0 To DConexionesXML.Count - 1
            nomCon = DConexionesXML.Keys(t)

            If DConexionesXML(nomCon).State = ConnectionState.Closed Then
                DConexionesXML(nomCon).Open()
            End If
        Next

        If sload = False Then
            cgXML = New Collection
            CreaXML("", True)
        End If
    End Sub

    Public Sub LoadCFDI()
        Dim dtshow As DataTable = New DataTable()
        Dim dtshow2 As DataTable = New DataTable()
        Dim dt As DataTable = New DataTable()
        Dim DDtable As Dictionary(Of String, DataTable)
        Dim t As Integer, nomCon As String, sQuery As String

        Dim DSQLCommand As Dictionary(Of String, SqlCommand)
        Dim DSQLCommandNoti As Dictionary(Of String, SqlCommand)


        DSQLCommand = New Dictionary(Of String, SqlCommand)
        DSQLCommandNoti = New Dictionary(Of String, SqlCommand)
        DSQLDepende2 = New Dictionary(Of String, SqlDependency)

        DDtable = New Dictionary(Of String, DataTable)

        For t = 0 To DConexionesCFDI.Count - 1
            nomCon = DConexionesCFDI.Keys(t)

            If DConexionesCFDI(nomCon).State = ConnectionState.Closed Then
                DConexionesCFDI(nomCon).Open()
            End If

            sQuery = "SELECT UUID FROM dbo.AsocCFDIs"
            DSQLCommandNoti(nomCon) = New SqlCommand(sQuery, DConexionesCFDI(nomCon))

            DSQLCommandNoti(nomCon).Notification = Nothing

            DSQLDepende2(nomCon) = New SqlDependency(DSQLCommandNoti(nomCon))

            DDtable(nomCon) = New DataTable


            DDtable(nomCon).Load(DSQLCommandNoti(nomCon).ExecuteReader(CommandBehavior.CloseConnection))


            AddHandler DSQLDepende2(nomCon).OnChange, AddressOf de_OnChangeCFDi

        Next


        For t = 0 To DConexionesCFDI.Count - 1
            nomCon = DConexionesCFDI.Keys(t)

            If DConexionesCFDI(nomCon).State = ConnectionState.Closed Then
                DConexionesCFDI(nomCon).Open()
            End If
        Next

        If sload = False Then
            CreaExpediente("", True)
        End If
    End Sub

    Public Sub de_OnChange(sender As Object, e As SqlNotificationEventArgs)
        Dim nomCon As String
        Dim dependency As SqlDependency = CType(sender, SqlDependency)

        For t = 0 To DConexionesXML.Count - 1
            nomCon = DConexionesXML.Keys(t)
            RemoveHandler DSQLDepende(nomCon).OnChange, AddressOf de_OnChange
        Next

        If dependency IsNot Nothing AndAlso dependency.HasChanges Then
            RaiseEvent OnNewHome()
        End If
    End Sub

    Public Sub de_OnChangeCFDi(sender As Object, e As SqlNotificationEventArgs)
        Dim nomCon As String
        Dim dependency As SqlDependency = CType(sender, SqlDependency)

        For t = 0 To DConexionesCFDI.Count - 1
            nomCon = DConexionesCFDI.Keys(t)
            RemoveHandler DSQLDepende2(nomCon).OnChange, AddressOf de_OnChangeCFDi
        Next

        If dependency IsNot Nothing AndAlso dependency.HasChanges Then
            RaiseEvent OnNewHome2()
        End If
    End Sub




    Private Sub frmprincipal_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If FormWindowState.Minimized = Me.WindowState Then
            'Me.WindowState = FormWindowState.Minimized
            Me.Hide()
            Me.Refresh()
            niClose.Visible = True
            niClose.ShowBalloonTip(1000, "Incloud", "En ejecución", ToolTipIcon.Info)
        End If
    End Sub

    Private Sub NiClose_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles niClose.MouseDoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        niClose.Visible = False
    End Sub

    Private Sub BtnSincro_Click(sender As Object, e As EventArgs)
        'Me.btnSincro.Enabled = False
        'Me.Hide()
        'Me.Refresh()
        'niClose.Visible = True
        'niClose.ShowBalloonTip(1000, "App", "XMl procesando", ToolTipIcon.Info)
        'esBoton = True
        'LoadData()
        'LoadCFDI()
        'esBoton = False
        'niClose.ShowBalloonTip(1000, "App", "XML, Fin Proceso Manual", ToolTipIcon.Info)
        'Me.btnSincro.Enabled = True
    End Sub

    Private Sub frmprincipal_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        KillAllExcels()
    End Sub

    Private Sub MConfig_Click(sender As Object, e As EventArgs) Handles MConfig.Click
        Dim frm As New frmConfig
        frm.ShowDialog()
        CargaConexiones()
        AddHandler OnNewHome, New NewHome(AddressOf Form1_OnNewHome)
        AddHandler OnNewHome2, New NewHome(AddressOf Form1_OnNewHome2)
        sload = True
        LoadData()
        LoadCFDI()
        sload = False
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim frm As New frmProcesar
        frm.ShowDialog()
    End Sub
End Class
