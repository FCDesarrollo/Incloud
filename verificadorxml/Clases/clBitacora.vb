Imports System.Data.SqlClient

Public Class clBitacora
    Private _idsubmenu As Integer
    Private _tipodocumento As String

    Private _regbitacora As New Collection
    Private _rfc As String
    Private _idusuarioentrega As Integer
    Private _idusuariosubida As Integer
    Private _status As Integer

    Public Property Idsubmenu As Integer
        Get
            Return _idsubmenu
        End Get
        Set(value As Integer)
            _idsubmenu = value
        End Set
    End Property

    Public Property Tipodocumento As String
        Get
            Return _tipodocumento
        End Get
        Set(value As String)
            _tipodocumento = value
        End Set
    End Property

    Public Property Regbitacora As Collection
        Get
            Return _regbitacora
        End Get
        Set(value As Collection)
            _regbitacora = value
        End Set
    End Property

    Public Property Rfc As String
        Get
            Return _rfc
        End Get
        Set(value As String)
            _rfc = value
        End Set
    End Property

    Public Property Idusuarioentrega As Integer
        Get
            Return _idusuarioentrega
        End Get
        Set(value As Integer)
            _idusuarioentrega = value
        End Set
    End Property

    Public Property Idusuariosubida As Integer
        Get
            Return _idusuariosubida
        End Get
        Set(value As Integer)
            _idusuariosubida = value
        End Set
    End Property

    Public Property Status As Integer
        Get
            Return _status
        End Get
        Set(value As Integer)
            _status = value
        End Set
    End Property

    Public Sub AgregaRegistro()
        Dim bQuery As String
        Dim reg As clRegistroBitacora

        bQuery = "IF NOT EXISTS (SELECT id FROM zIncContBitacora WHERE idsubmenu=@idsub AND " &
                        "tipodocumento=@tipo AND periodo=@periodo AND ejercicio=@ejercicio) " &
                        "BEGIN INSERT INTO zIncContBitacora(idsubmenu, tipodocumento, periodo, " &
                "ejercicio, fecha, fechamodificacion, archivo, nombrearchivo, sincronizado) " &
                "VALUES(@idsub,@tipo,@periodo,@ejercicio,@fecha,@fechamod,@archi,@nomarch,@sincro) END"
        For Each reg In _regbitacora
            Using cCom = New SqlCommand(bQuery, DConexiones("CON"))
                cCom.Parameters.AddWithValue("@idsub", _idsubmenu)
                cCom.Parameters.AddWithValue("@tipo", _tipodocumento)
                cCom.Parameters.AddWithValue("@periodo", reg.Periodo)
                cCom.Parameters.AddWithValue("@ejercicio", reg.Ejercicio)
                cCom.Parameters.AddWithValue("@fecha", Date.Now.Date)
                cCom.Parameters.AddWithValue("@fechamod", Date.Now)
                cCom.Parameters.AddWithValue("@archi", reg.Archivo)
                cCom.Parameters.AddWithValue("@nomarch", reg.Nombrearchivo)
                cCom.Parameters.AddWithValue("@sincro", 0)
                cCom.ExecuteNonQuery()
            End Using
        Next reg

    End Sub
End Class
