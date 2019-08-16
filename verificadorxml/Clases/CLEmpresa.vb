Public Class CLEmpresa
    Private _cNomEmpresa As String
    Private _cIDEmpresa As Integer
    Private _cRFCEmpresa As String
    Private _cGuidDSL As String
    Private _cDireccion As String
    Private _cCodigoPostal As String
    Private _cRegCamara As String
    Private _cRegEstatal As String

    Public Property CIDEmpresa As Integer
        Get
            Return _cIDEmpresa
        End Get
        Set(value As Integer)
            _cIDEmpresa = value
        End Set
    End Property

    Public Property CRFCEmpresa As String
        Get
            Return _cRFCEmpresa
        End Get
        Set(value As String)
            _cRFCEmpresa = value
        End Set
    End Property

    Public Property CGuidDSL As String
        Get
            Return _cGuidDSL
        End Get
        Set(value As String)
            _cGuidDSL = value
        End Set
    End Property

    Public Property CDireccion As String
        Get
            Return _cDireccion
        End Get
        Set(value As String)
            _cDireccion = value
        End Set
    End Property

    Public Property CCodigoPostal As String
        Get
            Return _cCodigoPostal
        End Get
        Set(value As String)
            _cCodigoPostal = value
        End Set
    End Property

    Public Property CRegCamara As String
        Get
            Return _cRegCamara
        End Get
        Set(value As String)
            _cRegCamara = value
        End Set
    End Property

    Public Property CRegEstatal As String
        Get
            Return _cRegEstatal
        End Get
        Set(value As String)
            _cRegEstatal = value
        End Set
    End Property

    Public Property CNomEmpresa As String
        Get
            Return _cNomEmpresa
        End Get
        Set(value As String)
            _cNomEmpresa = value
        End Set
    End Property
End Class
