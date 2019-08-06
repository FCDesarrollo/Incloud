Public Class CLXml
    Private _sMoneda As String
    Private _sFormaPago As String
    Private _sTipoCambio As String
    Private _sMetodoPago As String
    Private _sCuenta As String
    Private _sFecha As String
    Private _sVersion As String
    Private _sRFCEmisor As String
    Private _sNombreEmisor As String
    Private _sNombreReceptor As String
    Private _sSerie As String
    Private _sFolio As String
    Private _sTipo As String
    Private _sRegimenFiscalE As String
    Private _sUsoCFDI As String
    Private _sNoCertificado As String
    Private _sDescto As Double
    Private _sSubtotal As Double
    Private _sTotalXML As Double

    Private _sTotalIva As Double
    Private _sTotalIeps As Double
    Private _sTotalRetIva As Double
    Private _sTotalRetIsr As Double


    Private _sLugarExpedicion As String
    Private _sRFCReceptor As String
    Private _sDomicilioReceptor As String

    Private _sUUID As Guid
    Private _sFechaTimbrado As String
    Private _sCerSAT As String
    Private _sSelloDig As String
    Private _sSelloSAT As String
    Private _sVersionSello As String

    Private _sCodigoQr As String
    Private _movXml As New Collection

    Public Property SFormaPago As String
        Get
            Return _sFormaPago
        End Get
        Set(value As String)
            _sFormaPago = value
        End Set
    End Property

    Public Property STipoCambio As String
        Get
            Return _sTipoCambio
        End Get
        Set(value As String)
            _sTipoCambio = value
        End Set
    End Property

    Public Property SMetodoPago As String
        Get
            Return _sMetodoPago
        End Get
        Set(value As String)
            _sMetodoPago = value
        End Set
    End Property

    Public Property SCuenta As String
        Get
            Return _sCuenta
        End Get
        Set(value As String)
            _sCuenta = value
        End Set
    End Property

    Public Property SFecha As String
        Get
            Return _sFecha
        End Get
        Set(value As String)
            _sFecha = value
        End Set
    End Property

    Public Property SVersion As String
        Get
            Return _sVersion
        End Get
        Set(value As String)
            _sVersion = value
        End Set
    End Property

    Public Property SRFCEmisor As String
        Get
            Return _sRFCEmisor
        End Get
        Set(value As String)
            _sRFCEmisor = value
        End Set
    End Property

    Public Property SNombreEmisor As String
        Get
            Return _sNombreEmisor
        End Get
        Set(value As String)
            _sNombreEmisor = value
        End Set
    End Property

    Public Property SNombreReceptor As String
        Get
            Return _sNombreReceptor
        End Get
        Set(value As String)
            _sNombreReceptor = value
        End Set
    End Property

    Public Property SSerie As String
        Get
            Return _sSerie
        End Get
        Set(value As String)
            _sSerie = value
        End Set
    End Property

    Public Property SFolio As String
        Get
            Return _sFolio
        End Get
        Set(value As String)
            _sFolio = value
        End Set
    End Property

    Public Property STipo As String
        Get
            Return _sTipo
        End Get
        Set(value As String)
            _sTipo = value
        End Set
    End Property

    Public Property SRegimenFiscalE As String
        Get
            Return _sRegimenFiscalE
        End Get
        Set(value As String)
            _sRegimenFiscalE = value
        End Set
    End Property

    Public Property SUsoCFDI As String
        Get
            Return _sUsoCFDI
        End Get
        Set(value As String)
            _sUsoCFDI = value
        End Set
    End Property

    Public Property SNoCertificado As String
        Get
            Return _sNoCertificado
        End Get
        Set(value As String)
            _sNoCertificado = value
        End Set
    End Property

    Public Property SDescto As Double
        Get
            Return _sDescto
        End Get
        Set(value As Double)
            _sDescto = value
        End Set
    End Property

    Public Property SSubtotal As Double
        Get
            Return _sSubtotal
        End Get
        Set(value As Double)
            _sSubtotal = value
        End Set
    End Property

    Public Property STotalXML As Double
        Get
            Return _sTotalXML
        End Get
        Set(value As Double)
            _sTotalXML = value
        End Set
    End Property

    Public Property SLugarExpedicion As String
        Get
            Return _sLugarExpedicion
        End Get
        Set(value As String)
            _sLugarExpedicion = value
        End Set
    End Property

    Public Property SRFCReceptor As String
        Get
            Return _sRFCReceptor
        End Get
        Set(value As String)
            _sRFCReceptor = value
        End Set
    End Property

    Public Property SDomicilioReceptor As String
        Get
            Return _sDomicilioReceptor
        End Get
        Set(value As String)
            _sDomicilioReceptor = value
        End Set
    End Property



    Public Property SFechaTimbrado As String
        Get
            Return _sFechaTimbrado
        End Get
        Set(value As String)
            _sFechaTimbrado = value
        End Set
    End Property

    Public Property SCerSAT As String
        Get
            Return _sCerSAT
        End Get
        Set(value As String)
            _sCerSAT = value
        End Set
    End Property

    Public Property SSelloDig As String
        Get
            Return _sSelloDig
        End Get
        Set(value As String)
            _sSelloDig = value
        End Set
    End Property

    Public Property SSelloSAT As String
        Get
            Return _sSelloSAT
        End Get
        Set(value As String)
            _sSelloSAT = value
        End Set
    End Property

    Public Property SVersionSello As String
        Get
            Return _sVersionSello
        End Get
        Set(value As String)
            _sVersionSello = value
        End Set
    End Property

    Public Property SMoneda As String
        Get
            Return _sMoneda
        End Get
        Set(value As String)
            _sMoneda = value
        End Set
    End Property

    Public Property SUUID As Guid
        Get
            Return _sUUID
        End Get
        Set(value As Guid)
            _sUUID = value
        End Set
    End Property

    Public Property MovXml As Collection
        Get
            Return _movXml
        End Get
        Set(value As Collection)
            _movXml = value
        End Set
    End Property

    Public Property SCodigoQr As String
        Get
            Return _sCodigoQr
        End Get
        Set(value As String)
            _sCodigoQr = value
        End Set
    End Property

    Public Property STotalIva As Double
        Get
            Return _sTotalIva
        End Get
        Set(value As Double)
            _sTotalIva = value
        End Set
    End Property

    Public Property STotalIeps As Double
        Get
            Return _sTotalIeps
        End Get
        Set(value As Double)
            _sTotalIeps = value
        End Set
    End Property

    Public Property STotalRetIva As Double
        Get
            Return _sTotalRetIva
        End Get
        Set(value As Double)
            _sTotalRetIva = value
        End Set
    End Property

    Public Property STotalRetIsr As Double
        Get
            Return _sTotalRetIsr
        End Get
        Set(value As Double)
            _sTotalRetIsr = value
        End Set
    End Property
End Class
