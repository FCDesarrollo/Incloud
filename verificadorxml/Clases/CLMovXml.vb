Public Class CLMovXml
    Private _mImporte As Double
    Private _mValorUni As Double
    Private _mDescrip As String
    Private _mNoIentifi As String
    Private _mUnidad As String
    Private _mCantidad As Double
    Private _mCveProdSer As Double
    Private _mDesc As Double

    Private _mIva As Double
    Private _mIeps As Double


    Public Property MImporte As Double
        Get
            Return _mImporte
        End Get
        Set(value As Double)
            _mImporte = value
        End Set
    End Property

    Public Property MValorUni As Double
        Get
            Return _mValorUni
        End Get
        Set(value As Double)
            _mValorUni = value
        End Set
    End Property



    Public Property MNoIentifi As String
        Get
            Return _mNoIentifi
        End Get
        Set(value As String)
            _mNoIentifi = value
        End Set
    End Property

    Public Property MUnidad As String
        Get
            Return _mUnidad
        End Get
        Set(value As String)
            _mUnidad = value
        End Set
    End Property

    Public Property MCantidad As Double
        Get
            Return _mCantidad
        End Get
        Set(value As Double)
            _mCantidad = value
        End Set
    End Property

    Public Property MCveProdSer As Double
        Get
            Return _mCveProdSer
        End Get
        Set(value As Double)
            _mCveProdSer = value
        End Set
    End Property

    Public Property MDescrip As String
        Get
            Return _mDescrip
        End Get
        Set(value As String)
            _mDescrip = value
        End Set
    End Property



    Public Property MDesc As Double
        Get
            Return _mDesc
        End Get
        Set(value As Double)
            _mDesc = value
        End Set
    End Property

    Public Property MIva As Double
        Get
            Return _mIva
        End Get
        Set(value As Double)
            _mIva = value
        End Set
    End Property

    Public Property MIeps As Double
        Get
            Return _mIeps
        End Get
        Set(value As Double)
            _mIeps = value
        End Set
    End Property

End Class
