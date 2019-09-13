Public Class clRegistroBitacora
    Private _periodo As Integer
    Private _ejercicio As Integer
    Private _archivo As String
    Private _nombrearchivo As String



    Public Property Periodo As Integer
        Get
            Return _periodo
        End Get
        Set(value As Integer)
            _periodo = value
        End Set
    End Property

    Public Property Ejercicio As Integer
        Get
            Return _ejercicio
        End Get
        Set(value As Integer)
            _ejercicio = value
        End Set
    End Property

    Public Property Archivo As String
        Get
            Return _archivo
        End Get
        Set(value As String)
            _archivo = value
        End Set
    End Property

    Public Property Nombrearchivo As String
        Get
            Return _nombrearchivo
        End Get
        Set(value As String)
            _nombrearchivo = value
        End Set
    End Property





End Class
