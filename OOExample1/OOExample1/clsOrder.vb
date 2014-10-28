Public Class Order
    Private OrderCode As String
    Private CreateDate As DateTime
    Private CustomerId As Integer
    Private ProductId As Integer
    Private PriceExVat As Double
    Private dispatched As Boolean
    Private PaymentRecieved As Boolean

    Public Sub New(ByRef pCustomerId As Integer, ByRef pProductId As Integer)
        ProductId = pProductId
        CustomerId = pCustomerId
        CreateDate = Now()
        dispatched = False
        PaymentRecieved = False

    End Sub

    Public Sub New(ByRef OrderCode As String)

    End Sub

    Public ReadOnly Property IsDispatched As Boolean
        Get
            IsDispatched = dispatched
        End Get
    End Property

    Public Sub OrderDispatched()
        dispatched = True
    End Sub
End Class
