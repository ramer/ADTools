
Public Class clsEmailAddress
    Dim _address As String
    Dim _addressfull As String
    Dim _isprimary As Boolean

    Sub New(address As String, addressfull As String, isprimary As Boolean)
        _address = address
        _addressfull = addressfull
        _isprimary = isprimary
    End Sub

    Public ReadOnly Property Address() As String
        Get
            Return _address
        End Get
    End Property

    Public ReadOnly Property AddressFull() As String
        Get
            Return _addressfull
        End Get
    End Property

    Public ReadOnly Property IsPrimary() As Boolean
        Get
            Return _isprimary
        End Get
    End Property
End Class