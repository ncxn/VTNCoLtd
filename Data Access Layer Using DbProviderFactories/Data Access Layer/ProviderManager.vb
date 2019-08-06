Imports System.Data.Common

Public Class ProviderManager
    Public Property _ProviderName As String

    Public ReadOnly Property Factory As DbProviderFactory
        Get
            Dim _factory As DbProviderFactory = DbProviderFactories.GetFactory(_ProviderName)
            Return _factory
        End Get
    End Property

    Public Sub New()
        _ProviderName = GetProviderName(DefaultConnection)
    End Sub

    Public Sub New(ByVal providerName As String)
        _ProviderName = providerName
    End Sub
End Class
