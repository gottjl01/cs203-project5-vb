Imports System.Configuration

Namespace PFW.CSIST203.Project5.Config

    Public Class Urls
        Inherits System.Configuration.ConfigurationElementCollection

        Protected Overrides Function CreateNewElement() As ConfigurationElement
            Throw New NotImplementedException()
        End Function

        Protected Overrides Function GetElementKey(element As ConfigurationElement) As Object
            Throw New NotImplementedException()
        End Function

    End Class

End Namespace


