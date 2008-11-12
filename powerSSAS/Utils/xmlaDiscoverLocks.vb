Imports Gosbell.PowerSSAS.Types

Namespace Utils
    Public Class XmlaDiscoverLocks
        Inherits XmlaDiscover

        Protected Overrides Function GetPSObject(ByVal rowSchema As System.Collections.Generic.Dictionary(Of String, String), ByVal server As String, ByVal row As System.Xml.XmlNode) As Object
            Return New Lock(server, row)
        End Function

    End Class
End Namespace
