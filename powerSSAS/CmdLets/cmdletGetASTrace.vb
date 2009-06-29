Imports System.Management.Automation
Imports Microsoft.AnalysisServices.AdomdClient
Imports Gosbell.PowerSSAS.Cmdlets

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASTrace")> _
    Public Class CmdletGetASTrace
        Inherits CmdletDiscoverBase

        Public Overrides ReadOnly Property DiscoverRowset() As String
            Get
                Return "DISCOVER_TRACES"
            End Get
        End Property
    End Class
End Namespace