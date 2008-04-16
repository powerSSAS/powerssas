Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Gosbell.PowerSSAS.PowerSSASProvider
Imports System.Xml
Imports System.Globalization

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASSession", DefaultParameterSetName:="byObject")> _
    Public Class CmdletGetASSession
        Inherits CmdletDiscoverBase

        Protected Overrides Sub ProcessRecord()
            Dim xSess As New Utils.XmlaDiscoverSessions
            xSess.Discover(Me.DiscoverRowset, ServerName, XmlaRestrictions, XmlaProperties, AddressOf Me.OutputObject)
        End Sub


        Public Overrides ReadOnly Property DiscoverRowset() As String
            Get
                Return "DISCOVER_SESSIONS"
            End Get
        End Property
    End Class
End Namespace