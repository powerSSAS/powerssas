Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Gosbell.PowerSSAS.PowerSSASProvider
Imports System.Xml
Imports System.Globalization
Imports Gosbell.PowerSSAS.Types

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASSession")> _
    Public Class CmdletGetASSession
        Inherits CmdletDiscoverBase

        'Protected Overrides Sub ProcessRecord()
        'Dim xSess As New Utils.XmlaDiscoverSessions
        '    xSess.Discover(Me.DiscoverRowset, ServerName, XmlaRestrictions, "", AddressOf Me.OutputObject)
        'End Sub


        Public Overrides ReadOnly Property DiscoverRowset() As String
            Get
                Return "DISCOVER_SESSIONS"
            End Get
        End Property

        'Protected Overrides Sub OutputObject(ByVal output As Object)
        '    ' gets a collection of datarows
        '    '//Dim xsess As New Utils.XmlaDiscoverSessions
        '    '//Me.WriteObject(New Session(Me.ServerName, output))
        '    '//Debug.WriteLine("OutputObject")
        '    For Each dr As DataRow In CType(output, DataTable).Rows
        '        Me.WriteObject(dr)
        '    Next
        'End Sub
    End Class
End Namespace