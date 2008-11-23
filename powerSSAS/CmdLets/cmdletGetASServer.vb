Imports System.Management.Automation
Imports System.Xml
Imports Microsoft.AnalysisServices.Xmla
Imports Microsoft.AnalysisServices

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASServer")> _
    Public Class CmdletGetASServer
        Inherits Cmdlet

        Private mServerName As String = ""

        <Parameter(HelpMessage:="Analysis Services server name", Position:=0, ValueFromPipeline:=True, Mandatory:=True, ParameterSetName:="Default")> _
        Public Property ServerName() As String
            Get
                Return mServerName
            End Get
            Set(ByVal value As String)
                mServerName = value
            End Set
        End Property

        Protected Overrides Sub ProcessRecord()
            'svr.Connect(mServerName)
            WriteObject(ConnectionFactory.ConnectToServer(mServerName))
        End Sub

        'Protected Overrides Sub EndProcessing()
        '    MyBase.EndProcessing()
        '    'svr.Disconnect()
        'End Sub

        'Protected Overrides Sub StopProcessing()
        '    MyBase.StopProcessing()
        '    If svr.Connected Then
        '        svr.Disconnect()
        '    End If
        'End Sub

    End Class
End Namespace
