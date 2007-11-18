Imports System.Management.Automation

Namespace Cmdlets
    <Cmdlet(VerbsCommunications.Send, "XmlaDiscover", SupportsShouldProcess:=True)> _
    Public Class cmdletSendXmlaDiscover
        Inherits PSCmdlet

        Private mServerName As String = ""
        <Parameter(HelpMessage:="Analysis Services server name", Mandatory:=True, ParameterSetName:="Default", Position:=0, ValueFromPipeline:=True)> _
        Public Property ServerName() As String
            Get
                Return mServerName
            End Get
            Set(ByVal value As String)
                mServerName = value
            End Set
        End Property

        Private mRowSetName As String = ""

        <Parameter(Mandatory:=True, Position:=1, ParameterSetName:="Default")> _
        Public Property RowsetName() As String
            Get
                Return mRowSetName
            End Get
            Set(ByVal value As String)
                mRowSetName = value
            End Set
        End Property

        Private mRestrictions As String = ""
        <Parameter(Mandatory:=False)> _
        Public Property Restrictions() As String
            Get
                Return mRestrictions
            End Get
            Set(ByVal value As String)
                mRestrictions = value
            End Set
        End Property

        Private mProperties As String = ""
        <Parameter(Mandatory:=False)> _
        Public Property Properties() As String
            Get
                Return mProperties
            End Get
            Set(ByVal value As String)
                mProperties = value
            End Set
        End Property

        Protected Overrides Sub ProcessRecord()
            Dim xd As New Utils.xmlaDiscover()
            WriteObject(xd.Discover(RowsetName, ServerName, Restrictions, Properties))
        End Sub
    End Class
End Namespace
