Imports System.Management.Automation

Namespace Cmdlets
    <Cmdlet("Invoke", "AsDiscover", SupportsShouldProcess:=True)> _
    Public Class CmdletSendXmlaDiscover
        Inherits Cmdlet

        Private xd As New Utils.XmlaDiscover()

        Private mRaw As SwitchParameter
        <Parameter(HelpMessage:="Outputs the results as a raw string")> _
        Public Property Raw() As SwitchParameter
            Get
                Return mRaw
            End Get
            Set(ByVal value As SwitchParameter)
                mRaw = value
            End Set
        End Property

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

        <Parameter(HelpMessage:="Name of the Rowset to query", Mandatory:=True, Position:=1, ParameterSetName:="Default")> _
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
            Try
                If mRaw.IsPresent Then
                    xd.Discover(RowsetName, ServerName, Restrictions, Properties, AddressOf OutputObject, True)
                Else
                    xd.Discover(RowsetName, ServerName, Restrictions, Properties, AddressOf OutputObject)
                End If
            Catch ex As PipelineStoppedException
                'WriteError(New ErrorRecord(New OperationCanceledException("The operation has been cancelled"), "Cancelled", ErrorCategory.OperationStopped, Me))
            End Try

        End Sub

        Protected Sub OutputObject(ByVal output As Object)
            If Me.Stopping Then
                xd.Stop()
            End If
            WriteObject(output)
        End Sub

    End Class
End Namespace
