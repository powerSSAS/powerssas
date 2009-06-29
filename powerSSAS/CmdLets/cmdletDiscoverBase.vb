Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Gosbell.PowerSSAS.PowerSSASProvider
Imports System.Xml
Imports System.Globalization

Namespace Cmdlets

    Public MustInherit Class CmdletDiscoverBase
        Inherits Cmdlet

        Private mMajorObjectPath As String()
        Private mServerName As String = ""

        Protected Sub New()
        End Sub

        Public MustOverride ReadOnly Property DiscoverRowset() As String

        Private mXmlaRestrictions As New AdomdClient.AdomdRestrictionCollection
        <Parameter(HelpMessage:="Xmla Restrictions", Mandatory:=False, ParameterSetName:="byServer", Position:=1, ValueFromPipeline:=False)> _
        Public Overridable Property XmlaRestrictions() As AdomdClient.AdomdRestrictionCollection
            Get
                Return mXmlaRestrictions
            End Get
            Set(ByVal value As AdomdClient.AdomdRestrictionCollection)
                mXmlaRestrictions = value
            End Set
        End Property

        Private mDatabaseName As String = ""
        <Parameter(HelpMessage:="Xmla Properties", Mandatory:=False, ParameterSetName:="byServer", Position:=2, ValueFromPipeline:=False)> _
        Protected Overridable Property ConnectionDatabaseName() As String
            Get
                Return mDatabaseName
            End Get
            Set(ByVal value As String)
                mDatabaseName = value
            End Set
        End Property

        <Parameter(HelpMessage:="Analysis Services server name", ParameterSetName:="byServer", Position:=0, ValueFromPipeline:=False)> _
        Public Property ServerName() As String
            Get
                If mServerName.Length = 0 Then
                    WriteWarning("No server specified, using localhost")
                    Return "localhost"
                End If
                Return mServerName
            End Get
            Set(ByVal value As String)
                mServerName = value
            End Set
        End Property

        Private mServer As Microsoft.AnalysisServices.Server
        <[Alias]("Server")> _
        <Parameter(HelpMessage:="Analysis Services server object", ParameterSetName:="byObject", Position:=0, ValueFromPipeline:=True)> _
        Public Property ASServer() As Microsoft.AnalysisServices.Server
            Get
                Return mServer
            End Get
            Set(ByVal value As Microsoft.AnalysisServices.Server)
                mServer = value
                If Not value Is Nothing Then
                    mServerName = mServer.Name
                End If
            End Set
        End Property

        Protected Overrides Sub BeginProcessing()
            MyBase.BeginProcessing()

            'If Me.ServerName.Length = 0 Then
            '    Dim psDrive As PSDriveInfo = Me.SessionState.Path.CurrentLocation.Drive

            '    If TypeOf psDrive Is AmoDriveInfo Then
            '        Dim p As PSObject
            '        Dim itmCol As System.Collections.ObjectModel.Collection(Of PSObject)
            '        itmCol = Me.InvokeProvider.Item.Get(Me.SessionState.Path.CurrentLocation.Path)

            '        If itmCol.Count = 1 Then
            '            p = itmCol.Item(0)
            '            If TypeOf p.BaseObject Is Microsoft.AnalysisServices.MajorObject Then
            '                Dim majObj As MajorObject
            '                majObj = CType(p.BaseObject, MajorObject)
            '                Dim amoDrive As AmoDriveInfo
            '                amoDrive = CType(psDrive, AmoDriveInfo)
            '                Me.ServerName = amoDrive.AmoServer.Name
            '            End If
            '        Else
            '            Me.WriteWarning(String.Format("There were {0} objects found at {1}.", itmCol.Count, Me.SessionState.Path.CurrentLocation.Path))
            '            Me.StopProcessing()
            '        End If

            '    Else
            '        Me.WriteError(New ErrorRecord(New ArgumentException("You must specify a server name or run this command from a powerSSAS drive"), "Invalid Argument", ErrorCategory.InvalidArgument, Me))
            '        Me.StopProcessing()
            '    End If
            'End If

        End Sub

        Protected Overrides Sub ProcessRecord()
            Dim xd As New Utils.XmlaDiscover()

            xd.Discover(DiscoverRowset, ServerName, XmlaRestrictions, ConnectionDatabaseName, AddressOf OutputObject)
        End Sub

        Protected Overridable Sub OutputObject(ByVal output As Object)
            Me.WriteObject(output)
        End Sub
    End Class
End Namespace