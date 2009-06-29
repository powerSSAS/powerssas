Imports System.Management.Automation
Imports System.Xml
Imports Microsoft.AnalysisServices

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASDatabase", DefaultParameterSetName:="byObject")> _
    Public Class CmdletGetASDatabase
        Inherits Cmdlet

        Private mSvr As Server

        Private mServerName As String = ""
        <Parameter(ParameterSetName:="byName", HelpMessage:="Analysis Services server name", Position:=0, ValueFromPipeline:=False)> _
        Public Property ServerName() As String
            Get
                Return mServerName
            End Get
            Set(ByVal value As String)
                mServerName = value
            End Set
        End Property

        Private mDatabaseName As String = ""
        <[Alias]("Database")> _
        <Parameter(ParameterSetName:="byName", HelpMessage:="Analysis Services database name", Position:=1, ValueFromPipeline:=False)> _
        Public Property DatabaseName() As String
            Get
                Return mDatabaseName
            End Get
            Set(ByVal value As String)
                mDatabaseName = value
            End Set
        End Property

        Private mServer As Microsoft.AnalysisServices.Server
        <[Alias]("Server")> _
        <Parameter(ParameterSetName:="byObject", HelpMessage:="Analysis Services server object", Position:=0, ValueFromPipeline:=True)> _
        Public Property ASServer() As Microsoft.AnalysisServices.Server
            Get
                Return mServer
            End Get
            Set(ByVal value As Microsoft.AnalysisServices.Server)
                mServer = value
            End Set
        End Property


        'Protected Overrides Sub EndProcessing()
        '    '// if the server was connected inside this cmdlet, we should disconnect it.
        '    If mServer Is Nothing AndAlso mServerName.Length > 0 Then
        '        mSvr.Disconnect()
        '    End If
        '    MyBase.EndProcessing()
        'End Sub

        Protected Overrides Sub ProcessRecord()

            If mServer IsNot Nothing Then
                '// if the server object is passed in, then use that
                mSvr = mServer
            ElseIf mServer Is Nothing AndAlso mServerName.Length > 0 Then
                '// if just the name is passed in, then we need to connect
                mSvr = ConnectionFactory.ConnectToServer(mServerName)
            Else
                WriteError(New ErrorRecord(New ArgumentException("You must pass in a server object or server name"), "MissingParam", ErrorCategory.InvalidArgument, Nothing))
            End If

            If mDatabaseName.Length = 0 Then
                '// if no specific database is requested, return the whole collection.
                For Each db As Database In mSvr.Databases
                    WriteObject(db)
                Next db
            Else
                If WildcardPattern.ContainsWildcardCharacters(mDatabaseName) Then
                    '// if we have a wildcard pattern, search for matches
                    Dim wc As New System.Management.Automation.WildcardPattern(mDatabaseName, WildcardOptions.IgnoreCase)
                    For Each db As Database In mSvr.Databases
                        If wc.IsMatch(db.Name) Then
                            WriteObject(db)
                        End If
                    Next
                Else
                    '// if a single database has been selected, try and return that.
                    WriteObject(mSvr.Databases.GetByName(mDatabaseName))
                End If
            End If
        End Sub

    End Class
End Namespace
