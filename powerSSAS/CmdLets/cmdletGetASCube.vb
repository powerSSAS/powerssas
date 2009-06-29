Imports System.Management.Automation
Imports System.Xml
Imports Microsoft.AnalysisServices

Namespace Cmdlets


    <Cmdlet(VerbsCommon.Get, "ASCube", DefaultParameterSetName:="byObject")> _
    Public Class CmdletGetASCube
        Inherits Cmdlet

        Private mSvr As Server
        Private mDB As Database

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

        Private mCubeName As String = ""
        <[Alias]("Cube")> _
        <Parameter(ParameterSetName:="byName", HelpMessage:="Analysis Services cube name", Position:=2, ValueFromPipeline:=False)> _
        Public Property CubeName() As String
            Get
                Return mCubeName
            End Get
            Set(ByVal value As String)
                mCubeName = value
            End Set
        End Property

        Private mParamDB As Microsoft.AnalysisServices.Database
        <[Alias]("Server")> _
        <Parameter(ParameterSetName:="byObject", HelpMessage:="Analysis Services server object", Position:=0, ValueFromPipeline:=True)> _
        Public Property ASDatabase() As Microsoft.AnalysisServices.Database
            Get
                Return mParamDB
            End Get
            Set(ByVal value As Microsoft.AnalysisServices.Database)
                mParamDB = value
            End Set
        End Property

        Protected Overrides Sub ProcessRecord()

            If mParamDB IsNot Nothing Then
                '// if the server object is passed in, then use that
                mDB = mParamDB
            ElseIf mParamDB Is Nothing AndAlso mServerName.Length > 0 Then
                '// if just the name is passed in, then we need to connect
                mSvr = ConnectionFactory.ConnectToServer(mServerName)
                mDB = mSvr.Databases.GetByName(mDatabaseName)
            Else
                WriteError(New ErrorRecord(New ArgumentException("You must pass in a Database object or server & database name"), "MissingParam", ErrorCategory.InvalidArgument, Nothing))
            End If

            If mCubeName.Length = 0 Then
                '// if no specific cube is requested, return the whole collection.
                For Each c As Cube In mDB.Cubes
                    WriteObject(c)
                Next c
            Else
                If WildcardPattern.ContainsWildcardCharacters(mCubeName) Then
                    '// if we have a wildcard pattern, search for matches
                    Dim wc As New System.Management.Automation.WildcardPattern(mCubeName, WildcardOptions.IgnoreCase)
                    For Each c As Cube In mDB.Cubes
                        If wc.IsMatch(c.Name) Then
                            WriteObject(c)
                        End If
                    Next
                Else
                    '// if a single cube has been selected, try and return that.
                    WriteObject(mDB.Cubes.GetByName(mCubeName))
                End If
            End If
        End Sub

    End Class
End Namespace
