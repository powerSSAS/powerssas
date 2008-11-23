Imports System.Management.Automation
Imports System.Xml
Imports Microsoft.AnalysisServices.Xmla
Imports Microsoft.AnalysisServices

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASDimension", DefaultParameterSetName:="byObject")> _
    Public Class CmdletGetASDimension
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

        Private mDimName As String = ""
        <[Alias]("Dimension")> _
        <Parameter(ParameterSetName:="byName", HelpMessage:="Analysis Services Dimension name", position:=2, ValueFromPipeline:=False, Mandatory:=False)> _
        Public Property DimensionName() As String
            Get
                Return mDimName
            End Get
            Set(ByVal value As String)
                mDimName = value
            End Set
        End Property

        Private mParamCubeName As String = ""
        <Parameter(ParameterSetName:="byName", HelpMessage:="Analysis Services Cube Name", Position:=3, ValueFromPipeline:=True)> _
        Public Property CubeName() As String
            Get
                Return mParamCubeName
            End Get
            Set(ByVal value As String)
                mParamCubeName = value
            End Set
        End Property

        Private mParamDB As Microsoft.AnalysisServices.Database
        <[Alias]("Server")> _
        <Parameter(ParameterSetName:="byDbObject", HelpMessage:="Analysis Services server object", Position:=0, ValueFromPipeline:=True)> _
        Public Property ASDatabase() As Microsoft.AnalysisServices.Database
            Get
                Return mParamDB
            End Get
            Set(ByVal value As Microsoft.AnalysisServices.Database)
                mParamDB = value
            End Set
        End Property

        Private mParamCube As Microsoft.AnalysisServices.Cube
        <[Alias]("Cube")> _
        <Parameter(ParameterSetName:="byCubeObject", HelpMessage:="Analysis Services Cube object", Position:=0, ValueFromPipeline:=True)> _
        Public Property ASCube() As Microsoft.AnalysisServices.Cube
            Get
                Return mParamCube
            End Get
            Set(ByVal value As Microsoft.AnalysisServices.Cube)
                mParamCube = value
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

            If mParamCube Is Nothing AndAlso mParamCubeName.Length = 0 Then

                '// if no specific cube is requested, return the whole collection.
                If WildcardPattern.ContainsWildcardCharacters(mDimName) Then
                    '// if we have a wildcard pattern, search for matches
                    Dim wc As New System.Management.Automation.WildcardPattern(mDatabaseName, WildcardOptions.IgnoreCase)
                    For Each d As Dimension In mDB.Dimensions
                        If wc.IsMatch(d.Name) Then
                            WriteObject(d)
                        End If
                    Next
                ElseIf mDimName.Length = 0 Then
                    For Each dbdim As Dimension In mDB.Dimensions
                        WriteObject(dbdim)
                    Next dbdim
                Else
                    '// if a single database has been selected, try and return that.
                    WriteObject(mDB.Dimensions.GetByName(mDimName))
                End If
            Else
                If mParamCube Is Nothing AndAlso mParamCubeName.Length > 0 Then
                    mDB.Cubes.FindByName(mParamCubeName)
                End If

                If WildcardPattern.ContainsWildcardCharacters(mDimName) Then
                    '// if we have a wildcard pattern, search for matches
                    Dim wc As New System.Management.Automation.WildcardPattern(mDatabaseName, WildcardOptions.IgnoreCase)
                    For Each d As Dimension In mParamCube.Dimensions
                        If wc.IsMatch(d.Name) Then
                            WriteObject(d)
                        End If
                    Next
                ElseIf mDimName.Length = 0 Then
                    For Each cubeDim As Dimension In mParamCube.Dimensions
                        WriteObject(cubeDim)
                    Next cubeDim
                Else
                    '// if a single database has been selected, try and return that.
                    WriteObject(mParamCube.Dimensions.GetByName(mDatabaseName))
                End If
            End If
        End Sub

    End Class
End Namespace
