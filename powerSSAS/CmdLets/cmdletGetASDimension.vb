Imports System.Management.Automation
Imports System.Xml
Imports Microsoft.AnalysisServices.Xmla
Imports Microsoft.AnalysisServices

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASDimension", SupportsShouldProcess:=True, DefaultParameterSetName:="byObject")> _
    Public Class cmdletGetASDimension
        Inherits Cmdlet

        Private mSvr As New Server()
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


        Protected Overrides Sub EndProcessing()
            '// if the server was connected inside this cmdlet, we should disconnect it.
            If mParamDB Is Nothing AndAlso mServerName.Length > 0 Then
                mSvr.Disconnect()
            End If
            MyBase.EndProcessing()
        End Sub

        Protected Overrides Sub ProcessRecord()

            If mParamDB IsNot Nothing Then
                '// if the server object is passed in, then use that
                mDB = mParamDB
            ElseIf mParamDB Is Nothing AndAlso mServerName.Length > 0 Then
                '// if just the name is passed in, then we need to connect
                mSvr.Connect(mServerName)
                mDB = mSvr.Databases.GetByName(mDatabaseName)
            Else
                WriteError(New ErrorRecord(New ArgumentException("You must pass in a Database object or server & database name"), "MissingParam", ErrorCategory.InvalidArgument, Nothing))
            End If

            'If mDatabaseName.Length = 0 Then

            '// if no specific cube is requested, return the whole collection.
            WriteObject(mDB.Dimensions)

            'Else
            'If WildcardPattern.ContainsWildcardCharacters(mDatabaseName) Then
            '    '// if we have a wildcard pattern, search for matches
            '    Dim wc As New System.Management.Automation.WildcardPattern(mDatabaseName, WildcardOptions.IgnoreCase)
            '    For Each db As Database In mSvr.Databases
            '        If wc.IsMatch(db.Name) Then
            '            WriteObject(db)
            '        End If
            '    Next
            'Else
            '    '// if a single database has been selected, try and return that.
            '    WriteObject(mSvr.Databases.GetByName(mDatabaseName))
            'End If
            'End If
        End Sub

    End Class
End Namespace
