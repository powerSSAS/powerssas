Imports System.Management.Automation
Imports Microsoft.AnalysisServices

<Cmdlet(VerbsCommon.Remove, "ASClrAssembly", SupportsShouldProcess:=True)> _
Public Class CmdletRemoveASClsAssembly
    Inherits PSCmdlet

    '<Parameter(Position:=0, Mandatory:=False)> _
    'Public Property Name() As String
    '    Get
    '        Return ""
    '    End Get
    '    Set(ByVal value As String)
    '    End Set
    'End Property

    '                 strings    Db-obj  Svr-Obj  
    ' AssemblyName       x         x        x
    ' DatabaseName      (x)                (x)
    ' Database                     x
    ' ServerName         x
    ' ServerObject                          x

    Private mAssemblyName As String = ""
    <Parameter(Position:=0, Mandatory:=True, ParameterSetName:="Names")> _
    <Parameter(Position:=0, Mandatory:=True, ParameterSetName:="DbObj")> _
    <Parameter(Position:=0, Mandatory:=True, ParameterSetName:="ServerObj")> _
    Public Property AssemblyName() As String
        Get
            Return mAssemblyName
        End Get
        Set(ByVal value As String)
            mAssemblyName = value
        End Set
    End Property

    Private mDatabaseName As String = ""
    <Parameter(Position:=1, Mandatory:=False, ParameterSetName:="Names")> _
    Public Property DatabaseName() As String
        Get
            Return mDatabaseName
        End Get
        Set(ByVal value As String)
            mDatabaseName = value
        End Set
    End Property


    Private mServerName As String = ""
    <Parameter(Position:=2, Mandatory:=False, ParameterSetName:="Names")> _
    Public Property ServerName() As String
        Get
            Return mServerName
        End Get
        Set(ByVal value As String)
            mServerName = value
        End Set
    End Property

    Private mDatabaseObj As Microsoft.AnalysisServices.Database
    <Parameter(Position:=1, Mandatory:=False, ParameterSetName:="DbObj")> _
    Public Property Database() As Database
        Get
            Return mDatabaseObj
        End Get
        Set(ByVal value As Database)
            mDatabaseObj = value
        End Set
    End Property

    Private mServerObj As Microsoft.AnalysisServices.Server
    <Parameter(Position:=1, Mandatory:=False, ParameterSetName:="ServerObj")> _
    Public Property Server() As Server
        Get
            Return mServerObj
        End Get
        Set(ByVal value As Server)
            mServerObj = value
        End Set
    End Property


    Protected Overrides Sub ProcessRecord()
        Dim db As Database = Nothing
        Dim svr As Server = Nothing
        Dim ass As Assembly
        Try
            Select Case Me.ParameterSetName
                Case "Names"
                    svr = ConnectionFactory.ConnectToServer(mServerName)
                    If mDatabaseName.Length > 0 Then
                        db = svr.Databases.GetByName(Me.mDatabaseName)
                        If db Is Nothing Then
                            WriteError(New ErrorRecord(New ArgumentException("Database not found on server"), "Database not found", ErrorCategory.InvalidArgument, svr))
                        End If
                    End If
                Case "DbObj"
                    db = mDatabaseObj
                Case "ServerObj"
                    svr = mServerObj
                    If mDatabaseName.Length > 0 Then
                        db = svr.Databases.GetByName(mDatabaseName)
                    End If
            End Select
            If db Is Nothing Then
                ass = svr.Assemblies.GetByName(Me.mAssemblyName)
                'TODO: throw error if assembly name not found
                '//svr.Assemblies.Remove(ass)
                ass.Drop()
            Else
                ass = db.Assemblies.GetByName(Me.mAssemblyName)
                'TODO: throw error if assembly name not found
                '//db.Assemblies.Remove(ass)
                ass.Drop()
            End If

        Catch ex As Exception

        End Try
    End Sub

End Class
