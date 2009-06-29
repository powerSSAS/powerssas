Imports System.Management.Automation
Imports Microsoft.AnalysisServices

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Add, "ASClrAssembly")> _
    Public Class CmdletAddASClrAssembly
        Inherits PSCmdlet

        '<Parameter(Position:=0, Mandatory:=False)> _
        'Public Property Name() As String
        '    Get
        '        Return ""
        '    End Get
        '    Set(ByVal value As String)
        '    End Set
        'End Property

        '                 Strings  db-obj  svr-Obj
        ' AssemblyName       x       x        x
        ' FilePath           x       x        x
        ' PermissionSet      x       x        x
        ' DatabaseName      (x)              (x)
        ' DatabaseObject             x 
        ' ServerName         x
        ' ServerObject                        x

        ' AssemblyName
        ' FilePath
        ' PermissionSet
        ' DatabaseName
        ' DatabaseObject
        ' ServerName 
        ' ServerObject

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

        Private mFilePath As String
        <Parameter(Mandatory:=True, Position:=1, ParameterSetName:="Names")> _
        <Parameter(Mandatory:=True, Position:=1, ParameterSetName:="DbObj")> _
        <Parameter(Mandatory:=True, Position:=1, ParameterSetName:="ServerObj")> _
        Public Property FilePath() As String
            Get
                Return mFilePath
            End Get
            Set(ByVal value As String)
                mFilePath = value
            End Set
        End Property

        Private mPermissionSet As PermissionSet
        <Parameter(Position:=2, Mandatory:=True, ParameterSetName:="Names")> _
        <Parameter(Position:=2, Mandatory:=True, ParameterSetName:="DbObj")> _
        <Parameter(Position:=2, Mandatory:=True, ParameterSetName:="ServerObj")> _
        Public Property PermissionSet() As PermissionSet
            Get
                Return mPermissionSet
            End Get
            Set(ByVal value As PermissionSet)
                mPermissionSet = value
            End Set
        End Property


        Private mDatabaseName As String = ""
        <Parameter(Position:=4, Mandatory:=False, ParameterSetName:="Names")> _
        <Parameter(Position:=4, Mandatory:=False, ParameterSetName:="ServerObj")> _
        Public Property DatabaseName() As String
            Get
                Return mDatabaseName
            End Get
            Set(ByVal value As String)
                mDatabaseName = value
            End Set
        End Property

        Private mServerName As String = ""
        <Parameter(Position:=3, Mandatory:=True, ParameterSetName:="Names")> _
        Public Property ServerName() As String
            Get
                Return mServerName
            End Get
            Set(ByVal value As String)
                mServerName = value
            End Set
        End Property

        Private mDatabaseObj As Microsoft.AnalysisServices.Database
        <Parameter(Position:=3, Mandatory:=False, ParameterSetName:="DbObj")> _
        Public Property Database() As Database
            Get
                Return mDatabaseObj
            End Get
            Set(ByVal value As Database)
                mDatabaseObj = value
            End Set
        End Property


        Private mServerObj As Server
        <Parameter(ParameterSetName:="ServerObj", Position:=3)> _
        Public Property Server() As Server
            Get
                Return mServerObj
            End Get
            Set(ByVal value As Server)
                mServerObj = value
            End Set
        End Property

        Private mLoadPdb As Boolean '= False
        'TODO - add parameter

        Private mImpInfo As New ImpersonationInfo(ImpersonationMode.Default)
        'TODO - add parameter

        Protected Overrides Sub ProcessRecord()
            Dim assCol As AssemblyCollection = Nothing
            Try
                Select Case Me.ParameterSetName
                    Case "Names"
                        mServerObj = ConnectionFactory.ConnectToServer(mServerName)
                        If mDatabaseName.Length > 0 Then
                            mDatabaseObj = mServerObj.Databases.GetByName(mDatabaseName)
                            assCol = mDatabaseObj.Assemblies
                        Else
                            assCol = mServerObj.Assemblies
                        End If
                    Case "DbObj"
                        assCol = mDatabaseObj.Assemblies
                    Case "ServerObj"
                        If mDatabaseName.Length = 0 Then
                            assCol = mServerObj.Assemblies
                        Else
                            If mServerObj.Databases.ContainsName(mDatabaseName) Then
                                assCol = mServerObj.Databases.GetByName(mDatabaseName).Assemblies
                            Else
                                WriteError(New ErrorRecord(New ArgumentException(String.Format("database '{0}' not found", mDatabaseName)), "Database not found", ErrorCategory.InvalidArgument, mDatabaseName))
                            End If
                        End If
                End Select

                Dim ass As ClrAssembly = assCol.Add(mAssemblyName)
                ass.PermissionSet = mPermissionSet
                ass.ImpersonationInfo = mImpInfo
                ass.LoadFiles(mFilePath, mLoadPdb)
                ass.Update()
            Catch ex As Exception
                WriteError(New ErrorRecord(ex, "error", ErrorCategory.NotSpecified, Nothing))
            End Try
        End Sub

    End Class
End Namespace