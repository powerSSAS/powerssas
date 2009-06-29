Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Gosbell.PowerSSAS.PowerSSASProvider
Imports System.Xml
Imports System.Globalization

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASRole")> _
    Public Class CmdletGetASRole
        Inherits Cmdlet

        Private mRoleID As String = ""
        <Parameter(HelpMessage:="Only returns the Role with the specified RoleID")> _
        Public Property RoleID() As String
            Get
                Return mRoleID
            End Get
            Set(ByVal value As String)
                mRoleID = value
            End Set
        End Property

        Private mRoleName As String = ""
        <Parameter(HelpMessage:="Only return role with the specified name")> _
        Public Property RoleName() As String
            Get
                Return mRoleName
            End Get
            Set(ByVal value As String)
                mRoleName = value
            End Set
        End Property

        Private mServerName As String = ""
        <Parameter(HelpMessage:="The name of the server to query for the role information", position:=0)> _
        Public Property ServerName() As String
            Get
                Return mServerName
            End Get
            Set(ByVal value As String)
                mServerName = value
            End Set
        End Property

        Private mDatabaseName As String = ""
        <Parameter(HelpMessage:="The name of the database to query for the role information", position:=1)> _
        Public Property DatabaseName() As String
            Get
                Return mDatabaseName
            End Get
            Set(ByVal value As String)
                mDatabaseName = value
            End Set
        End Property


        Protected Overrides Sub ProcessRecord()
            Dim svr As Server
            Dim roles As RoleCollection

            svr = ConnectionFactory.ConnectToServer(mServerName)
            roles = svr.Databases.GetByName(mDatabaseName).Roles
            If mRoleID.Length > 0 Then
                WriteObject(roles.Item(mRoleID))
            ElseIf mRoleName.Length > 0 Then
                WriteObject(roles.GetByName(mRoleName))
            Else
                WriteObject(roles, True)
            End If

        End Sub

    End Class
End Namespace