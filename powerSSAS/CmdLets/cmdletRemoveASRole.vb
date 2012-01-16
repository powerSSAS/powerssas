Imports System.Management.Automation
Imports Microsoft.AnalysisServices

<Cmdlet(VerbsCommon.Remove, "ASRole")> _
Public Class cmdletRemoveASRole
    Inherits Cmdlet

    Private mServerName As String = ""
    <Parameter(HelpMessage:="The name of the Server with the Database containing the Role to drop", ParameterSetName:="byName", Position:=0)> _
    Public Property ServerName() As String
        Get
            Return mServerName
        End Get
        Set(ByVal value As String)
            mServerName = value
        End Set
    End Property

    Private mDatabaseName As String = ""
    <Parameter(HelpMessage:="The name of the Database containing the Role to drop", ParameterSetName:="byName", Position:=1)> _
    Public Property DatabaseName() As String
        Get
            Return mDatabaseName
        End Get
        Set(ByVal value As String)
            mDatabaseName = value
        End Set
    End Property


    Private mRoleName As String = ""
    <Parameter(HelpMessage:="The name of the Role to drop", ParameterSetName:="byName", Position:=2)> _
    Public Property RoleName() As String
        Get
            Return mRoleName
        End Get
        Set(ByVal value As String)
            mRoleName = value
        End Set
    End Property

    Private mRole As Microsoft.AnalysisServices.Role
    <Parameter(HelpMessage:="The name of the Role to drop", ParameterSetName:="byName", ParameterSetName:="byObject", Position:=0)> _
    Public Property Role() As Role
        Get
            Return mRole
        End Get
        Set(ByVal value As Role)
            mRole = value
        End Set
    End Property


    Protected Overrides Sub ProcessRecord()
        MyBase.ProcessRecord()

        Dim db As Database = Nothing

        If mRole Is Nothing Then
            Dim Svr As Server = ConnectionFactory.ConnectToServer(mServerName)
            If Not Svr Is Nothing Then
                db = Svr.Databases.FindByName(mDatabaseName)
                If Not db Is Nothing Then
                    mRole = db.Roles.FindByName(mRoleName)
                End If
            End If
        End If

        If Not mRole Is Nothing Then
            db = CType(mRole.Parent, Database)

            For Each c As Cube In db.Cubes
                For Each cd As CubeDimension In c.Dimensions
                    'TODO
                Next
                For Each cp As CubePermission In c.CubePermissions
                    'TODO - If mRole.
                Next
            Next

            For Each d As Dimension In db.Dimensions
                'TODO 
            Next

        Else
            If mRoleName.Length > 0 Then
                Throw New ArgumentException("The role called '" & mRoleName & "' could not be found.")
            Else
                Throw New ArgumentException("You must pass either a role object or a role name to this cmdlet")
            End If
        End If

    End Sub

End Class
