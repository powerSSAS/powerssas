Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Gosbell.PowerSSAS.PowerSSASProvider
'Imports System.Xml
'Imports System.Globalization

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASCubePermission")> _
    Public Class CmdletGetASCubePermission
        Inherits Cmdlet

        Private mRole As Role
        <Parameter(HelpMessage:="The Role for which to find CubePermissions", ValueFromPipeline:=True, Position:=0)> _
        Public Property RoleObject() As Role
            Get
                Return mRole
            End Get
            Set(ByVal value As Role)
                mRole = value
            End Set
        End Property

        'Private mRoleName As String = ""
        '<Parameter(HelpMessage:="Only return role with the specified name")> _
        'Public Property RoleName() As String
        '    Get
        '        Return mRoleName
        '    End Get
        '    Set(ByVal value As String)
        '        mRoleName = value
        '    End Set
        'End Property

        'Private mServerName As String = ""
        '<Parameter(HelpMessage:="The name of the server to query for the role information", position:=0)> _
        'Public Property ServerName() As String
        '    Get
        '        Return mServerName
        '    End Get
        '    Set(ByVal value As String)
        '        mServerName = value
        '    End Set
        'End Property

        'Private mDatabaseName As String = ""
        '<Parameter(HelpMessage:="The name of the database to query for the role information", position:=0)> _
        'Public Property DatabaseName() As String
        '    Get
        '        Return mDatabaseName
        '    End Get
        '    Set(ByVal value As String)
        '        mDatabaseName = value
        '    End Set
        'End Property


        Protected Overrides Sub ProcessRecord()
            Dim db As Database
            db = CType(mRole.OwningCollection.Parent, Database)

            For Each c As Cube In db.Cubes
                WriteObject(c.CubePermissions.GetByRole(mRole.ID))
            Next

        End Sub

    End Class
End Namespace