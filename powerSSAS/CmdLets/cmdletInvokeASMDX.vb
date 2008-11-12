Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Microsoft.AnalysisServices.AdomdClient

<Cmdlet("Invoke", "ASMDX", SupportsShouldProcess:=True)> _
Public Class cmdletInvokeASMDX
    Inherits Cmdlet

    Private mQuery As String = ""
    <Parameter(Position:=1, Mandatory:=False)> _
    Public Property Query() As String
        Get
            Return mQuery
        End Get
        Set(ByVal value As String)
            mQuery = value
        End Set
    End Property

    Private mServerName As String = ""
    <Parameter(Position:=0, Mandatory:=False)> _
    Public Property ServerName() As String
        Get
            Return mServerName
        End Get
        Set(ByVal value As String)
            mServerName = value
        End Set
    End Property

    Protected Overrides Sub ProcessRecord()
        Dim connStr As String = ConnectionFactory.ConnectToServer(mServerName).ConnectionString
        Dim conn As AdomdConnection = New AdomdConnection(connStr)
        conn.Open()
        Dim cmd As New AdomdCommand(mQuery, conn)
        Dim rdr As AdomdDataReader = cmd.ExecuteReader()
        Dim pso As PSObject
        While rdr.Read
            pso = New PSObject
            For i As Integer = 0 To rdr.FieldCount - 1
                pso.Properties.Add(New PSNoteProperty(rdr.GetName(i), rdr.GetValue(i)))
            Next
            WriteObject(pso)
        End While
        conn.Close()
    End Sub

End Class
