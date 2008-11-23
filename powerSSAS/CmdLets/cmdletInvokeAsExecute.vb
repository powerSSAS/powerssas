Imports System.Management.Automation
Imports Microsoft.AnalysisServices

<Cmdlet("Invoke", "AsExecute", SupportsShouldProcess:=True)> _
Public Class cmdletInvokeAsExecute
    Inherits Cmdlet

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

    Private mXmla As String = ""
    <Parameter(Position:=1, Mandatory:=True, HelpMessage:="")> _
    Public Property Xmla() As String
        Get
            Return mXmla
        End Get
        Set(ByVal value As String)
            mXmla = value
        End Set
    End Property

    Protected Overrides Sub ProcessRecord()
        Dim svr As Server
        svr = ConnectionFactory.ConnectToServer(mServerName)
        Dim results As XmlaResultCollection
        results = svr.Execute(mXmla)
        For Each res As XmlaResult In results
            WriteObject(res.Messages, True)
        Next
    End Sub

End Class
