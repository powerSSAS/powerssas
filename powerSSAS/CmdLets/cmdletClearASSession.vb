Imports System.Management.Automation
Imports System.Xml
Imports Microsoft.AnalysisServices.Xmla

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Clear, "ASSession", SupportsShouldProcess:=True)> _
    Public Class cmdletClearSession
        Inherits Cmdlet

        Private mSessionID As String = ""
        <Parameter(HelpMessage:="Unique Session ID (Guid)", Position:=1, Mandatory:=True, ValueFromPipeline:=True)> _
        Public Property SessionID() As String
            Get
                Return mSessionID
            End Get
            Set(ByVal value As String)
                mSessionID = value
            End Set
        End Property

        Private mServerName As String = ""
        <Parameter(HelpMessage:="Analysis Services server name", Mandatory:=True, Position:=0, ValueFromPipeline:=True)> _
        Public Property ServerName() As String
            Get
                Return mServerName
            End Get
            Set(ByVal value As String)
                mServerName = value
            End Set
        End Property

        Protected Overrides Sub BeginProcessing()
            Dim cancelCmd As String = "<Cancel xmlns=""http://schemas.microsoft.com/analysisservices/2003/engine""><SessionID>"
            cancelCmd &= Me.SessionID & "</SessionID></Cancel> "

            WriteVerbose("Cancelling SessionID: " & Me.SessionID)

            Dim xc As New Microsoft.AnalysisServices.Xmla.XmlaClient
            xc.Connect("Data Source=" & ServerName)
            Dim res As String = ""
            Try
                res = xc.Send(cancelCmd, Nothing)
            Finally
                If Not xc Is Nothing Then
                    xc.Disconnect()
                End If
            End Try

            '// check for errors
            Dim tr As New IO.StringReader(res)
            Dim xmlRdr As New XmlTextReader(tr)
            Dim msg As Object = xmlRdr.NameTable.Add("Error")
            While xmlRdr.Read
                If xmlRdr.NodeType = XmlNodeType.Element AndAlso Object.ReferenceEquals(xmlRdr.LocalName, msg) Then
                    If xmlRdr.MoveToAttribute("Description") Then
                        WriteError(New ErrorRecord(New ArgumentException(xmlRdr.Value), "Error Cancelling", ErrorCategory.InvalidOperation, Me))
                    End If
                End If
            End While
        End Sub
    End Class
End Namespace