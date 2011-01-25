Imports System.Management.Automation
Imports System.Xml

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Clear, "ASSession", SupportsShouldProcess:=True, DefaultParameterSetName:="bySession")> _
    Public Class CmdletClearSession
        Inherits Cmdlet

        Private mSessionID As String = ""
        <[Alias]("ID")> _
        <AllowNull()> _
        <Parameter(ParameterSetName:="bySessionID")> _
        <Parameter(HelpMessage:="Unique Session ID (Guid)", Position:=1)> _
        Public Property SessionId() As String
            Get
                Return mSessionID
            End Get
            Set(ByVal value As String)
                mSessionID = value
            End Set
        End Property

        Private mServerName As String = ""
        <[Alias]("Server")> _
        <AllowNull(), AllowEmptyString()> _
        <Parameter(ParameterSetName:="bySessionID")> _
        <Parameter(ParameterSetName:="bySPID")> _
        <Parameter(HelpMessage:="Analysis Services server name", Position:=0, ValueFromPipeline:=False)> _
        Public Property ServerName() As String
            Get
                Return mServerName
            End Get
            Set(ByVal value As String)
                mServerName = value
            End Set
        End Property

        Private mSPID As Integer
        <AllowNull()> _
        <Parameter(ParameterSetName:="bySPID")> _
        <Parameter(HelpMessage:="Unique Session ID (Guid)", Position:=1)> _
        Public Property SPID() As Integer
            Get
                Return mSPID
            End Get
            Set(ByVal value As Integer)
                mSPID = value
            End Set
        End Property


        Private mSession As PowerSSAS.Types.Session()
        <AllowNull()> _
        <Parameter(HelpMessage:="Analysis Services server name", Position:=0, ParameterSetName:="bySession", ValueFromPipeline:=True)> _
        Public Property InputObject() As PowerSSAS.Types.Session()
            Get
                Return mSession
            End Get
            Set(ByVal value As PowerSSAS.Types.Session())
                mSession = value
            End Set
        End Property

        Protected Overrides Sub ProcessRecord()
            If Not InputObject Is Nothing Then
                For Each s As PowerSSAS.Types.Session In InputObject
                    cancelSession(s.ServerName, "<SessionID>" & s.Id & "</SessionID>")
                Next
            ElseIf SessionID.Length > 0 Then
                cancelSession(ServerName, "<SessionID>" & SessionID & "</SessionID>")
            ElseIf SPID <> 0 Then
                cancelSession(ServerName, "<SPID>" & Me.SPID & "</SPID>")
            Else
                Throw New ArgumentException("no session identified")
            End If

        End Sub

        Private Sub CancelSession(ByVal servername As String, ByVal restriction As String)

            Dim cancelCmd As String = "<Cancel xmlns=""http://schemas.microsoft.com/analysisservices/2003/engine"">"
            cancelCmd &= restriction
            cancelCmd &= "</Cancel> "

            WriteVerbose("Cancelling Session where: " & restriction)

            'TODO - do we need to check that the user does not cancel themselves?

            'Dim xc As New Microsoft.AnalysisServices.Xmla.XmlaClient
            'xc.Connect("Data Source=" & servername)
            Dim svr As New Microsoft.AnalysisServices.Server()
            svr.Connect(servername)

            Dim res As String = ""
            Try
                If Me.ShouldProcess(String.Format("Server: {0} Restriction: {1}", servername, restriction), "Clear-ASSession") Then
                    'TODO - should i use svr.CancelSession(...) instead?
                    svr.Execute(cancelCmd)
                End If
            Finally
                If Not svr Is Nothing Then
                    svr.Disconnect()
                End If
            End Try

            '// the result will be empty if the cmdlet was run with the -whatif parameter
            If res.Length > 0 Then
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
            End If
        End Sub
    End Class
End Namespace