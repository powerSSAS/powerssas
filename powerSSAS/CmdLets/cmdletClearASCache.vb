Imports System.Management.Automation
Imports System.Xml

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Clear, "ASCache", SupportsShouldProcess:=True, DefaultParameterSetName:="byDatabase")> _
    Public Class CmdletClearCache
        Inherits Cmdlet

        Private mDatabaseID As String = ""
        <[Alias]("ID")> _
        <AllowNull()> _
        <Parameter(ParameterSetName:="byDatabase")> _
        <Parameter(HelpMessage:="Database ID", Position:=1)> _
        Public Property DatabaseId() As String
            Get
                Return mDatabaseID
            End Get
            Set(ByVal value As String)
                mDatabaseID = value
            End Set
        End Property

        Private mServerName As String = ""
        <[Alias]("Server")> _
        <AllowNull(), AllowEmptyString()> _
        <Parameter(ParameterSetName:="byDatabase")> _
        <Parameter(HelpMessage:="Analysis Services server name", Position:=0, ValueFromPipeline:=False)> _
        Public Property ServerName() As String
            Get
                Return mServerName
            End Get
            Set(ByVal value As String)
                mServerName = value
            End Set
        End Property


        Protected Overrides Sub ProcessRecord()
            If DatabaseId.Length > 0 Then
                ClearCache(ServerName, "<DatabaseID>" & DatabaseId & "</DatabaseID>")
            Else
                Throw New ArgumentException("no session identified")
            End If

        End Sub

        Private Sub ClearCache(ByVal servername As String, ByVal restriction As String)

            Dim clearCacheCmd As String = "<ClearCache xmlns=""http://schemas.microsoft.com/analysisservices/2003/engine""><Object>"
            clearCacheCmd &= restriction
            clearCacheCmd &= "</Object></ClearCache> "

            WriteVerbose("Clearing Cache where: " & restriction)

            'Dim xc As New Microsoft.AnalysisServices.Xmla.XmlaClient
            'xc.Connect("Data Source=" & servername)
            Dim svr As New Microsoft.AnalysisServices.Server()
            svr.Connect(servername)

            Dim res As String = ""
            Try
                If Me.ShouldProcess(String.Format("Server: {0} Command: {1}", servername, clearCacheCmd), "Clear-ASCache") Then
                    svr.Execute(clearCacheCmd)
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