Imports Microsoft.AnalysisServices

Public NotInheritable Class ConnectionFactory
    Private Shared mServers As New SortedList(Of String, Microsoft.AnalysisServices.Server)

    Private Sub New()
    End Sub

    Public Shared Function ConnectToServer(ByVal serverName As String) As Microsoft.AnalysisServices.Server
        If mServers.ContainsKey(serverName) Then
            Return mServers.Item(serverName)
        ElseIf serverName.Length = 0 And mServers.Count = 1 Then
            '// Return the first item in the collection if no server name has been specified 
            '// and there is only one in the cached collection
            Return mServers.Item(mServers.Keys(0))
        Else
            Dim svr As Server = New Server()
            Dim connStr As String = "Data Source=" & serverName & ";Application Name=PowerSSAS SSAS PowerShell provider"
            svr.Connect(connStr)
            mServers.Add(serverName, svr)
            Return svr
        End If
    End Function

    Public Shared Function GetServerFromObject(ByVal majObj As MajorObject) As Server
        Dim mObj As IModelComponent = majObj
        While Not TypeOf (mObj) Is Server
            mObj = mObj.Parent
        End While
        '// will return null if mboj is not a server object
        Return TryCast(mObj, Server)
    End Function

End Class
