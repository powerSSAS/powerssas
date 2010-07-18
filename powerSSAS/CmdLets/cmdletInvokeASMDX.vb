Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Microsoft.AnalysisServices.AdomdClient

Namespace Cmdlets
    <Cmdlet("Invoke", "ASMDX")> _
    Public Class cmdletInvokeASMDX
        Inherits Cmdlet

        Private mQuery As String = ""
        <Parameter(Position:=0, Mandatory:=True)> _
        Public Property Query() As String
            Get
                Return mQuery
            End Get
            Set(ByVal value As String)
                mQuery = value
            End Set
        End Property

        Private mServerName As String = ""
        <Parameter(Position:=1, ParameterSetName:="byServer")> _
        Public Property ServerName() As String
            Get
                Return mServerName
            End Get
            Set(ByVal value As String)
                mServerName = value
            End Set
        End Property

        Private mDatabaseName As String = ""
        <Parameter(Position:=2, ParameterSetName:="byServer")> _
        Public Property DatabaseName() As String
            Get
                Return mDatabaseName
            End Get
            Set(ByVal value As String)
                mDatabaseName = value
            End Set
        End Property

        Private mAsDataTable As SwitchParameter
        <Parameter(HelpMessage:="Returns the results of the query as a DataTable object")> _
        Public Property AsDataTable() As SwitchParameter
            Get
                Return mAsDataTable
            End Get
            Set(ByVal value As SwitchParameter)
                mAsDataTable = value
            End Set
        End Property

        Private mBenchmark As SwitchParameter
        <Parameter(HelpMessage:="Returns benchmark information instead of data.")> _
        Public Property Benchmark() As SwitchParameter
            Get
                Return mBenchmark
            End Get
            Set(ByVal value As SwitchParameter)
                mBenchmark = value
            End Set
        End Property

        Private mConnStr As String = ""
        <Parameter(Position:=1, ParameterSetName:="byConnStr", HelpMessage:="Sets the connection string for the connection that will run the MDX query")> _
        Public Property ConnectionString() As String
            Get
                Return mConnStr
            End Get
            Set(ByVal value As String)
                mConnStr = value
            End Set
        End Property

        Protected Overrides Sub ProcessRecord()
            Dim connStr As String = ""
            If mConnStr.Length > 0 Then
                connStr = mConnStr
            Else
                connStr = ConnectionFactory.ConnectToServer(mServerName).ConnectionString
                connStr &= ";Initial Catalog=" & mDatabaseName

            End If
            Dim conn As AdomdConnection = New AdomdConnection(connStr)
            conn.Open()
            Try
                Dim cmd As New AdomdCommand(mQuery, conn)
                If mBenchmark.IsPresent Then
                    ReturnBenchmarkResults(cmd)
                Else
                    If mAsDataTable.IsPresent Then
                        ReturnDataTable(cmd)
                    Else
                        ReturnCollectionOfPsObjects(cmd)
                    End If
                End If
            Finally
                conn.Close()
            End Try
        End Sub

        Private Sub ReturnBenchmarkResults(ByVal cmd As AdomdCommand)
            Dim iRowCnt As Integer
            Dim startTime As DateTime
            Dim endTime As DateTime
            Dim duration As TimeSpan

            '@"<Envelope xmlns="http://schemas.xmlsoap.org/soap/envelope/">
            '  <Body>
            '    <Execute xmlns="urn:schemas-microsoft-com:xml-analysis">
            '      <Command>
            '        <Statement>
            '          Select [Measures].[Internet Sales Amount] ON 0 FROM [Adventure Works]
            '        </Statement>
            '      </Command>
            '      <Properties>
            '        <PropertyList>
            '          <Catalog>Adventure Works DW 2008</Catalog>
            '          <!-- ### the default Format is MultiDimensional, Tabular gives you a flattened resultset -->
            '          <!-- <Format>Tabular</Format> -->
            '          <!-- ### the default Content is to return DataAndSchema, but you can turn either of these off-->
            '          <!--<Content>Data</Content> -->
            '        </PropertyList>
            '      </Properties>
            '    </Execute>
            '  </Body>
            '</Envelope>"

            Dim svr As Server = ConnectionFactory.ConnectToServer(ServerName)
            Dim trc As SessionTrace = svr.SessionTrace
            startTime = DateTime.Now()
            Dim rdr As AdomdDataReader = cmd.ExecuteReader()
            While rdr.Read
                iRowCnt += 1
            End While
            endTime = DateTime.Now()
            duration = endTime - startTime
            Dim pso As New PSObject
            pso.Properties.Add(New PSNoteProperty("StartTime", startTime))
            pso.Properties.Add(New PSNoteProperty("EndTime", endTime))
            pso.Properties.Add(New PSNoteProperty("Duration", duration.TotalMilliseconds))
            pso.Properties.Add(New PSNoteProperty("Columns", rdr.FieldCount))
            pso.Properties.Add(New PSNoteProperty("Rows", iRowCnt))

            WriteObject(pso)
        End Sub

        Private Sub ReturnCollectionOfPsObjects(ByVal cmd As AdomdCommand)
            Dim rdr As AdomdDataReader = cmd.ExecuteReader()
            Dim pso As PSObject
            While rdr.Read
                pso = New PSObject
                For i As Integer = 0 To rdr.FieldCount - 1
                    pso.Properties.Add(New PSNoteProperty(rdr.GetName(i), rdr.GetValue(i)))
                Next
                WriteObject(pso)
            End While
        End Sub

        Private Sub ReturnDataTable(ByVal cmd As AdomdCommand)
            Dim da As New AdomdDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            WriteObject(dt)
        End Sub

    End Class
End Namespace