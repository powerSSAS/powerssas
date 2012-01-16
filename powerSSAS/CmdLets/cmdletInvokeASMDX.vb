Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Microsoft.AnalysisServices.AdomdClient
Imports System
Imports System.Threading

Namespace Cmdlets
    <Cmdlet("Invoke", "ASMDX")> _
    Public Class cmdletInvokeASMDX
        Inherits Cmdlet

        Private evStat As New AutoResetEvent(False) 'ManualResetEvent(False)
        Private evExit As New AutoResetEvent(False)
        Private evts As WaitHandle()
        Private pr_sync As New Object()
        Private m_event As TraceEventArgs
        Private m_benchmarkresult As New MDXBenchmarkResult()

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

        Private mQueryName As String = ""
        <Parameter(Mandatory:=False)> _
        Public Property QueryName() As String
            Get
                Return mQueryName
            End Get
            Set(ByVal value As String)
                mQueryName = value
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

        Private mClientStatistics As SwitchParameter
        <Parameter(HelpMessage:="Returns benchmark information instead of data.")> _
        Public Property ClientStatistics() As SwitchParameter
            Get
                Return mClientStatistics
            End Get
            Set(ByVal value As SwitchParameter)
                mClientStatistics = value
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
                Dim svr As Server = ConnectionFactory.ConnectToServer(mServerName)
                connStr = svr.ConnectionString & ";Initial Catalog=" & mDatabaseName '& ";SessionId=" & svr.SessionID
            End If
            Dim conn As AdomdConnection = New AdomdConnection(connStr)
            Try
                conn.Open()
            
                Try
                    Dim cmd As New AdomdCommand(mQuery, conn)

                    If mBenchmark.IsPresent Then
                        ReturnBenchmarkResults(cmd)
                    ElseIf mClientStatistics.IsPresent Then
                        ReturnClientStatistics(cmd)
                    Else

                        If mAsDataTable.IsPresent Then
                            ReturnDataTable(cmd)
                        Else
                            ReturnCollectionOfPsObjects(cmd)
                        End If
                    End If
                Finally
                    conn.Close(False)
                End Try
            Catch ex As AdomdConnectionException
                '\\ the inner exception has more information about the specific error
                '\\ so display that if it is present (as by default PowerShell only 
                '\\ shows the outermost exception message)
                If Not ex.InnerException Is Nothing Then
                    WriteError(New ErrorRecord(ex.InnerException, "", ErrorCategory.NotSpecified, Me))
                Else
                    WriteError(New ErrorRecord(ex, "", ErrorCategory.NotSpecified, Me))
                End If

            Catch ex2 As Exception
                WriteError(New ErrorRecord(ex2, "", ErrorCategory.NotSpecified, Me))
            End Try
        End Sub

        Private Sub ReturnClientStatistics(ByVal cmd As AdomdCommand)
            ' Start Stopwatch
            Dim startTime As DateTime = DateTime.Now()
            ' execute cellset
            Dim cs As CellSet
            cs = cmd.ExecuteCellSet()
            ' End Stopwatch
            Dim endTime As DateTime = DateTime.Now()
            ' write results
            Dim results As New PSObject
            results.Properties.Add(New PSNoteProperty("StartTime", startTime))
            results.Properties.Add(New PSNoteProperty("EndTime", endTime))
            results.Properties.Add(New PSNoteProperty("Duration", endTime - startTime))
            results.Properties.Add(New PSNoteProperty("Columns", GetMemberCount(cs, 0)))
            results.Properties.Add(New PSNoteProperty("Rows", GetMemberCount(cs, 1)))
            WriteObject(results)
        End Sub

        Private Sub ReturnBenchmarkResults(ByVal cmd As AdomdCommand)
            Dim svr As Server = ConnectionFactory.ConnectToServer(ServerName)
            'Dim trc As SessionTrace = svr.SessionTrace
            Const EXIT_HANDLE As Integer = 0
            'TODO - implement IsTabular parameter when benchmarking so that we can 
            '       compare flattened vs. Native recordsets.
            Dim asyncCmd As CellsetCommand = New CellsetCommand(svr, cmd, False)

            AddHandler asyncCmd.QueryStatus, AddressOf OnQueryStatus
            AddHandler asyncCmd.TraceEvent, AddressOf OnTraceEvent
            AddHandler asyncCmd.Completed, AddressOf OnCompleted
            ReDim evts(1)
            evts(EXIT_HANDLE) = evExit
            evts(1) = evStat

            '// Start the Async command
            asyncCmd.Execute()

            '// This loop executes each time the evStat waithandle is signaled
            '// inbetween that the WaitAny call is blocking. When the exit handle
            '// is signalled then the loop exits and the rest of the code is executed.
            While WaitHandle.WaitAny(evts) <> EXIT_HANDLE
                SyncLock pr_sync
                    AggregateEvent(m_event)
                End SyncLock
            End While

            'Dim rdr As AdomdDataReader = asyncCmd.AdomdReader
            'While rdr.Read
            '    iRowCnt += 1
            'End While

            Dim cs As CellSet = asyncCmd.AdomdCellset
            m_benchmarkresult.Columns = GetMemberCount(cs, 0)
            m_benchmarkresult.Rows = GetMemberCount(cs, 1)
            m_benchmarkresult.CellsCalculated = asyncCmd.CellsCalculated
            m_benchmarkresult.FlatCacheInserts = asyncCmd.FlatCacheInserts
            m_benchmarkresult.StablePerformanceCounters = asyncCmd.PerfCountersStable
            m_benchmarkresult.NonEmptyUnoptimized = asyncCmd.NonEmptyUnoptimized
            m_benchmarkresult.QueryName = mQueryName
            WriteObject(m_benchmarkresult)

            'endTime = DateTime.Now()
            'duration = endTime - startTime
            'Dim pso As New PSObject
            'pso.Properties.Add(New PSNoteProperty("StartTime", startTime))
            'pso.Properties.Add(New PSNoteProperty("EndTime", endTime))
            'pso.Properties.Add(New PSNoteProperty("Duration", duration.TotalMilliseconds))
            'pso.Properties.Add(New PSNoteProperty("Columns", rdr.FieldCount))
            'pso.Properties.Add(New PSNoteProperty("Rows", iRowCnt))

            'WriteObject(pso)
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

        Private Sub OnTraceEvent(ByVal stat As TraceEventArgs)
            SyncLock pr_sync
                m_event = stat
                evStat.Set()
            End SyncLock
        End Sub

        Private Sub OnCompleted()
            SyncLock pr_sync
                evExit.Set()
            End SyncLock
        End Sub

        Private Sub OnQueryStatus(ByVal status As String)
            Me.WriteProgress(New ProgressRecord(0, "MDX Benchmark", status))
        End Sub

        '// This routine aggregates the trace events that are captured during query execution
        Private Sub AggregateEvent(ByVal e As TraceEventArgs)
            Select Case e.EventClass
                Case TraceEventClass.QueryEnd
                    If e.EventSubclass = TraceEventSubclass.MdxQuery Then
                        m_benchmarkresult.QueryDurationMS = e.Duration
                        m_benchmarkresult.Endtime = e.CurrentTime
                    End If
                Case TraceEventClass.QuerySubcube
                    m_benchmarkresult.QuerySubcubeDurationMS += e.Duration
                    m_benchmarkresult.querySubcubeCount += 1
                Case TraceEventClass.GetDataFromAggregation
                    m_benchmarkresult.AggHit += 1
                Case TraceEventClass.GetDataFromCache
                    m_benchmarkresult.CacheHit += 1
                Case TraceEventClass.QueryBegin
                    m_benchmarkresult.StartTime = e.CurrentTime
            End Select
        End Sub

        Private Function GetMemberCount(ByVal cs As CellSet, ByVal axis As Integer) As Long
            Return cs.Axes(axis).Set.Tuples.Count
        End Function
    End Class

    ''TODO - formatting of results, should this be done with powershell format file?
    'Public Class BenchmarkResult
    '    Public QueryDurationMS As Long
    '    Public QuerySubcubeDurationMS As Long
    '    Public QuerySubcubeCount As Long
    '    Public CacheHit As Long
    '    Public AggHit As Long
    '    Public Rows As Long
    '    Public Columns As Long
    '    Public StartTime As Date
    '    Public Endtime As Date
    '    Public CellsCalculated As Long
    '    Public ReadOnly Property FERatio() As Double
    '        Get
    '            Return (QueryDurationMS - QuerySubcubeDurationMS) / QueryDurationMS
    '        End Get
    '    End Property
    '    'Public ReadOnly Property QueryDuration() As TimeSpan
    '    '    Get
    '    '        Return New TimeSpan(0, 0, 0, 0, QueryDurationMS)
    '    '    End Get
    '    'End Property
    'End Class
End Namespace