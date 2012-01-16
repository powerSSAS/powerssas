Imports Microsoft.AnalysisServices
Imports System.Management.Automation
Imports System.Threading
Imports System
Imports Microsoft.AnalysisServices.AdomdClient

Public Class CellsetCommand
    Implements ICommand, IDisposable

    '// Private variables
    Private Delegate Function ExecuteCellset() As CellSet
    Private m_invokeCellset As ExecuteCellset
    Private Delegate Function ExecuteReader() As AdomdDataReader
    Private m_invokeReader As ExecuteReader
    Private m_svr As Server
    Private m_cmd As AdomdClient.AdomdCommand
    Private m_sessionTrace As SessionTrace
    'Private m_oTraceEvent As TraceEventHandler
    'Private m_oTraceStoppedEvent As TraceStoppedEventHandler
    Private evQueryComplete As New AutoResetEvent(False)
    Private m_FirstTraceEventRecieved As Boolean = False
    Private m_trc As Trace
    Private m_Counters As New Dictionary(Of String, PerfCounter)
    Private m_IsTabular As Boolean
    '// Public variables (really should be properties)
    Public AdomdCellset As CellSet
    Private AdomdReader As AdomdDataReader
    Public CellsCalculated As Long
    Public MemoryUsage As Long
    Public PerfCountersStable As Boolean
    Public FlatCacheInserts As Long
    Public NonEmptyUnoptimized As Long

    '// Events
    Event TraceEvent(ByVal status As TraceEventArgs)
    Event Completed()
    Event QueryStatus(ByVal desc As String)

    '// Constructor
    Sub New(ByVal svr As Server, ByVal cmd As AdomdCommand, ByVal IsTabular As Boolean)
        m_svr = svr
        m_cmd = cmd
        m_IsTabular = IsTabular
        If IsTabular Then
            m_invokeReader = New ExecuteReader(AddressOf cmd.ExecuteReader)
        Else
            m_invokeCellset = New ExecuteCellset(AddressOf cmd.ExecuteCellSet)
        End If

    End Sub


    Public Sub Execute() Implements ICommand.Execute

        RaiseEvent QueryStatus("Creating Trace")
        m_trc = CreateCustomSessionTrace(m_svr, m_cmd.Connection.SessionID)

        '        m_sessionTrace = m_svr.SessionTrace()
        'm_oTraceEvent = New TraceEventHandler(AddressOf OnTraceEvent)
        'm_oTraceStoppedEvent = New TraceStoppedEventHandler(AddressOf OnTraceStoppedEvent)

        '        m_sessionTrace = m_svr.SessionTrace()
        '        m_oTraceEvent = New TraceEventHandler(AddressOf OnTraceEvent)
        '        m_oTraceStoppedEvent = New TraceStoppedEventHandler(AddressOf OnTraceStoppedEvent)

        '//configure trace    
        'AddHandler m_sessionTrace.OnEvent, m_oTraceEvent
        'AddHandler m_sessionTrace.Stopped, m_oTraceStoppedEvent
        'm_sessionTrace.Start()

        AddHandler m_trc.OnEvent, AddressOf OnTraceEvent
        AddHandler m_trc.Stopped, AddressOf OnTraceStoppedEvent
        m_trc.Start()

        RaiseEvent QueryStatus("Waiting for Trace subscription")

        'Dim res As AdomdRestrictionCollection = CreatePerfCounterRestriction(m_svr, "MDX\Total Cells Calculated")
        'Dim dt As DataTable

        m_Counters.Add("Memory\Memory Usage KB", New PerfCounter("Memory\Memory Usage KB", m_svr, m_cmd))
        m_Counters.Add("Storage Engine Query\Total queries answered", New PerfCounter("Storage Engine Query\Total queries answered", m_svr, m_cmd))
        m_Counters.Add("MDX\Total Cells Calculated", New PerfCounter("MDX\Total Cells Calculated", m_svr, m_cmd))
        m_Counters.Add("MDX\Total flat cache inserts", New PerfCounter("MDX\Total flat cache inserts", m_svr, m_cmd))
        m_Counters.Add("MDX\Total NON EMPTY unoptimized", New PerfCounter("MDX\Total NON EMPTY unoptimized", m_svr, m_cmd))

        'm_Counters("Memory\Memory Usage KB").InitialValue = GetPerfCounter(m_svr, m_cmd, "Memory\Memory Usage KB")
        'm_Counters("Storage Engine Query\Total queries answered").InitialValue = GetPerfCounter(m_svr, m_cmd, "Storage Engine Query\Total queries answered")
        'm_Counters("MDX\Total Cells Calculated").InitialValue = GetPerfCounter(m_svr, m_cmd, "MDX\Total Cells Calculated")
        'm_Counters("MDX\Total flat cache inserts").InitialValue = GetPerfCounter(m_svr, m_cmd, "MDX\Total flat cache inserts")
        'm_Counters("MDX\Total NON EMPTY unoptimized").InitialValue = GetPerfCounter(m_svr, m_cmd, "MDX\Total NON EMPTY unoptimized")

        '// loop until trace has started and we get a DISCOVER_BEGIN event from the event handler
        While Not m_FirstTraceEventRecieved
            System.Threading.Thread.Sleep(1000)
            m_Counters("MDX\Total Cells Calculated").UpdateInitialValue() '= GetPerfCounter(m_svr, m_cmd, "MDX\Total Cells Calculated")        
        End While

        RaiseEvent QueryStatus("Running Query")
        '// Start the Async query call
        If m_IsTabular Then
            m_invokeReader.BeginInvoke(AddressOf CallBack, Nothing)
        Else
            m_invokeCellset.BeginInvoke(AddressOf CallBack, Nothing)
        End If


    End Sub

    Private Sub CallBack(ByVal res As IAsyncResult)
        ' ##############################################################
        ' WARNING - do not call RaiseEvent from this routine
        '           it executes on a different thread and powershell
        '           will throw an exception.
        ' ##############################################################
        Try
            If m_IsTabular Then
                AdomdReader = m_invokeReader.EndInvoke(res)
            Else
                AdomdCellset = m_invokeCellset.EndInvoke(res)
            End If


            '// Wait upto 30 seconds for the last trace event to come in
            evQueryComplete.WaitOne(30000)

            RemoveHandler m_trc.OnEvent, AddressOf OnTraceEvent
            RemoveHandler m_trc.Stopped, AddressOf OnTraceStoppedEvent

            m_Counters("Memory\Memory Usage KB").UpdateFinalValue() '= GetPerfCounter(m_svr, m_cmd, "Memory\Memory Usage KB")
            m_Counters("Storage Engine Query\Total queries answered").UpdateFinalValue() '= GetPerfCounter(m_svr, m_cmd, "Storage Engine Query\Total queries answered")
            m_Counters("MDX\Total Cells Calculated").UpdateFinalValue() '= GetPerfCounter(m_svr, m_cmd, "MDX\Total Cells Calculated")
            m_Counters("MDX\Total flat cache inserts").UpdateFinalValue() '= GetPerfCounter(m_svr, m_cmd, "MDX\Total flat cache inserts")
            m_Counters("MDX\Total NON EMPTY unoptimized").UpdateFinalValue() '= GetPerfCounter(m_svr, m_cmd, "MDX\Total NON EMPTY unoptimized")


            CellsCalculated = m_Counters("MDX\Total Cells Calculated").Delta
            MemoryUsage = m_Counters("Memory\Memory Usage KB").Delta
            PerfCountersStable = m_Counters("MDX\Total Cells Calculated").IsStable
            NonEmptyUnoptimized = m_Counters("MDX\Total NON EMPTY unoptimized").Delta
            FlatCacheInserts = m_Counters("MDX\Total flat cache inserts").Delta

            m_trc.Stop()
            m_trc.Drop()

        Catch ex As Exception
            Throw
        Finally
            RaiseEvent Completed()
        End Try
    End Sub

    '// Event Handlers

    Private Sub OnTraceEvent(ByVal sender As Object, ByVal e As TraceEventArgs)
        If e.EventClass = TraceEventClass.DiscoverBegin _
        Or e.EventClass = TraceEventClass.ServerStateDiscoverBegin Then
            m_FirstTraceEventRecieved = True
            Exit Sub
        End If

        RaiseEvent TraceEvent(e)
        If (e.EventClass = TraceEventClass.QueryEnd) Then
            '// signal the waiting callback function that the query has finished
            evQueryComplete.Set()
        End If
    End Sub

    Private Sub OnTraceStoppedEvent(ByVal sender As Object, ByVal e As TraceStoppedEventArgs)
        RaiseEvent Completed()
        If Not e.Exception Is Nothing Then
            Throw e.Exception
        End If
    End Sub

    '// Private Methods
    Private Function CreateCustomSessionTrace(ByVal svr As Server, ByVal sessionId As String) As Trace
        Dim trc As Trace = svr.Traces.Add("PowerSSASBenchmark_" + Guid.NewGuid.ToString())
        trc.Events.Add(GetTraceEvent(TraceEventClass.QueryBegin)) '9
        trc.Events.Add(GetTraceEvent(TraceEventClass.QueryEnd)) '10
        trc.Events.Add(GetTraceEvent(TraceEventClass.QuerySubcube)) '11
        trc.Events.Add(GetTraceEvent(TraceEventClass.GetDataFromAggregation))  '60
        trc.Events.Add(GetTraceEvent(TraceEventClass.GetDataFromCache)) '61
        trc.Events.Add(GetTraceEvent(TraceEventClass.Error)) '17
        trc.Events.Add(GetTraceEvent(TraceEventClass.ProgressReportEnd)) '6
        'trc.Events.Add(GetTraceEvent(TraceEventClass.DiscoverBegin))
        trc.Events.Add(GetTraceEvent(TraceEventClass.ServerStateDiscoverBegin))
        trc.Events.Add(GetTraceEvent(TraceEventClass.CommandEnd))
        Dim xDoc As New Xml.XmlDocument()
        'Dim nsm As Xml.XmlNamespaceManager = New Xml.XmlNamespaceManager(xDoc.NameTable)
        'nsm.AddNamespace("", "http://schemas.microsoft.com/analysisservices/2003/engine")
        '// column 39 = SessionID
        xDoc.LoadXml("<Equal xmlns=""http://schemas.microsoft.com/analysisservices/2003/engine""><ColumnID>39</ColumnID><Value>" & sessionId & "</Value></Equal>")
        trc.Filter = xDoc

        'Dim sb As New Text.StringBuilder()
        'Dim tw As New IO.StringWriter(sb)
        'Dim xw As New Xml.XmlTextWriter(tw)
        'Microsoft.AnalysisServices.Scripter.WriteAlter(xw, trc, True, True)
        'Diagnostics.Debug.WriteLine(sb.ToString())

        trc.Update()

        Return trc
    End Function

    Private Function GetTraceEvent(ByVal eventClass As TraceEventClass) As TraceEvent
        Dim te As New TraceEvent(eventClass)

        te.Columns.Add(TraceColumn.EventClass) '0
        te.Columns.Add(TraceColumn.SessionID)           '39
        te.Columns.Add(TraceColumn.TextData) '42

        If eventClass <> TraceEventClass.Error _
        And eventClass <> TraceEventClass.GetDataFromAggregation Then
            te.Columns.Add(TraceColumn.EventSubclass)  '1
        End If

        If eventClass <> TraceEventClass.Error _
        And eventClass <> TraceEventClass.QueryBegin _
        And eventClass <> TraceEventClass.ServerStateDiscoverBegin _
        And eventClass <> TraceEventClass.DiscoverBegin Then
            te.Columns.Add(TraceColumn.CpuTime)    '6
            te.Columns.Add(TraceColumn.Duration) '5
        End If

        If eventClass <> TraceEventClass.Error Then
            te.Columns.Add(TraceColumn.CurrentTime)  '2
        End If

        If eventClass = TraceEventClass.QueryBegin Then
            te.Columns.Add(TraceColumn.RequestParameters)   '44
            te.Columns.Add(TraceColumn.RequestProperties)   '45
        End If

        Return te
    End Function

    'Private Function CreatePerfCounterRestriction(ByVal svr As Server, ByVal counterName As String) As AdomdRestrictionCollection
    '    Dim res As New AdomdRestrictionCollection()
    '    Dim cntrPrefix As String = "\MSAS 2008"
    '    If m_svr.Name.Split("\".ToCharArray()(0)).Length > 1 Then
    '        cntrPrefix = "\MSOLAP$" + svr.Name.Split("\".ToCharArray()(0))(1)
    '    End If
    '    res.Add("PERF_COUNTER_NAME", cntrPrefix + ":" + counterName)
    '    Return res
    'End Function

    'Private Function GetPerfCounter(ByVal svr As Server, ByVal cmd As AdomdCommand, ByVal counterName As String) As Integer
    '    Return CType(cmd.Connection.GetSchemaDataSet("DISCOVER_PERFORMANCE_COUNTERS", CreatePerfCounterRestriction(svr, counterName)).Tables(0).Rows(0).Item("PERF_COUNTER_VALUE"), Integer)
    'End Function

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                'free other state (managed objects).
                evQueryComplete.Close()
                evQueryComplete = Nothing
            End If

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

