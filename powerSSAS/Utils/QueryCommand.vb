Imports Microsoft.AnalysisServices
Imports System.Management.Automation
Imports System.Threading
Imports System
Imports Microsoft.AnalysisServices.AdomdClient

Public Class QueryCommand
    Implements ICommand

    '// Private variables
    Private Delegate Function ExecuteCommand() As AdomdDataReader
    Private m_invokeMe As ExecuteCommand
    Private m_svr As Server
    Private m_cmd As AdomdClient.AdomdCommand = Nothing
    Private m_sessionTrace As SessionTrace
    Private m_oTraceEvent As TraceEventHandler
    Private m_oTraceStoppedEvent As TraceStoppedEventHandler
    Private evQueryComplete As New AutoResetEvent(False)
    Public AdomdReader As AdomdClient.AdomdDataReader
    '// Events
    Event TraceEvent(ByVal status As TraceEventArgs)
    Event Completed()

    '// Constructor
    Sub New(ByVal svr As Server, ByVal cmd As AdomdCommand)
        m_svr = svr
        m_cmd = cmd
        m_invokeMe = New ExecuteCommand(AddressOf cmd.ExecuteReader)
    End Sub


    Public Sub Execute() Implements ICommand.Execute

        m_sessionTrace = m_svr.SessionTrace()
        m_oTraceEvent = New TraceEventHandler(AddressOf OnTraceEvent)
        m_oTraceStoppedEvent = New TraceStoppedEventHandler(AddressOf OnTraceStoppedEvent)

        '//configure trace    
        AddHandler m_sessionTrace.OnEvent, m_oTraceEvent
        AddHandler m_sessionTrace.Stopped, m_oTraceStoppedEvent
        m_sessionTrace.Start()

        While Not m_sessionTrace.IsStarted
            '// loop until trace has started
            System.Threading.Thread.Sleep(200)
        End While

        m_invokeMe.BeginInvoke(AddressOf CallBack, Nothing)

    End Sub

    Private Sub CallBack(ByVal res As IAsyncResult)
        Try
            AdomdReader = m_invokeMe.EndInvoke(res)

            '// Wait upto 30 seconds for the last trace event to come in
            evQueryComplete.WaitOne(30000)

            m_sessionTrace.Stop()
            RemoveHandler m_sessionTrace.OnEvent, m_oTraceEvent
            RemoveHandler m_sessionTrace.Stopped, m_oTraceStoppedEvent
        Catch ex As Exception
        Finally
            RaiseEvent Completed()
        End Try
    End Sub

    Private Sub OnTraceEvent(ByVal sender As Object, ByVal e As TraceEventArgs)
        'Dim stat As New ProcessEvent
        'stat.CurrentTime = e.CurrentTime
        'Try
        '    stat.Duration = e.Duration
        'Catch
        'End Try
        'stat.EventClass = e.EventClass
        'stat.EventSubClass = e.EventSubclass
        'stat.Text = e.TextData


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
End Class

'Public Class ProcessEvent
'    Public CurrentTime As Date
'    Public Text As String
'    Public Duration As Long
'    Public EventClass As TraceEventClass
'    Public EventSubClass As TraceEventSubclass
'End Class
