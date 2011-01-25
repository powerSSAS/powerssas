Imports Microsoft.AnalysisServices
Imports System.Management.Automation
Imports System.Threading
Imports System

Public Class ProcessCommand
    Implements ICommand

    '// Private variables
    Private Delegate Sub ProcessObject(ByVal pt As ProcessType)
    Private m_invokeMe As ProcessObject
    Private m_processType As ProcessType
    Private m_obj As ProcessableMajorObject
    Private m_sessionTrace As SessionTrace
    Private m_oTraceEvent As TraceEventHandler
    Private m_oTraceStoppedEvent As TraceStoppedEventHandler
    'Private m_cmdlet As Cmdlet
    Private evProcessComplete As New AutoResetEvent(False)

    '// Events
    Event StatusUpdate(ByVal status As ProcessEvent)
    Event Completed()

    '// Constructor
    Sub New(ByVal obj As ProcessableMajorObject, ByVal pt As ProcessType)
        'm_cmdlet = thisCmdlet
        m_obj = obj
        m_processType = pt
        m_invokeme = New ProcessObject(AddressOf obj.Process)
    End Sub


    Public Sub Execute() Implements ICommand.Execute
        Dim mServer As Server

        Dim mParent As MajorObject
        mParent = CType(m_obj.Parent, MajorObject)

        While (Not TypeOf mParent Is Server)
            mParent = CType(mParent.Parent, MajorObject)
        End While

        mServer = CType(mParent, Server)

        m_sessionTrace = mServer.SessionTrace()
        m_oTraceEvent = New TraceEventHandler(AddressOf OnTraceEvent)
        m_oTraceStoppedEvent = New TraceStoppedEventHandler(AddressOf OnTraceStoppedEvent)

        '//configure trace    
        AddHandler m_sessionTrace.OnEvent, m_oTraceEvent
        AddHandler m_sessionTrace.Stopped, m_oTraceStoppedEvent
        m_sessionTrace.Start()

        m_invokeme.BeginInvoke(m_processType, AddressOf CallBack, Nothing)

    End Sub

    Private Sub CallBack(ByVal res As IAsyncResult)
        Try
            m_invokeme.EndInvoke(res)

            '// Wait 30 seconds for the last trace event to come in
            evProcessComplete.WaitOne(30000)

            m_sessionTrace.Stop()
            'Threading.Thread.Sleep(100)
            RemoveHandler m_sessionTrace.OnEvent, m_oTraceEvent
            RemoveHandler m_sessionTrace.Stopped, m_oTraceStoppedEvent
        Catch ex As Exception
        Finally
            RaiseEvent Completed()
        End Try
    End Sub

    Sub OnTraceEvent(ByVal sender As Object, ByVal e As TraceEventArgs)
        Dim stat As New ProcessEvent
        stat.CurrentTime = e.CurrentTime
        Try
            stat.Duration = e.Duration
        Catch
        End Try
        stat.EventClass = e.EventClass
        stat.EventSubClass = e.EventSubclass
        stat.Text = e.TextData

        RaiseEvent StatusUpdate(stat)
        If (e.EventClass = TraceEventClass.CommandEnd And e.EventSubclass = TraceEventSubclass.Process) Then
            '// signal the waiting callback function that the processing has finished
            evProcessComplete.Set()
            'RaiseEvent Completed()
        End If
    End Sub

    Sub OnTraceStoppedEvent(ByVal sender As Object, ByVal e As TraceStoppedEventArgs)
        RaiseEvent Completed()
        'Me.WriteObject("Trace Stopped")
        If Not e.Exception Is Nothing Then
            'Me.WriteError(New ErrorRecord(e.Exception, "Processing Error", ErrorCategory.OperationStopped, sender))
            Throw e.Exception
        End If
    End Sub
End Class

Public Class ProcessEvent
    Public CurrentTime As Date
    Public Text As String
    Public Duration As Long
    Public EventClass As TraceEventClass
    Public EventSubClass As TraceEventSubclass
End Class
