Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Microsoft.AnalysisServices.AdomdClient

Namespace Cmdlets
    <Cmdlet(VerbsLifecycle.Start, "ASProcess")> _
    Public Class cmdletInvokeASProcess
        Inherits Cmdlet

        Private mMajorObj As ProcessableMajorObject = Nothing
        <Parameter(Position:=0, Mandatory:=True)> _
        Public Property MajorOjbect() As ProcessableMajorObject
            Get
                Return mMajorObj
            End Get
            Set(ByVal value As ProcessableMajorObject)
                mMajorObj = value
            End Set
        End Property

        Private mProcessType As Microsoft.AnalysisServices.ProcessType = ProcessType.ProcessDefault
        <Parameter(Position:=1, Mandatory:=True)> _
        Public Property ProcessType() As ProcessType
            Get
                Return mProcessType
            End Get
            Set(ByVal value As ProcessType)
                mProcessType = value
            End Set
        End Property

        Private mShowProgress As SwitchParameter
        <Parameter(HelpMessage:="Returns the results of the query as a DataTable object")> _
        Public Property ShowProgress() As SwitchParameter
            Get
                Return mShowProgress
            End Get
            Set(ByVal value As SwitchParameter)
                mShowProgress = value
            End Set
        End Property

        Protected Overrides Sub ProcessRecord()
            Dim mServer As Server

            Dim mParent As MajorObject
            mParent = CType(mMajorObj.Parent, MajorObject)

            While (Not TypeOf mParent Is Server)
                mParent = CType(mParent.Parent, MajorObject)
            End While

            mServer = CType(mParent, Server)
            Dim sessionTrace As SessionTrace
            sessionTrace = mServer.SessionTrace()
            Dim oTraceEvent As TraceEventHandler = New TraceEventHandler(AddressOf OnTraceEvent)
            Dim oTraceStoppedEvent As TraceStoppedEventHandler = New TraceStoppedEventHandler(AddressOf OnTraceStoppedEvent)

            Try
                '//configure trace    
                AddHandler sessionTrace.OnEvent, oTraceEvent
                AddHandler sessionTrace.Stopped, oTraceStoppedEvent
                sessionTrace.Start()
                '// start process
                mMajorObj.Process(mProcessType)

            Finally
                '// stop trace
                sessionTrace.Stop()
                RemoveHandler sessionTrace.OnEvent, oTraceEvent
                RemoveHandler sessionTrace.Stopped, oTraceStoppedEvent
            End Try
        End Sub

        Sub OnTraceEvent(ByVal sender As Object, ByVal e As TraceEventArgs)
            Me.CommandRuntime.Host.
            Me.WriteObject(e.TextData)
            Dim p As New ProgressRecord(0, e.EventClass.ToString(), e.TextData)
            Me.WriteProgress(p)
            'Dim q As Queue
            'q.
            

        End Sub

        Sub OnTraceStoppedEvent(ByVal sender As Object, ByVal e As TraceStoppedEventArgs)
            Me.WriteObject("Trace Stopped")
            If Not e.Exception Is Nothing Then
                Me.WriteError(New ErrorRecord(e.Exception, "Processing Error", ErrorCategory.OperationStopped, sender))
            End If
        End Sub

    End Class
End Namespace