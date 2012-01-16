Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Microsoft.AnalysisServices.AdomdClient
Imports System.Threading
Imports System

Namespace Cmdlets
    <Cmdlet(VerbsLifecycle.Invoke, "ASProcess")> _
    Public Class cmdletInvokeASProcess
        Inherits Cmdlet
        Private evStat As New AutoResetEvent(False) 'ManualResetEvent(False)
        Private evExit As New AutoResetEvent(False)
        Private evts As WaitHandle()
        Private pr_sync As New Object()
        Private m_output As ProcessEvent

        Private mMajorObj As ProcessableMajorObject = Nothing
        <Parameter(Position:=0, Mandatory:=True, ValueFromPipeline:=True)> _
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
            ReDim evts(1)
            evts(0) = evExit
            evts(1) = evStat

            Dim cmd As New ProcessCommand(MajorOjbect, Me.ProcessType)
            AddHandler cmd.StatusUpdate, AddressOf OnStatusUpdate
            AddHandler cmd.Completed, AddressOf OnCompleted

            '// This command launches the Async processing in another thread
            cmd.Execute()

            '// This loop waits until one of the waithandles are set
            '// it the exit handle is set (index = 0) then the loop exits
            '// otherwise the event output is written to the pipeline
            While WaitHandle.WaitAny(evts) <> 0
                SyncLock pr_sync
                    WriteObject(m_output)
                End SyncLock
            End While
        End Sub

        Private Sub OnStatusUpdate(ByVal stat As ProcessEvent)
            SyncLock pr_sync
                m_output = stat
                '// signal that a status has been recieved
                evStat.Set()
            End SyncLock
        End Sub

        Private Sub OnCompleted()
            '// signal the exit waithandle
            evExit.Set()
        End Sub
    End Class
End Namespace