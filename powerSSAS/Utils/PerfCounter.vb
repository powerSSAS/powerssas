Imports Microsoft.AnalysisServices
Imports Microsoft.AnalysisServices.AdomdClient

Public Class PerfCounter
    Private m_initValue As Long
    Private m_finalValue As Long
    Private m_stable As Boolean = True
    Private m_name As String
    Private m_svr As Server
    Private m_cmd As AdomdCommand

    Public Sub New(ByVal name As String, ByVal svr As Server, ByVal cmd As AdomdCommand)
        m_cmd = cmd
        m_svr = svr
        m_name = name
        m_initValue = 0
        m_initValue = GetPerfCounterValue()
    End Sub

    Public Property InitialValue() As Long
        Get
            Return m_initValue
        End Get
        Friend Set(ByVal value As Long)
            If m_initValue <> 0 And m_initValue <> value Then
                m_stable = False
            End If
            m_initValue = value
        End Set
    End Property

    Public ReadOnly Property FinalValue() As Long
        Get
            Return m_finalValue
        End Get
        'Set(ByVal value As Long)
        '    m_finalValue = value
        'End Set
    End Property

    Public Sub UpdateFinalValue()
        m_finalValue = GetPerfCounterValue()
    End Sub

    Public Sub UpdateInitialValue()
        InitialValue = GetPerfCounterValue()
    End Sub

    Public ReadOnly Property Delta() As Long
        Get
            Return m_finalValue - m_initValue
        End Get
    End Property

    Public Property CounterName() As String
        Get
            Return m_name
        End Get
        Set(ByVal value As String)
            m_name = value
        End Set
    End Property

    Public ReadOnly Property IsStable() As Boolean
        Get
            Return m_stable
        End Get
    End Property

    Private Function GetPerfCounterValue() As Long '(ByVal svr As Server, ByVal cmd As AdomdCommand, ByVal counterName As String) As Integer
        Return CType(m_cmd.Connection.GetSchemaDataSet("DISCOVER_PERFORMANCE_COUNTERS", CreatePerfCounterRestriction(m_svr, m_name)).Tables(0).Rows(0).Item("PERF_COUNTER_VALUE"), Integer)
    End Function

    Private Function CreatePerfCounterRestriction(ByVal svr As Server, ByVal counterName As String) As AdomdRestrictionCollection
        Dim res As New AdomdRestrictionCollection()
        Dim cntrPrefix As String = "\MSAS 2008"
        If m_svr.Name.Split("\".ToCharArray()(0)).Length > 1 Then
            cntrPrefix = "\MSOLAP$" + svr.Name.Split("\".ToCharArray()(0))(1)
        End If
        res.Add("PERF_COUNTER_NAME", cntrPrefix + ":" + counterName)
        Return res
    End Function
End Class
