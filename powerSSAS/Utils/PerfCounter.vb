Public Class PerfCounter
    Private m_initValue As Long
    Private m_finalValue As Long
    Private m_stable As Boolean = True
    Private m_name As String

    Public Sub New(ByVal initVal As Long)
        m_initValue = initVal
    End Sub

    Public Property InitialValue() As Long
        Get
            Return m_initValue
        End Get
        Set(ByVal value As Long)
            If m_initValue <> 0 And m_initValue <> value Then
                m_stable = False
            End If
            m_initValue = value
        End Set
    End Property

    Public Property FinalValue() As Long
        Get
            Return m_finalValue
        End Get
        Set(ByVal value As Long)
            m_finalValue = value
        End Set
    End Property

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
End Class
