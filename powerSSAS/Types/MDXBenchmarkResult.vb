'TODO - formatting of results, should this be done with powershell format file?
Public Class MDXBenchmarkResult
    Public QueryName As String
    Public QueryDurationMS As Long
    Public QuerySubcubeDurationMS As Long
    Public QuerySubcubeCount As Long
    Public CacheHit As Long
    Public AggHit As Long
    Public Rows As Long
    Public Columns As Long
    Public StartTime As Date
    Public Endtime As Date
    Public CellsCalculated As Long
    Public NonEmptyUnoptimized As Long
    Public FlatCacheInserts As Long
    Public StablePerformanceCounters As Boolean = False
    Public ReadOnly Property FERatio() As Double
        Get
            Return (QueryDurationMS - QuerySubcubeDurationMS) / QueryDurationMS
        End Get
    End Property
    'Public ReadOnly Property QueryDuration() As TimeSpan
    '    Get
    '        Return New TimeSpan(0, 0, 0, 0, QueryDurationMS)
    '    End Get
    'End Property
End Class
