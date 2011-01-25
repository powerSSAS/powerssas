

Public Module PowerSSASConstants
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1707:IdentifiersShouldNotContainUnderscores", MessageId:="Member")> Public Const TYPE_DATA_FOLDER As String = "TypeData\"
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1707:IdentifiersShouldNotContainUnderscores", MessageId:="Member")> Public Const FORMAT_DATA_FOLDER As String = "FormatData\"
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1707:IdentifiersShouldNotContainUnderscores", MessageId:="Member")> Public Const HELP_FOLDER As String = "Help\"


End Module

Public NotInheritable Class PowerSSASNouns
    Private Sub New()
    End Sub

    Public Shared Function ASSession() As String
        Return "ASSession"
    End Function
    Public Shared Function ASConnection() As String
        Return "ASConnection"
    End Function
End Class

<Flags()> _
Public Enum LockTypes
    LOCK_NONE = &H0               ' No lock.
    LOCK_SESSION_LOCK = &H1       ' Inactive session; does not interfere with other locks.
    LOCK_READ = &H2               ' Read lock during processing.
    LOCK_WRITE = &H4              ' Write lock during processing.
    LOCK_COMMIT_READ = &H8        ' Commit lock, shared.
    LOCK_COMMIT_WRITE = &H10      ' Commit lock, exclusive.
    LOCK_COMMIT_ABORTABLE = &H20  ' Abort at commit progress.
    LOCK_COMMIT_INPROGRESS = &H40 ' Commit in progress.
    LOCK_INVALID = &H80           ' Invalid lock.
End Enum
