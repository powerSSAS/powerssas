Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Gosbell.PowerSSAS.PowerSSASProvider
Imports System.Xml
Imports System.Globalization

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASLock")> _
    Public Class CmdletGetASLock
        Inherits CmdletDiscoverBase

        Private mLockTypeFilter As Integer
        Private mLockFilterSet As Boolean = False

        Private mReadLockSwitch As SwitchParameter
        <Parameter(HelpMessage:="Read Lock during processing")> _
        Public Property READ_LOCK() As SwitchParameter
            Get
                Return mReadLockSwitch
            End Get
            Set(ByVal value As SwitchParameter)
                mReadLockSwitch = value
                mLockTypeFilter = mLockTypeFilter Xor LockTypes.LOCK_READ
            End Set
        End Property

        Private mGrantedStatus As SwitchParameter
        <Parameter(HelpMessage:="Selects Locks with a status of GRANTED", Mandatory:=False)> _
        Public Property GRANTED() As SwitchParameter
            Get
                Return mGrantedStatus
            End Get
            Set(ByVal value As SwitchParameter)
                mGrantedStatus = value
            End Set
        End Property

        Private mWaitingStatus As SwitchParameter
        <Parameter(HelpMessage:="Selectes Locks with a status of WAITING", Mandatory:=False)> _
        Public Property WAITING() As SwitchParameter
            Get
                Return mWaitingStatus
            End Get
            Set(ByVal value As SwitchParameter)
                mWaitingStatus = value
            End Set
        End Property

        Private mWriteLockSwitch As SwitchParameter
        <Parameter(HelpMessage:="Read Lock during processing")> _
        Public Property WRITE_LOCK() As SwitchParameter
            Get
                Return mWriteLockSwitch
            End Get
            Set(ByVal value As SwitchParameter)
                mWriteLockSwitch = value
                mLockTypeFilter = mLockTypeFilter Xor LockTypes.LOCK_WRITE
                mLockFilterSet = True
            End Set
        End Property

        Private mSessionLockSwitch As SwitchParameter
        <Parameter(HelpMessage:="Session idle Lock")> _
        Public Property SESSION_LOCK() As SwitchParameter
            Get
                Return mSessionLockSwitch
            End Get
            Set(ByVal value As SwitchParameter)
                mSessionLockSwitch = value
                mLockTypeFilter = mLockTypeFilter Xor LockTypes.LOCK_SESSION_LOCK
            End Set
        End Property

        'TODO - other restrictions to code
        'LOCK_OBJECT_ID
        'LOCK_TYPE (partially done)
        '

        Private mMinimumMilliseconds As Integer
        <Parameter(HelpMessage:="Only return locks that have been granted for the specified minimum milliseconds")> _
        Public Property MinimumMilliseconds() As Integer
            Get
                Return mMinimumMilliseconds
            End Get
            Set(ByVal value As Integer)
                mMinimumMilliseconds = value
            End Set
        End Property

        Private mSPID As Integer
        <Parameter(HelpMessage:="Only return locks for the specified SPID")> _
        Public Property SPID() As Integer
            Get
                Return mSPID
            End Get
            Set(ByVal value As Integer)
                mSPID = value
            End Set
        End Property

        Private mTransactionID As String = ""
        <Parameter(HelpMessage:="Only return locks that have been granted as part of the specified Transaction")> _
        Public Property TransactionID() As String
            Get
                Return mTransactionID
            End Get
            Set(ByVal value As String)
                mTransactionID = value
            End Set
        End Property


        Protected Overrides Sub ProcessRecord()
            Dim xLock As New Utils.XmlaDiscoverLocks
            If mLockFilterSet And XmlaRestrictions.Find("LOCK_TYPE") Is Nothing Then
                XmlaRestrictions.Add("LOCK_TYPE", mLockTypeFilter.ToString())
            End If
            If XmlaRestrictions.Find("<LOCK_STATUS>") Is Nothing Then
                If mGrantedStatus.IsPresent And Not mWaitingStatus.IsPresent Then
                    XmlaRestrictions.Add("LOCK_STATUS", 1)
                ElseIf mWaitingStatus.IsPresent And Not mGrantedStatus.IsPresent Then
                    XmlaRestrictions.Add("LOCK_STATUS", 0)
                End If
            End If
            If mMinimumMilliseconds > 0 And XmlaRestrictions.Find("LOCK_MIN_TOTAL_MS") Is Nothing Then
                XmlaRestrictions.Add("LOCK_MIN_TOTAL_MS", mMinimumMilliseconds.ToString())
            End If
            If mSPID > 0 And XmlaRestrictions.Find("SPID") Is Nothing Then
                XmlaRestrictions.Add("SPID", mSPID.ToString())
            End If
            If mTransactionID.Length > 0 And XmlaRestrictions.Find("TRANSACTION_ID") Is Nothing Then
                XmlaRestrictions.Add("TRANSACTION_ID", mTransactionID)
            End If
            xLock.Discover(Me.DiscoverRowset, ServerName, XmlaRestrictions, "", AddressOf Me.OutputObject)
        End Sub


        Public Overrides ReadOnly Property DiscoverRowset() As String
            Get
                Return "DISCOVER_LOCKS"
            End Get
        End Property
    End Class
End Namespace