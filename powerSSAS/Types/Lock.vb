Imports System.Xml
Imports Microsoft.AnalysisServices.Xmla.XmlaClient
Imports Microsoft.AnalysisServices

Namespace Types
    ''' <summary>
    ''' This class is the strongly typed class for a powerSSAS lock.
    ''' The main feature this gives us is the get the actual object that is locked.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Lock

        Private mSpid As Integer
        Private mLockId As Guid
        Private mTransactionId As Guid
        Private mObjectId As XmlNode
        Private mLockStatus As Integer
        Private mLockType As Integer
        Private mCreationTime As Date
        Private mGrantTime As Date
        Private mMajorObject As IMajorObject

        Public ReadOnly Property Id() As Guid
            Get
                Return mLockId
            End Get
        End Property

        Public ReadOnly Property Spid() As Integer
            Get
                Return mSpid
            End Get
        End Property

        Public ReadOnly Property TransactionId() As Guid
            Get
                Return mTransactionId
            End Get
        End Property

        Public ReadOnly Property ObjectId() As String
            Get
                Return mObjectId.FirstChild.InnerXml
            End Get
        End Property

        Public ReadOnly Property LockStatus() As Integer
            Get
                Return mLockStatus
            End Get
        End Property

        Public ReadOnly Property LockType() As Integer
            Get
                Return mLockType
            End Get
        End Property

        Public ReadOnly Property CreationTime() As Date
            Get
                Return mCreationTime
            End Get
        End Property

        Public ReadOnly Property GrantTime() As Date
            Get
                Return mGrantTime
            End Get
        End Property

        Private mServerName As String
        Public ReadOnly Property ServerName() As String
            Get
                Return mServerName
            End Get
        End Property

        Public ReadOnly Property ParentServer() As Server
            Get
                If Not mMajorObject Is Nothing Then
                    Return mMajorObject.ParentServer
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public ReadOnly Property ParentDatabase() As Database
            Get
                If Not mMajorObject Is Nothing Then
                    Return mMajorObject.ParentDatabase
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Sub New(ByVal serverName As String, ByVal row As XmlNode)
            mServerName = serverName
            If Not row Is Nothing Then
                For Each n As XmlNode In row.ChildNodes
                    Select Case n.LocalName
                        Case "LOCK_ID"
                            mLockId = New Guid(n.InnerText)
                        Case "SPID"
                            mSpid = Integer.Parse(n.InnerText)
                        Case "LOCK_TRANSACTION_ID"
                            mTransactionId = New Guid(n.InnerText)
                        Case "LOCK_OBJECT_ID"
                            mObjectId = n
                            mLockObjectType = n.LastChild.LastChild.LocalName.Substring(0, n.LastChild.LastChild.LocalName.Length - 2)
                            mLockObjectName = n.LastChild.LastChild.InnerText
                        Case "LOCK_STATUS"
                            mLockStatus = Integer.Parse(n.InnerText)
                        Case "LOCK_TYPE"
                            mLockType = Integer.Parse(n.InnerText)
                        Case "LOCK_CREATION_TIME"
                            mCreationTime = Date.Parse(n.InnerText)
                        Case "LOCK_GRANT_TIME"
                            mGrantTime = Date.Parse(n.InnerText)
                    End Select
                Next
            End If
        End Sub

        Public ReadOnly Property LockTypeDescription() As String
            Get
                Dim mList As New List(Of String)
                If (mLockType And LockTypes.LOCK_WRITE) = LockTypes.LOCK_WRITE Then Return "WRITE"
                If (mLockType And LockTypes.LOCK_SESSION_LOCK) = LockTypes.LOCK_SESSION_LOCK Then Return "SESSION"
                If (mLockType And LockTypes.LOCK_READ) = LockTypes.LOCK_READ Then Return "READ"

                If (mLockType And LockTypes.LOCK_INVALID) = LockTypes.LOCK_INVALID Then Return "INVALID"
                If (mLockType And LockTypes.LOCK_COMMIT_WRITE) = LockTypes.LOCK_COMMIT_WRITE Then mList.Add("COMMIT_WRITE")
                If (mLockType And LockTypes.LOCK_COMMIT_READ) = LockTypes.LOCK_COMMIT_READ Then mList.Add("COMMIT_READ")
                If (mLockType And LockTypes.LOCK_COMMIT_INPROGRESS) = LockTypes.LOCK_COMMIT_INPROGRESS Then mList.Add("COMMIT_INPROGRESS")
                If (mLockType And LockTypes.LOCK_COMMIT_ABORTABLE) = LockTypes.LOCK_COMMIT_ABORTABLE Then Return "COMMIT_ABORTABLE"
                If (mLockType And LockTypes.LOCK_NONE) = LockTypes.LOCK_NONE Then Return "NONE"
                Return ""
            End Get
        End Property

        Public ReadOnly Property LockStatusDescription() As String
            Get
                If mLockStatus = 0 Then
                    Return "WAITING"
                Else
                    Return "GRANTED"
                End If
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return Me.Id.ToString()
        End Function

        Public Function GetLockObject() As Microsoft.AnalysisServices.IMajorObject
            Return MajorObjectFromIdPath(mObjectId)
        End Function

        Private mLockObjectType As String = ""
        Public ReadOnly Property LockObjectType() As String
            Get
                Return mLockObjectType
            End Get
        End Property

        Private mLockObjectName As String = ""
        Public ReadOnly Property LockObjectName() As String
            Get
                Return mLockObjectName
            End Get
        End Property

        Private Function MajorObjectFromIdPath(ByVal idPath As XmlNode) As IMajorObject
            Dim svr As Microsoft.AnalysisServices.Server = ConnectionFactory.ConnectToServer(mServerName)
            Dim doc As XmlDocument = New XmlDocument()
            Dim majObj As IMajorObject
            doc.LoadXml(idPath.InnerXml)
            '// Need to change the default namespace so that the reference can be resolved correctly
            '// TODO - need to check if there is a more efficient way than using an XmlDocument
            doc.DocumentElement.SetAttribute("xmlns", "http://schemas.microsoft.com/analysisservices/2003/engine")
            majObj = ObjectReference.ResolveObjectReference(svr, ObjectReference.Deserialize(doc.InnerXml, True))


            Return majObj
        End Function

    End Class
End Namespace