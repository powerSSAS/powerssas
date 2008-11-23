Imports System.Xml
Imports Microsoft.AnalysisServices.Xmla.XmlaClient

Namespace Types
    ''' <summary>
    ''' This class is the strongly typed class for a powerSSAS session.
    ''' The main feature this gives us is the ability to call the kill()
    ''' method on a given session.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Session

        Private mId As String
        Private mSpid As Integer
        Private mConnectionId As Integer
        Private mUserName As String
        Private mCurrentDatabase As String
        Private mStartTime As Date
        Private mElapsedTimeMs As Integer
        Private mLastCommandStartTime As Date
        Private mLastCommandEndTime As Date
        Private mLastCommandElapsedTimeMs As Integer
        Private mIdleTimeMs As Integer
        Private mCpuTimeMs As Integer
        Private mLastCommand As String
        Private mLastCommandCpuTimeMs As Integer

        Public ReadOnly Property Id() As String
            Get
                Return mId
            End Get
        End Property

        Public ReadOnly Property Spid() As Integer
            Get
                Return mSpid
            End Get
        End Property

        Public ReadOnly Property ConnectionId() As Integer
            Get
                Return mConnectionId
            End Get
        End Property

        Public ReadOnly Property UserName() As String
            Get
                Return mUserName
            End Get
        End Property

        Public ReadOnly Property CurrentDatabase() As String
            Get
                Return mCurrentDatabase
            End Get
        End Property

        Public ReadOnly Property StartTime() As Date
            Get
                Return mStartTime
            End Get
        End Property

        Public ReadOnly Property ElapsedTimeMS() As Integer
            Get
                Return mElapsedTimeMs
            End Get
        End Property

        Public ReadOnly Property LastCommandStartTime() As Date
            Get
                Return mLastCommandStartTime
            End Get
        End Property

        Public ReadOnly Property LastCommandEndTime() As Date
            Get
                Return mLastCommandEndTime
            End Get
        End Property

        Public ReadOnly Property LastCommandElapsedTimeMS() As Integer
            Get
                Return mLastCommandElapsedTimeMs
            End Get
        End Property

        Public ReadOnly Property IdleTimeMS() As Integer
            Get
                Return mIdleTimeMs
            End Get
        End Property

        Public ReadOnly Property CpuTimeMS() As Integer
            Get
                Return mCpuTimeMs
            End Get
        End Property

        Public ReadOnly Property LastCommand() As String
            Get
                Return mLastCommand
            End Get
        End Property
        Public ReadOnly Property LastCommandCpuTimeMS() As Integer
            Get
                Return mLastCommandCpuTimeMs
            End Get
        End Property

        Private mServerName As String
        Public ReadOnly Property ServerName() As String
            Get
                Return mServerName
            End Get
        End Property

        Sub New(ByVal serverName As String, ByVal row As XmlNode)
            mserverName = serverName
            If Not row Is Nothing Then
                For Each n As XmlNode In row.ChildNodes

                    Select Case n.LocalName
                        Case "SESSION_ID"
                            mId = n.InnerText
                        Case "SESSION_SPID"
                            mSpid = Integer.Parse(n.InnerText)
                        Case "SESSION_CONNECTION_ID"
                            mConnectionId = Integer.Parse(n.InnerText)
                        Case "SESSION_USER_NAME"
                            mUserName = n.InnerText
                        Case "SESSION_CURRENT_DATABASE"
                            mCurrentDatabase = n.InnerText
                        Case "SESSION_START_TIME"
                            mStartTime = Date.Parse(n.InnerText)
                        Case "SESSION_ELAPSED_TIME_MS"
                            mElapsedTimeMs = Integer.Parse(n.InnerText)
                        Case "SESSION_LAST_COMMAND_START_TIME"
                            mLastCommandStartTime = Date.Parse(n.InnerText)
                        Case "SESSION_LAST_COMMAND_END_TIME"
                            mLastCommandEndTime = Date.Parse(n.InnerText)
                        Case "SESSION_LAST_COMMAND_ELAPSED_TIME_MS"
                            mLastCommandElapsedTimeMs = Integer.Parse(n.InnerText)
                        Case "SESSION_IDLE_TIME_MS"
                            mIdleTimeMs = Integer.Parse(n.InnerText)
                        Case "SESSION_CPU_TIME_MS"
                            mCpuTimeMs = Integer.Parse(n.InnerText)
                        Case "SESSION_LAST_COMMAND"
                            mLastCommand = n.InnerText
                        Case "SESSION_LAST_COMMAND_CPU_TIME_MS"
                            mLastCommandCpuTimeMs = Integer.Parse(n.InnerText)
                    End Select
                Next
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return Me.Id
        End Function

        ''' <summary>
        ''' This method will cancel the current session.
        ''' </summary>
        ''' <remarks>It delegates the actual cancelling to the KillSession cmdlet.</remarks>
        <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId:="o")> Sub Kill()
            Dim cmdletKill As New Cmdlets.CmdletClearSession
            'cmdletKill.ServerName = Me.serverName
            'cmdletKill.SessionID = Me.Id
            Dim s As Session() = {Me}
            cmdletKill.InputObject = s

            '// Need to loop through the Invoke enumerator to execute the 
            '// kill session cmdlet
            For Each o As Object In cmdletKill.Invoke()

            Next


            'Dim cancelCmd As String = "<Cancel xmlns=""http://schemas.microsoft.com/analysisservices/2003/engine""><SessionID>"
            'cancelCmd &= Me.Id & "</SessionID></Cancel> "

            'Console.WriteLine("Cancelling SessionID: " & Me.Id)

            'Dim xc As New Microsoft.AnalysisServices.Xmla.XmlaClient
            'xc.Connect("Data Source=" & serverName)
            'Dim res As String = ""
            'Try
            '    res = xc.Send(cancelCmd, Nothing)
            'Finally
            '    If Not xc Is Nothing Then
            '        xc.Disconnect()
            '    End If
            'End Try

            ''// check for errors
            'Dim tr As New IO.StringReader(res)
            'Dim xmlRdr As New XmlTextReader(tr)
            'Dim msg As Object = xmlRdr.NameTable.Add("Error")
            'While xmlRdr.Read
            '    If xmlRdr.NodeType = XmlNodeType.Element AndAlso Object.ReferenceEquals(xmlRdr.LocalName, msg) Then
            '        If xmlRdr.MoveToAttribute("Description") Then
            '            Throw New ArgumentException(xmlRdr.Value)
            '        End If
            '    End If
            'End While

        End Sub
    End Class
End Namespace
