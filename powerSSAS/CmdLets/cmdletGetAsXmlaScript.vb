Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Gosbell.PowerSSAS.PowerSSASProvider

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASXmlaScript")> _
    Public Class cmdletGetAsXmlaScript
        Inherits PSCmdlet

        Private mMajorObjectPath As String()
        <Parameter(HelpMessage:="An Analysis Services Major Object" _
            , Mandatory:=False, Position:=0, ValueFromPipeline:=False _
            , ParameterSetName:="byPath")> _
        Public Property MajorObjectPath() As String()
            Get
                Return mMajorObjectPath
            End Get
            Set(ByVal value As String())
                mMajorObjectPath = value
            End Set
        End Property

        Private mCreate As SwitchParameter
        <Parameter(Mandatory:=False)> _
        Public Property Create() As SwitchParameter
            Get
                Return mCreate
            End Get
            Set(ByVal value As SwitchParameter)
                mCreate = value
            End Set
        End Property

        Private mParentObject As MajorObject
        <Parameter(Mandatory:=False, HelpMessage:="This parameter is only used when scripting a create operation.", ParameterSetName:="byObject", Position:=1)> _
        Public Property ParentObject() As MajorObject
            Get
                Return mParentObject
            End Get
            Set(ByVal value As MajorObject)
                mParentObject = value
            End Set
        End Property

        Private mInputObj As MajorObject
        <Parameter(mandatory:=False, valueFromPipeline:=True, ParameterSetName:="byObject", Position:=0)> _
        Public Property InputObject() As MajorObject
            Get
                Return mInputObj
            End Get
            Set(ByVal value As MajorObject)
                mInputObj = value
            End Set
        End Property

        Private mExcludeDependents As SwitchParameter
        <Parameter(Mandatory:=False)> _
        Public Property ExcludeDependents() As SwitchParameter
            Get
                Return mExcludeDependents
            End Get
            Set(ByVal value As SwitchParameter)
                mExcludeDependents = value
            End Set
        End Property

        'TODO: only valid for database, cube, dimension, cubedimension
        Private mExcludePermissions As SwitchParameter
        <Parameter(Mandatory:=False)> _
        Public Property ExcludePermissions() As SwitchParameter
            Get
                Return mExcludePermissions
            End Get
            Set(ByVal value As SwitchParameter)
                mExcludePermissions = value
            End Set
        End Property

        'TODO: only valid for database, cube, measuregroup
        Private mExcludePartitions As SwitchParameter
        <Parameter(Mandatory:=False)> _
        Public Property ExcludePartitions() As SwitchParameter
            Get
                Return mExcludePartitions
            End Get
            Set(ByVal value As SwitchParameter)
                mExcludePartitions = value
            End Set
        End Property

        Protected Overrides Sub EndProcessing()
            MyBase.EndProcessing()
        End Sub

        Protected Overrides Sub ProcessRecord()
            MyBase.ProcessRecord()
            Dim majObj As MajorObject = Nothing

            If (mCreate.IsPresent AndAlso mParentObject Is Nothing) _
            OrElse (Not mCreate.IsPresent AndAlso Not mParentObject Is Nothing) Then
                Throw New ArgumentException("You specify both the Create and ParentObject options.")
                Return
            End If

            Select Case Me.ParameterSetName
                Case "byPath"
                    WriteDebug("Path parameter passed in")
                    Dim psDrive As PSDriveInfo = Me.SessionState.Path.CurrentLocation.Drive
                    If TypeOf psDrive Is AmoDriveInfo Then
                        Me.WriteObject(Me.InvokeProvider.Item)
                        Dim p As PSObject
                        Dim itmCol As System.Collections.ObjectModel.Collection(Of PSObject)
                        itmCol = Me.InvokeProvider.Item.Get(Me.SessionState.Path.CurrentLocation.Path)

                        If itmCol.Count = 1 Then
                            p = itmCol.Item(0)
                            If TypeOf p.BaseObject Is Microsoft.AnalysisServices.MajorObject Then
                                majObj = CType(p.BaseObject, MajorObject)
                            End If
                        Else
                            Me.WriteWarning(String.Format("There were {0} objects found at {1}.", itmCol.Count, Me.SessionState.Path.CurrentLocation.Path))
                        End If
                    Else
                        Me.WriteError(New ErrorRecord(New NotSupportedException("You can only create an XMLA script from an AMO location"), "Invalid Location", ErrorCategory.InvalidOperation, psDrive))
                    End If

                Case "byObject"
                    WriteDebug("Object parameter passed in")
                    majObj = mInputObj
            End Select

            If Not majObj Is Nothing Then
                Dim sbXml As New Text.StringBuilder()
                Dim xws As New Xml.XmlWriterSettings()
                With xws
                    .Indent = True
                    .NewLineHandling = Xml.NewLineHandling.Entitize
                    .OmitXmlDeclaration = True
                End With
                Dim outXml As Xml.XmlWriter = Xml.XmlWriter.Create(sbXml, xws)
                Dim svr As Server = ConnectionFactory.GetServerFromObject(majObj)
                svr.BeginTransaction()

                Try
                    If mExcludePartitions.IsPresent Then
                        clearPartitions(majObj)
                    End If

                    If mExcludePermissions.IsPresent Then
                        clearPermissions(majObj)
                    End If

                    If mCreate.IsPresent Then
                        Microsoft.AnalysisServices.Scripter.WriteCreate(outXml, DirectCast(mParentObject, IMajorObject), DirectCast(majObj, IMajorObject), Not mExcludeDependents.IsPresent, True)
                    Else
                        Microsoft.AnalysisServices.Scripter.WriteAlter(outXml, DirectCast(majObj, IMajorObject), Not mExcludeDependents.IsPresent, True)
                    End If

                    outXml.Flush()
                Catch ex As Exception
                    WriteError(New ErrorRecord(ex, "get-xmlaScript", ErrorCategory.NotSpecified, majObj))
                Finally
                    outXml.Close()
                    '// reload the cube definition from the server
                    If mExcludePartitions.IsPresent OrElse mExcludePermissions.IsPresent Then

                        svr.RollbackTransaction()
                        '// only disconnecting and reconnecting seems to reset the AMO references properly
                        '// the rollback tran does not appear to work.
                        Dim connStr As String = svr.ConnectionString
                        svr.Disconnect()
                        svr.Connect(connStr)
                    End If
                End Try

                Me.WriteObject(sbXml.ToString)
            Else
                Me.WriteError(New ErrorRecord(New ArgumentException("You can only script objects that are implemented as Major Objects"), "Can only script Major Objects", ErrorCategory.InvalidArgument, majObj))
            End If
        End Sub

        Private Overloads Sub clearPartitions(ByVal majObj As MajorObject)
            Dim outCube As Cube = TryCast(majObj, Cube)
            Dim outDB As Database = TryCast(majObj, Database)
            Dim outMG As MeasureGroup = TryCast(majObj, MeasureGroup)
            If Not outDB Is Nothing Then
                clearPartitions(outDB)
            ElseIf Not outCube Is Nothing Then
                clearPartitions(outCube)
            ElseIf Not outMG Is Nothing Then
                clearPartitions(outMG)
            Else
                Throw New ArgumentException("You can only use the -ExcludePartitions option with Database, Cube or MeasureGroup objects")
            End If

        End Sub

        Private Overloads Sub clearPartitions(ByVal tempDB As Database)
            If Not tempDB Is Nothing Then
                For Each c As Cube In tempDB.Cubes
                    For Each mg As MeasureGroup In c.MeasureGroups
                        mg.Partitions.Clear()
                    Next
                Next
            End If
        End Sub

        Private Overloads Sub clearPartitions(ByVal tempCube As Cube)
            If Not tempCube Is Nothing Then
                For Each mg As MeasureGroup In tempCube.MeasureGroups
                    mg.Partitions.Clear()
                Next
            End If
        End Sub

        Private Overloads Sub clearPartitions(ByVal tempMG As MeasureGroup)
            tempMG.Partitions.Clear()
        End Sub

        Private Overloads Sub clearPermissions(ByVal majObj As MajorObject)
            If TypeOf majObj Is Cube Then
                clearPermissions(DirectCast(majObj, Cube))
            ElseIf TypeOf majObj Is Dimension Then
                clearPermissions(DirectCast(majObj, Dimension))
            ElseIf TypeOf majObj Is Database Then
                clearPermissions(DirectCast(majObj, Database))
            End If

        End Sub

        Private Overloads Sub clearPermissions(ByVal tmpCube As Cube)
            tmpCube.CubePermissions.Clear()
            For Each d As Dimension In tmpCube.Dimensions
                clearPermissions(d)
            Next
        End Sub

        Private Overloads Sub clearPermissions(ByVal tmpDimension As Dimension)
            tmpDimension.DimensionPermissions.Clear()
        End Sub

        Private Overloads Sub clearPermissions(ByVal tmpDB As Database)
            tmpDB.Roles.Clear()
            For Each d As Dimension In tmpDB.Dimensions
                clearPermissions(d)
            Next
            For Each c As Cube In tmpDB.Cubes
                clearPermissions(c)
            Next
        End Sub

    End Class

    
End Namespace