Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Gosbell.PowerSSAS.PowerSSASProvider

Namespace Cmdlets
    <Cmdlet(VerbsCommon.[New], "ASScript")> _
    Public Class cmdletNewAsScript
        Inherits PSCmdlet

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

        Private mInputObj As MajorObject
        <Parameter(mandatory:=True, valueFromPipeline:=True, Position:=0)> _
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



            majObj = mInputObj
            WriteDebug("Object parameter passed in :" & majObj.Name)

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
                        Dim mParent As IMajorObject = TryCast(majObj.Parent, IMajorObject)
                        If mParent Is Nothing Then
                            Throw New ArgumentException("Unable to find the parent for the input object")
                        Else
                            Microsoft.AnalysisServices.Scripter.WriteCreate(outXml, mParent, DirectCast(majObj, IMajorObject), Not mExcludeDependents.IsPresent, True)
                        End If
                    Else
                        Microsoft.AnalysisServices.Scripter.WriteAlter(outXml, DirectCast(majObj, IMajorObject), Not mExcludeDependents.IsPresent, True)
                    End If

                    outXml.Flush()
                Catch ex As Exception
                    WriteError(New ErrorRecord(ex, "New-AsScript", ErrorCategory.NotSpecified, majObj))
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