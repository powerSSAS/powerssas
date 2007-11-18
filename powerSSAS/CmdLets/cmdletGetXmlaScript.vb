Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Gosbell.PowerSSAS.PowerSSASProvider

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "XmlaScript")> _
    Public Class Scripter
        Inherits PSCmdlet
        Private mMajorObjectPath As String()

        <Parameter(HelpMessage:="An Analysis Services Major Object", Mandatory:=False, Position:=0, ValueFromPipeline:=True)> _
        Public Property MajorObjectPath() As String()
            Get
                Return mMajorObjectPath
            End Get
            Set(ByVal value As String())
                mMajorObjectPath = value

            End Set
        End Property

        Private moutputfile As String

        <Parameter(Mandatory:=False, Position:=1, ValueFromPipeline:=False)> _
        Public Property OutputFile() As String
            Get
                Return moutputfile
            End Get
            Set(ByVal value As String)
                moutputfile = value
            End Set
        End Property

        Protected Overrides Sub EndProcessing()
            MyBase.EndProcessing()
        End Sub

        Protected Overrides Sub ProcessRecord()
            MyBase.ProcessRecord()
            'Me.WriteObject(Me.CommandRuntime.Host.PrivateData.ImmediateBaseObject)
            Dim psDrive As PSDriveInfo = Me.SessionState.Path.CurrentLocation.Drive

            If TypeOf psDrive Is AmoDriveInfo Then
                'Me.WriteObject(Me.SessionState.Path.CurrentLocation.Path)
                'Me.WriteObject(Me.SessionState.Path.CurrentLocation.ProviderPath)

                'Dim amoDrive As powerSSAS.AmoDriveInfo
                'amoDrive = CType(Me.SessionState.Path.CurrentLocation.Drive, powerSSAS.AmoDriveInfo)
                'Me.GetResolvedProviderPathFromPSPath(Me.SessionState.Path.CurrentLocation.Path, Me.SessionState.Path.CurrentLocation.Provider)
                'Dim loc As PathInfo = Me.CurrentProviderLocation("AmoPs")
                Me.WriteObject(Me.InvokeProvider.Item)
                Dim p As PSObject
                Dim itmCol As System.Collections.ObjectModel.Collection(Of PSObject)
                itmCol = Me.InvokeProvider.Item.Get(Me.SessionState.Path.CurrentLocation.Path)
                'p = Me.InvokeProvider.ChildItem.Get(Me.SessionState.Path.CurrentLocation.Path, False)(0)
                'Dim p2 As Microsoft.AnalysisServices.MajorObject
                'p2 = CType(p.BaseObject, Microsoft.AnalysisServices.MajorObject)

                If itmCol.Count = 1 Then
                    p = itmCol.Item(0)
                    If TypeOf p.BaseObject Is Microsoft.AnalysisServices.MajorObject Then
                        Dim majObj As MajorObject
                        majObj = CType(p.BaseObject, MajorObject)
                        Dim amoDrive As AmoDriveInfo
                        amoDrive = CType(psDrive, AmoDriveInfo)
                        Dim sbXml As New Text.StringBuilder()
                        Dim xws As New Xml.XmlWriterSettings()
                        With xws
                            .Indent = True
                            .NewLineHandling = Xml.NewLineHandling.Entitize
                            .OmitXmlDeclaration = True
                        End With


                        Dim outXml As Xml.XmlWriter = Xml.XmlWriter.Create(sbXml, xws)

                        'Microsoft.AnalysisServices.Scripter.WriteStartBatch(outXml, False)
                        'Microsoft.AnalysisServices.Scripter.WriteCreate(outXml, CType(majObj.Parent, IMajorObject), CType(majObj, IMajorObject), False, True)
                        'Microsoft.AnalysisServices.Scripter.WriteEndBatch(outXml)

                        amoDrive.ScriptEngine.ScriptCreate(New MajorObject(0) {majObj}, outXml, True)

                        'Dim si As New Microsoft.AnalysisServices.ScriptInfo(majObj, ScriptAction.AlterWithAllowCreate, ScriptOptions.Default, True)
                        'amoDrive.ScriptEngine.Script(New Microsoft.AnalysisServices.ScriptInfo(0) {si}, outXml)
                        outXml.Flush()
                        outXml.Close()

                        'Dim xdoc As New Xml.XmlDocument
                        'xdoc.LoadXml(sbXml.ToString)
                        'Me.WriteObject(xdoc)

                        Me.WriteObject(sbXml.ToString)

                        'TODO

                    End If
                Else
                    Me.WriteWarning(String.Format("There were {0} objects found at {1}.", itmCol.Count, Me.SessionState.Path.CurrentLocation.Path))
                End If

            Else
                Me.WriteError(New ErrorRecord(New NotSupportedException("You can only create an XMLA script from an AMO location"), "Invalid Location", ErrorCategory.InvalidOperation, psDrive))
            End If


        End Sub

    End Class
End Namespace