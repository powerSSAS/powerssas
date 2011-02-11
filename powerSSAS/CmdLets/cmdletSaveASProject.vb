Imports System.IO
Imports System.Xml
Imports System.Xml.XPath
Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Microsoft.AnalysisServices.Utils
'Imports SSASHelper

Namespace Cmdlets
    <Cmdlet(VerbsData.Save, "ASProject")> _
    Public Class cmdletSaveASProject
        Inherits PSCmdlet

        Private mDatabase As Database = Nothing
        <Parameter(HelpMessage:="The AMO Database that is to be saved as a project", Mandatory:=True, Position:=0, ValueFromPipeline:=True)> _
        Public Property Database() As Database
            Get
                Return mDatabase
            End Get
            Set(ByVal value As Database)
                mDatabase = value
            End Set
        End Property

        Private mProjFolder As String = Nothing
        <Parameter(Mandatory:=True, Position:=1)> _
        Public Property ProjectFolder() As String
            Get
                Return mProjFolder
            End Get
            Set(ByVal value As String)
                Dim filePath As String = Me.GetUnresolvedProviderPathFromPSPath(value)

                mProjFolder = filePath
            End Set
        End Property

        Protected Overrides Sub ProcessRecord()
            MyBase.ProcessRecord()

            ProjectHelper.SerializeProject(mDatabase, mProjFolder)

            Me.WriteDebug("Database Name: " & mDatabase.Name)
            Me.WriteDebug("Database ID: " & mDatabase.ID)
            Me.WriteDebug("Project Folder: " & mProjFolder)

            Me.WriteVerbose("Database: " & mDatabase.Name & " saved to " & mProjFolder)
        End Sub


    End Class
End Namespace