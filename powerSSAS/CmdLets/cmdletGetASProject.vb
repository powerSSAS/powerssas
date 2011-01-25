Imports System.IO
Imports System.Xml
Imports System.Xml.XPath
Imports System.Management.Automation
Imports Microsoft.AnalysisServices
Imports Microsoft.AnalysisServices.Utils
'Imports SSASHelper

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASProject")> _
    Public Class cmdletGetASProject
        Inherits PSCmdlet

        Private mProjFile As String = Nothing
        <Parameter(Mandatory:=True)> _
        Public Property ProjectFile() As String
            Get
                Return mProjFile
            End Get
            Set(ByVal value As String)
                Dim filePath As String = Me.GetUnresolvedProviderPathFromPSPath(value)
                mProjFile = filePath
            End Set
        End Property

        Protected Overrides Sub ProcessRecord()
            MyBase.ProcessRecord()

            Dim db As Database = ProjectHelper.DeserializeProject(ProjectFile)  'LoadProject(ProjectFile)
            WriteObject(db)
        End Sub


        'Private Function LoadProject(ByVal fileName As FileInfo) As Database
        '    Dim xpd As XPathDocument = New XPathDocument(fileName.FullName)
        '    Dim xpn As XPathNavigator = xpd.CreateNavigator()
        '    Dim db As Database = New Database()
        '    Dim txtRdr As IO.StringReader = New IO.StringReader("")
        '    Dim xrdr As Xml.XmlReader = XmlReader.Create(txtRdr)
        '    Dim xni As XPathNodeIterator
        '    Dim asFile As FileInfo


        '    'Load Database
        '    xni = xpn.Select("/Project/Database/FullPath")
        '    xni.MoveNext()
        '    asFile = New FileInfo(Path.Combine(fileName.DirectoryName, xni.Current.Value))
        '    xrdr = XmlReader.Create(asFile.OpenRead())
        '    Microsoft.AnalysisServices.Utils.Deserialize(xrdr, db)
        '    xrdr.Close()

        '    'Load DataSources
        '    xni = xpn.Select("/Project/DataSources/ProjectItem/FullPath")
        '    xni.MoveNext()
        '    For i As Integer = 0 To xni.Count - 1
        '        asFile = New FileInfo(Path.Combine(fileName.DirectoryName, xni.Current.Value))
        '        xrdr = XmlReader.Create(asFile.OpenRead())

        '        Dim xrdr2 As XmlReader = XmlReader.Create(asFile.OpenRead)
        '        Dim xpd2 As New XPathDocument(xrdr2)
        '        Dim xpn2 As XPathNavigator = xpd2.CreateNavigator()
        '        Dim nsmgr As New XmlNamespaceManager(xpn2.NameTable)
        '        nsmgr.AddNamespace("a", "http://schemas.microsoft.com/analysisservices/2003/engine")
        '        Dim id As String = xpn2.SelectSingleNode("/a:DataSource/a:ID", nsmgr).InnerXml
        '        xrdr2.Close()

        '        Dim ds As Microsoft.AnalysisServices.DataSource = db.DataSources.AddNew(id)
        '        Microsoft.AnalysisServices.Utils.Deserialize(xrdr, ds)
        '        xrdr.Close()
        '        xni.MoveNext()
        '    Next

        '    'Load DSVs
        '    xni = xpn.Select("/Project/DataSourceViews/ProjectItem/FullPath")
        '    xni.MoveNext()
        '    For i As Integer = 0 To xni.Count - 1
        '        Dim dsv As Microsoft.AnalysisServices.DataSourceView
        '        asFile = New FileInfo(Path.Combine(fileName.DirectoryName, xni.Current.Value))

        '        Dim xrdr2 As XmlReader = XmlReader.Create(asFile.OpenRead)
        '        Dim xpd2 As New XPathDocument(xrdr2)
        '        Dim xpn2 As XPathNavigator = xpd2.CreateNavigator()
        '        Dim nsmgr As New XmlNamespaceManager(xpn2.NameTable)
        '        nsmgr.AddNamespace("a", "http://schemas.microsoft.com/analysisservices/2003/engine")
        '        Dim id As String = xpn2.SelectSingleNode("/a:DataSourceView/a:ID", nsmgr).InnerXml
        '        xrdr2.Close()

        '        xrdr = XmlReader.Create(asFile.OpenRead())
        '        dsv = db.DataSourceViews.AddNew(id)
        '        Microsoft.AnalysisServices.Utils.Deserialize(xrdr, dsv)
        '        xrdr.Close()
        '        xni.MoveNext()
        '    Next

        '    'Load Dimensions

        '    'Load Cubes/Partitions
        '    'TODO load Partitions
        '    xni = xpn.Select("/Project/Cubes/ProjectItem/FullPath")
        '    xni.MoveNext()
        '    For i As Integer = 0 To xni.Count - 1
        '        asFile = New FileInfo(Path.Combine(fileName.DirectoryName, xni.Current.Value))
        '        xrdr = XmlReader.Create(asFile.OpenRead())
        '        Dim c As New Cube
        '        Deserialize(xrdr, c)
        '        db.Cubes.Add(c)
        '        xni.MoveNext()
        '    Next

        '    'Load MiningModels

        '    'Load Roles

        '    Return db
        'End Function

    End Class
End Namespace