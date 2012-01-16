Imports System.Management.Automation
Imports Microsoft.AnalysisServices

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASTraceSubclass")> _
    Public Class cmdletGetASTraceSubclass
        Inherits Cmdlet

        Private mServerName As String = ""
        <Parameter(Position:=0, Mandatory:=False)> _
        Public Property ServerName() As String
            Get
                Return mServerName
            End Get
            Set(ByVal value As String)
                mServerName = value
            End Set
        End Property

        Protected Overrides Sub ProcessRecord()
            'Dim res As String = ""
            'Dim x As New Microsoft.AnalysisServices.Xmla.XmlaClient
            'Dim Svr As New Server()
            Dim cli As New Microsoft.AnalysisServices.AdomdClient.AdomdConnection("Data Source=" & ServerName)
            Dim xpathQry As String = "EVENTCATEGORY/EVENTLIST/EVENT/EVENTCOLUMNLIST/EVENTCOLUMN[ID=""1""]/EVENTCOLUMNSUBCLASSLIST/EVENTCOLUMNSUBCLASS"
            cli.Open()
            Dim ds As DataSet = cli.GetSchemaDataSet("DISCOVER_TRACE_EVENT_CATEGORIES", Nothing)

            'x.Discover("DISCOVER_TRACE_EVENT_CATEGORIES", "", "", res, False, False, False)
            Dim strColl As List(Of String) = GetData(ds)
            For Each strXml As String In strColl
                Dim xdoc As New Xml.XPath.XPathDocument(New IO.StringReader(strXml))
                Dim xnav As Xml.XPath.XPathNavigator = xdoc.CreateNavigator()
                Dim iter As Xml.XPath.XPathNodeIterator
                iter = xnav.Select(xpathQry)
                While iter.MoveNext()
                    Dim esc As New EventSubClass
                    esc.SubclassId = CType(iter.Current.Evaluate("string(ID/text())"), String)
                    esc.SubclassName = CType(iter.Current.Evaluate("string(NAME/text())"), String)
                    esc.EventID = CType(iter.Current.Evaluate("string(ancestor::EVENT/ID/text())"), String)
                    esc.EventName = CType(iter.Current.Evaluate("string(ancestor::EVENT/NAME/text())"), String)
                    esc.EventDescription = CType(iter.Current.Evaluate("string(ancestor::EVENT/DESCRIPTION/text())"), String)

                    'esc.SubclassId = iter.Current.Select("ID").Current.Value
                    'esc.SubclassName = iter.Current.Select("NAME").Current.Value
                    'esc.EventID = iter.Current.Select("ancestor::EVENT/ID").Current.Value
                    'esc.EventName = iter.Current.Select("ancestor::EVENT/NAME").Current.Value
                    'esc.EventDescription = iter.Current.Select("ancestor::EVENT/DESCRIPTION").Current.Value
                    WriteObject(esc)
                End While
            Next

        End Sub

        Private Function GetData(ByVal x As String) As List(Of String)
            Dim xdoc As New Xml.XmlDocument()
            xdoc.LoadXml(x)
            Dim lst As New List(Of String)
            For iCnt As Integer = 1 To xdoc.ChildNodes(0).ChildNodes(0).ChildNodes.Count - 1
                lst.Add(xdoc.ChildNodes(0).ChildNodes(0).ChildNodes(iCnt).InnerText)
            Next iCnt

            Return lst
        End Function

        Private Function GetData(ByVal ds As DataSet) As List(Of String)
            Dim xdoc As New Xml.XmlDocument()
            'xdoc.LoadXml(x)
            Dim lst As New List(Of String)
            For Each dr As DataRow In ds.Tables(0).Rows
                xdoc.LoadXml(CType(dr.Item("Data"), String))
                For iCnt As Integer = 1 To xdoc.ChildNodes(0).ChildNodes.Count - 1
                    lst.Add(xdoc.ChildNodes(0).ChildNodes(iCnt).InnerText)
                Next iCnt
            Next


            Return lst
        End Function
    End Class
End Namespace

Public Class EventSubclass
    Public EventName As String
    Public EventID As String
    Public EventDescription As String
    Public SubclassName As String
    Public SubclassId As String
End Class
