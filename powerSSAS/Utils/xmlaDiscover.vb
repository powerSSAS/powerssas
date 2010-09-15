Imports System.Xml
Imports System.Management.Automation
Imports System.Globalization
Imports Microsoft.AnalysisServices.AdomdClient

Namespace Utils

    Public Delegate Sub OutputObjectCallback(ByVal output As Object)

    Public Class XmlaDiscover

        Friend Sub [Stop]()
            mStopping = True
        End Sub

        Private mStopping As Boolean '= False
        Public ReadOnly Property Stopping() As Boolean
            Get
                Return mStopping
            End Get
        End Property

        'Public Sub Discover(ByVal command As String, ByVal serverName As String, ByVal restrictions As AdomdRestrictionCollection, ByVal databaseName As String, ByVal outputCallback As OutputObjectCallback)
        '    Discover(command, serverName, restrictions, databaseName, outputCallback, False)
        'End Sub

        Public Sub Discover(ByVal command As String, ByVal serverName As String, ByVal restrictions As AdomdRestrictionCollection, ByVal databaseName As String, ByVal outputCallback As OutputObjectCallback) ', ByVal rawResultSet As Boolean)
            If outputCallback Is Nothing Then
                Throw New ArgumentException("The ouputCallback argument must be a valid delegate")
            End If

            If Not command Is Nothing Then
                'Dim x As New Microsoft.AnalysisServices.Xmla.XmlaClient
                'x.Connect("data source=" & serverName)
                Dim ds As DataSet
                Dim connStr As String = "Data Source=" & serverName
                If databaseName.Length > 0 Then
                    connStr &= ";Initial Catalog=" & databaseName
                End If
                Dim conn As New AdomdConnection(connStr)
                Try
                    Dim res As String = ""
                    If command.Contains("<") Then
                        'res = x.Send(command, Nothing)
                        Throw (New ArgumentException("Invalid command Schema Rowset name: " & command))
                    Else
                        conn.Open()
                        ds = conn.GetSchemaDataSet(command.ToUpper(), restrictions)
                        outputCallback(ds.Tables.Item(0))
                    End If

                    'If rawResultSet Then
                    '    outputCallback(res)
                    'Else
                    'ProcessDataSet(serverName, ds, outputCallback)
                    'End If
                Finally
                    conn.Close()
                End Try
            End If
        End Sub

        Private Sub ProcessDataSet(ByVal serverName As String, ByVal ds As DataSet, ByVal OutputCallback As OutputObjectCallback)
            For Each dr As DataRow In ds.Tables(0).Rows

                OutputCallback(GetPSObject(dr, serverName))
            Next
        End Sub

        '// This routine is fairly "brute force", creating an XmlDocument and 
        '// then iterating through the nodes
        'Private Sub ProcessXmlaResult(ByVal serverName As String, ByVal xmlaResult As String, ByVal OutputCallback As OutputObjectCallback)

        '    Dim rowSchema As New Dictionary(Of String, String)
        '    Dim doc As XmlDocument = New XmlDocument()
        '    doc.LoadXml(xmlaResult)

        '    '//               return         root       schema/rows       Look for element with attribute "name" = row
        '    '//                  ^             ^             ^             ^
        '    '//xmlDom.ChildNodes[0].ChildNodes[0].ChildNodes[0].ChildNodes[0].Attributes.Count
        '    For Each n As XmlNode In doc.ChildNodes(0).ChildNodes(0).ChildNodes
        '        If ((n.NodeType = XmlNodeType.Element) AndAlso (n.LocalName = "schema")) Then
        '            rowSchema = buildPSObjectFromSchema(n)
        '        End If
        '        If ((n.NodeType = XmlNodeType.Element) AndAlso (n.LocalName = "row")) Then
        '            OutputCallback(GetPSObject(rowSchema, serverName, n))
        '        End If
        '        If Me.Stopping Then
        '            Exit For
        '        End If
        '    Next 'n            

        'End Sub


        Public Function Discover(ByVal command As String, ByVal serverName As String, ByVal restrictions As AdomdRestrictionCollection, ByVal databaseName As String) As System.Collections.ObjectModel.Collection(Of Object)
            'Dim x As New Microsoft.AnalysisServices.Xmla.XmlaClient
            'x.Connect("data source=" & serverName)
            Dim conn As New Microsoft.AnalysisServices.AdomdClient.AdomdConnection("Data Source=" & serverName)
            conn.Open()
            Dim ds As DataSet
            Try
                Dim res As String = ""
                If command.Contains("<") Then
                    'res = x.Send(command, Nothing)
                    Throw New ArgumentException("Invalid Schema Rowset: " & command)
                Else
                    ds = conn.GetSchemaDataSet(command, restrictions)
                End If
            Finally
                conn.Close()
            End Try

            Return GetObjectListFromDataSet(serverName, ds)
        End Function

        Private Function GetObjectListFromDataSet(ByVal serverName As String, ByVal ds As DataSet) As System.Collections.ObjectModel.Collection(Of Object)
            'TODO Implement GetObjectListFromDataSet
            Dim coll As New System.Collections.ObjectModel.Collection(Of Object)
            For Each dr As DataRow In ds.Tables(0).Rows
                Dim o As New PSObject
                For Each dc As DataColumn In ds.Tables(0).Columns
                    o.Properties.Add(New PSNoteProperty(dc.ColumnName, dr(dc.ColumnName)))
                Next
                If Not ds.Tables(0).Columns.Contains("Name") Then
                    o.Properties.Add(New PSNoteProperty("Name", dr(0)))
                End If
                coll.Add(o)
            Next
            Return coll
            'Throw New NotImplementedException("GetObjectListFromDataSet")
        End Function

        '// This routine is fairly "brute force", creating an XmlDocument and 
        '// then iterating through the nodes
        Private Function GetObjectListFromXmla(ByVal serverName As String, ByVal xmlaResult As String) As System.Collections.ObjectModel.Collection(Of Object)

            Dim rowSchema As New Dictionary(Of String, String)
            Dim doc As XmlDocument = New XmlDocument()
            doc.LoadXml(xmlaResult)
            Dim lst As New System.Collections.ObjectModel.Collection(Of Object)

            '//               return         root       schema/rows       Look for element with attribute "name" = row
            '//                  ^             ^             ^             ^
            '//xmlDom.ChildNodes[0].ChildNodes[0].ChildNodes[0].ChildNodes[0].Attributes.Count
            For Each n As XmlNode In doc.ChildNodes(0).ChildNodes(0).ChildNodes
                If ((n.NodeType = XmlNodeType.Element) AndAlso (n.LocalName = "schema")) Then
                    rowSchema = buildPSObjectFromSchema(n)
                End If
                If ((n.NodeType = XmlNodeType.Element) AndAlso (n.LocalName = "row")) Then
                    lst.Add(GetPSObject(rowSchema, serverName, n))
                End If
            Next 'n            
            Return lst
        End Function


        Private Shared Function buildPSObjectFromSchema(ByVal n As XmlNode) As Dictionary(Of String, String)
            Dim schema As New Dictionary(Of String, String)

            For Each n2 As XmlNode In n.ChildNodes
                System.Diagnostics.Debug.WriteLine(n.Value)
                If ((n2.NodeType = XmlNodeType.Element) AndAlso (n2.LocalName = "complexType")) Then
                    For Each a As XmlAttribute In n2.Attributes
                        If ((a.Name = "name") AndAlso (a.Value = "row")) Then
                            '// we have the row definition
                            '//                                  Sequence
                            '//                                   ^
                            For Each n3 As XmlNode In n2.ChildNodes(0).ChildNodes
                                Dim fld As String = ""
                                Dim typ As String = ""
                                For Each a2 As XmlAttribute In n3.Attributes
                                    Select Case a2.Name
                                        Case "sql:field"
                                            fld = a2.Value
                                        Case "type"
                                            typ = a2.Value
                                    End Select
                                Next
                                If ((Not typ Is Nothing) AndAlso (Not fld Is Nothing)) Then
                                    schema.Add(fld, typ)
                                End If
                                '// reset field and type variables
                                fld = Nothing
                                typ = Nothing
                            Next
                        End If
                    Next
                End If
            Next

            Return schema
        End Function

        Protected Overridable Function GetPSObject(ByVal dr As DataRow, ByVal server As String) As Object
            Dim pso As New PSObject()
            For Each dc As DataColumn In dr.Table.Columns
                pso.Properties.Add(New PSNoteProperty(NormalizeName(dc.ColumnName), dr.Item(dc.Ordinal)))
            Next
            Return pso
        End Function

        Protected Overridable Function GetPSObject(ByVal rowSchema As Dictionary(Of String, String), ByVal server As String, ByVal row As XmlNode) As Object
            'Protected Overridable Function GetPSObject(ByVal rowSchema As Dictionary(Of String, String), ByVal server As String, ByVal inputRow As XPath.IXPathNavigable) As Object
            Dim pso As New PSObject()
            'Dim row As XPath.XPathNavigator = inputRow.CreateNavigator()

            For Each n As XmlNode In row.ChildNodes
                If (rowSchema.ContainsKey(n.LocalName)) Then
                    'pso.Properties.Add(New PSNoteProperty(e.LocalName, convertObj(e.InnerText, row.Item(e.LocalName))))
                    Dim propName As String = NormalizeName(n.LocalName)
                    Dim prop As PSNoteProperty = CType(pso.Properties.Item(propName), PSNoteProperty)
                    If prop Is Nothing OrElse TypeOf prop Is IList Then
                        pso.Properties.Add(New PSNoteProperty(propName, convertObj(n, rowSchema.Item(n.LocalName))))
                    Else
                        If TypeOf prop.Value Is IList Then
                            CType(prop.Value, IList).Add(convertObj(n, rowSchema.Item(n.LocalName)))
                        Else
                            Dim lst As New List(Of Object)
                            lst.Add(prop.Value)
                            lst.Add(convertObj(n, rowSchema.Item(n.LocalName)))
                            pso.Properties.Remove(prop.Name)
                            prop = New PSNoteProperty(prop.Name, lst)
                            pso.Properties.Add(prop)
                        End If
                    End If
                End If
            Next
            Return pso
        End Function


        Private Shared Function convertObj(ByVal val As XmlNode, ByVal t As String) As Object
            Select Case t
                Case "xsd:int"
                    Return Integer.Parse(val.InnerText)
                Case "xsd:long", "xsd:unsignedInt"
                    Return Long.Parse(val.InnerText)
                Case "xsd:unsignedLong"
                    Return Convert.ToUInt64(val.InnerText)
                Case "xsd:dateTime"
                    Return Date.Parse(val.InnerText)
                Case "xsd:string", "uuid"
                    Return val.InnerText
                Case "xsd:boolean"
                    Return Convert.ToBoolean(val.InnerText)
                Case "xsd:short"
                    Return Convert.ToInt16(val.InnerText)
                Case "xsd:unsignedShort"
                    Return Convert.ToUInt16(val.InnerText)
                Case Else '// return the whole XmlNode
                    Return val
            End Select
        End Function

        <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:ValidateArgumentsOfPublicMethods")> Protected Shared Function NormalizeName(ByVal name As String) As String
            If Not name Is Nothing AndAlso name.Length > 0 Then
                If name.Contains("_") Or name.Contains(" ") Then
                    Dim ci As CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
                    Dim ti As TextInfo = ci.TextInfo
                    Return ti.ToTitleCase(name.ToLower.Replace("_", " ")).Replace(" ", "")
                End If
            End If
            Return name

        End Function
    End Class
End Namespace