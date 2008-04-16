Imports System.Xml
Imports System.Management.Automation
Imports Microsoft.AnalysisServices.Xmla
Imports System.Globalization

Namespace Utils

    Public Delegate Sub OutputObjectDelegate(ByVal output As Object)

    Public Class XmlaDiscover

        Friend Sub [Stop]()
            mStopping = True
        End Sub

        Private mStopping As Boolean = False
        Public ReadOnly Property Stopping() As Boolean
            Get
                Return mStopping
            End Get
        End Property

        Public Sub Discover(ByVal command As String, ByVal serverName As String, ByVal restrictions As String, ByVal properties As String, ByVal OutputCallback As OutputObjectDelegate)
            Discover(command, serverName, restrictions, properties, OutputCallback, False)
        End Sub

        Public Sub Discover(ByVal command As String, ByVal serverName As String, ByVal restrictions As String, ByVal properties As String, ByVal OutputCallback As OutputObjectDelegate, ByVal rawResultSet As Boolean)
            Dim x As New Microsoft.AnalysisServices.Xmla.XmlaClient
            x.Connect("data source=" & serverName)
            Try
                Dim res As String = ""
                If command.Contains("<") Then
                    res = x.Send(command, Nothing)
                Else
                    x.Discover(command.ToUpper(), restrictions, properties, res, False, False, False)
                End If
                x.Disconnect(True)
                If rawResultSet Then
                    OutputCallback(res)
                Else
                    ProcessXmlaResult(serverName, res, OutputCallback)
                End If
            Finally
                x.Disconnect()
            End Try

        End Sub


        '// This routine is fairly "brute force", creating an XmlDocument and 
        '// then iterating through the nodes
        Private Sub ProcessXmlaResult(ByVal serverName As String, ByVal xmlaResult As String, ByVal OutputCallback As OutputObjectDelegate)

            Dim rowSchema As New Dictionary(Of String, String)
            Dim doc As XmlDocument = New XmlDocument()
            doc.LoadXml(xmlaResult)

            '//               return         root       schema/rows       Look for element with attribute "name" = row
            '//                  ^             ^             ^             ^
            '//xmlDom.ChildNodes[0].ChildNodes[0].ChildNodes[0].ChildNodes[0].Attributes.Count
            For Each n As XmlNode In doc.ChildNodes(0).ChildNodes(0).ChildNodes
                If ((n.NodeType = XmlNodeType.Element) AndAlso (n.LocalName = "schema")) Then
                    rowSchema = buildPSObjectFromSchema(n)
                End If
                If ((n.NodeType = XmlNodeType.Element) AndAlso (n.LocalName = "row")) Then
                    OutputCallback(GetPSObject(rowSchema, serverName, n))
                End If
                If Me.Stopping Then
                    Exit For
                End If
            Next 'n            

        End Sub



        Public Function Discover(ByVal command As String, ByVal serverName As String, ByVal restrictions As String, ByVal properties As String) As List(Of Object)
            Dim x As New Microsoft.AnalysisServices.Xmla.XmlaClient
            x.Connect("data source=" & serverName)
            Dim res As String = ""
            If command.Contains("<") Then
                res = x.Send(command, Nothing)
            Else
                x.Discover(command, restrictions, properties, res, False, False, False)
            End If
            x.Disconnect(True)

            Return GetObjectListFromXmla(serverName, res)
        End Function

        '// This routine is fairly "brute force", creating an XmlDocument and 
        '// then iterating through the nodes
        Private Function GetObjectListFromXmla(ByVal serverName As String, ByVal xmlaResult As String) As List(Of Object)
            Dim pso As PSObject = Nothing
            Dim rowSchema As New Dictionary(Of String, String)
            Dim doc As XmlDocument = New XmlDocument()
            doc.LoadXml(xmlaResult)
            Dim lst As New List(Of Object)

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


        Private Function buildPSObjectFromSchema(ByVal n As XmlNode) As Dictionary(Of String, String)
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

        Protected Overridable Function GetPSObject(ByVal rowSchema As Dictionary(Of String, String), ByVal server As String, ByVal row As XmlNode) As Object
            Dim pso As New PSObject()
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


        Private Function convertObj(ByVal val As XmlNode, ByVal t As String) As Object
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

        Protected Shared Function NormalizeName(ByVal name As String) As String
            If name.Contains("_") Or name.Contains(" ") Then
                Dim ci As CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
                Dim ti As TextInfo = ci.TextInfo
                Return ti.ToTitleCase(name.ToLower.Replace("_", " ")).Replace(" ", "")
            Else
                Return name
            End If
        End Function
    End Class
End Namespace