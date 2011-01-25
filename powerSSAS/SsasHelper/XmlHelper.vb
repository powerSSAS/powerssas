Imports System
Imports System.Collections.Generic
'Imports System.Linq
Imports System.Text
Imports System.Xml

Namespace SsasHelper

    '/// <summary>
    '/// Functions to help when manipulating XML documents.
    '/// </summary>
    Class XmlHelper

        '/// <summary>
        '/// Check to see if a node exists under a given node.
        '/// </summary>
        '/// <param name="parentNode">Parent node to check under</param>
        '/// <param name="nodeToCheckFor">Name of the node to look for</param>
        '/// <returns>true/false based on existence</returns>
        Public Shared Function NodeExists(ByVal parentNode As XmlNode, ByVal nodeToCheckFor As String) As Boolean
            Dim ret As Boolean = False

            For Each sibling As XmlNode In parentNode.ChildNodes
                If (sibling.Name = nodeToCheckFor) Then
                    ret = True
                    Exit For
                End If
            Next

            Return ret
        End Function

        '/// <summary>
        '/// Remove all nodes in the given XmlNodeList.
        '/// </summary>
        '/// <param name="nodeList">List of nodes to remove</param>
        '/// <returns>Number of nodes removed</returns>
        Public Shared Function RemoveNodes(ByVal nodeList As XmlNodeList) As Integer
            Dim count As Integer = 0

            For Each node As XmlNode In nodeList
                node.ParentNode.RemoveChild(node)
                count += 1
            Next

            Return count
        End Function

        '/// <summary>
        '/// Remove the specified attribute from a node.
        '/// </summary>
        '/// <param name="nodeList">List of nodes to remove the attribute from</param>
        '/// <param name="attributeName">Attribute to remove</param>
        '/// <returns></returns>
        Public Shared Function RemoveAttributes(ByVal nodeList As XmlNodeList, ByVal attributeName As String) As Integer
            Dim count As Integer = 0

            For Each node As XmlElement In nodeList
                node.RemoveAttribute(attributeName)
                count += 1
            Next node

            Return count
        End Function
    End Class
end Namespace
