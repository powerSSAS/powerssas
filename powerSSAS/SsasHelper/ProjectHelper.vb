Imports System
Imports System.Collections.Generic
'Imports System.Linq
Imports System.Text

Imports System.Xml
Imports System.Xml.Xsl
Imports Microsoft.AnalysisServices
Imports System.IO
Imports System.Diagnostics
Imports Gosbell.PowerSSAS.SsasHelper

'/// <summary>
'/// This class contains methods to work with SSAS Projects.  
'/// </summary>
'/// <remarks>
'/// Author:  DDarden
'/// Date  :  200905031313
'/// Use this at your own risk.  It has been tested and runs successfully, but results may vary.
'/// Works on my machine... :)
'/// 
'/// KNOWN ISSUES
'/// - Partitions are reordered when De-Serialized/Serialized.  Makes it a pain to validate,
'///   but I've seen no ill effects.
'/// </remarks>
Public Class ProjectHelper


    Public Overloads Shared Function DeserializeProject(ByVal ssasProjectFile As String) As Database

        '// Verify inputs
        If (Not File.Exists(ssasProjectFile)) Then
            Throw New ArgumentException(String.Format("'{0}' does not exist", ssasProjectFile))
        End If

        '// Get the directory to load all project files
        Dim fi As FileInfo = New FileInfo(ssasProjectFile)
        Return DeserializeProject(fi)
    End Function

    '/// <summary>
    '/// Load a SSAS project file into a SSAS Database
    '/// </summary>
    '/// <remarks>
    '/// TODO:  Doesn't support Assemblies, or possibly some other types
    '/// </remarks>
    '/// <param name="ssasProjectFile">Path to the .dwproj file for a SSAS Project</param>
    '/// <returns>AMO Database built from the SSAS project file</returns>
    Public Overloads Shared Function DeserializeProject(ByVal ssasProjectFile As FileInfo) As Database

        Dim database As Database = New Database()
        Dim innerReader As XmlReader = Nothing
        Dim nodeList As XmlNodeList = Nothing
        Dim dependencyNodeList As XmlNodeList = Nothing
        Dim fullPath As String = Nothing

        Dim ssasProjectDirectory As String = ssasProjectFile.Directory.FullName + "\\"

        '// Load the SSAS Project File
        Dim reader As XmlReader = New XmlTextReader(ssasProjectFile.FullName)
        Dim doc As XmlDocument = New XmlDocument()
        doc.Load(reader)

        '// Load the Database
        nodeList = doc.SelectNodes("//Database/FullPath")
        fullPath = nodeList(0).InnerText
        innerReader = New XmlTextReader(ssasProjectDirectory + fullPath)
        Microsoft.AnalysisServices.Utils.Deserialize(innerReader, CType(database, MajorObject))

        '// Load all the Datasources
        nodeList = doc.SelectNodes("//DataSources/ProjectItem/FullPath")
        Dim ds As DataSource = Nothing
        For Each node As XmlNode In nodeList
            fullPath = node.InnerText
            innerReader = New XmlTextReader(ssasProjectDirectory + fullPath)
            ds = New RelationalDataSource()
            Microsoft.AnalysisServices.Utils.Deserialize(innerReader, CType(ds, MajorObject))
            database.DataSources.Add(ds)
        Next node

        '// Load all the DatasourceViews
        nodeList = doc.SelectNodes("//DataSourceViews/ProjectItem/FullPath")
        Dim dsv As DataSourceView = Nothing
        For Each node As XmlNode In nodeList
            fullPath = node.InnerText
            innerReader = New XmlTextReader(ssasProjectDirectory + fullPath)
            dsv = New DataSourceView()
            Microsoft.AnalysisServices.Utils.Deserialize(innerReader, CType(dsv, MajorObject))
            database.DataSourceViews.Add(dsv)
        Next node

        '// Load all the Roles
        nodeList = doc.SelectNodes("//Roles/ProjectItem/FullPath")
        Dim r As Role = Nothing
        For Each node As XmlNode In nodeList
            fullPath = node.InnerText
            innerReader = New XmlTextReader(ssasProjectDirectory + fullPath)
            r = New Role()
            Microsoft.AnalysisServices.Utils.Deserialize(innerReader, CType(r, MajorObject))
            database.Roles.Add(r)
        Next node

        '// Load all the Dimensions
        nodeList = doc.SelectNodes("//Dimensions/ProjectItem/FullPath")
        Dim d As Dimension = Nothing
        For Each node As XmlNode In nodeList
            fullPath = node.InnerText
            innerReader = New XmlTextReader(ssasProjectDirectory + fullPath)
            d = New Dimension()
            Microsoft.AnalysisServices.Utils.Deserialize(innerReader, CType(d, MajorObject))
            database.Dimensions.Add(d)
        Next node

        '// Load all the Mining Models
        nodeList = doc.SelectNodes("//MiningModels/ProjectItem/FullPath")
        Dim ms As MiningStructure = Nothing
        For Each node As XmlNode In nodeList
            fullPath = node.InnerText
            innerReader = New XmlTextReader(ssasProjectDirectory + fullPath)
            ms = New MiningStructure()
            Microsoft.AnalysisServices.Utils.Deserialize(innerReader, CType(ms, MajorObject))
            database.MiningStructures.Add(ms)
        Next node

        '// Load all the Cubes
        nodeList = doc.SelectNodes("//Cubes/ProjectItem/FullPath")
        Dim c As Cube = Nothing
        For Each node As XmlNode In nodeList
            fullPath = node.InnerText
            innerReader = New XmlTextReader(ssasProjectDirectory + fullPath)
            c = New Cube()
            Microsoft.AnalysisServices.Utils.Deserialize(innerReader, CType(c, MajorObject))
            database.Cubes.Add(c)

            '// Process cube dependencies (i.e. partitions
            '// Little known fact:  The Serialize/Deserialize methods DO handle partitions... just not when 
            '// paired with anything else in the cube.  We have to do this part ourselves
            dependencyNodeList = node.ParentNode.SelectNodes("Dependencies/ProjectItem/FullPath")
            For Each dependencyNode As XmlNode In dependencyNodeList
                fullPath = dependencyNode.InnerText
                innerReader = ProjectHelper.FixPartitionsFileForDeserialize(ssasProjectDirectory + fullPath, c)
                Dim partitionsCube As Cube = New Cube()
                Microsoft.AnalysisServices.Utils.Deserialize(innerReader, CType(partitionsCube, MajorObject))
                MergePartitionCube(c, partitionsCube)
            Next dependencyNode

        Next node

        Return database
    End Function

    ''/// <summary>
    ''/// Save a Database object to the individual BIDS files.
    ''/// The filename for each object will be based on its Name.
    ''/// </summary>
    ''/// <remarks>
    ''/// TODO:  Doesn't support Assemblies, or possibly some other types
    ''/// 
    ''/// Some attributes are re-ordered in DSVs.
    ''/// 
    ''/// The following information will be lost when saving a project:
    ''/// (these are not required for correct operation)
    ''/// - State (Processed, Unprocessed)
    ''/// - CreatedTimestamp
    ''/// - LastSchemaUpdate
    ''/// - LastProcessed
    ''/// - dwd:design-time-name
    ''/// - CurrentStorageMode
    ''/// </remarks>
    ''/// <param name="database">Database to output files for</param>
    ''/// <param name="targetDirectory">Directory to create the files in</param>
    Public Shared Sub SerializeProject(ByVal database As Database, ByVal targetDirectory As String)

        '// Validate inputs
        If (database Is Nothing) Then
            Throw New ArgumentException("Please provide a database object")
        End If

        If (String.IsNullOrEmpty(targetDirectory)) Then
            Throw New ArgumentException("Please provide a directory to write the files to")
        End If

        If (Not Directory.Exists(targetDirectory)) Then
            Directory.CreateDirectory(targetDirectory)
        End If

        If (Not targetDirectory.EndsWith("\\", StringComparison.InvariantCulture)) Then
            targetDirectory &= "\\"
        End If

        Dim writer As XmlTextWriter = Nothing

        '// Iterate through all objects in the database and serialize them
        For Each ds As DataSource In database.DataSources
            writer = New XmlTextWriter(targetDirectory + ds.Name + ".ds", Encoding.UTF8)
            writer.Formatting = Formatting.Indented
            Microsoft.AnalysisServices.Utils.Serialize(writer, CType(ds, MajorObject), False)
            writer.Close()
        Next ds

        For Each dsView As DataSourceView In database.DataSourceViews
            writer = New XmlTextWriter(targetDirectory + dsView.Name + ".dsv", Encoding.UTF8)
            writer.Formatting = Formatting.Indented
            Microsoft.AnalysisServices.Utils.Serialize(writer, CType(dsView, MajorObject), False)
            writer.Close()
        Next dsView

        For Each r As Role In database.Roles
            writer = New XmlTextWriter(targetDirectory + r.Name + ".role", Encoding.UTF8)
            writer.Formatting = Formatting.Indented
            Microsoft.AnalysisServices.Utils.Serialize(writer, CType(r, MajorObject), False)
            writer.Close()
        Next r

        For Each DBdim As Dimension In database.Dimensions
            writer = New XmlTextWriter(targetDirectory + DBdim.Name + ".dim", Encoding.UTF8)
            writer.Formatting = Formatting.Indented
            Microsoft.AnalysisServices.Utils.Serialize(writer, CType(DBdim, MajorObject), False)
            writer.Close()
        Next DBdim

        For Each miningStruct As MiningStructure In database.MiningStructures
            writer = New XmlTextWriter(targetDirectory + miningStruct.Name + ".dmm", Encoding.UTF8)
            writer.Formatting = Formatting.Indented
            Microsoft.AnalysisServices.Utils.Serialize(writer, CType(miningStruct, MajorObject), False)
            writer.Close()
        Next miningStruct

        '// Special case:  The cube serialization won't work for partitions when Partion/AggregationDesign
        '// objects are mixed in with other objects.  Serialize most of the cube, then split out
        '// Partion/AggregationDesign objects into their own cube to serialize, then clean up
        '// a few tags
        For Each c As Cube In database.Cubes

            writer = New XmlTextWriter(targetDirectory + c.Name + ".cube", Encoding.UTF8)
            writer.Formatting = Formatting.Indented
            Microsoft.AnalysisServices.Utils.Serialize(writer, CType(c, MajorObject), False)
            writer.Close()

            '// Partitions and AggregationDesigns may be written to the Cube file, and we want
            '// to keep them all in the Partitions file; strip them from the cube file
            FixSerializedCubeFile(targetDirectory + c.Name + ".cube")

            Dim partitionCube As Cube = SplitPartitionCube(c)
            writer = New XmlTextWriter(targetDirectory & c.Name & ".partitions", Encoding.UTF8)
            writer.Formatting = Formatting.Indented
            Microsoft.AnalysisServices.Utils.Serialize(writer, CType(partitionCube, MajorObject), False)
            writer.Close()

            '// The partitions file gets serialized with a few extra nodes... remove them
            FixSerializedPartitionsFile(targetDirectory + c.Name + ".partitions")
        Next
    End Sub

    '/// <summary>
    '/// Generate a .ASDatabase file based on a database object.
    '/// </summary>
    '/// <remarks>
    '/// The following information will be lost from the .ASDatabase file
    '/// - State (Processed, Unprocessed)
    '/// - CreatedTimestamp
    '/// - LastSchemaUpdate
    '/// - LastProcessed
    '/// - dwd:design-time-name
    '/// - CurrentStorageMode
    '/// </remarks>
    '/// <param name="database">Database to build</param>
    '/// <param name="targetFilename">File to generate</param>
    Public Shared Sub GenerateASDatabaseFile(ByVal database As Database, ByVal targetFilename As String)

        '// Validate inputs
        If (database Is Nothing) Then
            Throw New ArgumentException("Please provide a database object")
        End If

        '// Create the directory to put the file in if it doesn't exist
        Dim fi As FileInfo = New FileInfo(targetFilename)
        If (Not fi.Directory.Exists) Then
            fi.Directory.Create()
        End If

        '// Build the ASDatabase file...
        Dim writer As XmlTextWriter = New XmlTextWriter(targetFilename, Encoding.UTF8)
        writer.Formatting = Formatting.Indented
        Microsoft.AnalysisServices.Utils.Serialize(writer, CType(database, MajorObject), False)
        writer.Close()
    End Sub

    '/// <summary>
    '/// Generate a .ASDatabase file based on a database object.
    '/// </summary>
    '/// <remarks>
    '/// The following information will be lost from the .ASDatabase file
    '/// - State (Processed, Unprocessed)
    '/// - CreatedTimestamp
    '/// - LastSchemaUpdate
    '/// - LastProcessed
    '/// - dwd:design-time-name
    '/// - CurrentStorageMode
    '/// </remarks>
    '/// <param name="ssasProjectFile">Path the the .dwproj file for a SSAS Project</param>
    '/// <param name="targetFilename">File to generate</param>
    Public Shared Sub GenerateASDatabaseFile(ByVal ssasProjectFile As String, ByVal targetFilename As String)
        Dim database As Database
        database = ProjectHelper.DeserializeProject(ssasProjectFile)
        ProjectHelper.GenerateASDatabaseFile(database, targetFilename)
    End Sub

    '/// <summary>
    '/// Clean non-essential, highly volatile attributes from SSAS project files.
    '/// </summary>
    '/// <remarks>
    '/// Clean only the top level directory, using the default files, creating backups,
    '/// not removing dimension annotations,not removing design-time-name, and throwing away file counts
    '/// </remarks>
    '/// <param name="ssasProjectPath">SSAS project directory</param>
    Public Shared Sub CleanSsasProjectDirectory(ByVal ssasProjectPath As String)

        Dim filesInspectedCount As Integer = 0
        Dim filesAlteredCount As Integer = 0
        Dim filesCleanedCount As Integer = 0

        '// Clean the SSAS directory
        '// Only the top level directory, using the default files, creating backups, not removing dimension annotations,
        '// not removing design-time-name, and throwing away file counts
        CleanSsasProjectDirectory(ssasProjectPath, String.Empty, SearchOption.TopDirectoryOnly, False, False, True, filesInspectedCount, filesCleanedCount, filesAlteredCount)
    End Sub

    '/// <summary>
    '/// Clean non-essential, highly volatile attributes from SSAS project files.
    '/// </summary>
    '/// <remarks>
    '/// Clean the directory(ies) with a higher degree of control.
    '/// </remarks>
    '/// <param name="ssasProjectPath">SSAS project directory</param>
    '/// <param name="searchPatterns">A CSV list of files to inspect; use an empty string to use "*.cube,*.partitions,*.dsv,*.dim,*.ds,*.role"</param>
    '/// <param name="searchOption">Top level directory only or entire subtree</param>
    '/// <param name="removeDesignTimeNames">Remove "design-time-name" attributes?  These are regenerated each time the SSAS solution is created, reverse engineered from a server, etc.</param>
    '/// <param name="removeDimensionAnnotations">Remove annotations from dimensions?</param>
    '/// <param name="createBackup">Create a backup of any file that is modified?</param>
    '/// <param name="filesInspectedCount">Number of files that were inspected based on the search pattern</param>
    '/// <param name="filesCleanedCount">Number of writeable files that were analyzed</param>
    '/// <param name="filesAlteredCount">Number of files whose contents were modified</param>
    Public Shared Sub CleanSsasProjectDirectory(ByVal ssasProjectPath As String, ByVal searchPatterns As String, ByVal searchOption As SearchOption _
        , ByVal removeDesignTimeNames As Boolean, ByVal removeDimensionAnnotations As Boolean, ByVal createBackup As Boolean _
        , ByRef filesInspectedCount As Integer, ByRef filesCleanedCount As Integer, ByRef filesAlteredCount As Integer)

        '// Validate inputs
        If (Not Directory.Exists(ssasProjectPath)) Then
            Throw New ArgumentException("Please provide SSAS project directory")
        End If

        '// Provide the default list of file extensions
        If (searchPatterns.Trim().Length = 0) Then
            searchPatterns = "*.cube,*.partitions,*.dsv,*.dim,*.ds,*.dmm,*.role"
        End If

        '// Keep track of the changes that we're making
        filesInspectedCount = 0
        filesCleanedCount = 0
        filesAlteredCount = 0
        Dim fileChanged As Boolean = False

        For Each searchPattern As String In searchPatterns.Split(",".ToCharArray())

            '// Iterate over all the files that match the search pattern
            For Each filename As String In Directory.GetFiles(ssasProjectPath, searchPattern, searchOption)

                fileChanged = False
                filesInspectedCount += 1

                '// Check if the file is Read-only; ignore if so
                If (File.GetAttributes(filename) And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                    Debug.WriteLine(String.Format("{0} is read-only; skipping", filename))
                    Continue For
                End If

                filesCleanedCount += 1

                '// Load the file we're cleaning
                Dim document As XmlDocument = New XmlDocument()
                document.Load(filename)

                '// Prepare for our XPath queries
                Dim xmlnsManager As XmlNamespaceManager = LoadSsasNamespaces(document)

                Dim nodeList As XmlNodeList

                '// Clean the elements in the SSAS files that are volatile but unimportant
                '// Stripping these out makes comparing, merging, and analyzing the source files
                '// substantially easier.  These fields are not required for correct operation.
                nodeList = document.SelectNodes("//AS:CreatedTimestamp", xmlnsManager)
                If (XmlHelper.RemoveNodes(nodeList) > 0) Then
                    fileChanged = True
                End If

                nodeList = document.SelectNodes("//AS:LastSchemaUpdate", xmlnsManager)
                If (XmlHelper.RemoveNodes(nodeList) > 0) Then fileChanged = True
                nodeList = document.SelectNodes("//AS:LastProcessed", xmlnsManager)
                If (XmlHelper.RemoveNodes(nodeList) > 0) Then fileChanged = True
                nodeList = document.SelectNodes("//AS:State", xmlnsManager)
                If (XmlHelper.RemoveNodes(nodeList) > 0) Then fileChanged = True
                nodeList = document.SelectNodes("//AS:CurrentStorageMode", xmlnsManager)
                If (XmlHelper.RemoveNodes(nodeList) > 0) Then fileChanged = True

                '// Dimension annotations don't tend to be as volatile as other annotations
                '// We might not want to remove them
                If ((removeDimensionAnnotations) Or (Not filename.EndsWith(".dim", StringComparison.InvariantCulture))) Then

                    nodeList = document.SelectNodes("//AS:Annotations", xmlnsManager)
                    If (XmlHelper.RemoveNodes(nodeList) > 0) Then fileChanged = True
                End If

                '// Remove the 'design-time-name' element, which is regenerated each time the
                '// SSAS project files are regenerated.  This element is not required.
                If (removeDesignTimeNames) Then

                    nodeList = document.SelectNodes("//@dwd:design-time-name/..", xmlnsManager)
                    If (XmlHelper.RemoveAttributes(nodeList, "dwd:design-time-name") > 0) Then fileChanged = True
                    nodeList = document.SelectNodes("//@msprop:design-time-name/..", xmlnsManager)
                    If (XmlHelper.RemoveAttributes(nodeList, "msprop:design-time-name") > 0) Then fileChanged = True
                End If

                '// If we actually changed the file, update the count and save it
                If (fileChanged) Then

                    filesAlteredCount += 1
                    Debug.WriteLine(String.Format("Altered '{0}'", filename))

                    '// Create a backup of the file since we're mucking with it
                    If (createBackup) Then
                        Dim backupFilename As String = filename + "." + DateTime.Now.ToString("yyyyMMddHHmm") + ".bak"
                        File.Copy(filename, backupFilename)
                    End If

                    document.Save(filename)
                End If

                Debug.Write(String.Format("Cleaned '{0}'", filename))

            Next
        Next
        Debug.WriteLine(String.Format("Inspected {0:n0} files", filesInspectedCount))
        Debug.WriteLine(String.Format("Cleaned {0:n0} files", filesCleanedCount))
        Debug.WriteLine(String.Format("Altered {0:n0} files", filesAlteredCount))
    End Sub


    '/// <summary>
    '/// Transform the XML that makes up a given SSAS project file to sort the elements,
    '/// and remove non-essential attributes.  This makes it easy to compare and validate
    '/// files.
    '/// </summary>
    '/// <param name="inputFilename">File to read and transform</param>
    '/// <param name="outputFilename">Transformed file to output</param>
    Public Shared Sub SortSsasFile(ByVal inputFilename As String, ByVal outputFilename As String)

        '// Validate inputs
        If (Not File.Exists(inputFilename)) Then
            Throw New ArgumentException(String.Format("'{0}' does not exist", inputFilename))
        End If

        '// Load an XML document to transform
        Dim xmlDoc As XmlDocument = New XmlDocument()
        xmlDoc.Load(inputFilename)

        '// Load the XSLT to sort the file; pull from a property into an XmlReader
        Dim xslt As String = My.Resources.SsasSortXslt
        Dim xsltReader As StringReader = New StringReader(xslt)
        Dim xmlXslt As XmlReader = XmlReader.Create(xsltReader)

        '// Create the transform
        Dim xslTran As XslCompiledTransform = New XslCompiledTransform()
        xslTran.Load(xmlXslt)

        '// Create a TextWriter to output the Transformed XML document so it will be correctly formated
        Dim output As XmlTextWriter = New XmlTextWriter(outputFilename, System.Text.Encoding.UTF8)
        output.Formatting = Formatting.Indented

        '// Transform and output
        xslTran.Transform(xmlDoc, Nothing, output)
        output.Close()
    End Sub

#Region "Private Helper Methods"
    '/// <summary>
    '/// Split out the Partitions and AggregationDesignss from a base cube
    '/// into their own cube for deserialization into a Partitions file
    '/// </summary>
    '/// <param name="baseCube">Cube to split</param>
    '/// <returns>Cube containing only partitions and aggregations</returns>
    Private Shared Function SplitPartitionCube(ByVal baseCube As Cube) As Cube

        Dim partitionCube As Cube = New Cube()

        For Each mg As MeasureGroup In baseCube.MeasureGroups

            Dim newMG As MeasureGroup = New MeasureGroup(mg.Name, mg.ID)

            If ((mg.Partitions.Count = 0) AndAlso (mg.AggregationDesigns.Count = 0)) Then
                Continue For
            End If
            partitionCube.MeasureGroups.Add(newMG)

            '// Heisenberg principle in action with these objects; use 'for' instead of 'foreach'
            If (mg.Partitions.Count > 0) Then
                For i As Integer = 0 To mg.Partitions.Count - 1
                    Dim partitionCopy As Partition = mg.Partitions(i).Clone()
                    newMG.Partitions.Add(partitionCopy)
                Next i
            End If

            '// Heisenberg principle in action with these objects; use 'for' instead of 'foreach'
            If (mg.AggregationDesigns.Count > 0) Then
                For i As Integer = 0 To mg.AggregationDesigns.Count - 1
                    Dim aggDesignCopy As AggregationDesign = mg.AggregationDesigns(i).Clone()
                    newMG.AggregationDesigns.Add(aggDesignCopy)
                Next i
            End If
        Next mg

        Return partitionCube
    End Function

    '/// <summary>
    '/// Merge a cube containing only Partitions and AggregationDesigns with
    '/// a 'base cube' that contains all Measure Groups present in the 'partition cube'.
    '/// </summary>
    '/// <param name="baseCube">A fully populated cube</param>
    '/// <param name="partitionCube">A cube containing only partitions and aggregation designs</param>
    Private Shared Sub MergePartitionCube(ByVal baseCube As Cube, ByVal partitionCube As Cube)

        Dim baseMG As MeasureGroup = Nothing

        For Each mg As MeasureGroup In partitionCube.MeasureGroups

            baseMG = baseCube.MeasureGroups.Find(mg.ID)

            '// Heisenberg principle in action with these objects; use 'for' instead of 'foreach'
            If (mg.Partitions.Count > 0) Then
                For i As Integer = 0 To mg.Partitions.Count - 1
                    Dim partitionCopy As Partition = mg.Partitions(i).Clone()
                    baseMG.Partitions.Add(partitionCopy)
                Next i
            End If

            '// Heisenberg principle in action with these objects; use 'for' instead of 'foreach'
            If (mg.AggregationDesigns.Count > 0) Then
                For i As Integer = 0 To mg.AggregationDesigns.Count - 1
                    Dim aggDesignCopy As AggregationDesign = mg.AggregationDesigns(i).Clone()
                    baseMG.AggregationDesigns.Add(aggDesignCopy)
                Next i
            End If
        Next
    End Sub

    '/// <summary>
    '/// Load a .Partitions file and the cube object it belongs
    '/// to.  It will add a <Name></Name> node (populated by matching the Partitions
    '/// MeasureGroup ID to the cube's MeasureGroup ID).
    '/// </summary>
    '/// <param name="partitionFilename">Name of the partitions file</param>
    '/// <param name="sourceCube">Cube the .partitions file belongs to</param>
    '/// <returns>XmlReader containing the file</returns>
    Private Shared Function FixPartitionsFileForDeserialize(ByVal partitionFilename As String, ByVal sourceCube As Cube) As XmlReader

        '// Validate inputs
        If (sourceCube Is Nothing) Then
            Throw New ArgumentException("Provide a Cube object that matches the partitions file")
        End If

        If (String.IsNullOrEmpty(partitionFilename)) Then
            Throw New ArgumentException("Provide a partitions file")
        End If

        '// I am NOT validating the extention to provide some extra flexibility here
        Dim document As XmlDocument = New XmlDocument()
        document.Load(partitionFilename)

        '// Setup for XPath queries
        Dim xmlnsManager As XmlNamespaceManager = LoadSsasNamespaces(document)
        Dim defaultNamespaceURI As String = "http://schemas.microsoft.com/analysisservices/2003/engine"

        '// Get all the MeasureGroup IDs
        Dim nodeList As XmlNodeList = document.SelectNodes("/AS:Cube/AS:MeasureGroups/AS:MeasureGroup/AS:ID", xmlnsManager)
        Dim newNode As XmlNode = Nothing

        '// Add a Name node underneath the ID node if one doesn't exist, using the MeasureGroup's real name
        For Each node As XmlNode In nodeList
            '// Verify the node doesn't exist
            If (XmlHelper.NodeExists(node.ParentNode, "Name")) Then
                Continue For
            End If
            newNode = document.CreateNode(XmlNodeType.Element, "Name", defaultNamespaceURI)
            '// Lookup the MG name from the cube based on the ID in the file
            newNode.InnerText = sourceCube.MeasureGroups.Find(node.InnerText).Name
            node.ParentNode.InsertAfter(newNode, node)
        Next

        '// Return this as an XmlReader, so it can be manipulated
        Return New XmlTextReader(New StringReader(document.OuterXml))
    End Function

    '/// <summary>
    '/// Remove the <Name></Name> nodes stored by the serialize method that
    '/// don't belong.
    '/// </summary>
    '/// <param name="partitionFilename">Name of the partitions file</param>
    Private Shared Sub FixSerializedPartitionsFile(ByVal partitionFilename As String)

        '// Validate inputs
        If (String.IsNullOrEmpty(partitionFilename)) Then
            Throw New ArgumentException("Provide a partitions file")
        End If

        '// I am NOT validating the extention to provide some extra flexibility here

        Dim document As XmlDocument = New XmlDocument()
        document.Load(partitionFilename)

        Dim xmlnsManager As XmlNamespaceManager = LoadSsasNamespaces(document)

        Dim nodeList As XmlNodeList = Nothing

        '// Remove the MeasureGroup Names
        nodeList = document.SelectNodes("/AS:Cube/AS:MeasureGroups/AS:MeasureGroup/AS:Name", xmlnsManager)
        XmlHelper.RemoveNodes(nodeList)

        '// Remove the StorageModes
        nodeList = document.SelectNodes("/AS:Cube/AS:MeasureGroups/AS:MeasureGroup/AS:StorageMode", xmlnsManager)
        XmlHelper.RemoveNodes(nodeList)

        '// Remove the ProcessingModes
        nodeList = document.SelectNodes("/AS:Cube/AS:MeasureGroups/AS:MeasureGroup/AS:ProcessingMode", xmlnsManager)
        XmlHelper.RemoveNodes(nodeList)

        document.Save(partitionFilename)
    End Sub

    '/// <summary>
    '/// Remove the <Partitions></Partitions> and <AggregationDesigns></AggregationDesigns>
    '/// elements from the Cube file if they exists; these will be serialized in the Partitions file.
    '/// </summary>
    '/// <param name="cubeFilename">Name of the cube file to</param>
    Private Shared Sub FixSerializedCubeFile(ByVal cubeFilename As String)

        '// Validate inputs
        If (String.IsNullOrEmpty(cubeFilename)) Then
            Throw New ArgumentException("Provide a cube file")
        End If
        '// I am NOT validating the extention to provide some extra flexibility here

        Dim document As XmlDocument = New XmlDocument()
        document.Load(cubeFilename)

        Dim xmlnsManager As XmlNamespaceManager = LoadSsasNamespaces(document)

        Dim nodeList As XmlNodeList = Nothing

        '// Remove the MeasureGroup Names
        nodeList = document.SelectNodes("/AS:Cube/AS:MeasureGroups/AS:MeasureGroup/AS:Partitions", xmlnsManager)
        XmlHelper.RemoveNodes(nodeList)

        '// Remove the StorageModes
        nodeList = document.SelectNodes("/AS:Cube/AS:MeasureGroups/AS:MeasureGroup/AS:AggregationDesigns", xmlnsManager)
        XmlHelper.RemoveNodes(nodeList)

        document.Save(cubeFilename)
    End Sub

    '/// <summary>
    '/// Load the SSAS namespaces into a XmlNameSpaceManager.  These are used for XPath
    '/// queries into a SSAS file
    '/// </summary>
    '/// <param name="document">XML Document to load namespaces for</param>
    '/// <returns>XmlNamespaceManager loaded with SSAS namespaces</returns>
    Private Shared Function LoadSsasNamespaces(ByVal document As XmlDocument) As XmlNamespaceManager

        Dim xmlnsManager As XmlNamespaceManager = New System.Xml.XmlNamespaceManager(document.NameTable)
        '//xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
        xmlnsManager.AddNamespace("xsd", "http://www.w3.org/2001/XMLSchema")
        '//xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
        xmlnsManager.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance")
        '//xmlns:ddl2="http://schemas.microsoft.com/analysisservices/2003/engine/2" 
        xmlnsManager.AddNamespace("ddl2", "http://schemas.microsoft.com/analysisservices/2003/engine/2")
        '//xmlns:ddl2_2="http://schemas.microsoft.com/analysisservices/2003/engine/2/2" 
        xmlnsManager.AddNamespace("ddl2_2", "http://schemas.microsoft.com/analysisservices/2003/engine/2/2")
        '//xmlns:ddl100_100="http://schemas.microsoft.com/analysisservices/2008/engine/100/100" 
        xmlnsManager.AddNamespace("ddl100_100", "http://schemas.microsoft.com/analysisservices/2008/engine/100/100")
        '//xmlns:dwd="http://schemas.microsoft.com/DataWarehouse/Designer/1.0" 
        xmlnsManager.AddNamespace("dwd", "http://schemas.microsoft.com/DataWarehouse/Designer/1.0")
        '//xmlns="http://schemas.microsoft.com/analysisservices/2003/engine"
        xmlnsManager.AddNamespace("AS", "http://schemas.microsoft.com/analysisservices/2003/engine")
        xmlnsManager.AddNamespace("msprop", "urn:schemas-microsoft-com:xml-msprop")
        xmlnsManager.AddNamespace("xs", "http://www.w3.org/2001/XMLSchema")
        xmlnsManager.AddNamespace("msdata", "urn:schemas-microsoft-com:xml-msdata")

        Return xmlnsManager
    End Function
#End Region
End Class

