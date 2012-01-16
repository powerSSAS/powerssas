Imports System.Reflection
Imports System.ComponentModel
Imports System.Management.Automation
Imports System.Management.Automation.Provider
Imports System.Management.Automation.runspaces
Imports Microsoft.AnalysisServices
Imports Gosbell.PowerSSAS.Types
Imports System.Collections.ObjectModel

<CmdletProvider("powerSSAS", ProviderCapabilities.ShouldProcess)> _
Public Class PowerSSASProvider
    Inherits NavigationCmdletProvider

    Private Const PATH_SEPARATOR As String = "\"

    'Protected Overrides Sub [Stop]()
    '    MyBase.[Stop]()
    '    If Not Me.PSDriveInfo Is Nothing Then
    '        Dim amoServer As Server = CType(PSDriveInfo, AmoDriveInfo).AmoServer
    '        If amoServer.Connected Then
    '            amoServer.Disconnect()
    '        End If
    '    End If
    'End Sub

    Protected Overrides Function NewDrive(ByVal drive As PSDriveInfo) As PSDriveInfo

        If drive Is Nothing Then
            WriteError(New ErrorRecord(New ArgumentNullException("drive"), "NullDrive", ErrorCategory.InvalidArgument, Nothing))
            Return Nothing
        End If

        If String.IsNullOrEmpty(drive.Root) Then
            WriteError(New ErrorRecord(New ArgumentException("drive.Root"), "NoRoot", ErrorCategory.InvalidArgument, drive))
            Return Nothing
        End If
        WriteProgress(New ProgressRecord(0, "new-psdrive", "creating server object"))
        Dim amoDrive As AmoDriveInfo = New AmoDriveInfo(drive)
        Dim amoServer As New Microsoft.AnalysisServices.Server
        WriteProgress(New ProgressRecord(0, "new-psdrive", "connecting to server"))
        amoServer.Connect("Data Source=" & drive.Root)
        amoDrive.AmoServer = amoServer

        WriteProgress(New ProgressRecord(0, "new-psdrive", "connected to " & drive.Root))
        Dim func As String = "function " & amoDrive.Name & ": { set-location " & amoDrive.Name & ": }"
        Me.InvokeCommand.InvokeScript(func, False, PipelineResultTypes.None, Nothing, Nothing)

        Return amoDrive
    End Function

    Protected Overrides Function RemoveDrive(ByVal drive As PSDriveInfo) As PSDriveInfo
        If drive Is Nothing Then
            WriteError(New ErrorRecord(New ArgumentNullException("drive"), "NullDrive", ErrorCategory.InvalidArgument, Nothing))
            Return Nothing
        End If
        Dim amoDrive As AmoDriveInfo = CType(drive, AmoDriveInfo)

        ' removes the function 
        Dim func As String = "remove-item function:\" & amoDrive.Name & ":"
        Me.InvokeCommand.InvokeScript(func, False, PipelineResultTypes.None, Nothing, Nothing)

        ' disconnect from AMO
        amoDrive.AmoServer.Disconnect()
        amoDrive.AmoServer.Dispose()

        Return amoDrive
    End Function

    Protected Overrides Sub InvokeDefaultAction(ByVal path As String)
        Try
            MyBase.InvokeDefaultAction(path)
        Catch ex As Exception
            Me.WriteError(New ErrorRecord(ex, "InvokeDefaultAction", ErrorCategory.NotSpecified, path))
        End Try
    End Sub

    Protected Overrides Sub SetItem(ByVal path As String, ByVal value As Object)
        Try
            MyBase.SetItem(path, value)
        Catch ex As Exception
            Me.WriteError(New ErrorRecord(ex, "setItem", ErrorCategory.NotSpecified, path))
        End Try
    End Sub


    Protected Overrides Function IsValidPath(ByVal path As String) As Boolean
        Return True
    End Function

    Protected Overrides Sub GetItem(ByVal path As String)
        Dim isContainer As Boolean = False
        Dim oItem As Object = GetCurrentItem(path)
        If TypeOf oItem Is ICollection Then isContainer = True
        WriteItemObject(oItem, path, isContainer)
    End Sub

    Protected Overrides Function ItemExists(ByVal path As String) As Boolean

        If pathIsDrive(path) Then Return True

        Dim itm As Object
        itm = GetCurrentItem(path)
        If itm Is Nothing Then
            Return False
        Else
            Return True
        End If
    End Function

    Protected Overrides Sub GetChildItems(ByVal path As String, ByVal recurse As Boolean)
        Dim itm As Object
        Dim isCont As Boolean = False
        itm = GetCurrentItem(path)

        Dim itmColl As ICollection = TryCast(itm, ICollection)
        If Not itmColl Is Nothing Then
            For Each o As Object In itmColl
                If TypeOf o Is ICollection _
                OrElse TypeOf o Is MajorObject _
                OrElse TypeOf o Is Microsoft.AnalysisServices.Command Then
                    isCont = True
                Else
                    isCont = False
                End If
                WriteItemObject(o, path & PATH_SEPARATOR & o.ToString, isCont)
                If recurse And isCont Then
                    GetChildItems(path & PATH_SEPARATOR & o.ToString, recurse)
                End If
            Next
        Else
            Dim col As Dictionary(Of String, CollectionProperty) = GetPropertyCollection(itm)
            Dim isContainer As Boolean = False

            For Each skey As String In col.Keys
                If TypeOf col.Item(skey).Value Is ICollection Then
                    isContainer = True
                Else
                    isContainer = False
                End If

                '// The Owning collection property points at the parent collection
                '// but it just creates confusion in the context of a provider.
                If skey <> "OwningCollection" Then
                    WriteItemObject(col.Item(skey), path & PATH_SEPARATOR & skey, isContainer)
                End If
            Next

        End If
    End Sub

    Protected Overrides Sub GetChildNames(ByVal path As String, ByVal returnContainers As System.Management.Automation.ReturnContainers)
        MyBase.GetChildNames(path, returnContainers)
        'TODO - GetChildNames
    End Sub

    Protected Overrides Function HasChildItems(ByVal path As String) As Boolean
        Dim itm As Object = GetCurrentItem(path)
        Dim coll As ICollection = TryCast(itm, ICollection)
        If Not coll Is Nothing Then
            If coll.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function


    Protected Overrides Function Start(ByVal providerInfo As System.Management.Automation.ProviderInfo) As System.Management.Automation.ProviderInfo
        Dim myProviderInfo As ProviderInfo = MyBase.Start(providerInfo)

        With myProviderInfo
            .Description = "AMO PowerShell Provider"
        End With

        Return myProviderInfo
    End Function

    Protected Overrides Function IsItemContainer(ByVal path As String) As Boolean
        'Return MyBase.IsItemContainer(path)
        'Dim curItem As Object = GetCurrentItem(path)
        'If pathIsDrive(path) _
        'OrElse TypeOf curItem Is ICollection _
        'OrElse TypeOf curItem Is Database _
        'orlese typeof curitem is Then

        '// Always return true so that we can navigate down to an objects properties
        '// ,particularly when using a collection property using PowerShell Analyzer.
        Return True

        'Else
        '    Return False
        'End If
    End Function

    Protected Overrides Function MakePath(ByVal parent As String, ByVal child As String) As String
        Try
            Return MyBase.MakePath(parent, child)
        Catch ex As Exception
            Me.WriteError(New ErrorRecord(ex, "MakePath", ErrorCategory.NotSpecified, parent))
            Return ""
        End Try
    End Function

    Protected Overrides Function NormalizeRelativePath(ByVal path As String, ByVal basePath As String) As String
        Try
            Return MyBase.NormalizeRelativePath(path, basePath)
        Catch ex As Exception
            Me.WriteError(New ErrorRecord(ex, "NormalizeRelativePath", ErrorCategory.NotSpecified, path))
            Return path
        End Try
    End Function

    Public Shared Function GetXmlaCreate(ByVal majorobj As MajorObject) As Xml.XPath.IXPathNavigable ' Xml.XmlDocument
        Dim scr As New Microsoft.AnalysisServices.Scripter
        Dim sbOut As New Text.StringBuilder
        Dim xmlOut As Xml.XmlWriter = Xml.XmlWriter.Create(sbOut)
        scr.ScriptCreate(New MajorObject() {majorobj}, xmlOut, True)
        Dim xmlFrag As New Xml.XmlDocument
        xmlFrag.LoadXml(sbOut.ToString)
        Return xmlFrag
    End Function

#Region "Unsupported operations"
    Protected Overrides Sub NewItem(ByVal path As String, ByVal itemTypeName As String, ByVal newItemValue As Object)
        'MyBase.NewItem(path, type, newItemValue)
        ThrowTerminatingError(New ErrorRecord(New NotImplementedException("NewItem not implemented (" & path & ")"), "powerSSAS.NewItem", ErrorCategory.NotImplemented, Nothing))
    End Sub

    Protected Overrides Sub CopyItem(ByVal path As String, ByVal copyPath As String, ByVal recurse As Boolean)
        'MyBase.CopyItem(path, copyPath, recurse)
        ThrowTerminatingError(New ErrorRecord(New NotImplementedException("CopyItem not implemented (" & path & ")"), "powerSSAS.CopyItem", ErrorCategory.NotImplemented, Nothing))
    End Sub

    Protected Overrides Sub RemoveItem(ByVal path As String, ByVal recurse As Boolean)
        Dim o As Object = GetCurrentItem(path)
        'Dim w As XmlaWarningCollection
        'Dim res As ImpactDetailCollection

        Dim mo As MajorObject = TryCast(o, MajorObject)
        If Not mo Is Nothing Then
            'TODO: implement Me.ShouldProcess()
            If recurse Then
                Me.WriteWarning("Recurse not implemented - Would delete: " & mo.Name & " (ID: " & mo.ID & ")")
                'mo.Drop(DropOptions.AlterOrDeleteDependents, w, res)
                'return
            Else
                If ShouldProcess(path, "Drop") Then
                    mo.Drop()
                End If
                Return
            End If
        End If
        ThrowTerminatingError(New ErrorRecord(New NotImplementedException("Only Major  (" & path & ")"), "powerSSAS.RemoveItem", ErrorCategory.NotImplemented, Nothing))
    End Sub
#End Region

#Region " Helper Functions "

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")> _
    Private Function GetAmoDriveInternal(ByVal path As String) As AmoDriveInfo
        Dim pathBits() As String = path.Split(PATH_SEPARATOR.ToCharArray())
        If Me.PSDriveInfo Is Nothing Then
            For Each drv As PSDriveInfo In Me.ProviderInfo.Drives
                '\\ Check for default instance
                If drv.Root = pathBits(0) Then
                    Return CType(drv, AmoDriveInfo)
                End If
                '\\ check for named instances
                If pathBits.Length > 1 AndAlso drv.Root = pathBits(0) & "\" & pathBits(1) Then
                    Return CType(drv, AmoDriveInfo)
                End If
            Next
            Throw New ArgumentException("The drive for the path " & path & " could not be found")
        Else
            Return CType(Me.PSDriveInfo, AmoDriveInfo)
        End If
    End Function

    Private Function GetCurrentItem(ByVal path As String) As Object
        Dim pathChunks() As String = chunkPath(path)
        Dim itm As Object = GetAmoDriveInternal(path).AmoServer
        For i As Integer = 0 To pathChunks.Length - 1
            '\\ the root of the drive will return a blank pathChunk
            If pathChunks(i).Length > 0 Then
                If TypeOf itm Is ICollection Then
                    itm = GetCollectionItem(itm, pathChunks(i))
                ElseIf TypeOf itm Is ClrAssembly _
                AndAlso CaseInsensitiveMatch(pathChunks(i), "functions") Then
                    itm = Utils.FunctionLister.ListFunctions(CType(itm, Microsoft.AnalysisServices.Assembly))
                ElseIf TypeOf itm Is Server _
                AndAlso CaseInsensitiveMatch(pathChunks(i), "sessions") Then
                    Dim xds As New Utils.XmlaDiscoverSessions
                    itm = xds.Discover("DISCOVER_SESSIONS", CType(itm, Server).Name, Nothing, Nothing)
                Else
                    itm = GetProperty(itm, pathChunks(i))
                End If
            End If
        Next
        Return itm
    End Function

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic")> Private Function CaseInsensitiveMatch(ByVal string1 As String, ByVal string2 As String) As Boolean
        Return (String.Compare(string1, string2, True, System.Threading.Thread.CurrentThread.CurrentUICulture) = 0)
    End Function

    'Private Function GetDriveRootInternal(ByVal path As String) As String
    '    Dim driveRoot As String = ""
    '    If Me.PSDriveInfo Is Nothing Then

    '    Else
    '        driveRoot = Me.PSDriveInfo.Root
    '    End If


    '    Return driveRoot

    '    'Dim pathBits As String()
    '    'If path.EndsWith(PATH_SEPARATOR) Then
    '    '    path = path.Substring(0, path.Length - PATH_SEPARATOR.Length)
    '    'End If
    '    'pathBits = path.Split(CType(PATH_SEPARATOR, Char))

    '    'If pathBits.Length >= 2 AndAlso pathBits(1) <> "Databases" Then
    '    '    Return pathBits(0) & PATH_SEPARATOR & pathBits(1)
    '    'Else
    '    '    Return pathBits(0)
    '    'End If
    'End Function

    Private Function chunkPath(ByVal path As String) As String()
        Dim s As String = path.Replace((GetAmoDriveInternal(path).Root + PATH_SEPARATOR), "")
        Return s.Split(PATH_SEPARATOR.ToCharArray)
    End Function

    Private Function pathIsDrive(ByVal path As String) As Boolean
        Dim s As String = ""
        Try
            s = path.Replace((GetAmoDriveInternal(path).Root + PATH_SEPARATOR), "")
        Catch
            s = ""
        End Try
        If (s.Length = 0) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Shared Function NormalizePath(ByVal path As String) As String
        Dim result As String = path
        If Not String.IsNullOrEmpty(path) Then
            result = path.Replace("/", PATH_SEPARATOR)
        End If
        Return result
    End Function

    Private Function GetCollectionItem(ByVal obj As Object, ByVal itemName As String) As Object
        Dim methGetByName As MethodInfo
        Dim propIndexOf As PropertyInfo
        Dim propItem As PropertyInfo

        methGetByName = obj.GetType.GetMethod("GetByName", New Type() {GetType(String)})
        If Not methGetByName Is Nothing Then
            '\\ We have a GetByName method with a single string parameter
            Return methGetByName.Invoke(obj, New Object() {itemName})
        Else
            propItem = obj.GetType.GetProperty("Item", Nothing, New Type() {GetType(String)}, Nothing)
            If Not propItem Is Nothing Then
                '\\ We have an Item property with a string parameter
                Return propItem.GetValue(obj, New Object() {itemName})
            Else

                propItem = obj.GetType.GetProperty("Item", Nothing, New Type() {GetType(Integer)}, Nothing)
                propIndexOf = obj.GetType.GetProperty("IndexOf", GetType(Integer), New Type() {GetType(String)}, Nothing)
                If (Not propItem Is Nothing) AndAlso (Not propIndexOf Is Nothing) Then
                    '\\ we have an Item property and an indexOf property
                    Dim itemIndex As Integer = CType(propIndexOf.GetValue(obj, New Object() {itemName}), Integer)
                    Return propItem.GetValue(obj, New Object() {itemIndex})
                End If
            End If
        End If


        Dim lst As IList = TryCast(obj, IList)
        If Not lst Is Nothing Then
            If lst.Count = 1 Then Return lst.Item(0)

            Return lst.Item(lst.IndexOf("itemName"))
        ElseIf TypeOf obj Is ICollection Then
            ThrowTerminatingError(New ErrorRecord(New NotImplementedException(itemName & " (ICollection) : " & obj.ToString()), "powerSSAS.GetCollectionItem", ErrorCategory.NotImplemented, Nothing))
        Else
            ThrowTerminatingError(New ErrorRecord(New NotImplementedException(itemName & " : " & obj.ToString()), "powerSSAS.GetCollectionItem", ErrorCategory.NotImplemented, Nothing))
        End If
        Return Nothing
    End Function

    Private Function GetPropertyCollection(ByVal obj As Object) As Dictionary(Of String, CollectionProperty)
        Dim props As New Dictionary(Of String, CollectionProperty)
        Dim propCol As PropertyInfo()
        propCol = obj.GetType.GetProperties(BindingFlags.Instance Or BindingFlags.Public)

        Dim clrAss As ClrAssembly = TryCast(obj, ClrAssembly)
        If Not clrAss Is Nothing Then
            If Not props.ContainsKey("Functions") Then
                Dim Functions As ICollection(Of FunctionSignature) = Utils.FunctionLister.ListFunctions(clrAss)
                props.Add("Functions", New CollectionProperty("Functions", Functions, GetType(FunctionSignature)))
                Return props
            End If
        End If

        If TypeOf obj Is Server Then
            If Not props.ContainsKey("Sessions") Then
                Dim xds As New Utils.XmlaDiscoverSessions

                Dim sess As Collection(Of Object) = xds.Discover("DISCOVER_SESSIONS", CType(Me.PSDriveInfo, AmoDriveInfo).AmoServer.Name, Nothing, Nothing)
                props.Add("Sessions", New CollectionProperty("Sessions", sess, GetType(Session)))
            End If
        End If

        For Each prop As PropertyInfo In propCol
            'If TypeOf prop.GetValue(obj, Nothing) Is ICollection Then
            If Not props.ContainsKey(prop.Name) Then
                props.Add(prop.Name, New CollectionProperty(prop.Name, prop.GetValue(obj, Nothing), prop))
            End If
            'End If
        Next

        Return props

    End Function

    Private Shared Function GetProperty(ByVal obj As Object, ByVal propertyName As String) As Object
        Dim propInfo As PropertyInfo = obj.GetType.GetProperty(propertyName)
        If propInfo Is Nothing Then
            '\\ GetProperty is case sensitive so we should search the 
            '\\ entire properties collection to double check
            Dim propCol As PropertyInfo()
            propCol = obj.GetType.GetProperties(BindingFlags.Instance Or BindingFlags.Public)
            For Each prop As PropertyInfo In propCol
                If String.Compare(prop.Name, propertyName, StringComparison.CurrentCultureIgnoreCase) = 0 Then
                    Return prop.GetValue(obj, Nothing)
                End If
            Next

            Return Nothing

        Else
            Return propInfo.GetValue(obj, Nothing)
        End If

    End Function

#End Region

#Region "Class AmoDriveInfo"

    Friend Class AmoDriveInfo
        Inherits System.Management.Automation.PSDriveInfo
        Private mAmoServer As Microsoft.AnalysisServices.Server
        'Private mScripter As Microsoft.AnalysisServices.Scripter

        Public Sub New(ByVal amoDriveInfo As System.Management.Automation.PSDriveInfo)
            MyBase.New(amoDriveInfo)
        End Sub


        Public Property AmoServer() As Microsoft.AnalysisServices.Server
            Get
                Return mAmoServer
            End Get
            Set(ByVal value As Microsoft.AnalysisServices.Server)
                mAmoServer = value
            End Set
        End Property


        'Public ReadOnly Property ScriptEngine() As Microsoft.AnalysisServices.Scripter
        '    Get
        '        If mScripter Is Nothing Then
        '            mScripter = New Microsoft.AnalysisServices.Scripter
        '        End If

        '        Return mScripter
        '    End Get
        'End Property

    End Class

#End Region

#Region "Class CollectionProperty"
    Private Class CollectionProperty
        Private mName As String = ""
        Private mValue As Object
        Private mType As Type


        Public Sub New(ByVal propertyName As String, ByVal propertyValue As Object, ByVal propInfo As Type)
            mName = propertyName
            mValue = propertyValue
            mType = propInfo
        End Sub
        Public Sub New(ByVal propertyName As String, ByVal propertyValue As Object, ByVal propInfo As System.Reflection.PropertyInfo)
            mName = propertyName
            mValue = propertyValue
            mType = propInfo.PropertyType
        End Sub

        <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")> Public ReadOnly Property Name() As String
            Get
                Return mName
            End Get
        End Property

        Public ReadOnly Property Value() As Object
            Get
                Return mValue
            End Get
        End Property

        <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")> Public ReadOnly Property ObjectType() As Type
            Get
                Return mType
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return mName
        End Function

    End Class
#End Region

End Class

