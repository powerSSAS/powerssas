
Imports System.Data
Imports System
Imports Microsoft.AnalysisServices
Imports System.Reflection
Imports Gosbell.PowerSSAS.Types

Namespace Utils
    Public NotInheritable Class FunctionLister

        Private Sub New()
        End Sub

        Private Shared lstFuncs As List(Of FunctionSignature)

        Public Shared Function ListFunctions(ByVal ass As Microsoft.AnalysisServices.Assembly) As ICollection(Of FunctionSignature)
            Return getFunctionList(ass)
        End Function

        Private Shared Function getFunctionList(ByVal ass As Microsoft.AnalysisServices.Assembly) As List(Of FunctionSignature)

            If lstFuncs Is Nothing Then
                lstFuncs = New List(Of FunctionSignature)

                Dim clrAss As ClrAssembly = TryCast(ass, ClrAssembly)
                If Not clrAss Is Nothing Then '// we can only use reflection against .Net assemblies

                    For Each f As ClrAssemblyFile In clrAss.Files '// an assembly can have multiple files

                        '// We only want to get the "main" asembly file and only files which have data
                        '// (Some of the system assemblies appear to be registrations only and do not
                        '// have any data.
                        If (f.Data.Count > 0 AndAlso f.Type = ClrAssemblyFileType.Main) Then

                            '// assembly the assembly back into a single byte from the blocks of base64 strings
                            Dim rawAss(0) As Byte
                            Dim iPos As Integer = 0
                            Dim buff As Byte()

                            For Each block As String In f.Data
                                buff = System.Convert.FromBase64String(block)
                                System.Array.Resize(rawAss, rawAss.Length + buff.Length)
                                buff.CopyTo(rawAss, iPos)
                                iPos += buff.Length
                            Next

                            '// use reflection to extract the public types and methods from the 
                            '// re-assembled assembly.
                            Dim asAss As System.Reflection.Assembly = System.Reflection.Assembly.ReflectionOnlyLoad(rawAss)

                            AddHandler AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve, AddressOf CurrentDomain_ReflectionOnlyAssemblyResolve

                            Dim assTypes As Type() = asAss.GetTypes()
                            For i As Integer = 0 To assTypes.Length - 1

                                Dim t2 As Type = assTypes(i)
                                If (t2.IsPublic) Then

                                    Dim methods As MethodInfo()
                                    methods = t2.GetMethods(BindingFlags.DeclaredOnly Or BindingFlags.Public Or BindingFlags.Static Or BindingFlags.Instance)
                                    Dim paramCnt As Integer = 0
                                    For Each meth As MethodInfo In methods

                                        '// build the parameter signature as a string
                                        Dim params() As ParameterInfo
                                        Dim paramList As System.Text.StringBuilder = New System.Text.StringBuilder
                                        Dim returnTypeName As String = String.Empty
                                        Try
                                            params = meth.GetParameters()
                                            paramList = New System.Text.StringBuilder()
                                            paramCnt = params.Length
                                            '// add the first parameter
                                            If (paramCnt > 0) Then
                                                paramList.Append(params(0).ToString())
                                            End If
                                            '// add subsequent parameters, inserting a comma before each new one.
                                            For j As Integer = 1 To paramCnt - 1
                                                paramList.Append(", ")
                                                paramList.Append(params(j).ToString())
                                            Next j
                                            returnTypeName = StripNamespace(meth.ReturnType.ToString())
                                        Catch ex As Exception
                                            paramList = New System.Text.StringBuilder
                                            paramList.Append(" ** NOT AVAILABLE **")

                                        End Try

                                        Try
                                            returnTypeName = StripNamespace(meth.ReturnType.ToString())
                                        Catch ex As Exception
                                            returnTypeName = " ** NOT AVAILABLE **"
                                        End Try

                                        Dim funcSig As FunctionSignature = New FunctionSignature
                                        With funcSig
                                            .Assembly = ass.Name
                                            .ClassName = t2.Name
                                            .Method = meth.Name
                                            .ReturnType = returnTypeName
                                            .Parameters = paramList.ToString()
                                        End With
                                        If Not lstFuncs.Contains(funcSig) Then
                                            lstFuncs.Add(funcSig)
                                        End If
                                    Next meth
                                End If
                            Next i

                        End If

                    Next f
                End If

                Return lstFuncs
            Else
                Return lstFuncs
            End If
        End Function

        Private Shared Function CurrentDomain_ReflectionOnlyAssemblyResolve(ByVal sender As Object, ByVal args As ResolveEventArgs) As System.Reflection.Assembly
            Try
                Dim ass As System.Reflection.Assembly

                ass = System.Reflection.Assembly.ReflectionOnlyLoad(args.Name)
                If Not ass Is Nothing Then
                    Return ass
                End If
            Catch

            End Try
            Try
                Dim parts As String() = args.Name.Split(CType(",", Char))
                Dim file As String = System.IO.Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) & "\" & parts(0).Trim & ".dll"
                Return System.Reflection.Assembly.ReflectionOnlyLoadFrom(file)
            Catch
                Return Nothing
            End Try
        End Function

        '// This function strips off any leading namespaces from a type name
        '// to make it more concise.
        Private Shared Function StripNamespace(ByVal typeName As String) As String
            Dim lastDot As Integer = typeName.LastIndexOf(".", StringComparison.InvariantCulture)
            If (lastDot > 0) Then
                Return typeName.Substring(lastDot + 1)
            Else
                Return typeName
            End If
        End Function


    End Class

End Namespace