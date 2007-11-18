Namespace Types
    Public Class FunctionSignature
        Private mAssembly As String
        Private mClassName As String
        Private mMethod As String
        Private mReturnType As String
        Private mParameters As String

        Public Property Assembly() As String
            Get
                Return mAssembly
            End Get
            Set(ByVal value As String)
                mAssembly = value
            End Set
        End Property
        Public Property ClassName() As String
            Get
                Return mClassName
            End Get
            Set(ByVal value As String)
                mClassName = value
            End Set
        End Property

        Public Property Method() As String
            Get
                Return mMethod
            End Get
            Set(ByVal value As String)
                mMethod = value
            End Set
        End Property
        Public Property ReturnType() As String
            Get
                Return mReturnType
            End Get
            Set(ByVal value As String)
                mReturnType = value
            End Set
        End Property
        Public Property Parameters() As String
            Get
                Return mParameters
            End Get
            Set(ByVal value As String)
                mParameters = value
            End Set
        End Property
        Public Overrides Function ToString() As String
            Return mMethod
        End Function
    End Class
End Namespace