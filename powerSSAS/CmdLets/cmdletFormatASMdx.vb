Imports System.Management.Automation
Imports System.Xml
Imports Microsoft.AnalysisServices
Imports MDXParser

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Format, "ASMdx")> _
    Public Class cmdletFormatASMdx
        Inherits Cmdlet

        Private mMDX As String = ""
        <Parameter()> _
        Public Property MDX() As String
            Get
                Return mMDX
            End Get
            Set(ByVal value As String)
                mMDX = value
            End Set
        End Property

        Private mLineWidth As Integer = 80
        <Parameter()> _
        Public Property LineWidth() As Integer
            Get
                Return mLineWidth
            End Get
            Set(ByVal value As Integer)
                mLineWidth = value
            End Set
        End Property

        Private mCommaBeforeNewLine As SwitchParameter
        <Parameter()> _
        Public Property CommaBeforeNewLine() As SwitchParameter
            Get
                Return mCommaBeforeNewLine
            End Get
            Set(ByVal value As SwitchParameter)
                mCommaBeforeNewLine = value
            End Set
        End Property

        Private mOutput As MDXParser.OutputFormat
        <Parameter()> _
        Public Property Output() As MDXParser.OutputFormat
            Get
                Return mOutput
            End Get
            Set(ByVal value As MDXParser.OutputFormat)
                mOutput = value
            End Set
        End Property

        Protected Overrides Sub ProcessRecord()

            Dim parser As New MDXParser.MDXParser(mMDX)
            Dim options As New FormatOptions()
            options.CommaBeforeNewLine = mCommaBeforeNewLine.IsPresent
            options.Output = mOutput
            options.LineWidth = mLineWidth
            'options.Indent = 
            parser.Parse()
            WriteObject(parser.FormatMDX(options))

        End Sub
    End Class
End Namespace
