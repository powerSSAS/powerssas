Imports MDXParser.MDXParser
Imports Microsoft.AnalysisServices
Imports System.Management.Automation

<Cmdlet(VerbsLifecycle.Invoke, "AsMdxAnalyze")> _
Public Class cmdletInvokeAsMdxAnalyze
    Inherits Cmdlet

    Private mMdx As String
    <Parameter()> _
    Public Property MDX() As String
        Get
            Return mMdx
        End Get
        Set(ByVal value As String)
            mMdx = value
        End Set
    End Property

    Protected Overrides Sub ProcessRecord()
        MyBase.ProcessRecord()

        Dim parser As New MDXParser.MDXParser(mMdx)
        parser.Parse()
        Dim analyzer As MDXParser.Analyzer = parser.Analyze()
        For Each msg As MDXParser.Message In analyzer.Messages
            Dim res As New AnalyzeResult
            res.Text = msg.Text
            res.URL = msg.URL
            res.Severity = msg.Severity
            res.Type = msg.Type.ToString()
            res.Line = msg.Location.Line
            res.Column = msg.Location.Column
            WriteObject(res)
        Next
    End Sub

End Class

Friend Class AnalyzeResult
    Public Severity As Integer
    Public Text As String
    Public Line As Integer
    Public Column As Integer
    Public URL As String
    Public Type As String
End Class