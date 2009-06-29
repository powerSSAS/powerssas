Imports System.Management.Automation
Imports Microsoft.AnalysisServices.AdomdClient
Imports Gosbell.PowerSSAS.Cmdlets

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASConnection")> _
    Public Class CmdletGetASConnection
        Inherits CmdletDiscoverBase

        Private mID As String = ""
        <Parameter(HelpMessage:="Returns the connections with the given ID", Mandatory:=False, Position:=2, ValueFromPipeline:=True)> _
        Public Property Id() As String
            Get
                Return mID
            End Get
            Set(ByVal value As String)
                mID = value
            End Set
        End Property

        Public Overrides Property XmlaRestrictions() As AdomdRestrictionCollection
            Get
                If MyBase.XmlaRestrictions.Count = 0 AndAlso mID.Length > 0 Then
                    MyBase.XmlaRestrictions.Add("CONNECTION_ID", mID)
                    Return MyBase.XmlaRestrictions
                Else
                    Return Nothing
                End If
            End Get
            Set(ByVal value As AdomdRestrictionCollection)
                MyBase.XmlaRestrictions = value
            End Set
        End Property

        Public Overrides ReadOnly Property DiscoverRowset() As String
            Get
                Return "DISCOVER_CONNECTIONS"
            End Get
        End Property
    End Class
End Namespace