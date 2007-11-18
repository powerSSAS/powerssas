Imports System.Management.Automation
Imports Gosbell.PowerSSAS.Cmdlets

Namespace Cmdlets
    <Cmdlet(VerbsCommon.Get, "ASConnection")> _
    Public Class cmdletGetASConnection
        Inherits CmdletDiscoverBase

        Private mID As String = ""
        <Parameter(HelpMessage:="Returns the connections with the given ID", Mandatory:=False, Position:=2, ValueFromPipeline:=True)> _
        Public Property ID() As String
            Get
                Return mID
            End Get
            Set(ByVal value As String)
                mID = value
            End Set
        End Property

        Public Overrides Property XmlaRestrictions() As String
            Get
                If MyBase.XmlaRestrictions.Length = 0 AndAlso mID.Length > 0 Then
                    Return "<CONNECTION_ID>" & mID & "</CONNECTION_ID>"
                Else
                    Return ""
                End If
            End Get
            Set(ByVal value As String)
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