Imports System.ComponentModel
Imports System.Management.Automation

<RunInstaller(True)> _
Public Class powerSSASSnapin
    Inherits CustomPSSnapIn


    Public Overrides ReadOnly Property Description() As String
        Get
            Return My.Resources.SnapinDescription
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            Return "powerSSAS"
        End Get
    End Property

    Public Overrides ReadOnly Property Vendor() As String
        Get
            Return "Darren Gosbell (c) 2006 http://geekswithblogs.net/darrengosbell"
        End Get
    End Property


    Public Overrides ReadOnly Property Formats() As System.Collections.ObjectModel.Collection(Of System.Management.Automation.Runspaces.FormatConfigurationEntry)
        Get
            Dim fmts As New System.Collections.ObjectModel.Collection(Of System.Management.Automation.Runspaces.FormatConfigurationEntry)
            fmts.Add(New System.Management.Automation.Runspaces.FormatConfigurationEntry(FORMAT_DATA_FOLDER & "powerSSAS.Format.ps1xml"))
            Return fmts
        End Get
    End Property

    Public Overrides ReadOnly Property Providers() As System.Collections.ObjectModel.Collection(Of System.Management.Automation.Runspaces.ProviderConfigurationEntry)
        Get
            Dim pe As New Management.Automation.Runspaces.ProviderConfigurationEntry("powerSSAS", GetType(PowerSSASProvider), HELP_FOLDER & "about_powerSSAS.help.text")
            Dim coll As New System.Collections.ObjectModel.Collection(Of Runspaces.ProviderConfigurationEntry)
            coll.Add(pe)
            Return coll
        End Get
    End Property

    Public Overrides ReadOnly Property Cmdlets() As System.Collections.ObjectModel.Collection(Of System.Management.Automation.Runspaces.CmdletConfigurationEntry)
        Get
            '// Use reflection to load all the cmdlets in this assembly
            Dim coll As New System.Collections.ObjectModel.Collection(Of System.Management.Automation.Runspaces.CmdletConfigurationEntry)
            Dim c As System.Management.Automation.Runspaces.CmdletConfigurationEntry
            Dim ass As System.Reflection.Assembly
            ass = System.Reflection.Assembly.GetExecutingAssembly
            For Each t As Type In ass.GetExportedTypes
                If (GetType(System.Management.Automation.PSCmdlet).IsAssignableFrom(t) _
                    OrElse GetType(System.Management.Automation.Cmdlet).IsAssignableFrom(t)) _
                    AndAlso Not t.IsAbstract Then

                    Dim a As CmdletAttribute = CType(Attribute.GetCustomAttribute(t, GetType(CmdletAttribute)), CmdletAttribute)
                    Dim name As String = a.VerbName & "-" & a.NounName

                    c = New System.Management.Automation.Runspaces.CmdletConfigurationEntry(name, t, HELP_FOLDER & "powerSSAS.Dll-Help.xml")
                    coll.Add(c)
                End If
            Next
            Return coll
        End Get
    End Property

    Public Overrides ReadOnly Property Types() As System.Collections.ObjectModel.Collection(Of System.Management.Automation.Runspaces.TypeConfigurationEntry)
        Get
            Dim typColl As New ObjectModel.Collection(Of Runspaces.TypeConfigurationEntry)
            typColl.Add(New Runspaces.TypeConfigurationEntry(TYPE_DATA_FOLDER & "powerSSAS.Types.Ps1Xml"))
            Return typColl
        End Get
    End Property

End Class
