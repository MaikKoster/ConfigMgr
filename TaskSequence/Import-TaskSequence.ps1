#Requires -Version 3.0

<#
    .SYNOPSIS
        Imports a Task Sequence from a XML file

    .DESCRIPTION 
        Imports a Task Sequence from a XML file as used by ConfigMgr 2007 and older.
        
        ConfigMgr 2012 and above use a different format for the default export via the ConfigMgr console
        which consists of several files and folders compressed into an archive which can also contain copies 
        of the referenced packages.

    .LINK
        http://maikkoster.com/

    .NOTES
        Copyright (c) 2015 Maik Koster

        Author:  Maik Koster
        Version: 1.0
        Date:    15.03.2016

        Version History:
            1.0 - 15.03.2016 - Published script
            

#>
[CmdLetBinding(SupportsShouldProcess)]
PARAM (
    # Specifies the name and path of the Task Sequence xml file.
    [Parameter(Mandatory, ParameterSetName="ID")]
    [Parameter(Mandatory, ParameterSetName="Create")]
    [Parameter(Mandatory, ParameterSetName="Name")]
    [ValidateScript({Test-Path $_ -PathType 'Leaf'})]
    [Alias("FilePath")]
    [string]$Path,

    # Specifies the Task Sequence ID (PackageID).
    # Use either ID or Name to select the Task Sequence.
    [Parameter(Mandatory, ParameterSetName="ID")]
    [ValidateNotNullOrEmpty()]
    [Alias("PackageID")]
    [string]$ID,

    # Specifies if a new Task Sequence Package shall be created.
    # Use the Name parameter to specify the name of the new Task Sequence.
    [Parameter(Mandatory,ParameterSetName="Create")]
    [switch]$Create,

    # Specifies the name of an existing or the new Task Sequence
    [Parameter(Mandatory, ParameterSetName="Name")]
    [Parameter(ParameterSetName="Create")]
    [ValidateNotNullOrEmpty()]
    [string]$Name,

    # Specifies the description of the new Task Sequence
    [Parameter(ParameterSetName="Create")]
    [string]$Description,

    # Specifies the ConfigMgr Provider Server name. 
    # If no value is specified, the script assumes to be executed on the Site Server.
    [Alias("SiteServer", "ServerName")]
    [string]$ProviderServer = $env:COMPUTERNAME,

    # Specifies the ConfigMgr Provider Site Code. 
    # If no value is specified, the script will evaluate it from the Site Server.
    [string]$SiteCode,

    # Specifies the credentials to connect to the ConfigMgr Provider Server.
    [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty
)


Process {

    ###############################################################################
    # Start Script
    ###############################################################################
    
    # Ensure this isn't processed when dot sourced by e.g. Pester Test trun
    if ($MyInvocation.InvocationName -ne '.') {
        # Path has been validated already
        # TODO: Add Replace logic
        $TaskSequencePackageXML = [xml](Get-Content -Path $Path)
        
        if ($TaskSequencePackageXML -ne $null) {
            # Create a connection to the ConfigMgr Provider Server
            $ConnParams = @{
                ServerName = $ProviderServer;
                SiteCode = $SiteCode;
            }

            if ($PSBoundParameters["Credential"]) {
                $connParams.Credential = $Credential
            }
        
            New-CMConnection @ConnParams
        
            switch ($PSCmdLet.ParameterSetName) {
                "Create" {Start-Import -Path $Path -Create -Name $Name -Description $Description}
                "ID" {Start-Import -Path $Path -ID $ID}
                "Name" {Start-Import -Path $Path -Name}
            }
        } else {
            Write-Error "File $Path is empty or not a valid xml file."
        }
    }

}

Begin {

    Set-StrictMode -Version Latest

    # Import ConfigMgr Module
    if (-not (Get-Module ConfigMgr)) {Import-Module ConfigMgr}
    
    function Start-Import {
        [CmdLetBinding(SupportsShouldProcess)]
        PARAM(
            [Parameter(Mandatory, ParameterSetName="ID")]
            [Parameter(Mandatory, ParameterSetName="Create")]
            [Parameter(Mandatory, ParameterSetName="Name")]
            [ValidateNotNullOrEmpty()]
            [xml]$TaskSequencePackageXML,

            # TaskSequencePackage ID
            [Parameter(Mandatory, ParameterSetName="ID")]
            [ValidateNotNullOrEmpty()]
            [string]$ID,

            # If specified, a new TaskSequencePackage with the supplied Name will be created
            [Parameter(Mandatory,ParameterSetName="Create")]
            [switch]$Create,

            # Name of the existing or new TaskSequencePackage
            [Parameter(Mandatory, ParameterSetName="Name")]
            [Parameter(ParameterSetName="Create")]
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            [Parameter(ParameterSetName="Create")]
            [string]$Description
        )

        # Get the Sequence node
        # Try to find a sequence node (ConfigMgr 2007 and Export-TaskSequence.ps1 format
        $TaskSequenceXML = $TaskSequencePackageXML.SelectSingleNode('//sequence')

        if ($TaskSequenceXML -eq $null) {
            # ConfigMgr 2012 format (object.xml from expanded zip file)
            $TaskSequenceXML.SelectSingleNode('//PROPERTY[@NAME="Sequence"]').VALUE
        }

        if ($TaskSequenceXML -ne $null) {
            # Convert XML to Sequence
            $TaskSequence = Convert-XMLToTaskSequence -TaskSequenceXML $TaskSequenceXML

            # Create a new Task Sequence package if requested
            if ($Create.IsPresent){
                if ([string]::IsNullOrEmpty($Name)) {
                    $Name = Get-XmlNodeText -XmlDocument $TaskSequencePackageXML -NodePath "SmsTaskSequencePackage.Name"
                    if ([string]::IsNullOrEmpty($Name)) {
                        $Name = "New Task Sequence"   
                    }
                }

                if ([string]::IsNullOrEmpty($Description)) {
                    $Description = Get-XmlNodeText -XmlDocument $TaskSequencePackageXML -NodePath "SmsTaskSequencePackage.Description"
                }

                $Properties = @{}
                    
                # Add additional info if available
                $BootImageID = Get-XmlNodeText -XmlDocument $TaskSequencePackageXML -NodePath "SmstaskSequencePackage.BootImageID"
                if (!([string]::IsNullOrEmpty($BootImageID))) {
                    $Properties.BootImageID = $BootImageID
                }
                    
                $Category = Get-XmlNodeText -XmlDocument $TaskSequencePackageXML -NodePath "SmstaskSequencePackage.Category"
                if (!([string]::IsNullOrEmpty($Category))) {
                    $Properties.Category = $Category
                }
                    
                $Duration = Get-XmlNodeText -XmlDocument $TaskSequencePackageXML -NodePath "SmstaskSequencePackage.Duration"
                if (!([string]::IsNullOrEmpty($Duration))) {
                    $Properties.Duration = $Duration
                }
                    
                $DependentProgram = Get-XmlNodeText -XmlDocument $TaskSequencePackageXML -NodePath "SmstaskSequencePackage.DependentProgram"
                if (!([string]::IsNullOrEmpty($DependentProgram))) {
                    $Properties.DependentProgram = $DependentProgram
                }

                $TaskSequencePackage = New-TaskSequencePackage -Name $Name -Description $Description -TaskSequence $TaskSequence -Properties $Properties

            } else {
                if (!([string]::IsNullOrEmpty($ID))) {
                    $TaskSequencePackage = Get-TaskSequencePackage -ID $ID
                } elseif (!([string]::IsNullOrEmpty($Name))) {
                    $TaskSequencePackage = Get-TaskSequencePackage -Name $Name
                }

                if (($TaskSequence -ne $null) -and ($TaskSequencePackage -ne $null)) {
                    Set-TaskSequence -TaskSequencePackage $TaskSequencePackage -TaskSequence $TaskSequence
                }
            }
        }
    }

    function Get-XmlNode([ xml ]$XmlDocument, [string]$NodePath, [string]$NamespaceURI = "", [string]$NodeSeparatorCharacter = '.') {
        # If a Namespace URI was not given, use the Xml document's default namespace.
        if ([string]::IsNullOrEmpty($NamespaceURI)) { $NamespaceURI = $XmlDocument.DocumentElement.NamespaceURI }   
     
        # In order for SelectSingleNode() to actually work, we need to use the fully qualified node path along with an Xml Namespace Manager, so set them up.
        $XmlNsManager = New-Object System.Xml.XmlNamespaceManager($XmlDocument.NameTable)
        $XmlNsManager.AddNamespace("ns", $NamespaceURI)
        $FullyQualifiedNodePath = "/ns:$($NodePath.Replace($($NodeSeparatorCharacter), '/ns:'))"
     
        # Try and get the node, then return it. Returns $null if the node was not found.
        $Node = $XmlDocument.SelectSingleNode($FullyQualifiedNodePath, $XmlNsManager)
        return $Node
    }

    function Get-XmlNodeText([ xml ]$XmlDocument, [string]$NodePath, [string]$NamespaceURI = "", [string]$NodeSeparatorCharacter = '.') {
        # Try and get the node
        $Node = Get-XmlNode -XmlDocument $XmlDocument -NodePath $NodePath -NamespaceURI $NamespaceURI -NodeSeparatorCharacter $NodeSeparatorCharacter

        if ($Node -ne $null) {
            return $Node.InnerText
        } else {
            return ""
        }
        
    }
}