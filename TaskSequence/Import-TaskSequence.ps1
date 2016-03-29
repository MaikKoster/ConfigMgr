#Requires -Version 3.0
#Requires -Modules ConfigMgr

<#
    .SYNOPSIS
        Imports a Task Sequence from a XML file

    .DESCRIPTION 
        Imports a Task Sequence from a XML file as used by ConfigMgr 2007 and older.
        
        ConfigMgr 2012 and above use a different format for the default export via the ConfigMgr console
        which consists of several files and folders compressed into an archive which can also contain copies 
        of the referenced packages.

    .EXAMPLE
        .\Import-TaskSequence.ps1 -Path "C:\Backup\TST00001_001.xml" -ID "TST00001"

        Import and replace an existing Task Sequence.

    .EXAMPLE
        .\Import-TaskSequence.ps1 -Path "C:\Backup\TST00001_001.xml" -ID "TST00001" -ShowProgress -PassThru

        Import and replace an existing Task Sequence while showing the progress of the operation and returning the updated Task Sequence package object.

    .EXAMPLE
        .\Import-TaskSequence.ps1 -Path "C:\Backup\TST00001_001.xml" -Create -Name "Save the World"

        Import Task Sequence and create a new Task Sequence Package.

    .EXAMPLE
        .\Import-TaskSequence.ps1 -Path "C:\Backup\TST00001_001.xml" -Create -Name "Save the World" -ProviderServer "CM01" -SiteCode "TST" -Credential (Get-Credential)

        Import Task Sequence and create a new Task Sequence Package using specific ConfigMgr connection.

    .LINK
        http://maikkoster.com/
        https://github.com/MaikKoster/ConfigMgr/blob/master/TaskSequence/Import-TaskSequence.ps1

    .NOTES
        Copyright (c) 2016 Maik Koster

        Author:  Maik Koster
        Version: 1.0
        Date:    22.03.2016

        Version History:
            1.0 - 22.03.2016 - Published script
            
#>
[CmdLetBinding(SupportsShouldProcess,DefaultParameterSetName="ID")]
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

    # Enable the Progress output. 
    [switch]$ShowProgress,

    # Specifies if the script should pass through the updated or created task sequence package
    [switch]$PassThru,

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
        Set-Progress -ShowProgress ($ShowProgress.IsPresent) -Activity "Import Task Sequence '$Path'" -Status "Get content from '$Path'" -TotalSteps 4 -Step 1
        $TaskSequencePackageXML = [xml](Get-Content -Path $Path)
        
        if ($TaskSequencePackageXML -ne $null) {
            # Create a connection to the ConfigMgr Provider Server
            $ConnParams = @{
                ServerName = $ProviderServer;
                SiteCode = $SiteCode;
            }

            if ($PSBoundParameters["Credential"]) {$connParams.Credential = $Credential}

            $ProcessParams = @{TaskSequencePackageXML = $TaskSequencePackageXML}
            if ($ShowProgress.IsPresent) {$ProcessParams.ShowProgress = $true}
            if ($PassThru.IsPresent) {$ProcessParams.PassThru = $true}
        
            New-CMConnection @ConnParams
        
            switch ($PSCmdLet.ParameterSetName) {
                "Create" {Start-Import @ProcessParams -Create -Name $Name -Description $Description}
                "ID" {Start-Import @ProcessParams -ID $ID}
                "Name" {Start-Import @ProcessParams -Name $Name}
            }
        } else {
            Write-Error "File '$Path' is empty or not a valid xml file."
        }
    }
}

Begin {
    Set-StrictMode -Version Latest

    # Import ConfigMgr Module
    if (-not (Get-Module ConfigMgr)) {Import-Module ConfigMgr}
    
    function Start-Import {
        [CmdLetBinding(SupportsShouldProcess,DefaultParameterSetName="ID")]
        PARAM(
            # Specifies the xml document of the exported Task Sequence Package
            [Parameter(Mandatory, ParameterSetName="ID")]
            [Parameter(Mandatory, ParameterSetName="Create")]
            [Parameter(Mandatory, ParameterSetName="Name")]
            [ValidateNotNullOrEmpty()]
            [xml]$TaskSequencePackageXML,

            # Specifies the Task Sequence PackageID
            [Parameter(Mandatory, ParameterSetName="ID")]
            [ValidateNotNullOrEmpty()]
            [Alias("PackageID")]
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
            [string]$Description,

            # Enables the Progress output. 
            [switch]$ShowProgress,

            # Specifies if the script should pass through the updated or created task sequence package
            [switch]$PassThru
        )

        Begin {
            # Prepare Progress parameters
            $ProgressParams = @{
                Activity = "Import Task Sequence $ID$Name"
                TotalSteps = 4
                ShowProgress = $ShowProgress.IsPresent
            }
        }
        
        Process {
            Write-Verbose "Start Import Process for Task Sequence '$ID$Name'."
            # Get the Sequence node
            # Try to find a sequence node (ConfigMgr 2007 and Export-TaskSequence.ps1 format
            $TaskSequenceXML = $TaskSequencePackageXML.SelectSingleNode('//sequence')

            if ($TaskSequenceXML -eq $null) {
                # ConfigMgr 2012 format (object.xml from expanded zip file)
                $TaskSequenceXML.SelectSingleNode('//PROPERTY[@NAME="Sequence"]').VALUE
            }

            if ($TaskSequenceXML -ne $null) {
                # Convert XML to Sequence
                Set-Progress @ProgressParams -Status "Convert xml to Task Sequence." -Step 2
                $TaskSequence = Convert-XMLToTaskSequence -TaskSequenceXML $TaskSequenceXML

                # Create a new Task Sequence package if requested
                if ($Create.IsPresent){
                    Write-Verbose "Create new Task Sequence Package."
                    if ([string]::IsNullOrEmpty($Name)) {
                        $Name = Get-XmlNodeText -XmlDocument $TaskSequencePackageXML -NodePath "SmsTaskSequencePackage.Name"
                        if ([string]::IsNullOrEmpty($Name)) {
                            $Name = "New Task Sequence"   
                            Write-Verbose "No name supplied. Use '$Name'."
                        } else {
                            Write-Verbose "Use Name '$Name'."
                        }
                    }

                    if ([string]::IsNullOrEmpty($Description)) {
                        $Description = Get-XmlNodeText -XmlDocument $TaskSequencePackageXML -NodePath "SmsTaskSequencePackage.Description"
                        Write-Verbose "Use Description '$Description'."
                    }

                    $Properties = @{}
                    
                    # Add additional info if available
                    $BootImageID = Get-XmlNodeText -XmlDocument $TaskSequencePackageXML -NodePath "SmstaskSequencePackage.BootImageID"
                    if (!([string]::IsNullOrEmpty($BootImageID))) {
                        $Properties.BootImageID = $BootImageID
                        Write-Verbose "Use BootImageID '$BootImageID'."
                    }
                    
                    $Category = Get-XmlNodeText -XmlDocument $TaskSequencePackageXML -NodePath "SmstaskSequencePackage.Category"
                    if (!([string]::IsNullOrEmpty($Category))) {
                        $Properties.Category = $Category
                        Write-Verbose "Use Category '$Category'."
                    }
                    
                    $Duration = Get-XmlNodeText -XmlDocument $TaskSequencePackageXML -NodePath "SmstaskSequencePackage.Duration"
                    if (!([string]::IsNullOrEmpty($Duration))) {
                        $Properties.Duration = $Duration
                        Write-Verbose "Use Duration '$Duration'."
                    }
                    
                    $DependentProgram = Get-XmlNodeText -XmlDocument $TaskSequencePackageXML -NodePath "SmstaskSequencePackage.DependentProgram"
                    if (!([string]::IsNullOrEmpty($DependentProgram))) {
                        $Properties.DependentProgram = $DependentProgram
                        Write-Verbose "Use DependentProgram '$DependentProgram'."
                    }

                    Set-Progress @ProgressParams -Status "Create new Task Sequence Package." -Step 4

                    $Params = @{
                        Name = $Name
                        Description = $Description
                        TaskSequence = $TaskSequence
                        Property = $Properties
                    }
                    if ($PassThru.IsPresent) {$Params.PassThru = $true}
                    $TaskSequencePackage = New-TaskSequencePackage @Params

                    if ($PassThru.IsPresent) {
                        # Drop TaskSequencePackage back to pipeline
                        $TaskSequencePackage
                    }
                } else {
                    Set-Progress @ProgressParams -Status "Get Task Sequence Package." -Step 3
                    if (!([string]::IsNullOrEmpty($ID))) {
                        $TaskSequencePackage = Get-TaskSequencePackage -ID $ID
                    } elseif (!([string]::IsNullOrEmpty($Name))) {
                        $TaskSequencePackage = Get-TaskSequencePackage -Name $Name
                    }

                    if (($TaskSequence -ne $null) -and ($TaskSequencePackage -ne $null)) {
                        Set-Progress @ProgressParams -Status "Update Task Sequence Package." -Step 4
                        $Result = Set-TaskSequence -TaskSequencePackage $TaskSequencePackage -TaskSequence $TaskSequence -PassThru

                        if (($Result -ne $null) -and ($Result.ReturnValue -eq 0)) {
                            Write-Verbose "Successfully updated Task Sequence Package '$($TaskSequencePackage.PackageID)'."
                            if ($PassThru.IsPresent) {
                                # Drop TaskSequencePackage back to pipeline
                                $TaskSequencePackage = Get-TaskSequencePackage -ID $ID
                                $TaskSequencePackage
                            }
                        } else {
                            Write-Error "Failed to update Task Sequence Package '$($TaskSequencePackage.PackageID)'."
                        }
                    } else{
                        Write-Error "Failed to get Task Sequence Package and/or Task Sequence."
                    }
                }
            } else {
                Write-Error "Unable to find Task Sequence in supplied xml object."
            }
        }
    }

    # Returns a xml node based on the supplied path or null if the node does not exist
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

    # Returns the inner text of a xml node based on the supplied path or an empty string if the node does not exist.
    function Get-XmlNodeText([ xml ]$XmlDocument, [string]$NodePath, [string]$NamespaceURI = "", [string]$NodeSeparatorCharacter = '.') {
        # Try and get the node
        $Node = Get-XmlNode @PSBoundParameters

        if ($Node -ne $null) {
            return $Node.InnerText
        } else {
            return ""
        }
    }

    # Wraps the Write-Progress CmdLet
    Function Set-Progress {
        [CmdLetBinding()]
        PARAM(
            # Specifies the current ID of the Progress output
            [int]$ID = 1,

            # Specifies the Parent ID of the Progress output
            [int]$ParentID = -1,

            # Specifies the total amount of steps
            [int]$TotalSteps,

            # Specifies the current step
            [int]$Step = 1,

            # Specifies the current status
            [string]$Status,

            # Specifies the current Activity
            [string]$Activity,

            # Specifies if Progress shall be shown
            [bool]$ShowProgress = $False
        )

        if ($ShowProgress) {
            $PercentComplete = (100/$TotalSteps*$Step)
            if ($ParentID -gt 0) {
                Write-Progress -ParentId $ParentID -Id $ID -Status $Status -Activity $Activity -PercentComplete $PercentComplete
            } else {
                Write-Progress -Id $ID -Status $Status -Activity $Activity -PercentComplete $PercentComplete
            }
        }
    }
}