#Requires -Version 3.0

<#
    .SYNOPSIS
        Exports a Task Sequence to a XML file

    .DESCRIPTION 
        Exports a Task Sequence to a XML file as used by ConfigMgr 2007 and older.
        
        ConfigMgr 2012 and above use a different format for the default export via the ConfigMgr console
        which consists of several files and folders compressed into an archive which can also contain copies 
        of the referenced packages. This script creates only a XML file which can only get imported back 
        using the corresponding "Import-TaskSequence.ps1" script.

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
    # Specifies the path for the exported Task Sequence xml file.
    [Parameter(Mandatory, ParameterSetName="ID")]
    [Parameter(Mandatory, ParameterSetName="Name")]
    [Alias("FilePath")]
    [string]$Path,

    # Specifies the template for the filename. 
    # Default is "#ID\#ID.xml", where #ID will be replaced with the PackageID and an incrementing number will be added.
    # Additional options that will be replaced automatically are:
    #     - #ID -> Task Sequence PackageID
    #     - #Name -> Task Sequence Name
    #     - #0, #00, #000, #0000, ... -> incrementing number based on the same name
    # If no incrementing number is specified, the parameter Force need to be set 
    # to overwrite an existing file.
    [string]$Filename = "#ID\#ID_#000.xml",

    # Specifies the Task Sequence ID (PackageID).
    # Use either ID or Name to select the Task Sequence.
    [Parameter(Mandatory, ParameterSetName="ID")]
    [ValidateNotNullOrEmpty()]
    [Alias("PackageID")]
    [string[]]$ID,

    # Specifies the Task Sequence Name.
    # Use either Name or ID to select the Task Sequence.
    [Parameter(Mandatory, ParameterSetName="Name")]
    [ValidateNotNullOrEmpty()]
    [string[]]$Name,

    # Specifies if secret information (some passwords and product keys) shall be kept in the XML file.
    [switch]$KeepSecretInformation,

    # Overrides the restriction that prevent the command from succeeding.
    # On default, any existing file will not be overwritten.
    [switch]$Force,

    # Disables the Progress output. 
    # Usefull for automated Export.
    [switch]$NoProgress,

    # Specifies the ConfigMgr Provider Server name. 
    # If no value is specified, the script assumes to be executed on the Site Server.
    [Alias("SiteServer", "ServerName")]
    [string]$ProviderServer = $env:COMPUTERNAME,

    # Specifies the ConfigMgr provider Site Code. 
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

        # Create a connection to the ConfigMgr Provider Server
        if (!($NoProgress.IsPresent)) {Write-Progress -Id 1 -Activity "Exporting Task Sequences .." -Status "Connecting to ConfigMgr ProviderServer" -PercentComplete 0}
        $ConnParams = @{ServerName = $ProviderServer;SiteCode = $SiteCode;}

        if ($PSBoundParameters["Credential"]) {$connParams.Credential = $Credential}
        
        New-CMConnection @ConnParams

        # Prepare parameters for splatting
        $ExportParams = @{Path = $Path; Filename = $Filename}
        if ($KeepSecretInformation.IsPresent) {$ExportParams.KeepSecretInformation = $true}
        if ($NoProgress.IsPresent) {$ExportParams.NoProgress = $true}
        if ($Force.IsPresent) {$ExportParams.Force = $true}

        # Start export based on parameter set
        $Count = 0        
        switch ($PSCmdLet.ParameterSetName) {
            "ID" {
                    ForEach ($TSID In $ID) {
                        if (!($NoProgress.IsPresent)) {Write-Progress -Id 1 -Activity "Exporting Task Sequences .." -Status "Processing Task Sequence $TSID .." -PercentComplete ($Count / $ID.Count * 100)}
                        Start-Export -ID $TSID @ExportParams
                        $Count++
                    }
                }
            "Name" {
                    Foreach ($TSName In $Name) {
                        if (!($NoProgress.IsPresent)) {Write-Progress -Id 1 -Activity "Exporting Task Sequences .." -Status "Processing Task Sequence $TSName .." -PercentComplete ($Count / $Name.Count * 100)}
                        Start-Export -Path $Path -Name $TSName -KeepSecretInformation:$KeepSecretInformation -NoProgress:$NoProgress
                        $Count++
                    }
                }
        }
    }
}

Begin {

    Set-StrictMode -Version Latest

    # Import ConfigMgr Module
    if (-not (Get-Module ConfigMgr)) {Import-Module ConfigMgr}
    
    function Start-Export {
        [CmdLetBinding(SupportsShouldProcess)]
        PARAM(
            # Specifies the path for the exported Task Sequence xml file.
            [Parameter(Mandatory,ParameterSetName="ID")]
            [Parameter(Mandatory,ParameterSetName="Name")]
            [Alias("FilePath")]
            [string]$Path,

            # Specifies the template for the filename. 
            # Default is "#ID\#ID.xml", where #ID will be replaced with the PackageID.
            # Additional options that will be replaced automatically are:
            #     - #ID -> Task Sequence PackageID
            #     - #Name -> Task Sequence Name
            #     - #0, #00, #000, #0000, ... -> incrementing number based on the same name
            # More complex option could be e.g. "#ID\#ID_#000.xml"
            # If no incrementing number is specified, the parameter Force need to be set 
            # to overwrite an existing file.
            [string]$Filename = "#ID\#ID_#000.xml",

            # Specifies the Task Sequence ID (PackageID).
            # Use either ID or Name to select the Task Sequence
            [Parameter(Mandatory,ParameterSetName="ID",ValueFromPipelineByPropertyName)]
            [ValidateNotNullOrEmpty()]
            [Alias("PackageID")]
            [string]$ID,

            # Specifies the Task Sequence Name.
            # Use either Name or ID to select the Task Sequence
            [Parameter(Mandatory,ParameterSetName="Name",ValueFromPipelineByPropertyName)]
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            # Specifies if secret information (mainly passwords and product keys) shall be kept in the XML file
            [switch]$KeepSecretInformation,

            # Overrides the restriction that prevent the command from succeeding.
            # On default, any existing file will not be overwritten.
            [switch]$Force,

            # Specifies the ID of the parent progress 
            [int]$ParentProgressID = -1,

            # Disables the Progress output. 
            # Usefull for automated Export.
            [switch]$NoProgress
        )

        Begin {
            $ProgressParams = @{}

            if (($ParentProgressID -ne $null) -and ($ParentProgressID -gt $null)) {
                $ProgressParams.ParentId = $ParentProgressID
                $ProgressParams.Id = $ParentProgressID + 1
            }
        }

        Process {
            # Get Task Sequence Package
            $ProgressParams.Activity = "Exporting Task Sequence $ID$Name"
            if (!($NoProgress.IsPresent)) {Write-Progress @ProgressParams -Status "Getting Task Sequence Package" -PercentComplete (100/6)}
            if (!([string]::IsNullOrEmpty("ID"))) {
                $TaskSequencePackage = Get-TaskSequencePackage -ID $ID
            } else {
                $TaskSequencePackage = Get-TaskSequencePackage -Name $Name
            }

            # Get Task Sequence
            if ($TaskSequencePackage -ne $null){
                $ProgressParams.Activity = "Exporting Task Sequence ""$($TaskSequencePackage.Name)"" ($($TaskSequencePackage.PackageID))"
                #if (!($NoProgress.IsPresent) -and ($ParentProgressID -gt 0)) {Write-Progress -Id $ParentProgressID -Status "Processing Task Sequence $($TaskSequencePackage.Name) ($($TaskSequencePackage.PackageID)) .."}
                if (!($NoProgress.IsPresent)) {Write-Progress @ProgressParams -Status "Getting Task Sequence from Task Sequence Package" -PercentComplete (100/6*2)}
                $TaskSequence = Get-TaskSequence -TaskSequencePackage $TaskSequencePackage
            }

            if ($TaskSequence -ne $null) {
                # Convert to xml
                if (!($NoProgress.IsPresent)) {Write-Progress @ProgressParams -Status "Converting Task Sequence to XML" -PercentComplete (100/6*3)}
                $TaskSequenceXML = [xml](Convert-TaskSequenceToXML -TaskSequence $TaskSequence -KeepSecretInformation:$KeepSecretInformation)

                if ($TaskSequenceXML -ne $null) {
                    if (!($NoProgress.IsPresent)) {Write-Progress @ProgressParams -Status "Adding Package properties to XML" -PercentComplete (100/6*4)}
                    $Result = Add-PackageProperties -TaskSequencePackage $TaskSequencePackage -TaskSequenceXML $TaskSequenceXML
                    if (!($NoProgress.IsPresent)) {Write-Progress @ProgressParams -Status "Saving xml file to $Path" -PercentComplete (100/6*5)}

                    $TSFilename = Get-Filename -Path $Path -Filename $Filename -PackageID $TaskSequencePackage.PackageID -PackageName $TaskSequencePackage.Name

                    # Ensure Path exits
                    $FullName = Join-Path -Path $Path -ChildPath $TSFilename
                    $ParentPath = Split-Path $FullName -Parent
                    if (!(Test-path($ParentPath))) { New-Item -ItemType Directory -Force -Path $ParentPath | Out-Null}

                    # Save XML file
                    $Result.Save($FullName)
                    if (!($NoProgress.IsPresent)) { Write-Progress @ProgressParams -Status "Done" -PercentComplete 100}
                }
            }
        }
    }

    # Extends the Task Sequence xml document with some default properties of the Task Sequence Package
    function Add-PackageProperties {
        [CmdLetBinding()]
        PARAM (
            # Specifies the Task Sequence Package. 
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [object]$TaskSequencePackage,

            # Specifies the Task Sequence xml document
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [xml]$TaskSequenceXML
        )

        # Some of the properties we use are lazy properties, so we need to get a different instance of the TS Package
        $TaskSequencePackageWMI = Get-TaskSequencePackage -ID ($TaskSequencePackage.PackageID) -IncludeLazyProperties

        # Generate a new xml file
        # To ease the job, lets create a string first

        $Output = "<SmsTaskSequencePackage><BootImageID>"
        $Output += $TaskSequencePackageWMI.BootImageID
        $Output += "</BootImageID><Category>"
        $Output += $TaskSequencePackageWMI.Category
        $Output += "</Category><DependentProgram>"
        $Output += $TaskSequencePackageWMI.DependentProgram
        $Output += "</DependentProgram><Description>"
        $Output += $TaskSequencePackageWMI.Description
        $Output += "</Description><Duration>"
        $Output += $TaskSequencePackageWMI.Duration
        $Output += "</Duration><Name>"
        $Output += $TaskSequencePackageWMI.Name
        $Output += "</Name><ProgramFlags>"
        $Output += $TaskSequencePackageWMI.ProgramFlags
        $Output += "</ProgramFlags><SequenceData>"
        $Output += $TaskSequenceXML.sequence.OuterXml
        $Output += "</SequenceData><SourceDate>"
        $SourceDate = [Management.ManagementDateTimeConverter]::ToDateTime($TaskSequencePackageWMI.SourceDate).ToString("yyy-MM-ddThh:mm:ss")
        $Output += $SourceDate
        $Output += "</SourceDate><SupportedOperatingSystems>"
        #$Output += $TaskSequencePackageWMI.SupportedOperatingSystems
        $Output += "</SupportedOperatingSystems><IconSize>"
        $Output += $TaskSequencePackageWMI.IconSize
        $Output += "</IconSize></SmsTaskSequencePackage>"

        return [xml]$Output
    }

    # Generates a file name for the Task Sequence xml file
    function Get-Filename {
        [CmdLetBinding()]
        PARAM(
            # Specifies the path for the exported Task Sequence xml file.
            [string]$Path,

            # Specifies the template for the filename.
            [string]$Filename,

            # Specifies the Task Sequence PackageID
            [string]$PackageID,

            # Specifies the Task Sequence name
            [string]$PackageName
        )

        # Build Filename
        # Replace #ID and #Name
        $TSFilename = $Filename.Replace("#ID", ($PackageID)).Replace("#Name", ($Name))

        # Replace #0... if necessary
        if ($TSFilename.Contains("#0")) {
            # Get amount of 0
            $StartPos = $TSFilename.IndexOf("#0")
            $MaxLength = $TSFilename.Length - $StartPos - 1
            $ZeroCount = 1
            for ($Count = 1; $Count -le $MaxLength; $Count++) {
                $ZeroCount = $Count
                if ($TSFilename.Substring($StartPos+$Count,1) -ne "0") {
                    $ZeroCount--
                    Break
                }
            }

            # Get files that match the pattern
            $TSFiles = @(Get-ChildItem -Path $Path -Filter $TSFilename.Replace($TSFilename.Substring($StartPos,$ZeroCount + 1), "*") -ErrorAction SilentlyContinue)

            if ($TSFiles -ne $null) {
                $TSFileCount = $TSFiles.Count
            } else {
                $TSFileCount = 0
            }
            $TSFileCount ++

            # Generate Filename
            # escape certain characters for formatting
            $TSFilename = $TSFilename.Replace("0.", "0'.'").Replace("\","'\'").Replace("#0", "0")
            $TSFilename = "{0:$TSFilename}" -f $TSFileCount

        }

        $TSFilename
    }
}