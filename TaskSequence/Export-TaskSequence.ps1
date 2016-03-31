#Requires -Version 3.0
#Requires -Modules ConfigMgr

<#
    .SYNOPSIS
        Exports a Task Sequence as an XML file

    .DESCRIPTION 
        Exports a Task Sequence as an XML file as used by ConfigMgr 2007 and older.
        
        ConfigMgr 2012 and above use a different format for the default export via the ConfigMgr console
        which consists of several files and folders compressed into an archive which can also contain copies 
        of the referenced packages. This script creates only a XML file which can only get imported back 
        using the corresponding "Import-TaskSequence.ps1" script.

    .EXAMPLE
        .\Export-TaskSequence.ps1 -Path "%Temp%\Export" -ID "TST00001"
    
        Export a Task Sequence to a subfolder with incrementing suffix.

    .EXAMPLE
        .\Export-TaskSequence.ps1 -Path "%Temp%\Export" -ID "TST00001" -Filename "#ID.xml" -Force
    
        Export a Task Sequence and overwrite the last exported file.

    .EXAMPLE
        .\Export-TaskSequence.ps1 -Path "%Temp%\Export" -ID "TST00001", "TST00002", "TST00003" -NoProgress
    
        Export several Task Sequences without progress information. E.g. as regular backup task.

    .EXAMPLE
        .\Export-TaskSequence.ps1 -Path "%Temp%\Export" -ID "TST00001" -ProviderServer TSTCM01 -SiteCode TST -Credentials (Get-Credential)
        
        Export a Task Sequence using different credentials and Provider server settings

    .LINK
        http://maikkoster.com/
        https://github.com/MaikKoster/ConfigMgr/blob/master/TaskSequence/Export-TaskSequence.ps1

    .NOTES
        Copyright (c) 2016 Maik Koster

        Author:  Maik Koster
        Version: 1.0
        Date:    31.03.2016

        Version History:
            1.0 - 31.03.2016 - Published script
            
#>
[CmdLetBinding(SupportsShouldProcess,DefaultParameterSetName="ID")]
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
    [Alias("PackageID", "TaskSequenceID")]
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

    # Enables the Progress output. 
    [switch]$ShowProgress,

    # Specifies if the script should pass through the path to the export file
    [switch]$PassThru,

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
        if (!($ShowProgress.IsPresent)) {Write-Progress -Id 1 -Activity "Exporting Task Sequences .." -Status "Connecting to ConfigMgr ProviderServer" -PercentComplete 0}
        $ConnParams = @{ServerName = $ProviderServer;SiteCode = $SiteCode;}

        if ($PSBoundParameters["Credential"]) {$connParams.Credential = $Credential}
        
        New-CMConnection @ConnParams

        # Prepare parameters for splatting
        $ExportParams = @{Path = $Path; Filename = $Filename}
        if ($KeepSecretInformation.IsPresent) {$ExportParams.KeepSecretInformation = $true}
        if ($Force.IsPresent) {$ExportParams.Force = $true}
        if ($ShowProgress.IsPresent) {$ExportParams.ShowProgress = $true}
        if ($PassThru.IsPresent) {$ExportParams.PassThru = $true}

        # Start export based on parameter set
        $Count = 0        
        switch ($PSCmdLet.ParameterSetName) {
            "ID" {
                    ForEach ($TSID In $ID) {
                        Set-Progress -ShowProgress:($ShowProgress.IsPresent) -Activity "Exporting Task Sequences .." -Status "Processing Task Sequence $TSID .." -TotalSteps $ID.Count -Step $Count
                        Start-Export -ID $TSID @ExportParams
                        $Count++
                    }
                }
            "Name" {
                    Foreach ($TSName In $Name) {
                        Set-Progress -ShowProgress:($ShowProgress.IsPresent) -Activity "Exporting Task Sequences .." -Status "Processing Task Sequence $TSName .." -TotalSteps $Name.Count -Step $Count
                        Start-Export -Name $TSName @ExportParams
                        $Count++
                    }
                }
        }
    }
}

Begin {

    Set-StrictMode -Version Latest
    
    # Starts the Export process. 
    # Moved to separate function to properly test the execution using Pester.
    Function Start-Export {
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
            [Alias("PackageID", "TaskSequenceID")]
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

            # Enables the Progress output. 
            [switch]$ShowProgress,

            # Specifies if the script should pass through the path to the export file
            [switch]$PassThru
        )

        Begin {
            $ProgressParams = @{
                TotalSteps = 6
                ShowProgress = $ShowProgress.IsPresent
            }

            if (($ParentProgressID -ne $null) -and ($ParentProgressID -gt 0)) {
                $ProgressParams.ParentId = $ParentProgressID
                $ProgressParams.Id = $ParentProgressID + 1
            }
        }

        Process {
            # Get Task Sequence Package
            $ProgressParams.Activity = "Exporting Task Sequence $ID$Name"
            Set-Progress @ProgressParams -Status "Getting Task Sequence Package" -Step 1
            if (!([string]::IsNullOrEmpty("ID"))) {
                Write-Verbose "Start Export Process for Task Sequence Package $ID."
                $TaskSequencePackage = Get-TaskSequencePackage -ID $ID -IncludeLazy
            } else {
                Write-Verbose "Start Export Process for Task Sequence Package '$Name'."
                $TaskSequencePackage = Get-TaskSequencePackage -Name $Name -IncludeLazy
            }

            # Get Task Sequence
            if ($TaskSequencePackage -ne $null){
                $ProgressParams.Activity = "Exporting Task Sequence '$($TaskSequencePackage.Name)' ($($TaskSequencePackage.PackageID))"
                #if (!($NoProgress.IsPresent) -and ($ParentProgressID -gt 0)) {Write-Progress -Id $ParentProgressID -Status "Processing Task Sequence $($TaskSequencePackage.Name) ($($TaskSequencePackage.PackageID)) .."}
                Set-Progress @ProgressParams -Status "Getting Task Sequence from Task Sequence Package" -Step 2
                $TaskSequence = Get-TaskSequence -TaskSequencePackage $TaskSequencePackage
            }

            if ($TaskSequence -ne $null) {
                # Convert to xml
                Set-Progress @ProgressParams -Status "Converting Task Sequence to XML" -Step 3
                $TaskSequenceXML = [xml](Convert-TaskSequenceToXML -TaskSequence $TaskSequence -KeepSecretInformation:$KeepSecretInformation)

                if ($TaskSequenceXML -ne $null) {
                    Set-Progress @ProgressParams -Status "Adding Package properties to XML" -Step 4
                    $Result = Add-PackageProperties -TaskSequencePackage $TaskSequencePackage -TaskSequenceXML $TaskSequenceXML
                    Set-Progress @ProgressParams -Status "Saving xml file to $Path" -Step 5

                    $TSFilename = Get-Filename -Path $Path -Filename $Filename -PackageID $TaskSequencePackage.PackageID -PackageName $TaskSequencePackage.Name

                    # Ensure Path exits
                    $FullName = Join-Path -Path $Path -ChildPath $TSFilename
                    $ParentPath = Split-Path $FullName -Parent
                    if (!(Test-path($ParentPath))) { New-Item -ItemType Directory -Force -Path $ParentPath | Out-Null}

                    # Save XML file
                    if ($PSCmdLet.ShouldProcess("Write Task Sequence xml to file '$FullName'.", "Write XML File")) {
                        Write-Verbose "Write Task Sequence xml to file '$Fullname'."
                        if ((Test-Path($FullName)) -and (-not ($Force.IsPresent))) {
                            Write-Warning "File '$Fullname' exists already and Force isn't set. Won't overwrite existing file."
                        } else {
                            $Result.Save($FullName)
                        }
                        Set-Progress @ProgressParams -Status "Done" -Step 6
                    }
                    if ($PassThru.IsPresent) {$FullName}
                }
            }
        }
    }

    # Extends the Task Sequence xml document with some default properties of the Task Sequence Package
    Function Add-PackageProperties {
        [CmdLetBinding(SupportsShouldProcess)]
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

        # Generate a new xml file
        # To ease the job, lets create a string first

        $Output = "<SmsTaskSequencePackage><BootImageID>"
        $Output += $TaskSequencePackage.BootImageID
        $Output += "</BootImageID><Category>"
        $Output += $TaskSequencePackage.Category
        $Output += "</Category><DependentProgram>"
        $Output += $TaskSequencePackage.DependentProgram
        $Output += "</DependentProgram><Description>"
        $Output += $TaskSequencePackage.Description
        $Output += "</Description><Duration>"
        $Output += $TaskSequencePackage.Duration
        $Output += "</Duration><Name>"
        $Output += $TaskSequencePackage.Name
        $Output += "</Name><ProgramFlags>"
        $Output += $TaskSequencePackage.ProgramFlags
        $Output += "</ProgramFlags><SequenceData>"
        $Output += $TaskSequenceXML.sequence.OuterXml
        $Output += "</SequenceData><SourceDate>"
        $Output += $TaskSequencePackage.SourceDate.ToString("yyyy-MM-ddThh:mm:ss")
        $Output += "</SourceDate><SupportedOperatingSystems>"
        #TODO Implement support for SupportedOperatingSystems
        #$Output += $TaskSequencePackageWMI.SupportedOperatingSystems
        $Output += "</SupportedOperatingSystems><IconSize>"
        $Output += $TaskSequencePackage.IconSize
        $Output += "</IconSize></SmsTaskSequencePackage>"

        return [xml]$Output
    }

    # Generates a file name for the Task Sequence xml file
    Function Get-Filename {
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
            $Pattern = $TSFilename.Replace($TSFilename.Substring($StartPos,$ZeroCount + 1), "*")
            $TSFiles = @(Get-ChildItem -Path $Path -Filter $Pattern -ErrorAction SilentlyContinue)

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