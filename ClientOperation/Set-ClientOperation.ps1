#Requires -Version 3.0
#Requires -Modules ConfigMgr

<#
    .SYNOPSIS
        Starts, cancels or deletes a Client Operation

    .DESCRIPTION 
        Starts, cancels or deletes a System Center Configuration Manager (SCCM) Client Operation
        like a SCEP FullScan, Software Update Evaluation or Computer Policy download and evaluation.

    .EXAMPLE
        .\Set-ClientOperation.ps1 -Start FullScan -ResourceName TestComputer
    
    Starts a Full Scan on one computer

    .EXAMPLE
        .\Set-ClientOperation.ps1 -Start QuickScan -CollectionName "All Windows 10 machines"
        
        Starts a QuickScan on all members of a collection

    .EXAMPLE
        .\Set-ClientOperation.ps1 -Cancel -All -ProviderServer CMServer01 -SiteCode XYZ
        
        Cancels All Client Operations from a remote computer

    .EXAMPLE
        .\Set-ClientOperation.ps1 -Delete -Expired -ProviderServer CMServer01 -SiteCode XYZ -Credential (Get-Credential CMAdmin)

        Deletes all Expired Client Operations from a remote computer with different credentials

    .LINK
        http://maikkoster.com/start-cancel-or-delete-a-sccm-client-operation/
        https://github.com/MaikKoster/ConfigMgr/blob/master/ClientOperation/Set-ClientOperation.ps1

    .NOTES
        Copyright (c) 2015 Maik Koster

        Author:  Maik Koster
        Version: 1.1
        Date:    08.12.2015

        Version History:
            1.0 - 07.12.2015 - Published script
            1.1 - 08.12.2015 - Updated Help. Fixed Whatif issue
            1.2 - 25.03.2015 - Moved common ConfigMgr functions to ConfigMgr Module
                             - Added standalone version

#>
[CmdletBinding(SupportsShouldProcess)]
Param (
    # Starts the specified Client Operation. Valid values are:
    #    FullScan                ->  to intiate a SCEP Full Scan
    #    QuickScan               ->  to initiate a SCEP Quick Scan
    #    DownloadDefinition      ->  to initiate a SCEP definition download
    #    EvaluateSoftwareUpdates ->  to initiate a software Update evaluation cycle
    #    RequestComputerPolicy   ->  to initiate a donwload and process of the current computer policy
    #    RequestUserPolicy       ->  to initiate a download and process of the current User policy
    [Parameter(Mandatory,ParameterSetname="Start")]
    [ValidateSet("FullScan", "QuickScan", "DownloadDefinition", "EvaluateSoftwareUpdates", "RequestComputerPolicy", "RequestUserPolicy")]
    [string]$Start,

    # Cancels a running client Operation. OperationID must be supplied.
    [Parameter(Mandatory,ParameterSetname="Cancel")]
    [Parameter(Mandatory,ParameterSetname="CancelAll")]
    [switch]$Cancel,

    # Deletes a Client Operation. OperationID must be supplied.
    [Parameter(Mandatory,ParameterSetname="Delete")]
    [Parameter(Mandatory,ParameterSetname="DeleteAll")]
    [Parameter(Mandatory,ParameterSetname="DeleteExpired")]
    [switch]$Delete,

    # CollectionID of the target(s) for the Client Operation.
    # Supply either CollectionID or CollectionName. 
    # If no ResourceID/ResourceName is specified, all Collection members will be targeted.
    # If no CollectionID is supplied, "SMS00001" (All Systems) will be used,
    # but at least one ResourceID/ResourceName must be specified.
    [Parameter(ParameterSetname="Start")]
    [ValidateNotNullOrEmpty()]
    [string]$CollectionID,

    # Collection name of the targets for the Client Operation.
    # Supply either CollectionName or CollectionID
    # If no ResourceID/ResourceName is specified, all Collection members will be targeted.
    # If no CollectionName is supplied, "All Systems" (SMS00001) will be used,
    # but at least one ResourceID/ResourceName must be specified.
    [Parameter(ParameterSetname="Start")]
    [ValidateNotNullOrEmpty()]
    [string]$CollectionName,

    # ResourceID of the target for the Client operation. 
    # Supply either ResourceId or ResourceName.
    # Multiple ResourceIDs can be supplied.
    [Parameter(ParameterSetname="Start")]
    [uint32[]]$ResourceID,

    # ResourceName of the target for the client operation.
    # Supply either ResourceName or ResourceID.
    # Multiple ResourceNames can be supplied.
    [Parameter(ParameterSetname="Start")]
    [string[]]$ResourceName,
    
    # The OperationID of the Client Operation that shall be canceled or deleted.
    # Multiple OperationIDs can be supplied.
    [Parameter(Mandatory,ParameterSetname="Cancel")]
    [Parameter(Mandatory,ParameterSetname="Delete")]
    [uint32[]]$OperationID,

    [Parameter(Mandatory,ParameterSetname="CancelAll")]
    [Parameter(Mandatory,ParameterSetname="DeleteAll")]
    # If supplied all Client Operations will be canceled or deleted.
    [switch]$All,

    [Parameter(Mandatory,ParameterSetname="DeleteExpired")]
    # If supplied all expired Client Operations will be deleted.
    [switch]$Expired,


    # The ConfigMgr Provider Server name. 
    # If no value is specified, the script assumes to be executed on the Site Server.
    [Alias("SiteServer", "ServerName")]
    [string]$ProviderServer = ".",

    # The ConfigMgr provider Site Code. 
    # If no value is specified, the script will evaluate it from the Site Server.
    [string]$SiteCode,

    # Credentials to connect to the Provider Server.
    [System.Management.Automation.CredentialAttribute()]$Credential
)


Process {

    ###############################################################################
    # Start Script
    ###############################################################################
    
    # Ensure this isn't processed when dot sourced by e.g. Pester Test trun
    if ($MyInvocation.InvocationName -ne '.') {
        # Prepare parameters for splatting
        $MainParams = $PSBoundParameters
        $MainParams.Remove("ProviderServer")
        $MainParams.Remove("SiteCode")

        # Create a connection to the ConfigMgr Provider Server
        $ConnParams = @{ServerName = $ProviderServer;SiteCode = $SiteCode;}

        if ($PSBoundParameters["Credential"]) {
            $connParams.Credential = $Credential
            $MainParams.Remove("Credential")
        }
        
        New-CMConnection @ConnParams

        # Start processing
        Main @PSBoundParameters
    }
}

Begin {

    Set-StrictMode -Version Latest

    ###############################################################################
    # Main code
    # This has been moved to a separate function to test the script logic for
    # different parameters by Pester.
    ###############################################################################
    function Main {

        [CmdLetBinding(SupportsShouldProcess)]
        PARAM(
            [string]$Start,
            [switch]$Cancel,
            [switch]$Delete,
            [string]$CollectionID,
            [string]$CollectionName,
            [uint32[]]$ResourceID,
            [string[]]$ResourceName,
            [uint32[]]$OperationID,
            [switch]$All,
            [switch]$Expired,
            [string]$ProviderServer,
            [string]$SiteCode,
            [System.Management.Automation.CredentialAttribute()]$Credential
        )

        if (!([string]::IsNullOrEmpty($Start))) {
            # Resolve Collection name into CollectionID if necessary
            if (!([System.String]::IsNullOrEmpty($CollectionName))) {
                #Get collection ID for Collection
                $Collection = Get-Collection -Type Device -Name $CollectionName

                if ($Collection -ne $null) {                     
                    $CollectionID = $Collection.CollectionID 
                } else {
                    Write-Error "Unable to find Collection $CollectionName." -ErrorAction Stop
                }
            }

            # Use "SMS00001" (All Systems) if no Collection is supplied
            if ([string]::IsNullOrEmpty($CollectionID)) {
                if (($ResourceID -ne $null) -or ($ResourceName -ne $null)) {
                    $CollectionID = "SMS00001"
                } else {
                    Throw "At least one Resource needs to be specified if no CollectionID/CollectionName is supplied."
                }
            }

            # Resolve Resource names into ResourceIDs if necessary
            if ($ResourceName -ne $null ) {
                $ResourceID = @(Get-Device -Name $ResourceName | Select-Object -ExpandProperty ResourceID)
            }

            Start-ClientOperation -Operation $Start -CollectionID $CollectionID -ResourceID $ResourceID

        } else {        
            if ($All.IsPresent) {
                $OperationID = @(Get-ClientOperation | Select-Object -ExpandProperty ID)            
            } elseif ($Expired.IsPresent) {
                $OperationID = @(Get-ClientOperation -Expired | Select-Object -ExpandProperty ID)
            }

            if ($OperationID -ne $null) {    
                if ($Cancel.IsPresent) {
                    Stop-ClientOperation -OperationID $OperationID
                } elseif ($Delete.IsPresent) {
                    Remove-ClientOperation -OperationID $OperationID
                }
            } else {
                Write-Warning "No Client Operation available."
            }
        }
    }
}
