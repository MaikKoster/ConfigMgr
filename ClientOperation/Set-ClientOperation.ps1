#Requires -Version 3.0

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

    .NOTES
        Copyright (c) 2015 Maik Koster

        Author:  Maik Koster
        Version: 1.1
        Date:    08.12.2015

        Version History:
            1.0 - 07.12.2015 - Published script
            1.1 - 08.12.2015 - Updated Help. Fixed Whatif issue

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

        # Create a connection to the ConfigMgr Provider Server
        if ($Script:Credential -ne $null) {
         
            New-CMConnection -ServerName $Script:ProviderServer -SiteCode $Script:SiteCode -Credential $Script:Credential

        } else {

            New-CMConnection -ServerName $Script:ProviderServer -SiteCode $Script:SiteCode

        }

        if (!([string]::IsNullOrEmpty($Start))) {

            # Resolve Collection name into CollectionID if necessary
            if (!([System.String]::IsNullOrEmpty($CollectionName))) {

                #Get collection ID for Collection
                $Collection = Get-DeviceCollection -Name $CollectionName

                if ($Collection -ne $null) { 
                    
                    $CollectionID = $Collection.CollectionID 

                } else {

                    Write-Error "Unable to find Collection $script:CollectionName." -ErrorAction Stop

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

                $ResourceID = @($ResourceName | Get-Computer | select -ExpandProperty ResourceID)

            }

            Start-ClientOperation -Operation $Start -CollectionID $CollectionID -ResourceID $ResourceID

        } else {
        
            if ($All.IsPresent) {

                $OperationID = @(Get-ClientOperation | select -ExpandProperty ID)
            
            } elseif ($Expired.IsPresent) {

                $OperationID = @(Get-ClientOperation -Expired | select -ExpandProperty ID)

            }

            if ($OperationID -ne $null) {
    
                if ($Cancel.IsPresent) {

                    Cancel-ClientOperation -OperationID $OperationID

                    ,$OperationID | Cancel-ClientOperation

                } elseif ($Delete.IsPresent) {

                    ,$OperationID | Delete-ClientOperation

                }

            } else {

                Write-Warning "No Client Operation available."

            }

        }

    }

    ###############################################################################
    # Function definitions
    ###############################################################################

    #   Starts a ConfigMgr Client Operation.
    function Start-ClientOperation {

        [CmdLetBinding()]
        PARAM( 

            # Type of Client Operation to start. Valid values are:
            #    
            #    FullScan               ->  to intiate a SCEP Full Scan
            #    QuickScan              ->  to initiate a SCEP Quick Scan
            #    DownloadDefinition     ->  to initiate a SCEP definition download
            #    EvaluateSoftwareUpdates ->  to initiate a software Update evaluation cycle
            #    RequestComputerPolicy  ->  to initiate a donwload and process of the current computer policy
            #    RequestUserPolicy      ->  to initiate a download and process of the current User policy
            [Parameter(Mandatory)]
            [ValidateSet("FullScan", "QuickScan", "DownloadDefinition", "EvaluateSoftwareUpdates", "RequestComputerPolicy", "RequestUserPolicy")]
            [string]$Operation,

            # CollectionID of the target(s) for the Client Operation.
            # If no ResourceID is specified, all collection members will be targeted
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$CollectionID,

            # ResourceID of the target for the Client operation. 
            # Multiple ResourceIDs can be supplied.
            [uint32[]]$ResourceID

            #[uint32]$RandomizationWindow

        )

        Write-Verbose "Starting Client operation $Operation on Collection $CollectionID"

        # Convert operation to proper type value
        switch ($Operation) {

            "FullScan" { $Type = 1 }
            "QuickScan" { $Type = 2 }
            "DownloadDefinition" { $Type = 3 }
            "EvaluateSoftwareUpdates" { $Type = 4}
            "RequestComputerPolicy" { $Type = 8 }
            "RequestUserPolicy" { $Type = 9 }
            default { $Type = 0 }
        }

        # TODO Evaluate usage of RandomizationWindow
        Invoke-CMMethod -Class "SMS_ClientOperation" -Name "InitiateClientOperation" -ArgumentList @($null, $CollectionID, $ResourceID, $Type)

    }


    # Cancels a Client Operation.
    function Cancel-ClientOperation {

        [CmdLetBinding()]
        PARAM (
            
            # The OperationID of the Client Operation that shall be canceled.
            # Multiple OperationIDs can be supplied.
            [Parameter(Mandatory,ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [Alias("ID")]
            [uint32[]]$OperationID

        )
        
        Process {
            
            foreach ($Op in $OperationID) {

                Invoke-CMMethod -Class "SMS_ClientOperation" -Name "CancelClientOperation" -ArgumentList @($Op)

            }

        }
    }

    # Deletes a Client Operation.
    function Delete-ClientOperation {

        [CmdLetBinding()]
        PARAM (

            # The OperationID of the Client Operation that shall be deleted.
            # Multiple OperationIDs can be supplied.
            [Parameter(Mandatory,ValueFromPipeline)]
            [Alias("ID")]
            [uint32[]]$OperationID

        )
        
        Process {
            
            foreach ($Op in $OperationID) {

                Invoke-CMMethod -Class "SMS_ClientOperation" -Name "DeleteClientOperation" -ArgumentList @($Op)

            }

        }
    }


    ###############################################################################
    # Standard ConfigMgr functions
    ###############################################################################

    # Creates a new connection to the ConfigMgr Provider server
    # If no Servername is supplied, localhost is assumed    
    function New-CMConnection {

        [CmdletBinding()]
        PARAM (

            # The ConfigMgr Provider Server name. 
            # If no value is specified, the script assumes to be executed on the Site Server.
            [Alias("ServerName", "Name")]
            [string]$ProviderServerName,

            # The ConfigMgr provider Site Code. 
            # If no value is specified, the script will evaluate it from the Site Server.
            [string]$SiteCode,

            # Credentials to connect to the Provider Server.
            [System.Management.Automation.CredentialAttribute()]$Credential

        )

        if ([System.String]::IsNullOrEmpty($ProviderServerName)) { $ProviderServerName = "." }

        # Get Provider Location
        # Throw if connection can't be established
        $ProviderLocation = $null
        if ($SiteCode -eq $null -or $SiteCode -eq "") {

            Write-Verbose "Get provider location for default site on server $ProviderServerName"
        
            if ($Credential -ne $null) {

                $ProviderLocation = Get-WmiObject -Query "SELECT * FROM SMS_ProviderLocation WHERE ProviderForLocalSite = true" -Namespace "root\sms" -ComputerName $ProviderServerName -ErrorAction Stop -Credential $Credential 

            } else {

                $ProviderLocation = Get-WmiObject -Query "SELECT * FROM SMS_ProviderLocation WHERE ProviderForLocalSite = true" -Namespace "root\sms" -ComputerName $ProviderServerName -ErrorAction Stop

            }

        } else {

            Write-Verbose "Get provider location for site $SiteCode on server $ProviderServerName"
    

            if ($Credential -ne $null) {

                $ProviderLocation = Get-WmiObject -Query "SELECT * FROM SMS_ProviderLocation WHERE SiteCode = '$SiteCode'" -Namespace "root\sms" -ComputerName $ProviderServerName -ErrorAction Stop -Credential $Credential

            } else {
        
                $ProviderLocation = Get-WmiObject -Query "SELECT * FROM SMS_ProviderLocation WHERE SiteCode = '$SiteCode'" -Namespace "root\sms" -ComputerName $ProviderServerName -ErrorAction Stop
        
            }
        }

        if ($ProviderLocation -ne $null) {

            # Split up the namespace path
            $Parts = $ProviderLocation.NamespacePath -split "\\", 4
            Write-Verbose "Provider is located on $($ProviderLocation.Machine) in namespace $($Parts[3])"

            # Set Script variables used by ConfigMgr related functions
            $script:CMProviderServer = $ProviderLocation.Machine
            $script:CMNamespace = $Parts[3]
            $script:CMSiteCode = $ProviderLocation.SiteCode

            # Keep credentials for further execution
            $script:CMCredential = $Credential

        } else {

            Throw "Unable to connect to specified provider"

        }

    }

    # Ensures that the ConfigMgr Provider information is available in script scope
    # If the script is . sourced,  this information will not be available if New-CMConnection isn't called explicitly.
    function Test-CMConnection {
        
        if (([string]::IsNullOrWhiteSpace($script:CMProviderServer)) -or ([string]::IsNullOrWhiteSpace($script:CMSiteCode)) -or ([string]::IsNullOrWhiteSpace($script:CMNamespace)) ) {

            New-CMConnection
            $true

        } else {
            
            $true
        }

    }


    # Returns a ConfigMgr object
    function Get-CMObject {

        [CmdletBinding()]
        PARAM (

            # ConfigMgr WMI provider Class
            [Parameter(Mandatory)] 
            [ValidateNotNullOrEmpty()]
            [string]$Class, 

            # Where clause to filter the specified ConfigMgr WMI provider class.
            # If no filter is supplied, all objects will be returned.
            [string]$Filter

        )

        if ([string]::IsNullOrWhiteSpace($Class)) { throw "Class is not specified" }

        # Ensure ConfigMgr Provider information is available
        if (Test-CMConnection) {

            if ([string]::IsNullOrWhiteSpace($Filter)) {

                Write-Verbose "Get all objects of class $Class"

                if ($CMCredential -ne $null) {

                    Get-WmiObject -ComputerName $CMProviderServer -Class $Class  -Namespace $CMNamespace -Credential $CMCredential

                } else {

                    Get-WmiObject -ComputerName $CMProviderServer -Class $Class -Namespace $CMNamespace

                }

            } else {
        
                Write-Verbose "Get objects of class $Class filtered by ""$Filter"""

                if ($CMCredential -ne $null) {

                    Get-WmiObject -Query "SELECT * FROM $Class WHERE $Filter" -ComputerName $CMProviderServer -Namespace $CMNamespace -Credential $CMCredential

                } else {
         
                    Get-WmiObject -Query "SELECT * FROM $Class WHERE $Filter" -ComputerName $CMProviderServer -Namespace $CMNamespace

                }
            }
        }
    }

    # Invokes a ConfigMgr Provider method
    function Invoke-CMMethod {

        [CmdletBinding(SupportsShouldProcess)]
        PARAM (

            # ConfigMgr WMI provider Class
            [Parameter(Mandatory)] 
            [ValidateNotNullOrEmpty()]
            [string]$Class, 

            # Method name
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            # Arguments to be supplied to the method.
            # Should be an array of objects.
            # Arguments must be in the correct order!
            [psobject]$ArgumentList

        )
        
        if ($PSCmdlet.ShouldProcess("$CMProviderServer", "Invoke $Name")) {  

            # Ensure ConfigMgr Provider information is available
            if (Test-CMConnection) {

                if ($CMCredential -ne $null) {

                    ($Result = Invoke-WmiMethod -ComputerName $CMProviderServer -Class $Class -Namespace $CMNamespace -Name $Name -ArgumentList $ArgumentList  -Credential $CMCredential) | Out-Null

                } else {

                    ($Result = Invoke-WmiMethod -ComputerName $CMProviderServer -Class $Class -Namespace $CMNamespace -Name $Name -ArgumentList $ArgumentList) | Out-Null

                }

                if ($Result -ne $null) {

                    if ($Result.ReturnValue -eq 0) {

                        Write-Verbose "Successfully invoked $Name on $CMProviderServer."

                    } else {

                        Write-Verbose "Failed to invoked $Name on $CMProviderServer. ReturnValue: $($Result.ReturnValue)"

                    }

                }

                Return $Result

            }
        }
    }

    # Returns a list of ConfigMgr Device Collections
    function Get-DeviceCollection {

        [CmdLetBinding()]
        PARAM(

            # Collection Name
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Name

        )

        # Get Device collection by name
        Get-CMObject -Class "SMS_Collection" -Filter "((Name = '$Name') AND (CollectionType = 2))"

    }

    # Returns a list of ConfigMgr SMS_R_System records
    function Get-Computer {

        [CmdLetBinding()]
        PARAM(

            # Computer Name
            [Parameter(Mandatory, ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [string[]]$Name

        )

        Process {

            foreach ($Computer in $Name) {

                Get-CMObject -Class "SMS_R_System" -Filter "((Name = '$Computer') OR (NetbiosName = '$Computer'))"

            }
        }

    }

    # Returns a list of ConfigMgr Client Operations
    function Get-ClientOperation {

        [CmdLetBinding()]
        PARAM(

            # If supplied, only expired Client Operations will be returned
            [switch]$Expired

        )

        if ($Expired.IsPresent) {
        
            Get-CMObject -Class "SMS_ClientOperation" -Filter "((State = 0) OR (State = 2))" 

        } else {

            Get-CMObject -Class "SMS_ClientOperation"

        }
 
    }

}
