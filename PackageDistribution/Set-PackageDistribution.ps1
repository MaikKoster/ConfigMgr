<#

.SYNOPSIS
This script allows to manage ConfigMgr package distribution

.DESCRIPTION
This script allows to manage several aspects of System Center Configuration Manager Package distribution.
This includes cancel running distributions, redistribute and update packages.

All types of packages including Applications, Boot Image packages, Driver Packages, etc. are supported.

Targeted Packages can be limited by package ID and/or Distribution Point.

The script will properly handle PullDPs as well.


.PARAMETER Cancel
Used to cancel running distributions.

.PARAMETER Redistribute
Used to redistribute ConfigMgr package(s).
Must be called in combination with at least one of the following parameters : 
    - Failed
    - PackageID
    - DistributionPoint

.PARAMETER Update
Used to update ConfigMgr package(s). 
Must be called in combination with the parameter PackageID.

.PARAMETER PackageID
Limits the script to the specified ConfigMgr Package(s). Single or multiple packages can be supplied.

.PARAMETER DistributionPoint
Limits the script to the specified Distribution Point(s). Single or multipe Distribution Points can be supplied
Should be the FQDN of the Distribution Point.

.PARAMETER Failed
Limits the script to failed packages only.

.PARAMETER DoNotCleanupPullDP
Prevents the script from cleaning up Pull Distribution Points when cancelling or redistributing a package.

.PARAMETER ProviderServer
The ConfigMgr Provider Server name. If no value is specified, the script assumes to be executed on the Site Server.

.PARAMETER SiteCode
The ConfigMgr provider Site Code. If no value is specified, the script will evaluate it from the Site Server.

.PARAMETER Credential
Credentials to connect to the Provider Server and Distribution Points.

.NOTES
    Author: Maik Koster
    Version: 0.1
    Date: 27.11.2015
    History:
        27.11.2015 - Published initial script

#>
#Requires -version 3.0
[CmdletBinding(SupportsShouldProcess)]
Param (

    [Parameter(Mandatory,ParameterSetName = "Cancel")]
    [switch]$Cancel,

    [Parameter(Mandatory,ParameterSetName = "Redistribute")]
    [switch]$Redistribute,

    [Parameter(Mandatory,ParameterSetName = "Update")]
    [switch]$Update,
    
    [Parameter(Position=0,ParameterSetName = "Cancel")]
    [Parameter(Position=0,ParameterSetName = "Redistribute")]
    [Parameter(Position=0,Mandatory,ParameterSetName = "Update")]
    [string[]]$PackageID,

    [Parameter(Position=1,ParameterSetName = "Cancel")]
    [Parameter(Position=1,ParameterSetName = "Redistribute")]
    [string[]]$DistributionPoint,

    [Parameter(ParameterSetName = "Redistribute")]
    [switch]$Failed,

    [Parameter(ParameterSetName = "Cancel")]
    [Parameter(ParameterSetName = "Redistribute")]
    [Parameter(ParameterSetName = "Update")]
    [switch]$DoNotCleanUpPullDP,

    [Alias("SiteServer", "ServerName")]
    [string]$ProviderServer = ".",

    [string]$SiteCode,

    [System.Management.Automation.CredentialAttribute()]$Credential

)


Process {
   
    ###############################################################################
    # Start Script
    ###############################################################################

    Set-StrictMode -Version Latest

    # Ensure this isn't processed when dot sourced by e.g. Pester
    if ($MyInvocation.InvocationName -ne '.') {
    
        # Create a connection to the ConfigMgr Provider Server
        if ($Credential -ne $null) {
         
            New-CMConnection -ServerName $ProviderServer -SiteCode $SiteCode -Credential $Credential

        } else {

            New-CMConnection -ServerName $ProviderServer -SiteCode $SiteCode

        }

        # Decide what to do
        if ($Cancel.IsPresent) {

            # On Cancel, we are only interested in running or failed distributions
            # State 1 = Install Pending, State 2 = Install retrying, State 3 = Install failed, State 7 = source files not reachable
            # TODO: Check if removal jobs should be cancelled as well
            $PackageStatus = Get-PackageStatus -PackageID $PackageID -DistributionPoint $DistributionPoint -State @(1,2,3,7)
        
            # Cancel PullDPs if not omitted
            if (!$DoNotCleanUpPullDP.IsPresent) { $PackageStatus | Cleanup-PullDP } 

            # Cancel all In-Progress distributions
            $PackageStatus | Cancel-PackageDistribution
        
        } elseif ($Redistribute.IsPresent) {

            # On Redistribute, we either get failed packages with optional filter on PackageID and/or DistributionPoint
            # Or we filter by PackageID and/or Distribution Point
            if ($Failed.IsPresent) {
        
                # Get failed packages
                $PackageStatus = Get-PackageStatusFailed -PackageID $PackageID -DistributionPoint $DistributionPoint

            } else {

                # Ensure at least one of PackageID or DistributionPoint is set, otherwise the redistribute would
                # affect all packages on all Distribution Points. That will bring down every environment.
                # If that's really required, use an array of all packages or all Distribution Points as value
                # Or better, DON'T DO IT!!!
                if ((![System.String]::IsNullOrWhiteSpace($PackageID)) -or (![System.String]::IsNullOrWhiteSpace($PackageID))) {
                
                    $PackageStatus = Get-PackageStatus -PackageID $PackageID -DistributionPoint $DistributionPoint
                
                } else {

                    Write-Error "Operation ""Redistribute"" needs to be limited by either Failed, PackageID or DistributionPoint parameters!"
                    Return
                }

            }

            # Cancel PullDPs if not omitted
            if (!$DoNotCleanUpPullDP.IsPresent) { $PackageStatus | Cleanup-PullDP } 

            # Cancel all In-Progress distributions
            $PackageStatus | Where-Object {(($_.State -eq 2) -or ($_.State -eq 3))} | Cancel-PackageDistribution

            # Redistribute the package
            $PackageStatus | Redistribute-Package

        } elseif ($Update.IsPresent) {

            # On Update, we get a list of packages that are still in progress or failed and stop them before the package is updated
            # State 1 = Install Pending, State 2 = Install retrying, State 3 = Install failed, State 7 = source files not reachable
            # TODO: Check if removal jobs should be cancelled as well
            $PackageStatus = Get-PackageStatus -PackageID $PackageID -DistributionPoint $DistributionPoint -State @(1,2,3,7)
        
            # Cancel PullDPs if not omitted
            if (!$DoNotCleanUpPullDP.IsPresent) { $PackageStatus | Cleanup-PullDP } 

            #TODO update Package

        }
    }
}

Begin {

    ###############################################################################
    # Function definitions
    ###############################################################################


    # Creates a new connection to the ConfigMgr Provider server
    # If no Servername is supplied, localhost is assumed    
    function New-CMConnection {

        [CmdletBinding()]
        PARAM (

            [Alias("ServerName", "Name")]
            [string]$ProviderServerName,

            [string]$SiteCode,

            [System.Management.Automation.CredentialAttribute()]$Credential

        )

        if ([System.String]::IsNullOrEmpty($ProviderServerName)) { $ProviderServerName = "." }


        # Get the pointer to the provider for the site code
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
            $script:CMProviderServer = $ProviderLocation.Machine
            $script:CMNamespace = $Parts[3]
            $script:CMSiteCode = $ProviderLocation.SiteCode

            # Keep credentials for further execution
            $script:CMCredential = $Credential

            # Make sure we can get a connection
            #$script:CMConnection = [wmi]"$($ProviderLocation.NamespacePath)"
            #Write-Verbose "Successfully connected to the specified provider"

        } else {

            Throw "Unable to connect to specified provider"

        }

    }

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

            [Parameter(Mandatory,ValueFromPipeline)] 
            [string]$Class, 

            [Parameter(ValueFromPipeline)]
            [string]$Filter

        )

        if ([string]::IsNullOrWhiteSpace($Class)) { throw "Class is not specified" }
        if (Test-CMConnection) {


            if ([string]::IsNullOrWhiteSpace($Filter)) {

                Write-Verbose "Get all objects of class $Class"

                if ($CMCredential -ne $null) {

                    Get-WmiObject -Class $Class -ComputerName $CMProviderServer -Namespace $CMNamespace -Credential $CMCredential

                } else {

                    Get-WmiObject -Class $Class -ComputerName $CMProviderServer -Namespace $CMNamespace

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

    function Get-PackageStatus {

        [CmdletBinding()]
        PARAM (

            [string[]]$PackageID,

            [String[]]$DistributionPoint,

            [ValidateRange(1,8)]
            [int[]]$State

        )

        # Create dynamic filter - PackageID
        $Filter = [System.String]::Empty
        if ($PackageID -ne $null) {

            if ($PackageID -is [System.Array]) {

                $Filter += "("
                $Count = 0
                foreach ($ID in $PackageID) {
                    if ($Count -gt 0) { $Filter += " OR " }
                    $Filter += "(PackageID = '$ID')"
                    $Count++
                }
                $Filter += ")"

            } else {

                $Filter = "(PackageID = '$PackageID')"

            }
        }

        # Create dynamic filter - DistributionPoint
        if ($DistributionPoint -ne $null) {

            if (!([System.String]::IsNullOrEmpty($Filter))) { $Filter += " AND " }

            if ($DistributionPoint -is [System.Array]) {

                $Filter += "("
                $Count = 0
                foreach ($DP in $DistributionPoint) {
                    if ($Count -gt 0) { $Filter += " OR " }
                    $Filter += "(ServerNalPath LIKE '%$DP%')"
                    $Count++
                }
                $Filter += ")"

            } else {

                $Filter += "(ServerNalPath LIKE '%$DistributionPoint%')"

            }
        }

        # Create dynamic filter - State
        if ($State -ne $null) {

            if (!([System.String]::IsNullOrEmpty($Filter))) { $Filter += " AND " }

            if ($State -is [System.Array]) {

                $Filter += "("
                $Count = 0
                foreach ($x in $State) {
                    if ($Count -gt 0) { $Filter += " OR " }
                    $Filter += "(State = $x)"
                    $Count++
                }
                $Filter += ")"

            } else {

                $Filter = "(State = $State)"

            }
        }

        Get-CMObject -Class "SMS_PackageStatusDistPointsSummarizer" -Filter $Filter
    
    }

    function Get-PackageStatusFailed {

        [CmdletBinding()]
        PARAM (
        
            [string[]]$PackageID,

            [string[]]$DistributionPoint

        )

        Get-PackageStatus -PackageID $PackageID -DistributionPoint $DistributionPoint -State @(3,7,8)
    }


    # Returns a list of Distribution Points
    # Can be filtered by PullDP
    function Get-DistributionPoint {

        [CmdletBinding()]
        PARAM (
            
            [switch]$PullDP
        )

        if ($PullDP.IsPresent) {

            Get-CMObject -Class SMS_DistributionPointInfo -Filter "IsPullDP = 1"

        } else {

            Get-CMObject -Class SMS_DistributionPointInfo

        }

    }

    # Returns a list of Pull Distribution Point Notifications
    function Get-PullDPNotification {

        [CmdletBinding()]
        PARAM (

            [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
            [Alias("NALPath")]
            [string]$ServerNALPath,

            [Parameter(ValueFromPipelineByPropertyName)]
            [string[]]$PackageID

        )

        Begin {
            Write-Verbose "Get Pull DP notifications ... "
        }

        Process {

            $DistributionPoint = Get-ServerNameFromNALPath($ServerNALPath)

            if (Test-Connection $DistributionPoint -Quiet) {

                if (![System.String]::IsNullOrEmpty($PackageID)) {
    
                    Write-Verbose "Get SMS_PullDPNotification objects from $DistributionPoint for PackageID $PackageID"
                    # Create dynamic filter - PackageID
                    $Filter = [System.String]::Empty
                    if ($PackageID -ne $null) {

                        if ($PackageID -is [System.Array]) {

                            $Filter += "("
                            $Count = 0
                            foreach ($ID in $PackageID) {
                                if ($Count -gt 0) { $Filter += " OR " }
                                $Filter += "(PackageID = '$ID')"
                                $Count++
                            }
                            $Filter += ")"

                        } else {

                            $Filter = "(PackageID = '$PackageID')"

                        }
                    }   

                } else {

                    Write-Verbose "Get SMS_PullDPNotification objects from $DistributionPoint"   
                    $Filter = ""
        
                }
                
                if ($CMCredential -ne $null) {

                    Get-WmiObject -ComputerName $DistributionPoint -Namespace "root\sccmdp" -Class "SMS_PullDPNotification" -Filter $Filter -Credential $CMCredential | Sort-Object -Property PackageID, PackageVersion

                } else {
        
                    Get-WmiObject -ComputerName $DistributionPoint -Namespace "root\sccmdp" -Class "SMS_PullDPNotification" -Filter $Filter | Sort-Object -Property PackageID, PackageVersion

                }

            } else {
                
                Write-Warning "Unable to connect to $DistributionPoint"

            }
        }
    }

    function Get-ServerNameFromNALPath {
    
        [CmdletBinding()]
        PARAM (
        
            [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
            [Alias("SourceNALPath", "NALPath")]
            [ValidateNotNullOrEmpty()]
            [string]$ServerNALPath

        )

        Process {
        
            if ($ServerNALPath.Contains("]")) {

                $ServerNALPath.Remove(0,$ServerNALPath.LastIndexOf("]")+1).Replace('\', '')

            } else {
                
                $ServerNALPath.Replace('\', '')
            }

        }
    }

    function Cleanup-PullDP {

        [CmdletBinding()]
        PARAM (

            [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
            [ValidateNotNullOrEmpty()]
            [Alias("NALPath")]
            [string]$ServerNALPath,

            [Parameter(ValueFromPipelineByPropertyName)]
            [string[]]$PackageID

        )

        Begin {
            
            Write-Verbose "Cleanup PullDP"

            # Get PullDP names

            $PullDPs = Get-DistributionPoint -PullDP | ForEach-Object {Get-ServerNameFromNALPath($_.NALPath)}

        }

        Process {
            
            $DistributionPoint = Get-ServerNameFromNALPath($ServerNALPath)

            if ($PullDPs -contains $DistributionPoint) {

                if (Test-Connection $DistributionPoint -Quiet) {

                    Get-PullDPNotification -ServerNALPath $ServerNALPath -PackageID $PackageID | Cancel-PullDPJob

                } else {

                    Write-Warning "Unable to connect to $DistributionPoint"
                }
            }

        }

    }

    function Cancel-PullDPJob {

        [CmdletBinding(SupportsShouldProcess)]
        PARAM (

            [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
            [ValidateNotNullOrEmpty()]
            [Alias("__SERVER")]
            [string]$DistributionPoint,

            [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
            [ValidateNotNullOrEmpty()]
            [string]$PackageID,

            [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
            [ValidateNotNullOrEmpty()]
            [Alias("SourceVersion")]
            [string]$PackageVersion

        )

        Begin {

            Write-Verbose "Cancelling PullDP jobs"

        }

        Process {
            
            if ($PSCmdlet.ShouldProcess("$DistributionPoint : $PackageID ($PackageVersion)", 'Invoke CancelPullDPJob')) {  

                if ($CMCredential -ne $null) {

                    ($Result = Invoke-WmiMethod -Class SMS_DistributionPoint -Namespace "root\sccmdp" -Name CancelPullDPJob -ArgumentList $PackageID,$PackageVersion -ComputerName $DistributionPoint -Credential $CMCredential) | Out-Null

                } else {

                    ($Result = Invoke-WmiMethod -Class SMS_DistributionPoint -Namespace "root\sccmdp" -Name CancelPullDPJob -ArgumentList $PackageID,$PackageVersion -ComputerName $DistributionPoint) | Out-Null

                }

                Write-Verbose "Invoked CancelPullDPJob on ""$DistributionPoint"" for Package ""$PackageID"" Version $PackageVersion. ReturnValue: $($Result.ReturnValue)"
            }
        }
    }


    function Cancel-PackageDistribution {

        [CmdletBinding(SupportsShouldProcess)]
        PARAM (

            [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
            [Alias("NALPath")]
            [string]$ServerNALPath,

            [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
            [string]$PackageID

        )

        Begin {

            Write-Verbose "Cancelling Package Distribution"

        }

        Process {

            $DistributionPoint = Get-ServerNameFromNALPath($ServerNALPath)

            if ($PSCmdlet.ShouldProcess("$DistributionPoint\: $PackageID", 'Invoke CancelDistribution')) {  

                if ($CMCredential -ne $null) {
                    
                    ($Result = Invoke-WmiMethod -Class SMS_DistributionPoint -Namespace $CMNamespace -Name CancelDistribution -ArgumentList $ServerNalPath,$PackageID -ComputerName $CMProviderServer -Credential $CMCredential) | Out-Null

                } else {

                    ($Result = Invoke-WmiMethod -Class SMS_DistributionPoint -Namespace $CMNamespace -Name CancelDistribution -ArgumentList $ServerNalPath,$PackageID -ComputerName $CMProviderServer) | Out-Null

                }

                Write-Verbose "Invoked CancelDistribution on ""$DistributionPoint"" for Package ""$PackageID"". ReturnValue: $($Result.ReturnValue)"

            }
        }
    }

    function Redistribute-Package {

        [CmdletBinding(SupportsShouldProcess)]
        PARAM (

            [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$True)]
            [Alias("NALPath")]
            [string]
            $ServerNALPath,

            [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$True)]
            [string]
            $PackageID

        )

        Begin {

            Write-Verbose "Redistributing Package"

        }

        Process {

            $DistributionPointName = Get-ServerNameFromNALPath($ServerNALPath)

            if ($PSCmdlet.ShouldProcess("$DistributionPointName\: $PackageID", 'Invoke RefreshNow')) {  

                $DistributionPoint = Get-CMObject -Class SMS_DistributionPoint -Filter "PackageID='$PackageID' AND ServerNALPath LIKE '%$DistributionPointName%'"

                if ($DistributionPoint -ne $null) {

                    $DistributionPoint.RefreshNow = $true
                    $DistributionPoint.Put() | Out-Null

                    Write-Verbose "Invoked RefreshNow on ""$DistributionPoint"" for Package ""$PackageID""."

                }
            }
        }
    }
}