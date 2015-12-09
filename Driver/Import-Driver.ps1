



Process {

    ###############################################################################
    # Start Script
    ###############################################################################
    
    # Ensure this isn't processed when dot sourced by e.g. Pester Test trun
    if ($MyInvocation.InvocationName -ne '.') {

        # Create a connection to the ConfigMgr Provider Server
        if ($Script:Credential -ne $null) {
         
            New-CMConnection -ServerName $Script:ProviderServer -SiteCode $Script:SiteCode -Credential $Script:Credential

        } else {

            New-CMConnection -ServerName $Script:ProviderServer -SiteCode $Script:SiteCode

        }

    }

}

Begin {

    Set-StrictMode -Version Latest

    ###############################################################################
    # Function definitions
    ###############################################################################



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

}