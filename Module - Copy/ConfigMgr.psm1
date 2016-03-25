


    # Creates a new connection to the ConfigMgr Provider server
    # If no Servername is supplied, localhost is assumed    
    function New-CMConnection {
        [CmdletBinding()]
        PARAM (
            # The ConfigMgr Provider Server name. 
            # If no value is specified, the script assumes to be executed on the Site Server.
            [Alias("ServerName", "Name")]
            [string]$ProviderServerName = $env:COMPUTERNAME,

            # The ConfigMgr provider Site Code. 
            # If no value is specified, the script will evaluate it from the Site Server.
            [string]$SiteCode,

            # Credentials to connect to the Provider Server.
            [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty
        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process {        
            # Get or Create session object to connect to currently provided Providerservername
            # Ensure processing stops if it fails to create a session
            $SessionParams = @{
                ErrorAction = "Stop"
                ComputerName = $ProviderServerName
            }

            if ($PSBoundParameters["Credential"]) {
                $SessionParams.Credential = $Credential
            }
        
            $CMSession = Get-CMSession @SessionParams

            # Get Provider location
            if ($CMSession -ne $null) {
                $ProviderLocation = $null
                if ($SiteCode -eq $null -or $SiteCode -eq "") {
                    Write-Verbose "Get provider location for default site on server $ProviderServerName"
                    $ProviderLocation = Invoke-CimCommand {Get-CimInstance -CimSession $CMSession -Namespace "root\sms" -ClassName SMS_ProviderLocation -Filter "ProviderForLocalSite = true" -ErrorAction Stop}
                } else {
                    Write-Verbose "Get provider location for site $SiteCode on server $ProviderServerName"
                    $ProviderLocation = Invoke-CimCommand {Get-CimInstance -CimSession $CMSession -Namespace "root\sms" -ClassName SMS_ProviderLocation -Filter "SiteCode = '$SiteCode'" -ErrorAction Stop}
                }

                if ($ProviderLocation -ne $null) {
                    # Split up the namespace path
                    $Parts = $ProviderLocation.NamespacePath -split "\\", 4
                    Write-Verbose "Provider is located on $($ProviderLocation.Machine) in namespace $($Parts[3])"

                    # Set Script variables used by ConfigMgr related functions
                    $global:CMProviderServer = $ProviderLocation.Machine
                    $global:CMNamespace = $Parts[3]
                    $global:CMSiteCode = $ProviderLocation.SiteCode
                    $global:CMCredential = $Credential

                    # Create and store session if necessary
                    if ($global:CMProviderServer -ne $ProviderServerName) {
                        $SessionParams.ComputerName = $global:CMProviderServer
                        $CMSession = Get-CMSession @SessionParams
                    }

                    if ($CMSession -eq $null) {
                        Throw "Unable to establish CIM session to $global:CMProviderServer"
                    } else {
                        $global:CMSession = $CMSession
                    }
                } else {
                    # Clear global variables
                    $global:CMProviderServer = [string]::Empty
                    $global:CMNamespace = [string]::Empty
                    $global:CMSiteCode = [string]::Empty
                    $global:CMCredential = $null

                    Throw "Unable to connect to specified provider"
                }
            } else {
                # Clear global variables
                $global:CMProviderServer = [string]::Empty
                $global:CMNamespace = [string]::Empty
                $global:CMSiteCode = [string]::Empty
                $global:CMCredential = $null

                Throw "Unable to create CIM session to $ProviderServerName"
            }
        }
    }

    # Used to catch some common errors due to RCP connection issues 
    # when using DCOM
    function Invoke-CimCommand {
        PARAM(
            [scriptblock]$Command
        )

        $RetryCount = 0
        Do {
            $Retry = $false

            Try {
                & $Command
            } Catch {
                if ($_.Exception -ne $null) {
                    if (($_.Exception.HResult -eq -2147023169 ) -or ($_.Exception.ErrorData.error_Code -eq 2147944127)) {
                        if ($RetryCount -ge 3) {
                            $Retry = $false
                        } else {
                            $RetryCount += 1
                            $Retry = $true
                            Write-Verbose "CIM/WMI command failed with Error 2147944127 (HRESULT 0x800706bf)."
                            Write-Verbose "Common RPC error, so retrying. Retry count $RetryCount"
                        }
                    } else {
                        throw $_.Exception
                    } 
                } else {
                    throw 
                }
            }
        } While ($Retry)
    }

    # Ensures that the ConfigMgr Provider information is available in script scope
    # If the script is . sourced,  this information will not be available if New-CMConnection isn't called explicitly.
    function Test-CMConnection {
        if ( ([string]::IsNullOrWhiteSpace($global:CMProviderServer)) -or 
             ([string]::IsNullOrWhiteSpace($global:CMSiteCode)) -or 
             ([string]::IsNullOrWhiteSpace($global:CMNamespace)) -or 
             ($global:CMSession -eq $null)) {

            New-CMConnection
            $true
        } else {
            $true
        }
    }

    # Returns a valid session to the specified computer
    # If session does not exist, new session is created
    # Falls back from WSMAN to DCOM for backwards compatibility
    function Get-CMSession {
        [CmdLetBinding()]
        PARAM (
            # The ComputerName to connect to. 
            [Parameter(Position=0)]
            [ValidateNotNullOrEmpty()]
            [string]$ComputerName = $env:COMPUTERNAME,
            
            # Credentials to connect to the Provider Server.
            [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty
        )

        Begin {

            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 

            $Opt = New-CimSessionOption -Protocol Dcom

            $SessionParams = @{
                ErrorAction = 'Stop'
            }

            if ($PSBoundParameters['Credential']) {
                $SessionParams.Credential = $Credential
            }
        }

        Process {
            # Check if there is an already existing session to the specified computer
            $Session = Get-CimSession | Where { $_.ComputerName -eq $ComputerName} | Select -First 1

            if ($Session -eq $null) {
            
                $SessionParams.ComputerName = $ComputerName

                $WSMan = Test-WSMan -ComputerName $ComputerName -ErrorAction SilentlyContinue

                if (($WSMan -ne $null) -and ($WSMan.ProductVersion -match 'Stack: ([3-9]|[1-9][0-9]+)\.[0-9]+')) {
                    try {
                        Write-Verbose -Message "Attempting to connect to $ComputerName using the WSMAN protocol."
                        $Session = New-CimSession @SessionParams
                    } catch {
                        Write-Verbose "Unable to connect to $ComputerName using the WSMAN protocol. Testing DCOM ..."
                        
                    }
                    try {
                        $SessionParams.SessionOption = $Opt
                        $Session = New-CimSession @SessionParams
                    } catch {
                        Write-Error -Message "Unable to connect to $ComputerName neither using the WSMAN nor DCOM protocol. Verify your credentials and try again."
                    }
                } else {
                    $SessionParams.SessionOption = $Opt
 
                    try {
                        Write-Verbose -Message "Attempting to connect to $ComputerName using the DCOM protocol."
                        $Session = New-CimSession @SessionParams
                    } catch {
                        Write-Error -Message "Unable to connect to $ComputerName using the WSMAN or DCOM protocol. Verify $ComputerName is online and try again."
                    }
                }
                
                If ($Session -eq $null) {
                    $Session = Get-CimSession | Where { $_.ComputerName -eq $ComputerName} | Select -First 1
                }
            }

            Return $Session
        }
    }

    # Returns a ConfigMgr object
    function Get-CMInstance {
        [CmdletBinding()]
        PARAM (
            # ConfigMgr WMI provider Class
            [Parameter(Mandatory)] 
            [ValidateNotNullOrEmpty()]
            [string]$ClassName, 

            # Where clause to filter the specified ConfigMgr WMI provider class.
            # If no filter is supplied, all objects will be returned.
            [string]$Filter,

            [switch]$ContainsLazy
        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process {
            if ([string]::IsNullOrWhiteSpace($ClassName)) { throw "Class is not specified" }

            # Ensure ConfigMgr Provider information is available
            if (Test-CMConnection) {

                if (($Filter.Contains(" JOIN ")) -or ($ContainsLazy.IsPresent)) {
                    Write-Verbose "Falling back to WMI cmdlets"                
                    $WMIParams = @{
                        ComputerName = $global:CMProviderServer;
                        Namespace = $CMNamespace;
                        Class = $ClassName;
                        Filter = $Filter
                    }
                    if ($global:CMCredential -ne [System.Management.Automation.PSCredential]::Empty) {
                        $WMIParams.Credential = $CMCredential
                    }
                    Invoke-CimCommand {Get-WmiObject @WMIParams -ErrorAction Stop}
                } else {
                    $InstanceParams = @{
                        CimSession = $global:CMSession
                        Namespace = $global:CMNamespace
                        ClassName = $ClassName
                    }
                    if ($Filter -ne "") {
                        $InstanceParams.Filter = $Filter
                    }

                    Invoke-CimCommand {Get-CimInstance @InstanceParams -ErrorAction Stop}
                }
            }
        }
    }
    
    
    # Creates a new ConfigMgr Object
    function New-CMInstance {
        [CmdletBinding(SupportsShouldProcess)]
        PARAM (
            # ConfigMgr WMI provider Class
            [Parameter(Mandatory)] 
            [ValidateNotNullOrEmpty()]
            [string]$ClassName,

            # Arguments to be supplied to the new object
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [hashtable]$Arguments,

            # Set EnforceWMI to enforce the deprecated WMI cmdlets
            # still required for e.g. embedded classes without key as they need to be handled differently
            [switch]$EnforceWMI
        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process {
            if ([string]::IsNullOrWhiteSpace($ClassName)) { throw "Class is not specified" }

            # Ensure ConfigMgr Provider information is available
            if (Test-CMConnection) {
                $ArgumentsString = ($Arguments.GetEnumerator() | foreach { "$($_.Key)=$($_.Value)" }) -join "; "
                Write-Debug "Create new $ClassName object. Arguments: $ArgumentsString"

                if ($EnforceWMI) {
                    $NewCMObject = New-CMWMIInstance -ClassName $ClassName
                    if ($NewCMObject -ne $null) {
                        #try to update supplied arguments
                        try {
                            Write-Debug "Adding Arguments to WMI class"
                            $Arguments.GetEnumerator() | foreach {
                                $Key = $_.Key
                                $Value = $_.Value
                                Write-Debug "$Key : $Value"
                                $NewCMObject[$Key] = $Value
                            }
                        } catch {
                            Write-Error "Unable to update Arguments on wmi class $ClassName"
                        }
                    }
                } else {
                    if ($PSCmdlet.ShouldProcess("Class: $ClassName", "Call New-CimInstance")) {
                        $NewCMObject = Invoke-CimCommand {New-CimInstance -CimSession $global:CMSession -Namespace $global:CMNamespace -ClassName $ClassName -Property $Arguments -ErrorAction Stop}

                        if ($NewCMObject -ne $null) {
                            # Ensure that generated properties are udpated
                            $hack = $NewCMObject.PSBase | select * | Out-Null
                        }
                    }
                }

                Return $NewCMObject
            }
        }
    }

    # Updates a ConfigMgr object
    function Set-CMInstance {
        [CmdletBinding(SupportsShouldProcess,DefaultParameterSetName="ClassInstance")]
        PARAM (
            # ConfigMgr WMI provider Class
            [Parameter(Mandatory,ParameterSetName="ClassName")] 
            [ValidateNotNullOrEmpty()]
            [string]$ClassName,

            # Filter
            [Parameter(Mandatory,ParameterSetName="ClassName")] 
            [ValidateNotNullOrEmpty()]
            [string]$Filter,
            
            # ConfigMgr WMI provider object
            [Parameter(Mandatory,ParameterSetName="ClassInstance")] 
            [ValidateNotNullOrEmpty()]
            [object]$ClassInstance,  

            # Arguments to be supplied to the method.
            # Should be a hashtable with key/name pairs.
            [hashtable]$Arguments
        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }
        
        Process {
            # Ensure ConfigMgr Provider information is available
            if (Test-CMConnection) {
                if ($ClassInstance -eq $null) {
                    $ClassInstance = Get-CMInstance -Class $ClassName -Filter $Filter
                }

                if ($ClassInstance -ne $null) {
                    if ($Arguments -eq $null) {
                        $ClassInstance | Invoke-CimCommand {Set-CimInstance -PassThru -ErrorAction Stop}
                    } else {
                        $ClassInstance | Invoke-CimCommand {Set-CimInstance -Property $Arguments -PassThru -ErrorAction Stop}
                    }
                } 
            }
        }
    }

    # Removes a ConfigMgr object
    function Remove-CMInstance {
        [CmdletBinding(SupportsShouldProcess,DefaultParameterSetName="ClassInstance")]
        PARAM (
            # ConfigMgr WMI provider Class
            [Parameter(Mandatory,ParameterSetName="ClassName")] 
            [ValidateNotNullOrEmpty()]
            [string]$ClassName,

            # Filter
            [Parameter(Mandatory,ParameterSetName="ClassName")] 
            [ValidateNotNullOrEmpty()]
            [string]$Filter,
            
            # ConfigMgr WMI provider object
            [Parameter(Mandatory,ParameterSetName="ClassInstance")] 
            [ValidateNotNullOrEmpty()]
            [object]$ClassInstance,  

            # Arguments to be supplied to the method.
            # Should be a hashtable with key/name pairs.
            [hashtable]$Arguments
        )
        
        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process {
            # Ensure ConfigMgr Provider information is available
            if (Test-CMConnection) {
                if ($ClassInstance -eq $null) {
                    $ClassInstance = Get-CMInstance -Class $ClassName -Filter $Filter
                }

                if ($ClassInstance -ne $null) {
                    Invoke-CimCommand {Remove-CimInstance -CimSession $global:CMSession -InputObject $ClassInstance -ErrorAction Stop}
                } 
            }
        }
    }

    
    # Creates a new configMgr WMI class instance
    # Uses Get-WMIObject as fall back
    # TODO: Figure out how to handle embedded classes with no key with the new CIM methods.
    function New-CMWMIInstance {
        [CmdletBinding(SupportsShouldProcess)]
        PARAM (
            # ConfigMgr WMI provider Class
            [Parameter(Mandatory)] 
            [ValidateNotNullOrEmpty()]
            [string]$ClassName
        )

        if ([string]::IsNullOrWhiteSpace($ClassName)) { throw "Class is not specified" }

        # Ensure ConfigMgr Provider information is available
        if (Test-CMConnection) {
            Write-Verbose "Create new WMI $ClassName instance."

            if ($global:CMCredential -ne [System.Management.Automation.PSCredential]::Empty) {
                if ($PSCmdlet.ShouldProcess("Class: $ClassName", "New-CMWMIInstance")) {
                    $CMClass = Get-WmiObject -List -Class $ClassName -ComputerName $CMProviderServer -Namespace $CMNamespace -Credential $global:CMCredential
                }

            } else {
                if ($PSCmdlet.ShouldProcess("Class: $ClassName", "New-CMWMIInstance")) {
                    $CMClass = Get-WmiObject -List -Class $ClassName -ComputerName $CMProviderServer -Namespace $CMNamespace
                }
            }

            if ($CMClass -ne $null) {
                $CMinstance = $CMClass.CreateInstance()

                Return $CMinstance
            }
        }
    }

    
    # Invokes a ConfigMgr Provider method
    function Invoke-CMMethod {
        [CmdletBinding(SupportsShouldProcess)]
        PARAM (
            # ConfigMgr WMI provider Class
            # Needs to be supplied for static class methods
            [Parameter(Mandatory,ParameterSetName="ClassName")] 
            [ValidateNotNullOrEmpty()]
            [string]$ClassName,
            
            # ConfigMgr WMI provider object
            # Needs to be supplied for instance methods
            [Parameter(Mandatory,ParameterSetName="ClassInstance")] 
            [ValidateNotNullOrEmpty()]
            [object]$ClassInstance,  

            # Method name
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$MethodName,

            # Arguments to be supplied to the method.
            # Should be a hashtable with key/name pairs.
            [hashtable]$Arguments,

            # If Set, ReturnValue will not be evaluated
            # Usefull if ReturnValue does not indicated successfull execution
            [switch]$SkipValidation
        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }
        
        Process {
            if ($PSCmdlet.ShouldProcess("$CMProviderServer", "Invoke $MethodName")) {  
                # Ensure ConfigMgr Provider information is available
                if (Test-CMConnection) {
                    
                    Do {
                        $retry = $false

                        if ($ClassInstance -ne $null) {
                            $Result = Invoke-CimCommand {Invoke-CimMethod -InputObject $ClassInstance -MethodName $MethodName -Arguments $Arguments -ErrorAction Stop}
                        } else {
                            $Result = Invoke-CimCommand {Invoke-CimMethod -CimSession $global:CMSession -Namespace $CMNamespace -ClassName $ClassName -MethodName $MethodName -Arguments $Arguments  -ErrorAction Stop}
                        }

                        if ((!($SkipValidation.IsPresent)) -and ($Result -ne $null)) {
                            if ($Result.ReturnValue -eq 0) {
                                Write-Verbose "Successfully invoked $MethodName on $CMProviderServer."
                            } else {
                                Write-Verbose "Failed to invoked $MethodName on $CMProviderServer. ReturnValue: $($Result.ReturnValue)"
                            }
                        } 
                    } While ($retry)

                    Return $Result
                }
            }
        }
    }

    
    # Moves a ConfigMgr Object to a different Folder
    Function Move-CMObject {            
        [CmdLetBinding()]
        PARAM (
            [Parameter(Mandatory)]        
            [uint32]$SourceFolderID,
            
            [Parameter(Mandatory)]
            [uint32]$TargetFolderID,

            [Parameter(Mandatory)]
            [string[]]$ObjectID
        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process {
            $Result = $false

            # Ensure ConfigMgr Provider information is available
            if (Test-CMConnection) {
                if ($SourceFolderID -eq $TargetFolderID) {
                    Write-Verbose "Move-CMObject: SourceFolder $SourceFolderID and TargetFolder $TargetFolderID are identical. No action necessary."
                    $Result = $true
                } else {
                    # Verify Folders exist
                    $SourceFolder = Get-Folder -ID $SourceFolderID
                    $TargetFolder = Get-Folder -ID $TargetFolderID
                    if (($SourceFolder -ne $null) -and ($TargetFolder -ne $null)) {
                        # Fix ObjectType if Source or Target are the root
                        if ($TargetFolderID -eq 0) { $TargetFolder.objectType = $SourceFolder.ObjectType}
                        if ($SourceFolderID -eq 0) { $SourceFolder.objectType = $TargetFolder.ObjectType}
                    
                        # Verify folders have the same ObjectType
                        if ($SourceFolder.ObjectType -eq $TargetFolder.ObjectType) {
                            $Arguments = @{
                                InstanceKeys = $ObjectID
                                ContainerNodeID=$SourceFolderID
                                TargetContainerNodeID=$TargetFolderID
                                ObjectType=$TargetFolder.ObjectType
                            }

                            $MethodResult = Invoke-CMMethod -ClassName SMS_ObjectContainerItem -MethodName MoveMembers -Arguments $Arguments
                            if (($MethodResult -ne $null) -and ($MethodResult.ReturnValue -eq 0)) {
                                Write-Verbose "Successfully moved object $ObjectID from $($SourceFolder.Name) to $($TargetFolder.Name)"
                                $Result = $true
                            }
                        } else {
                            Write-Warning "Unable to move object $ObjectID. SourceFolder $SourceFolderID ObjectType $($SourceFolder.ObjectType) and TargetFolder $TargetFolderID ObjectType $($TargetFolder.ObjectType) don't match."
                        }
                    } else {
                        $Message = "Unable to move object $ObjectID. "

                        if ($SourceFolder -eq $null) { $Message += "SourceFolder $SourceFolderID cannot be retrieved. "}
                        if ($TargetFolder -eq $null) { $Message += "TargetFolder $TargetFolderID cannot be retrieved. "}

                        Write-Warning $Message
                    }
                }
            }
            Return $Result
        }
    }

    # Returns a ConfigMgr Folder
    function Get-Folder {
        [CmdletBinding()]
        PARAM (
            # Folder ID
            [Parameter(Mandatory, ParameterSetName="ID")]
            [uint32]$ID,

            # Folder Name
            [Parameter(Mandatory, ParameterSetName="Name")]
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            # Folder type
            [Parameter(Mandatory, ParameterSetName="Name")]
            [ValidateSet("Package", "Advertisement", "Query", "Report", "MeteredProductRule", "ConfigurationItem", "OSInstallPackage", "StateMigration", "ImagePackage", "BootImagePackage", "TaskSequencePackage", "DeviceSettingPackage", "DriverPackage", "Driver", "SoftwareUpdate", "ConfigurationBaseline", "DeviceCollection", "UserCollection")]
            [string]$Type,

            # ParentFolder ID
            # To be able to differentiate between having no value supplied (which defaults to 0 on uint32)
            # and an explicit search on root level, the max value of uint32 (4294967295) is used.
            [Parameter(ParameterSetName="Name")]
            [uint32]$ParentFolderID = [uint32]::MaxValue
        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process {
            if (![string]::IsNullOrEmpty($Name)) {
                switch ($Type) {

                    "Package" {$TypeID = 2}
                    "Advertisement" {$TypeID = 3}
                    "Query" {$TypeID = 7}
                    "Report" {$TypeID = 8}
                    "MeteredProductRule" {$TypeID = 9}
                    "ConfigurationItem" {$TypeID = 11}
                    "OSInstallPackage" {$TypeID = 14}
                    "StateMigration" {$TypeID = 17}
                    "ImagePackage" {$TypeID = 18}
                    "BootImagePackage" {$TypeID = 19}
                    "TaskSequencePackage" {$TypeID = 20}
                    "DeviceSettingPackage" {$TypeID = 21}
                    "DriverPackage" {$TypeID = 23}
                    "Driver" {$TypeID = 25}
                    "SoftwareUpdate" {$TypeID = 1011}
                    "ConfigurationBaseline" {$TypeID = 2011}
                    "DeviceCollection" {$TypeID = 5000}
                    "UserCollection" {$TypeID = 5001}
                    default {Throw "Unsupported Folder Type"}
                }

                if ($ParentFolderID -eq [uint32]::MaxValue) {
                    Get-CMInstance -ClassName SMS_ObjectContainerNode -Filter "((Name = '$Name') AND (ObjectType=$TypeID))"
                } else {
                    Get-CMInstance -ClassName SMS_ObjectContainerNode -Filter "((Name = '$Name') AND (ObjectType=$TypeID) AND (ParentContainerNodeID = $ParentFolderID))"
                }
            } else {
                if ($ID -gt 0) {
                    Get-CMInstance -ClassName SMS_ObjectContainerNode -Filter "ContainerNodeID = $ID"
                } else {
                    # Return custom object to ease error handling
                    Write-Verbose "Get-Folder: ID 0 is the root folder. Returning custom object."
                    [PSCustomObject]@{Name = "Root"; ContainerNodeID = 0; ObjectType = 0}
                }
            }
        }
    }

    
    # Ceates a new ConfigMgr Folder
    function New-Folder {
        [CmdletBinding()]
        PARAM (
            # Folder Name
            [Parameter(Mandatory, ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            # Type of object to place in the console folder
            [Parameter(Mandatory, ValueFromPipeline)]
            [ValidateSet("Package", "Advertisement", "Query", "Report", "MeteredProductRule", "ConfigurationItem", "OSInstallPackage", "StateMigration", "ImagePackage", "BootImagePackage", "TaskSequencePackage", "DeviceSettingPackage", "DriverPackage", "Driver", "SoftwareUpdate", "ConfigurationBaseline", "DeviceCollection", "UserCollection")]
            [string]$Type,

            # The unique ID of the parent folder. 
            [Parameter(ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [string]$ParentFolderID = 0
        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process {
            Switch ($Type) {
                "Package" {$TypeID = 2}
                "Advertisement" {$TypeID = 3}
                "Query" {$TypeID = 7}
                "Report" {$TypeID = 8}
                "MeteredProductRule" {$TypeID = 9}
                "ConfigurationItem" {$TypeID = 11}
                "OSInstallPackage" {$TypeID = 14}
                "StateMigration" {$TypeID = 17}
                "ImagePackage" {$TypeID = 18}
                "BootImagePackage" {$TypeID = 19}
                "TaskSequencePackage" {$TypeID = 20}
                "DeviceSettingPackage" {$TypeID = 21}
                "DriverPackage" {$TypeID = 23}
                "Driver" {$TypeID = 25}
                "SoftwareUpdate" {$TypeID = 1011}
                "ConfigurationBaseline" {$TypeID = 2011}
                "DeviceCollection" {$TypeID = 5000}
                "UserCollection" {$TypeID = 5001}

                default {Throw "Unsupported Folder Type"}
            }

            $Arguments = @{
                Name = $Name; 
                ObjectType = $TypeID; 
                ParentContainerNodeid = $ParentFolderID
            }

            New-CMInstance -ClassName SMS_ObjectContainerNode -Arguments $Arguments
        }
    }

    ###########################
    # SMS_TaskSequencePackage #
    ###########################
    Function Get-TaskSequencePackage {
        [CmdletBinding(SupportsShouldProcess,DefaultParameterSetName="ID")]
        PARAM (
            # PackageID
            [Parameter(Mandatory,ParameterSetName="ID")] 
            [ValidateNotNullOrEmpty()]
            [string]$ID,

            # PackageName
            [Parameter(Mandatory,ParameterSetName="Name")] 
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            [switch]$IncludeLazyProperties

        )
        
        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process{
            if (!([string]::IsNullOrEmpty($ID))) {
                Get-CMInstance -ClassName "SMS_TaskSequencePackage" -Filter "PackageID='$ID'" -ContainsLazy:$IncludeLazyProperties
            } elseif (!([string]::IsNullOrEmpty($Name))) {
                Get-CMInstance -ClassName "SMS_TaskSequencePackage" -Filter "Name='$Name'" -ContainsLazy:$IncludeLazyProperties
            } 
        }
    }

    Function New-TaskSequencePackage {
        [CmdletBinding(SupportsShouldProcess)]
        PARAM (
            # Task Sequence Name
            [Parameter(Mandatory)] 
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            [Parameter(Mandatory)] 
            [object]$TaskSequence,

            # Description
            [string]$Description = "",

            # Additional Properties
            [Hashtable]$Properties

        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process{
            $Arguments = @{
                Name = $Name; 
                Description = $Description
            }

            if ($Properties -ne $null) {
                if ($Properties.ContainsKey("Name")) {$Properties.Remove("Name")}
                if (($Properties.ContainsKey("Description")) -and ([string]::IsNullOrEmpty($Description))) {$Properties.Remove("Description")}
                $Arguments = $Arguments + $Properties
            }

            # Create local instance of SMS_TaskSequencePackage
            $TaskSequencePackage = Invoke-CimCommand {New-CimInstance -ClassName "SMS_TaskSequencePackage" -Arguments $Arguments -Namespace $Global:CMNamespace -ClientOnly -ErrorAction Stop}

            # Invoke SetSequence to add the sequence and create the package
            [string]$Result = Set-TaskSequence -TaskSequencePackage $TaskSequencePackage -TaskSequence $TaskSequence

            # Result should contain "PackageID"
            if (!([string]::IsNullOrEmpty($Result)) -and ($Result.Contains("SMS_TaskSequencePackage.PackageID="))) {
                $PackageID = $Result.Replace("SMS_TaskSequencePackage.PackageID=","").Replace("""","")

                $TaskSequencePackage = Get-TaskSequencePackage -ID $PackageID

                $TaskSequencePackage
            } else {

                $null
            }
            
        }
    }
    
    ####################
    # SMS_TaskSequence #
    ####################

    Function Get-TaskSequence {
        [CmdletBinding(SupportsShouldProcess)]
        PARAM (
            [Parameter(Mandatory,ParameterSetName="ID")] 
            [ValidateNotNullOrEmpty()]
            [string]$ID,
            
            # PackageName
            [Parameter(Mandatory,ParameterSetName="Name")] 
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            [Parameter(Mandatory,ParameterSetName="Package")]
            [ValidateNotNullOrEmpty()]
            $TaskSequencePackage

        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process {

            if ($TaskSequencePackage -eq $null) {
                if (!([string]::IsNullOrEmpty($ID))) {
                    $TaskSequencePackage = Get-TaskSequencePackage -ID $ID
                } else {
                    $TaskSequencePackage = Get-TaskSequencePackage -Name $Name
                }
            }

            if ($TaskSequencePackage -ne $null) {

                $Result = Invoke-CMMethod -ClassName "SMS_TaskSequencePackage" -MethodName "GetSequence" -Arguments @{TaskSequencePackage=$TaskSequencePackage}

                if ($Result -ne $null) {
                    if ($Result.ReturnValue -eq 0) {

                        $TaskSequence = $Result.TaskSequence

                        $TaskSequence
                    }
                }
            }
        }
    }

    Function Set-TaskSequence {
        [CmdletBinding(SupportsShouldProcess)]
        PARAM (
            [Parameter(Mandatory,ParameterSetName="ID")] 
            [ValidateNotNullOrEmpty()]
            [string]$ID,

            [Parameter(Mandatory,ParameterSetName="Package")] 
            [ValidateNotNullOrEmpty()]
            [object]$TaskSequencePackage,

            [Parameter(Mandatory)] 
            [object]$TaskSequence

        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process {
            if (!([string]::IsNullOrEmpty($ID))) {
                $TaskSequencePackage = Get-TaskSequencePackage -ID $ID
            }

            if ($TaskSequencePackage -ne $null) {

                $Result = Invoke-CMMethod -ClassName "SMS_TaskSequencePackage" -MethodName "SetSequence" -Arguments @{TaskSequencePackage=$TaskSequencePackage;TaskSequence=$TaskSequence}

                if ($Result -ne $null) {
                    if ($Result.ReturnValue -eq 0) {

                        $Path = $Result.SavedTaskSequencePackagePath

                        $Path
                    } else {

                        Write-Error "Error $($Result.ReturnValue) on executing SetSequence."
                    }
                }
            }
        }
    }

    Function Convert-XMLToTaskSequence {
        [CmdletBinding(SupportsShouldProcess)]
        PARAM (
            # PackageID
            [Parameter(Mandatory)] 
            [ValidateNotNullOrEmpty()]
            $TaskSequenceXML

        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process {
            # TODO: Add ConfigMgr 2007 handling (LoadFromXML method in sms_TaskSequence)

            if ($TaskSequenceXML -ne $null) {
                $TSXMLString = $TaskSequenceXML.OuterXml

                $Result = Invoke-CMMethod -ClassName "SMS_TaskSequencePackage" -MethodName "ImportSequence" -Arguments @{SequenceXML=$TSXMLString}

                if ($Result -ne $null) {
                    if ($Result.ReturnValue -eq 0) {

                        $TaskSequence = $Result.TaskSequence

                        $TaskSequence
                    }
                }
            }
        }
    }

    Function Convert-TaskSequenceToXML {
        [CmdletBinding(SupportsShouldProcess)]
        PARAM (
            # PackageID
            [Parameter(Mandatory)] 
            [ValidateNotNullOrEmpty()]
            $TaskSequence,

            [switch]$KeepSecretInformation

        )

        Begin {
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState 
        }

        Process {

            if ($TaskSequence -ne $null) {

                $Result = Invoke-CMMethod -ClassName "SMS_TaskSequence" -MethodName "SaveToXml" -Arguments @{TaskSequence=$TaskSequence} -SkipValidation

                if ($Result -ne $null) {
                    $TaskSequenceString = $Result.ReturnValue
                }

                if ($KeepSecretInformation.IsPresent) {
                    $TaskSequenceXML = [xml]$TaskSequenceString
                } else {
                    $Result = Invoke-CMMethod -ClassName "SMS_TaskSequence" -MethodName "ExportXml" -Arguments @{Xml=$TaskSequenceString} -SkipValidation

                    if ($Result -ne $null) {
                        $TaskSequenceXML = [xml]($Result.ReturnValue)
                    }
                }

                $TaskSequenceXML
            }
        }
    }

    # Helper Tools
function Get-CallerPreference
{
    <#
    .Synopsis
       Fetches "Preference" variable values from the caller's scope.
    .DESCRIPTION
       Script module functions do not automatically inherit their caller's variables, but they can be
       obtained through the $PSCmdlet variable in Advanced Functions.  This function is a helper function
       for any script module Advanced Function; by passing in the values of $ExecutionContext.SessionState
       and $PSCmdlet, Get-CallerPreference will set the caller's preference variables locally.
    .PARAMETER Cmdlet
       The $PSCmdlet object from a script module Advanced Function.
    .PARAMETER SessionState
       The $ExecutionContext.SessionState object from a script module Advanced Function.  This is how the
       Get-CallerPreference function sets variables in its callers' scope, even if that caller is in a different
       script module.
    .PARAMETER Name
       Optional array of parameter names to retrieve from the caller's scope.  Default is to retrieve all
       Preference variables as defined in the about_Preference_Variables help file (as of PowerShell 4.0)
       This parameter may also specify names of variables that are not in the about_Preference_Variables
       help file, and the function will retrieve and set those as well.
    .EXAMPLE
       Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

       Imports the default PowerShell preference variables from the caller into the local scope.
    .EXAMPLE
       Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -Name 'ErrorActionPreference','SomeOtherVariable'

       Imports only the ErrorActionPreference and SomeOtherVariable variables into the local scope.
    .EXAMPLE
       'ErrorActionPreference','SomeOtherVariable' | Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

       Same as Example 2, but sends variable names to the Name parameter via pipeline input.
    .INPUTS
       String
    .OUTPUTS
       None.  This function does not produce pipeline output.
    .LINK
       about_Preference_Variables
    #>

    [CmdletBinding(DefaultParameterSetName = 'AllVariables')]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ $_.GetType().FullName -eq 'System.Management.Automation.PSScriptCmdlet' })]
        $Cmdlet,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.SessionState]
        $SessionState,

        [Parameter(ParameterSetName = 'Filtered', ValueFromPipeline = $true)]
        [string[]]
        $Name
    )

    Begin {
        $filterHash = @{}
    }
    
    Process {
        if ($null -ne $Name) {
            foreach ($string in $Name) {
                $filterHash[$string] = $true
            }
        }
    }

    End {
        # List of preference variables taken from the about_Preference_Variables help file in PowerShell version 4.0

        $vars = @{
            'ErrorView' = $null
            'FormatEnumerationLimit' = $null
            'LogCommandHealthEvent' = $null
            'LogCommandLifecycleEvent' = $null
            'LogEngineHealthEvent' = $null
            'LogEngineLifecycleEvent' = $null
            'LogProviderHealthEvent' = $null
            'LogProviderLifecycleEvent' = $null
            'MaximumAliasCount' = $null
            'MaximumDriveCount' = $null
            'MaximumErrorCount' = $null
            'MaximumFunctionCount' = $null
            'MaximumHistoryCount' = $null
            'MaximumVariableCount' = $null
            'OFS' = $null
            'OutputEncoding' = $null
            'ProgressPreference' = $null
            'PSDefaultParameterValues' = $null
            'PSEmailServer' = $null
            'PSModuleAutoLoadingPreference' = $null
            'PSSessionApplicationName' = $null
            'PSSessionConfigurationName' = $null
            'PSSessionOption' = $null

            'ErrorActionPreference' = 'ErrorAction'
            'DebugPreference' = 'Debug'
            'ConfirmPreference' = 'Confirm'
            'WhatIfPreference' = 'WhatIf'
            'VerbosePreference' = 'Verbose'
            'WarningPreference' = 'WarningAction'
        }

        foreach ($entry in $vars.GetEnumerator()) {
            if (([string]::IsNullOrEmpty($entry.Value) -or -not $Cmdlet.MyInvocation.BoundParameters.ContainsKey($entry.Value)) -and
                ($PSCmdlet.ParameterSetName -eq 'AllVariables' -or $filterHash.ContainsKey($entry.Name))) {
                $variable = $Cmdlet.SessionState.PSVariable.Get($entry.Key)
                
                if ($null -ne $variable) {
                    if ($SessionState -eq $ExecutionContext.SessionState) {
                        Set-Variable -Scope 1 -Name $variable.Name -Value $variable.Value -Force -Confirm:$false -WhatIf:$false
                    } else {
                        $SessionState.PSVariable.Set($variable.Name, $variable.Value)
                    }
                }
            }
        }

        if ($PSCmdlet.ParameterSetName -eq 'Filtered') {
            foreach ($varName in $filterHash.Keys) {
                if (-not $vars.ContainsKey($varName)) {
                    $variable = $Cmdlet.SessionState.PSVariable.Get($varName)
                
                    if ($null -ne $variable) {
                        if ($SessionState -eq $ExecutionContext.SessionState) {
                            Set-Variable -Scope 1 -Name $variable.Name -Value $variable.Value -Force -Confirm:$false -WhatIf:$false
                        } else {
                            $SessionState.PSVariable.Set($variable.Name, $variable.Value)
                        }
                    }
                }
            }
        }

    } # end

} # function Get-CallerPreference