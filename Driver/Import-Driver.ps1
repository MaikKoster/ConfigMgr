



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

    function Import-DriverPackageSourceFolder {

        [CmdLetBinding()]
        PARAM (

            # Source path of the drivers for the Driver Package
            [DirectoryInfo]$DriverSourcePath,

            # Source path of the package it needs to be created
            [DirectoryInfo]$PackageSourcePath,

            # Package name
            [string]$PackageName,

            # Category name
            [string]$CategoryName,

            # Cleanup content if new package path already exists
            [switch]$Cleanup

        )

        # Process this folder only if there are any changes
        $PackageHash = Get-FolderHash $DriverSourcePath.FullName

        if (Get-ChildItem $package.FullName -Filter "$($PackageHash).hash") {

		    Write-Verbose "No changes have been made to this Driver Package. Skipping."

	    } else {

            # Get Driver Package
            # TODO: Add logic for same package name in different folders
            $DriverPackage = Get-DriverPackage -Name $PackageName

            # Create the package if it doesn't already exist
            if ($DriverPackage -eq $null) {

                $DriverPackageSource = "$PackageSourcePath\$PackageName"
                if ((Get-Item $DriverPackageSource | %{$_.GetDirectories().Count + $_.GetFiles().Count}) -gt 0) {
					if ($cleanup) {

					    Write-Verbose "Folder already exists, removing content"
    				    dir $driverPackageSource | remove-item -recurse -force

					} else {

						Write-Error "Folder already exists, remove it manually."
						return
					}

				} else {

                    $null = MkDir "$DriverPackageSource"

                }

                $DriverPackage = New-DriverPackage -Name "$Name" -SourcePath $DriverPackageSource

                #Move-CMObject -TargetFolderID $folderID -ObjectID $CMPackage.PackageID -ObjectType 23

                Write-Host "Created driver package $($DriverPackage.PackageID)"

            } else {

                Write-Host "Existing driver package $($DriverPackage.PackageID) retrieved."

            }

            # Get current list of drivers
            $PackageID = $DriverPackage.PackageID
            $CurrentDrivers = Get-Driver -PackageID $PackageID
            $VerifiedDrivers = @()

            # Import all Drivers from the Driver Source
            Get-ChildItem $DriverSourcePath.FullName -Filter *.inf -Recurse | Import-Driver -CategoryName $CategoryName | % {

                # Add the driver to the driver package if necessary
                $DriverUniqueID = $_.CI_UniqueID
                $DriverName = $_.LocalizedDisplayName
                $DriverExistS = $CurrentDrivers | ? {$_.CI_UniqueID -eq $DriverUniqueID}

                if ($DriverExistS -eq $null) {

                    # Add the driver to the package since it's not already there
                    Write-Host "Adding driver $DriverName to package $($DriverPackage.Name)"
                    $null = Add-DriverToDriverPackage -Package $DriverPackage -Driver $_

                } else {

                    # Add driver to list of drivers still valid for that package
                    Write-Host "Driver $driverName is already in package $($driverPackage.Name)"
                    $VerifiedDrivers += $_

                }

            }

            # Enumerate all former drivers from the driver package and check if they are still available in the source
            foreach ($Driver in $CurrentDrivers) {

                $DriverVerified = $VerifiedDrivers | ? {$_.CI_UniqueID -eq $Driver.CI_UniqueID}

                if ($DriverVerified -eq $null) {

                    # Remove Driver from package
                    Write-Host "Driver $($driver.LocalizedDisplayName) is no longer available in the source for package $($driverPackage.Name)"

                    Remove-DriverFromDriverPackage -DriverPackage $DriverPackage -Driver $Driver

                    # Remove Category from Driver
                    if ($category -eq $null) { $category = GetCategory $categoryName }

                    Remove-CategoryFromDriver -Driver $Driver -Category 

                    # Verify ContentSourcePath

                }

            }


            # Update hash file
            Get-ChildItem $DriverSourcePath.FullName -Filter "*.hash"  | Remove-Item 
		    $null = New-Item "$($DriverSourcePath.FullName)\$($PackageHash).hash" -type file 

        }

    }


    # Returns ConfigMgr Driver Package
    function Get-DriverPackage {

        [CmdletBinding(DefaultParameterSetName="PackageID")]
        PARAM
        (
            # PackageID
            [Parameter(Mandatory, ParameterSetName="PackageID")]
            [ValidateNotNullOrEmpty()]
            [string]$PackageID,

            # Package Name
            [Parameter(Mandatory, ParameterSetName="Name")]
            [ValidateNotNullOrEmpty()]
            [string]$Name

        )

        if (![string]::IsNullOrEmpty($PackageID)) {

            Get-CMObject -Class SMS_DriverPackage -Filter "PackageID = '$PackageID'"

        } elseif (![string]::IsNullOrEmpty($Name)) {

            Get-CMObject -Class SMS_DriverPackage -Filter "Name = '$Name'"

        }

    }

    # Creates a new ConfigMgr Driver Package
    function New-DriverPackage
    {
        [CmdletBinding()]
        PARAM
        (
            # Package Name
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            # Package Source Path
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$SourcePath,

            # Package Description
            [ValidateNotNullOrEmpty()]
            [string]$Description = ""

        )

        # Build the parameters for creating the driver package
        $Arguments = @{Name = $Name; Description = $Description; PkgSourceFlag = 2; PkgSourcePath = $SourcePath}

        # Create new driver package
        $NewPackage = New-CMObject -Class SMS_DriverPackage -Arguments $Arguments
    
        # Return the package
        $NewPackage
    }

    # Returns the drivers
    function Get-Driver {

        [CmdletBinding()]
        PARAM
        (
            # Driver Package
            [Parameter(Mandatory,ParameterSetName="PackageID")]
            [ValidateNotNullOrEmpty()]
            [string]$PackageID,

            # Driver Package
            [Parameter(Mandatory,ParameterSetName="UniqueID")]
            [ValidateNotNullOrEmpty()]
            [string]$UniqueID
        )

        if (![string]::IsNullOrEmpty($PackageID)) {

            Get-CMObject -Class SMS_Driver -Filter "CI_ID IN (SELECT CTC.CI_ID FROM SMS_CIToContent AS CTC JOIN SMS_PackageToContent AS PTC ON CTC.ContentID=PTC.ContentID JOIN SMS_DriverPackage AS Pkg ON PTC.PackageID=Pkg.PackageID WHERE Pkg.PackageID='$($DriverPackage.PackageID)')"

        } elseif (![string]::IsNullOrEmpty($UniqueID)) {

            $UniqueID = $UniqueID.Replace("\", "/")
            Get-CMObject -Class SMS_Driver -Filter "CI_UniqueID = '$UniqueID'"

        }
    }


    # Imports the specified driver into ConfigMgr
    function Import-Driver
    {
        [CmdletBinding()]
        PARAM
        (
            [Parameter(Position=1, ValueFromPipelineByPropertyName=$true)]
            [Alias("FullName","Path")]
            $InfFile,

            [string]$CategoryName

        )

        Process
        {
            # Split the path
            $DriverINF = split-path $infFile -leaf 
            $DriverPath = split-path $infFile

            # Get the category if specified
            if ([string]::IsNullOrEmpty($CategoryName)) {

                Write-Verbose "No driver category specified"

            } else {

                $DriverCategory = Get-Category -Name $CategoryName

                if ($DriverCategory -eq $null) {

                    $DriverCategory = New-DriverCategory $CategoryName
                    Write-Verbose "Created new driver category $CategoryName"

                }

                $CategoryID = $DriverCategory.CategoryInstance_UniqueID
            }

            try {

                # Create a Driver object from Inf file
                $Result = Invoke-CMMethod -Class SMS_Driver -Name "CreateFromInf" -Arguments @($DriverPath, $DriverINF)

                # Get the display name out of the result
                $DriverXML = [XML]$Result.Driver.SDMPackageXML
                $DisplayName = $DriverXML.DesiredConfigurationDigest.Driver.Annotation.DisplayName.Text

                # Populate the localized settings to be used with the new driver instance
                $localizedSetting = New-CMInstance -Class SMS_CI_LocalizedProperties
                $localizedSetting.LocaleID =  1033 
                $localizedSetting.DisplayName = $DisplayName
                $localizedSetting.Description = ""
                [System.Management.ManagementObject[]] $ocalizedSettings += $localizedSetting

                # Prepare needed properties
                $Arguments = @{
                    SDMPackageXML = $Result.Driver.SDMPackageXML
                    ContentSourcePath = $Result.Driver.ContentSourcePath
                    IsEnabled = $true
                    LocalizedInformation = @($LocalizedSettings)
                }

                if ($CategoryID -ne $null) {

                    $Arguments | Add-Member -MemberType NoteProperty -Name CategoryInstance_UniqueIDs -Value @($CategoryID)
                }

                # Create new Driver 
                $Driver = New-CMObject -Class SMS_Driver -Arguments $Arguments
        
                Write-Verbose "Added new driver."

            } catch [System.Management.Automation.MethodInvocationException] {
                
                $exc = $_.Exception.GetBaseException()

                if ($exc.ErrorInformation.ErrorCode -eq 183) {

                    Write-Verbose "Duplicate driver found: $($exc.ErrorInformation.ObjectInfo)"

                    # Look for a match on the CI_UniqueID    
                    $CIUniqueID = $exc.ErrorInformation.ObjectInfo
                    $Driver = Get-Driver -UniqueID $CIUniqueID
                        
                    # Set the category
                    if ($CategoryID -ne $null) {
                        if (-not $Driver) {

                            Write-Warning "Unable to import and no existing driver found."
                            return

                        } elseif ($Driver.CategoryInstance_UniqueIDs -contains $CategoryID) {

                            Write-Verbose "Existing driver is already in the specified category."

                        } else {

                            $Driver.CategoryInstance_UniqueIDs += $CategoryID
                            $null = $Driver.Put()
                            Write-Verbose "Added category on existing driver."

                        }
                    }

                    # Check if ContentSourcePath is still valid
                    if (-not (Test-Path($driver.ContentSourcePath))) {

                        Write-Verbose "Existing driver path ""$($driver.ContentSourcePath)"" isn't valid. Updating with current driver path."
                        $Driver.ContentSourcePath = $driverPath
                        $null = $Driver.Put()

                    }

                } else {

                    Write-Warning "Unexpected error, skipping INF $($InfFile): $($exc.ErrorInformation.Description) $($exc.ErrorInformation.ErrorCode)"
                    return

                }
            }

            # Hack - for some reason without this we don't get the CollectionID value
            #$hack = $driver.PSBase | select * | out-null

            # Write the driver object to the pipeline
            $Driver
        }

    }


    # Adds a Driver to a Driver Package
    function Add-DriverToDriverPackage
    {
        [CmdletBinding()]
        PARAM
        (
            # ConfigMgr Driver Package
            [Parameter(Mandatory)]
            [System.Management.ManagementObject]$Package,

            # ConfigMgr Driver
            [Parameter(Mandatory)]
            [System.Management.ManagementObject] $Driver
        )

        # Get the list of content IDs
        $IdList = @()
        $CIID = $Driver.CI_ID

        $IDs = Get-CMObject -Class SMS_CIToContent -Filter "CI_ID = '$CIID'"

        if ($ids -eq $null) {

            Write-Warning "Driver not found in SMS_CIToContent"

        } else {

            foreach ($id in $ids) {

                $IdList += $id.ContentID

            }

            # Build a list of content source paths (one entry in the array)
            $Sources = @($Driver.ContentSourcePath)

            # Invoke the method
            try {

                $Package.AddDriverContent($IdList, $Sources, $false)

            } catch [System.Management.Automation.MethodInvocationException] {
                $exc = $_.Exception.GetBaseException()

                if ($exc.ErrorInformation.ErrorCode -eq 1078462229) {

                    Write-Verbose "Driver is already in the driver package (possibly because there are multiple INFs in the same folder or the driver already was added from a different location): $($exc.ErrorInformation.Description)"

                }

            }

        }

    }

    function Remove-DriverFromDriverPackage {

        [CmdLetBinding()]
        PARAM (

            [Parameter(Mandatory)]
            $DriverPackage,

            [Parameter(Mandatory, ValueFromPipeline)]
            $Driver
        )

        Process {

            $DriverCIID = $Driver.CI_ID

            $DriverContent = @(Get-CMObject -Class SMS_CIToContent -Filter "CI_ID = '$DriverCIID'")
            if ($DriverContent -eq $null) {

                Write-Warning "Warning: Driver not found in SMS_CIToContent"

            } else {
                    
                $IDs = @()

                foreach ($CI in $DriverContent) {
                    $IDs += $CI.ContentID    
                }

                if ($IDs -ne $null) {

                    Write-Verbose "Removing driver from Driver package"
                    $DriverPackage.RemoveDriverContent($IDs, $false)

                }

            }

        }

    }

    # Updates the DriverContentSourcePath property
    # Iterates through all inf files with the same name
    # and verifies if it has the same uniqueID as an 
    # existing driver.
    function Update-DriverContentSourcePath {
        
        [CmdLetBinding(SupportsShouldProcess)]
        PARAM(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            $Driver,

            [DirectoryInfo]$DriverSourcePath
        )


        # Check if driver points to current SourcePath
        # If not, it's a duplicate from a different package then no cleanup is necessary
        $DriverSourcePathName = $SourcePath.FullName
        if ($driver.ContentSourcePath -like "*$DriverSourcePathName*") {

            Write-Verbose "Driver ContentSourcePath has been removed. Searching for other possible content."
            $UpdatedContentSourcePath = $false

            # Driver Content Source path was from this package
            # Get all inf files with the same name
            # And check if any of them has the same CI_UniqueID
            $NewPathFound = $false
            $InfFile = $Driver.DriverInfFile

            $PossibleInfFiles = Get-ChildItem "$($DriverSourcePath.FullName)" -Filter "$InfFile)" -recurse
            foreach ($infFile in $PossibleInfFiles) {
            
                # Import each driver and see if the Unique ID matches
                $NewDriver = Import-Driver -InfFile $_.FullName 
                
                if ($NewDriver.CI_UniqueID -eq $Driver.CI_UniqueID) {     
                    
                    # Found a different path for the same driver
                    # Update ContentSourcePath
                    $Driver.ContentSourcePath = $_.Parent.FullName
                    $Driver.Put()

                    $NewPathFound = $true
                    break

                }

            }

            
            if (!$NewPathFound) {

                # Add "BrokenContenSourcePath" category to Driver
                $Broken = Get-Category -Name "BrokenContentSourcePath"
                
                if ($Broken -eq $null) {

                    $Broken = New-DriverCategory -Name "BrokenContentSourcePath"

                }
   
                $BrokenID = $Broken.CategoryInstance_UniqueID

                if ($Driver.CategoryInstance_UniqueIDs -notcontains $BrokenID) {
                                
                    $Driver.CategoryInstance_UniqueIDs += $BrokenID
                    $null = $Driver.Put()
                    Write-Host "Added category ""BrokenContentSourcePath"" on driver."

                }

            } else {

                # Remove "BrokenContentSourcePath" Category if necessary
                $Broken = Get-Category -Name "BrokenContentSourcePath"
                
                if ($Broken -ne $null) {

                    $BrokenID = $Broken.CategoryInstance_UniqueID

                    if ($Driver.CategoryInstance_UniqueIDs -contains $BrokenID) {
                                
                        $Driver.CategoryInstance_UniqueIDs -= $BrokenID
                        $null = $Driver.Put()
                        Write-Host "Removed category ""BrokenContentSourcePath"" on driver."

                    }

                }
                
            }

        }

    }

    # Returns a ConfigMgr Category
    function Get-Category {

        [CmdletBinding()]
        PARAM
        (
            # Category Name
            [Parameter(Mandatory, ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [string]$Name
        )

        Get-CMObject -Class SMS_CategoryInstance -Filter "LocalizedCategoryInstanceName = '$Name'"
    }

    # Ceates a new ConfigMgr Category
    function New-Category {

        [CmdletBinding()]
        PARAM
        (
            # Category Name
            [Parameter(Mandatory, ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            [Parameter(Mandatory)]
            [ValidateSet("Locale", "UpdateClassification", "Company", "ProductFamily", "UserCategories", "GlobalCondition", "Device", "Platform", "AppCategories", "Update Type", "CatalogCategories", "DriverCategories", "SettingsAndPolicy")]
            [string]$TypeName,

            [string]$LocaleID = 1033

        )

        # Create Localized Properties
        
        $LocalizedSettings =  New-CMInstance -Class SMS_Category_LocalizedProperties
        $LocalizedSettings.CategoryInstanceName = $Name
        $LocalizedSettings.LocaleID = $LocaleID

        # Build the parameters for creating the  Category
        $CategoryGuid = [System.Guid]::NewGuid().ToString()
        $Colon = ":"
        $UniqueID = "$TypeName$Colon$CategoryGuid"

        $Arguments = @{CategoryInstance_UniqueID = $UniqueID; LocalizedInformation = $LocalizedSettings; SourceSite = $script:CMSiteCode; CategoryTypeName = '$TypeName'}

        New-CMObject -Class SMS_CategoryInstance -Arguments $Arguments
    }


    # Ceates a new ConfigMgr Driver Category
    function New-DriverCategory {

        [CmdletBinding()]
        PARAM
        (
            # Category Name
            [Parameter(Mandatory, ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [string]$Name
        )

        New-Category -Name $Name -TypeName DriverCategories
    }

    # Removes a category from the driver
    function Remove-CategoryFromDriver {

        [CmdLetBinding(SupportsShouldProcess)]
        PARAM(
            
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            $Driver,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            $Category
        )


        Write-Verbose "Removing category ""$($Category.LocalizedCategoryInstanceName)"" (""$($Category.CategoryInstance_UniqueID)"")"

        $UpdatedCategories = @($Driver.CategoryInstance_UniqueIDs | ? {$_ -ne $Category.CategoryInstance_UniqueID})
                    
        if ($UpdatedCategories -eq $null) {

            Write-Host "Drive has no category left. Remove from Driver Store."
            if ($PSCmdlet.ShouldProcess("Driver $($Driver.LocalizedDisplayName)", "Remove Driver")) {             
                $Driver.Delete()
            }

        } else {
            
            if ($PSCmdlet.ShouldProcess("Driver $($Driver.LocalizedDisplayName)", "Remove Category $($Category.LocalizedCategoryInstanceName)")) {             

                $Driver.CategoryInstance_UniqueIDs = $UpdatedCategories
                $Driver.Put()

            }

        }

    }

    # Returns a ConfigMgr Folder
    function Get-Folder {

        [CmdletBinding()]
        PARAM
        (
            # Category Name
            [Parameter(Mandatory, ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [string]$Name
        )

        Get-CMObject -Class SMS_ObjectContainerNode -Filter "Name = '$Name'"
    }

    # Ceates a new ConfigMgr Folder
    function New-Folder {

        [CmdletBinding()]
        PARAM
        (
            # Folder Name
            [Parameter(Mandatory, ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            # Type of object to place in the console folder
            [Parameter(Mandatory, ValueFromPipeline)]
            [ValidateSet("Package", "Advertisement", "Query", "Report", "MeteredProductRule", "ConfigurationItem", "OperatingSystemInstallPackage", "StateMigration", "ImagePackage", "BootImagePackage", "TaskSequencePackage", "DeviceSettingPackage", "DriverPackage", "Driver", "SoftwareUpdate", "BaselineConfigurationItem")]
            [string]$Type,

            # The unique ID of the parent folder. 
            [Parameter(ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [string]$ParentFolderID = 0
        )

        Switch ($Type) {
            "Package" {$ObjectType = 2}
            "Advertisement" {$ObjectType = 3}
            "Query" {$ObjectType = 7}
            "Report" {$ObjectType = 8}
            "MeteredProductRule" {$ObjectType = 9}
            "ConfigurationItem" {$ObjectType = 11}
            "OperatingSystemInstallPackage" {$ObjectType = 14}
            "StateMigration" {$ObjectType = 17}
            "ImagePackage" {$ObjectType = 18}
            "BootImagePackage" {$ObjectType = 19}
            "TaskSequencePackage" {$ObjectType = 20}
            "DeviceSettingPackage" {$ObjectType = 21}
            "DriverPackage" {$ObjectType = 23}
            "Driver" {$ObjectType = 25}
            "SoftwareUpdate" {$ObjectType = 1011}
            "BaselineConfigurationItem" {$ObjectType = 2011}
            default {Throw "Unknown ObjectType $Type"}
        }

        $Arguments = @{Name = $Name; ObjectType = $ObjectType; ParentContainerNodeid = $ParentFolderID}

        New-CMObject -Class SMS_ObjectContainerNode -Arguments $Arguments
    }

    # Generates a Hash value for the specified File
    Function Get-ContentHash {

        [CmdLetBinding()]
        Param (

            [Parameter(Mandatory)]
            $File,

            [ValidateSet("sha1","md5")]
            [string]$Algorithm="md5"

        )
	
        $content = "$($file.Name)$($file.Length)"
        $algo = [type]"System.Security.Cryptography.md5"
	    $crypto = $algo::Create()
        $hash = [BitConverter]::ToString($crypto.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($content))).Replace("-", "")
        $hash
    }

    # Generates a Hash value for the specified folder
    # Used to detect changes
    Function Get-FolderHash {
        
        [CmdLetBinding()]
        Param (

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Folder,

            [ValidateSet("sha1","md5")]
            [string]$Algorithm="md5"
        )
    
	    Get-ChildItem $Folder -Recurse -Exclude "*.hash" | % { $content += Get-ContentHash $_ $Algorithm }
    
        $algo = [type]"System.Security.Cryptography.$Algorithm"
	    $crypto = $algo::Create()
	    $hash = [BitConverter]::ToString($crypto.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($content))).Replace("-", "")
    
        $hash
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

    # Creates a new ConfigMgr Object
    function New-CMObject {

        [CmdletBinding()]
        PARAM (

            # ConfigMgr WMI provider Class
            [Parameter(Mandatory)] 
            [ValidateNotNullOrEmpty()]
            [string]$Class, 

            # Arguments to be supplied to the new object
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [hashtable]$Arguments

        )

        if ([string]::IsNullOrWhiteSpace($Class)) { throw "Class is not specified" }

        # Ensure ConfigMgr Provider information is available
        if (Test-CMConnection) {

            $ArgumentsString = ($Arguments.GetEnumerator() | foreach { "$($_.Key)=$($_.Value)" }) -join "; "
            Write-Verbose "Create new $Class object. Arguments: $ArgumentsString"

            if ($CMCredential -ne $null) {

                $NewCMObject = Set-WmiInstance -Class $Class -Arguments $Arguments -ComputerName $CMProviderServer -Namespace $CMNamespace -Credential $CMCredential

            } else {
         
                $NewCMObject = Set-WmiInstance -Class $Class -Arguments $Arguments -ComputerName $CMProviderServer -Namespace $CMNamespace

            }

            if ($NewCMObject -ne $null) {

                # Ensure that generated properties are udpated
                $hack = $NewCMObject.PSBase | select * | Out-Null

            }

            Return $NewCMObject
        }
    }

    # Creates a new configMgr WMI class instance
    function New-CMInstance {

        [CmdletBinding()]
        PARAM (

            # ConfigMgr WMI provider Class
            [Parameter(Mandatory)] 
            [ValidateNotNullOrEmpty()]
            [string]$Class

        )

        if ([string]::IsNullOrWhiteSpace($Class)) { throw "Class is not specified" }

        # Ensure ConfigMgr Provider information is available
        if (Test-CMConnection) {

            Write-Verbose "Create new $Class instance."

            if ($CMCredential -ne $null) {

                $CMClass = Get-WmiObject -List -Class $Class -ComputerName $CMProviderServer -Namespace $CMNamespace -Credential $CMCredential

            } else {
         
                $CMClass = Get-WmiObject -List -Class $Class -ComputerName $CMProviderServer -Namespace $CMNamespace

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