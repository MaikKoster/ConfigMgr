
[CmdLetBinding()]
PARAM (
    
    [Parameter(Mandatory)]
    [ValidateScript({Test-Path $_ -PathType 'Container'})]
    [ValidateScript({$_.StartsWith("\\")})]
    [string]$DriverSourceRootFolder,

    [Parameter(Mandatory)]
    [ValidateScript({Test-Path $_ -PathType 'Container'})]
    [ValidateScript({$_.StartsWith("\\")})]
    [string]$DriverPackageRootFolder,

    # The total amount of levels in the driver folder hierarchy
    [int]$HierarchyLevels = 3,

    # The current level in the driver folder hierarchy
    [int]$CurrentHierarchyLevel = 0,
    
    [string]$PackageNameDelimiter = " - ",

    [String]$CategoryNameDelimiter = " - "

    

)


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

        # Evaluate Folder Hierarchy
        $FolderHierarchy = Get-FolderHierarchy -Path $DriverSourceRootFolder -MaxLevel $HierarchyLevels -CurrentLevel $CurrentHierarchyLevel

        Foreach ($Folder In $FolderHierarchy) {
            
            # Generate names
            $PackageName = [string]::Join($PackageNameDelimiter, ("$Folder.Folders\$Folder.Name").Split("\"))
            $CategoryName = [string]::Join($CategoryNameDelimiter, ("$Folder.Folders\$Folder.Name").Split("\"))

            # Verify the appropriate folders are created within ConfigMgr
            $FolderID = Verify-DriverPackageFolderStructure -Folders $Folder.Folders

            # Start Import
            Import-DriverPackageSourceFolder -DriverSourceRootFolder $Folder.Root -DriverPackageRootPath -Folders $Folder.Folders -PackageName $PackageName -CategoryName $CategoryName -CMFolderID $FolderID

        }

    }

}

Begin {

    Set-StrictMode -Version Latest

    ###############################################################################
    # Function definitions
    ###############################################################################

    function Get-FolderHierarchy {

        [CmdLetBinding()]
        PARAM (

            [Parameter(Mandatory,ValueFromPipeline)]
            [System.IO.DirectoryInfo]$Path,

            # Maximum amount of Hierarchy levels
            [int]$MaxLevel = 3,

            # The current Hierarchy level
            [int]$CurrentLevel = 0

        )

        Process {

            if ($CurrentLevel -eq $MaxLevel) {
                # This is supposed to be the content for the driver package
            
                # Get Root Folder
                $Root = $Path
                if ($MaxLevel -gt 0) { for ($x=1; $x -le $MaxLevel; $x++) { $Root = $Root.Parent } }
            
                # Get Subfolders
                $Folders = $Path.Parent.FullName.Replace("$($root.FullName)", "").Trim("\")
            
                # Create result
                $Result = @{
            
                    Name = $path.Name
                    Root = $Root.FullName
                    Folders = $Folders
            
                } 

                # Drop Result to pipeline
                $Result
        
            } elseif ($CurrentLevel -gt $MaxLevel) {

                Write-Error "Current level $CurrentLevel is larger than the Hierarchy Maximum level of $MaximumLevel. Aborting further processing ..." -ErrorAction Stop

            } else {

                # Iterate through subfolders
                Get-ChildItem $Path.FullName -Directory | Get-FolderHierarchy -CurrentLevel ($CurrentLevel + 1) -MaxLevel $MaxLevel
            
            }

        }

    }

    # Imports all drivers from the specified Source path to the specified Driver Package
    # If the Driver package does not exist, it will be created.
    function Import-DriverPackageSourceFolder {

        [CmdLetBinding()]
        PARAM (

            # Source path of the drivers for the Driver Package
            [Parameter(Mandatory)]
            [string]$DriverSourceRootFolder,

            # Source path of the package
            [Parameter(Mandatory)]
            [string]$DriverPackageRootFolder,

            # Current Subfolder structure
            [Parameter(Mandatory)]
            [string]$Folders,

            # Package name
            [Parameter(Mandatory)]
            [string]$PackageName,

            # Category name
            [Parameter(Mandatory)]
            [string]$CategoryName,

            # ConfigMgr Folder for the Driver Package
            [Parameter(Mandatory)]
            [uint32]$FolderID,

            # Cleanup content if new package path already exists
            [switch]$Cleanup

        )

        #
        $DriverSourcePath = Join-Path -Path $DriverSourceRootFolder -ChildPath $Folders
        $PackageSourcePath = Join-Path -Path $DriverPackageRootFolder -ChildPath $Folders

        # Process this folder only if there are any changes
        $PackageHash = Get-FolderHash $DriverSourcePath.FullName

        if (Get-ChildItem $package.FullName -Filter "$($PackageHash).hash") {

		    Write-Verbose "No changes have been made to this Driver Package. Skipping."

	    } else {

            # Get Driver Package
            # TODO: Add logic for same package name in different folders
            $DriverPackage = Get-DriverPackage -Name $PackageName -ParentFolderID $FolderID

            # Create the package if it doesn't already exist
            if ($DriverPackage -eq $null) {

                $DriverPackageSource = "$PackageSourcePath\$PackageName"
                if ((Get-Item $DriverPackageSource | %{$_.GetDirectories().Count + $_.GetFiles().Count}) -gt 0) {
					if ($cleanup) {

					    Write-Verbose "Folder already exists, removing content"
    				    dir $DriverPackageSource | Remove-Item -recurse -force

					} else {

						Write-Error "Folder already exists, remove it manually."
						return
					}

				} else {

                    MkDir "$DriverPackageSource" | Out-Null

                }

                $DriverPackage = New-DriverPackage -Name "$Name" -SourcePath $DriverPackageSource

                Move-CMObject -TargetFolderID $FolderID -ObjectID $DriverPackage.PackageID

                Write-Host "Created driver package $($DriverPackage.PackageID)"

            } else {

                Write-Host "Existing driver package $($DriverPackage.PackageID) retrieved."

            }

            # Get current list of drivers
            $PackageID = $DriverPackage.PackageID
            $CurrentDrivers = Get-Driver -PackageID $PackageID

            # Import all Drivers from the Driver Source and add to driver package
            $NewDrivers = Get-ChildItem $DriverSourcePath.FullName -Filter *.inf -Recurse | Import-Driver -CategoryName $CategoryName | Add-DriverToDriverPackage -Package $DriverPackage

            # Remove all drivers that are now longer available in the Driver source from the package and adjust the category 
            $RemovedDrivers = $CurrentDrivers | Where { $NewDrivers -notcontains $_ }  | Remove-DriverFromDriverPackage -Package $DriverPackage

            # Remove all Drivers from the Driver pool that have no category anymore, assuming they don't belong to any package now.
            $RemovedDrivers | Where { $_.CategoryInstance_UniqueIDs -eq $null } | Remove-Driver

            # Try to update ContentSourcePath of removed drivers that are in other driver packages, if it is pointing to the current package path
            # If it can't be updated the "BrokenContentSourcePath" category will be added
            $RemovedDrivers | Where { ($_.CategoryInstance_UniqueIDs -ne $null) -and ($_.DriverContentSourcePath -like "*$Folders\$PackageName*") } | Update-DriverContentSourcePath -DriverSourceRootFolder $DriverSourceRootFolder 

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
            [string]$Name,

            #ParentFolderID
            [Parameter(ParameterSetName="Name")]
            [uint32]$ParentFolderID = 0

        )

        if (![string]::IsNullOrEmpty($PackageID)) {

            Get-CMObject -Class SMS_DriverPackage -Filter "PackageID = '$PackageID'"

        } elseif (![string]::IsNullOrEmpty($Name)) {

            
            if ($ParentFolderID -gt 0) {

                Get-CMObject -Class SMS_DriverPackage -Filter "Name = '$Name' AND PackageID IN (SELECT InstanceKey FROM SMS_ObjectContainerItem WHERE ContainerNodeID = $ParentFolderID)"
            
            } else {

                Get-CMObject -Class SMS_DriverPackage -Filter "Name = '$Name'"

            }

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
    
        Write-Verbose "Created new Driver Package $($NewPackage.Name) ($($NewPackage.PackageID))"

        # Return the package
        $NewPackage

    }

    # Returns the drivers
    function Get-Driver {

        [CmdletBinding()]
        PARAM
        (

            # Driver PackageID
            [Parameter(Mandatory,ParameterSetName="PackageID")]
            [ValidateNotNullOrEmpty()]
            [string]$PackageID,

            # Driver Unique ID
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

    # Removes a Driver from the Configuration Manager Driver Store
    function Remove-Driver {

        [CmdletBinding(SupportsShouldProcess)]
        PARAM
        (
            [Parameter(Mandatory,ValueFromPipeline)]
            [System.Management.ManagementObject]$Driver

        )

        Process {

            Write-Verbose "Removing Driver $($Driver.LocalizedDisplayName) from Driver Store."
            if ($PSCmdlet.ShouldProcess("Driver $($Driver.LocalizedDisplayName)", "Remove Driver")) {
                         
                $Driver.Delete() | Out-Null

            }
            
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

        Begin {

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

        }

        Process
        {
            # Split the path
            $DriverINF = split-path $infFile -leaf 
            $DriverPath = split-path $infFile

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

                # Prepare driver properties
                $Arguments = @{
                    SDMPackageXML = $Result.Driver.SDMPackageXML
                    ContentSourcePath = $Result.Driver.ContentSourcePath
                    IsEnabled = $true
                    LocalizedInformation = @($LocalizedSettings)
                }

                # Add Category if available 
                if (![string]::IsNullOrEmpty($CategoryID)) { $Arguments | Add-Member -MemberType NoteProperty -Name CategoryInstance_UniqueIDs -Value @($CategoryID) }

                # Create new Driver 
                $Driver = New-CMObject -Class SMS_Driver -Arguments $Arguments
        
                Write-Verbose "Added new driver $($Driver.LocalizedDisplayname)."

            } catch [System.Management.Automation.MethodInvocationException] {
                
                $exc = $_.Exception.GetBaseException()

                if ($exc.ErrorInformation.ErrorCode -eq 183) {

                    Write-Verbose "Duplicate driver found: $($exc.ErrorInformation.ObjectInfo)"

                    # Look for a match on the CI_UniqueID    
                    $CIUniqueID = $exc.ErrorInformation.ObjectInfo
                    $Driver = Get-Driver -UniqueID $CIUniqueID
                        
                    # Set the category
                    if (![string]::IsNullOrEmpty($CategoryID)) {
                        if ($Driver -eq $null) {

                            Write-Warning "Unable to import and no existing driver found."
                            return

                        } elseif ($Driver.CategoryInstance_UniqueIDs -contains $CategoryID) {

                            Write-Verbose "Existing driver $($Driver.LocalizedDisplayname) is already in the specified category $($Category.LocalizedCategoryInstanceName)."

                        } else {

                            $Driver.CategoryInstance_UniqueIDs += $CategoryID
                            $Driver.Put() | Out-Null
                            Write-Verbose "Added category $($Category.LocalizedCategoryInstanceName) on existing driver $($Driver.LocalizedDisplayname)."

                        }
                    }

                    # Check if ContentSourcePath is still valid
                    if (!(Test-Path($Driver.ContentSourcePath))) {

                        Write-Verbose "Existing driver path ""$($Driver.ContentSourcePath)"" isn't valid. Updating with current driver path."
                        $Driver.ContentSourcePath = $DriverPath
                        $Driver.Put() | Out-Null

                    }

                } else {

                    Write-Warning "Unexpected error, skipping INF $($InfFile): $($exc.ErrorInformation.Description) $($exc.ErrorInformation.ErrorCode)"
                    return

                }
            }

            # Write the driver object back to the pipeline
            $Driver
        }

    }


    # Adds a Driver to a Driver Package
    function Add-DriverToDriverPackage
    {
        [CmdletBinding()]
        PARAM
        (
        
            # ConfigMgr Driver
            [Parameter(Mandatory, ValueFromPipeline)]
            [System.Management.ManagementObject]$Driver,

            # ConfigMgr Driver Package
            [Parameter(Mandatory)]
            [System.Management.ManagementObject]$Package

        )

        Begin {

            # Get current list of drivers for the Package
            # to only add drivers that aren't available yet
            $PackageID = $Package.PackageID
            $CurrentDrivers = Get-Driver -PackageID $PackageID

        }

        Process {

            if ($CurrentDrivers -notcontains $Driver) {

                Write-Verbose "Adding driver $($Driver.LocalizedDisplayName) to package $($Package.Name)"
                
                # Get the list of content IDs
                $ContentIDs = @(Get-CMObject -Class SMS_CIToContent -Filter "CI_ID = '$($Driver.CI_ID)'" | select -ExpandProperty ContentID)

                if (($ContentIDs -ne $null) -and ($ContentIDs.Count -gt 0)) {

                    # Build a list of content source paths (one entry in the array)
                    $Sources = @($Driver.ContentSourcePath)

                    # Invoke the method
                    try {

                        $Package.AddDriverContent($ContentIDs, $Sources, $false)

                    } catch [System.Management.Automation.MethodInvocationException] {

                        $exc = $_.Exception.GetBaseException()

                        if ($exc.ErrorInformation.ErrorCode -eq 1078462229) {

                            Write-Verbose "Driver is already in the driver package (possibly because there are multiple INFs in the same folder or the driver already was added from a different location): $($exc.ErrorInformation.Description)"

                        }

                    }

                } else {

                    Write-Warning "Driver not found in SMS_CIToContent"

                }

            } else {

                Write-Verbose "Driver $($Driver.LocalizedDisplayName) is already available in package $($Package.Name)"

            }

            # Write the Driver object back to the pipeline
            $Driver

        }
        
    }

    # Removes a Driver from the Driver Package
    function Remove-DriverFromDriverPackage {

        [CmdLetBinding()]
        PARAM (

            [Parameter(Mandatory, ValueFromPipeline)]
            $Driver,

            [Parameter(Mandatory)]
            $Package
            
        )

        Process {

            $ContentIDs = Get-CMObject -Class SMS_CIToContent -Filter "CI_ID = '$($Driver.CI_ID)'" | select -ExpandProperty ContentID

            if (($ContentIDs -ne $null) -and ($ContentIDs.Count -gt 0)) {

                Write-Verbose "Removing driver $($Driver.LocalizedDisplayName) from Driver package $($Package.Name)"
                $DriverPackage.RemoveDriverContent($ContentIDs, $false)

            } else {
                    
                Write-Warning "Driver $($Driver.LocalizedDisplayName) ($($Driver.CI_ID)) not found in SMS_CIToContent"
            }

            # Write the Driver object back to the pipeline
            $Driver

        }

    }

    # Updates the DriverContentSourcePath property
    # Iterates through all inf files with the same name
    # and verifies if it has the same uniqueID as an 
    # existing driver.
    function Update-DriverContentSourcePath {
        
        [CmdLetBinding(SupportsShouldProcess)]
        PARAM(
            [Parameter(Mandatory,ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            $Driver,

            [string]$DriverSourceRootFolder
        )

        Begin {

            $BrokenCategory = Get-Category -Name "BrokenContentSourcePath"
                
            if ($BrokenCategory -eq $null) { $BrokenCategory = New-DriverCategory -Name "BrokenContentSourcePath" }

            $BrokenCategoryID = $BrokenCategory.CategoryInstance_UniqueID

        }

        Process {

            # Check if driver points to current SourcePath
            # If not, it's a duplicate from a different package then no cleanup is necessary

            Write-Verbose "Driver ContentSourcePath has been removed. Searching for duplicate driver in $($DriverSourceRootFolder.Fullname)."
            $UpdatedContentSourcePath = $false

            # Driver Content Source path was from this package
            # Get all inf files with the same name
            # And check if any of them has the same CI_UniqueID
            $NewPathFound = $false
            $InfFile = $Driver.DriverInfFile

            $PossibleInfFiles = Get-ChildItem "$DriverSourceRootFolder" -Filter "$InfFile)" -recurse
            foreach ($infFile in $PossibleInfFiles) {
            
                # Import each driver and see if the Unique ID matches
                $NewDriver = Import-Driver -InfFile $_.FullName 
                
                if ($NewDriver.CI_UniqueID -eq $Driver.CI_UniqueID) {     
                    
                    # Found a different path for the same driver
                    # Update ContentSourcePath
                    Write-Verbose "Found alternative ContentSourcePath ""$($_.Parent.FullName)"" and updating Driver $($Driver.LocalizedDisplayname)."
                    $Driver.ContentSourcePath = $_.Parent.FullName
                    $Driver.Put() | Out-Null

                    $NewPathFound = $true
                    break

                }

            }
            
            if (!$NewPathFound) {

                # Add "BrokenContenSourcePath" category to Driver
                if ($Driver.CategoryInstance_UniqueIDs -notcontains $BrokenCategoryID) {
                                
                    Write-Warning "ContentSourcePath no longer valid for Driver $($Driver.LocalizedDisplayName). Adding Category ""BrokenContentSourcePath""."
                    $Driver.CategoryInstance_UniqueIDs += $BrokenCategoryID
                    $Driver.Put() | Out-Null

                }

            } else {

                # Remove "BrokenContentSourcePath" Category if necessary
                if ($Driver.CategoryInstance_UniqueIDs -contains $BrokenID) {
                            
                    Write-Verbose "ContentSourcePath for Driver $($Driver.LocalizedDisplayName) was updated. Removing Category ""BrokenContentSourcePath""."
                    $Driver.CategoryInstance_UniqueIDs -= $BrokenID
                    $Driver.Put() | Out-Null

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
            
            [Parameter(Mandatory, ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            $Driver,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            $Category
        )

        Begin {

            

        }


        Process {

            Write-Verbose "Removing category $($Category.LocalizedCategoryInstanceName) from Driver $($Driver.LocalizedDisplayname)"

            $UpdatedCategories = @($Driver.CategoryInstance_UniqueIDs | Where {$_ -ne $Category.CategoryInstance_UniqueID})

            if ($PSCmdlet.ShouldProcess("Driver $($Driver.LocalizedDisplayName)", "Remove Category $($Category.LocalizedCategoryInstanceName)")) {             

                $Driver.CategoryInstance_UniqueIDs = $UpdatedCategories
                $Driver.Put() | Out-Null

            }

            # Write the Driver object back to the pipeline
            $Driver

        }

    }

    # Returns a ConfigMgr Folder
    function Get-Folder {

        [CmdletBinding()]
        PARAM
        (

            # Folder ID
            [Parameter(Mandatory, ParameterSetName="ID")]
            [uint32]$ID,

            # Folder Name
            [Parameter(Mandatory, ParameterSetName="Name")]
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            # Folder type
            [Parameter(Mandatory, ParameterSetName="Name")]
            [ValidateSet("Package", "Advertisement", "Query", "Report", "MeteredProductRule", "ConfigurationItem", "OSInstallPackage", "StateMigration", "ImagePackage", "BootImagePackage", "TaskSequencePackage", "DeviceSettingPackage", "DriverPackage", "Driver", "SoftwareUpdate", "ConfigurationBaseline", "DeviceCollection", "UserCollectioN")]
            [string]$Type,

            # ParentID
            [Parameter(ParameterSetName="Name")]
            [uint32]$ParentFolderID = 0
        )

        switch ($Type) {

            "Package" {$TypeID = 2}
            "Advertisement" {$TypeID = 3}
            "Query" {$TypeID = 7}
            "Report" {$TypeID = 8}
            "MeteredProductRule" {$TypeID = 9}
            "ConfigurationItem" {$TypeID = 11}
            "OSInstallPackage" {$TypeID = 14}
            "StateMigratino" {$TypeID = 17}
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

        if ([string]::IsNullOrEmpty($Name)) {

            Get-CMObject -Class SMS_ObjectContainerNode -Filter "((Name = '$Name') AND (ObjectType=$TypeID) AND (ParentContainerNodeID = $ParentFolderID))"

        } else {

            Get-CMObject -Class SMS_ObjectContainerNode -Filter "ContainerNodeID = $ID"

        }
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
            [ValidateSet("Package", "Advertisement", "Query", "Report", "MeteredProductRule", "ConfigurationItem", "OSInstallPackage", "StateMigration", "ImagePackage", "BootImagePackage", "TaskSequencePackage", "DeviceSettingPackage", "DriverPackage", "Driver", "SoftwareUpdate", "ConfigurationBaseline", "DeviceCollection", "UserCollectioN")]
            [string]$Type,

            # The unique ID of the parent folder. 
            [Parameter(ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [string]$ParentFolderID = 0
        )

        Switch ($Type) {

            "Package" {$TypeID = 2}
            "Advertisement" {$TypeID = 3}
            "Query" {$TypeID = 7}
            "Report" {$TypeID = 8}
            "MeteredProductRule" {$TypeID = 9}
            "ConfigurationItem" {$TypeID = 11}
            "OSInstallPackage" {$TypeID = 14}
            "StateMigratino" {$TypeID = 17}
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

        $Arguments = @{Name = $Name; ObjectType = $ObjectType; ParentContainerNodeid = $ParentFolderID}

        New-CMObject -Class SMS_ObjectContainerNode -Arguments $Arguments
    }

    # Verifies that the supplied Driver Package Folder structure is created and returns the ID of the leaf node folder
    function Verify-DriverPackageFolderStructure {

        [CmdLetBinding()]
        PARAM (

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Folders

        )

        $ParentID = 0
        foreach ($FolderName In $Folders.Split("\")) {

            $CMFolder = Get-Folder -Name $FolderName -Type DriverPackage -ParentID $ParentID

            if ($CMFolder -eq $null) {

                $CMFolder = New-Folder -Name $FolderName -Type DriverPackage -ParentFolderID $ParentID

            }

            $ParentID = $CMFolder.ContainerNodeID

        }

        Return $ParentID
    }

    # Generates a Hash value for the specified File
    function Get-ContentHash {

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
    function Get-FolderHash {
        
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

    # Moves a ConfigMgr Object to a different Folder
    Function Move-CMObject {            
        
        [CmdLetBinding()]
        PARAM (
            [Parameter(Mandatory)]        
            [uint32]$SourceFolderID,
            
            [Parameter(Mandatory)]
            [uint32]$TargetFolderID,

            [Parameter(Mandatory)]
            [string[]]$ObjectID,

            [Parameter(Mandatory)]
            [string]$ObjectType           
        )
        
        # Verify Folders exist
        $SourceFolder = Get-Folder -ID $SourceFolderID
        $TargetFolder = Get-Folder -ID $TargetFolderID

        if (($SourceFolder -ne $null) -and ($TargetFolder -ne $null)) {
            
            # Verify folders have the same ObjectType
            if ($SourceFolder.ObjectType -eq $TargetFolder.ObjectType) {

                $Arguments = @($SourceFolderID, $ObjectID, $TargetFolder.ObjectType, $TargetFolderID)

                $Result = Invoke-CMMethod -Class SMS_ObjectContainerItem -Name MoveMembers -ArgumentList $Arguments
                
            } else {

                Write-Warning "Unable to move object $ObjectID. SourceFolder $SourceFolderID ObjectType $($SourceFolder.ObjectType) and TargetFolder $TargetFolderID ObjectType $($TargetFolder.ObjectType) don't match."

            }

        } else {

            Write-Warning "Unable to move object $ObjectID. SourceFolder $SourceFolderID or TargetFolder $TargetFolderID are not available."
        }
          
    }

}