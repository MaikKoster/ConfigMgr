###############################################################################
# 
# PESTER tests for Set-ClientOperation.ps1
# 
# Author:  Maik Koster
# Date:    09.12.2015 
# History:
#          09.12.2015 - Initial setup of tests
#
###############################################################################

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here\Import-Driver.ps1"


Describe "Get-FolderHierarchy" {

    Context "Default 3 Level Hierarchy" {

        New-Item -ItemType Directory -Path "$TestDrive\root\OS1\Make1\Model1"
        New-Item -ItemType Directory -Path "$TestDrive\root\OS1\Make1\Model2"
        New-Item -ItemType Directory -Path "$TestDrive\root\OS1\Make2\Model1"
        New-Item -ItemType Directory -Path "$TestDrive\root\OS1\Make2\Model2"
        New-Item -ItemType Directory -Path "$TestDrive\root\OS2\Make1\Model1"
        New-Item -ItemType Directory -Path "$TestDrive\root\OS2\Make1\Model2"
        New-Item -ItemType Directory -Path "$TestDrive\root\OS2\Make2\Model1"
        New-Item -ItemType Directory -Path "$TestDrive\root\OS2\Make2\Model2"

        $Testpath = "$TestDrive\root"

        It "Iterate through full Hierarchy" {

            $Result = Get-FolderHierarchy -Path "$Testpath"

            $Result.Count | Should Be 8
            $Result[0].Name | Should Be "Model1"
            $Result[0].Root | Should Be "$TestDrive\root"
            $Result[0].Folders | Should Be "OS1\Make1"
            $Result[1].Name | Should Be "Model2"
            $Result[1].Folders | Should Be "OS1\Make1"
            $Result[2].Name | Should Be "Model1"
            $Result[2].Folders | Should Be "OS1\Make2"
            $Result[3].Name | Should Be "Model2"
            $Result[3].Folders | Should Be "OS1\Make2"
            $Result[4].Name | Should Be "Model1"
            $Result[4].Folders | Should Be "OS2\Make1"
            $Result[5].Name | Should Be "Model2"
            $Result[5].Folders | Should Be "OS2\Make1"
            $Result[6].Name | Should Be "Model1"
            $Result[6].Folders | Should Be "OS2\Make2"
            $Result[7].Name | Should Be "Model2"
            $Result[7].Folders | Should Be "OS2\Make2"
        }

        It "Iterate through first level" {

            $Result = Get-FolderHierarchy -Path "$Testpath\OS1" -CurrentLevel 1

            $Result.Count | Should Be 4
            $Result[0].Name | Should Be "Model1"
            $Result[0].Root | Should Be "$TestDrive\root"
            $Result[0].Folders | Should Be "OS1\Make1"
            $Result[1].Name | Should Be "Model2"
            $Result[1].Folders | Should Be "OS1\Make1"
            $Result[2].Name | Should Be "Model1"
            $Result[2].Folders | Should Be "OS1\Make2"
            $Result[3].Name | Should Be "Model2"
            $Result[3].Folders | Should Be "OS1\Make2"

        }

        It "Iterate through second level" {

            $Result = Get-FolderHierarchy -Path "$Testpath\OS1\Make1" -CurrentLevel 2

            $Result.Count | Should Be 2
            $Result[0].Name | Should Be "Model1"
            $Result[0].Root | Should Be "$TestDrive\root"
            $Result[0].Folders | Should Be "OS1\Make1"
            $Result[1].Name | Should Be "Model2"
            $Result[1].Folders | Should Be "OS1\Make1"
        }

        It "Leaf node" {

            $Result = @(Get-FolderHierarchy -Path "$Testpath\OS1\Make1\Model2" -CurrentLevel 3)

            $Result.Count | Should Be 1
            $Result.Name | Should Be "Model2"
            $Result.Root | Should Be "$TestDrive\root"
            $Result.Folders | Should Be "OS1\Make1"

        }
    }

    Context "1 Level Hierarchy" {

        New-Item -ItemType Directory -Path "$TestDrive\root\Model1"
        New-Item -ItemType Directory -Path "$TestDrive\root\Model2"

        $Testpath = "$TestDrive\root"

        It "Iterate through full Hierarchy" {

            $Result = Get-FolderHierarchy -Path "$Testpath" -MaxLevel 1

            $Result.Count | Should Be 2
            $Result[0].Name | Should Be "Model1"
            $Result[0].Root | Should Be "$TestDrive\Root"
            $Result[0].Folders | Should Be ""
            $Result[1].Name | Should Be "Model2"
            $Result[1].Folders | Should Be ""

        }

        It "Leaf node" {

            $Result = @(Get-FolderHierarchy -Path "$Testpath\Model2" -CurrentLevel 1 -MaxLevel 1)

            $Result.Count | Should Be 1
            $Result[0].Name | Should Be "Model2"
            $Result[0].Root | Should Be "$TestDrive\root"
            $Result[0].Folders | Should Be ""

        }

    }
}

Describe "New-CMConnection" {

    Context "No server available" {

        It "Throw an exception if server cannot be contacted" {
            Mock Get-WmiObject {}
        
            { New-CMConnection } | Should Throw
        }
    }

    Context "Server available" {

        Mock Get-WmiObject { 
            [PSCustomObject]@{NamespacePath = "\\$ComputerName\root\sms\Site_XXX"; Machine = "$ComputerName"; SiteCode = "XXX"} 
        } 

        Mock Get-WmiObject { 
            [PSCustomObject]@{NamespacePath = "\\TestLocalServer\root\sms\Site_XXX"; Machine = "TestLocalServer"; SiteCode = "XXX"} 
        
        } -ParameterFilter {$ComputerName -eq "."}
        
        Mock Get-WmiObject {
            [PSCustomObject]@{NamespacePath = "\\$ComputerName\root\sms\Site_ZZZ"; Machine = "$ComputerName"; SiteCode = "ZZZ"}

        } -ParameterFilter {$Query -like "*WHERE SiteCode =*"}

        It "Use ""."" if no server name was supplied." {
            New-CMConnection
            
            $script:CMProviderServer | Should Be "TestLocalServer"
        }

        It "Evaluate SiteCode if no SiteCode is supplied." {
            New-CMConnection
            
            $script:CMSiteCode | Should Be "XXX"
        }

        It "Use supplied server name." {
            New-CMConnection -ProviderServerName "TestRemoteServer"
            
            $script:CMProviderServer | Should Be "TestRemoteServer"
        }

        It "Use supplied site code." {
            New-CMConnection -ProviderServerName "TestRemoteServer" -SiteCode "ZZZ"
            
            $script:CMSiteCode | Should Be "ZZZ"
        }
    }
}

Describe "Test-CMConnection" {
    
    $script:CMProviderServer = ""
    $script:CMSiteCode = ""
    $script:CMNamespace = ""


    It "Throw if connection cannot be created" {
        Mock Get-WmiObject { Throw }

        {Test-CMConnection}  | Should Throw

    }

    It "Connect to local Site Server" {
        Mock Get-WmiObject {
            [PSCustomObject]@{NamespacePath = "\\TestLocalServer\root\sms\Site_XXX"; Machine = "TestLocalServer"; SiteCode = "XXX"}

        } -ParameterFilter {$ComputerName -eq "."}
                
        Test-CMConnection 

        $script:CMProviderServer | Should Be "TestLocalServer"

    }

    It "Return $true if connection is established" {
        Test-CMConnection | Should Be $true
    }
}

Describe "Get-CMObject" {
    Context "No connection" {
        Mock Get-WmiObject {throw}

        It "Throw if no connection has been established" {
            {Get-CMObject -Class "TestMe"} | Should Throw
        }
    }

    Context "Connection established" {

        $script:CMProviderServer = "TestProviderServer"
        $script:CMSiteCode = "XXX"
        $script:CMNamespace = "root\sms\Site_XXX"

        Mock Get-WmiObject { 
            [PSCustomObject]@{Class = $Class; Query = $Query; Namespace = $Namespace; ComputerName = $ComputerName}
        }

        It "Throw if no class is supplied" {
            {Get-CMObject -Class ""} | Should Throw
        }
        
        It "Use Provider connection" {
            Get-CMObject -Class "TestClass" | select -ExpandProperty "Namespace" | Should Be "root\sms\site_XXX"
            Get-CMObject -Class "TestClass" | select -ExpandProperty "ComputerName" | Should Be "TestProviderServer"
        } 

        It "Use specified values for WMI Query" {
            Get-CMObject -Class "TestClass" | select -ExpandProperty "Class" | Should Be "TestClass"
            Get-CMObject -Class "TestClass" | select -ExpandProperty "Query" | Should Be ""
            Get-CMObject -Class "TestClass" -Filter "Name = 'TestFilter'" | select -ExpandProperty "Query" | Should Be "SELECT * FROM TestClass WHERE Name = 'TestFilter'"
        }

    }

}

Describe "New-CMObject" {
    Context "No connection" {
        Mock Test-CMConnection {throw}

        It "Throw if no connection has been established" {
            {New-CMObject -Class "TestClass" -Arguments @{TestArg=1} } | Should Throw
        }
    }

    Context "Connection established" {

        $script:CMProviderServer = "TestProviderServer"
        $script:CMSiteCode = "XXX"
        $script:CMNamespace = "root\sms\Site_XXX"

        Mock Set-WmiInstance { [PSCustomObject]@{Class = $Class; Arguments = $Arguments; Namespace = $Namespace; ComputerName = $ComputerName} }

        It "Throw if class or arguments are missing" {
            {New-CMObject -Class "" -Arguments @{TestArg=1}} | Should Throw
            {New-CMObject -Class "TestClass" -Arguments $null } | Should Throw
        }
        
        It "Use Provider connection" {
            $TestObject = New-CMObject -Class "TestClass" -Arguments @{TestArg=1} 
            
            $TestObject | select -ExpandProperty "Namespace" | Should Be "root\sms\site_XXX"
            $TestObject | select -ExpandProperty "ComputerName" | Should Be "TestProviderServer"
        } 

        It "Use specified values" {
            $TestObject = New-CMObject -Class "TestClass" -Arguments @{TestArg=1} 

            $TestObject | select -ExpandProperty "Class" | Should Be "TestClass"
            $TestObject | select -ExpandProperty "Arguments" | %{$_.TestArg} |  Should Be 1
        }

    }

}

Describe "New-CMInstance" {

    Context "No connection" {
        Mock Test-CMConnection {throw}

        It "Throw if no connection has been established" {
            {New-CMObject -Class "TestClass" -Arguments @{TestArg=1} } | Should Throw
        }
    }

    Context "Connection established" {

        $script:CMProviderServer = "TestProviderServer"
        $script:CMSiteCode = "XXX"
        $script:CMNamespace = "root\sms\Site_XXX"

        Mock Get-WmiObject { 
            $Result = [PSCustomObject]@{List = $List; Class = $Class; Namespace = $Namespace; ComputerName = $ComputerName} 
            $Result | Add-Member -MemberType ScriptMethod -Name CreateInstance -Value {
                    [PSCustomObject]@{List = $this.List; Class = $this.Class; Namespace = $this.Namespace; ComputerName = $this.ComputerName; CreateInstanceCalled=$True}
                }  
            $Result  
        }

        It "Throw if class is missing" {
            {New-CMInstance -Class ""} | Should Throw
        }
        
        It "Use Provider connection" {
            $TestObject = New-CMInstance -Class "TestClass" 
            
            $TestObject | select -ExpandProperty "Namespace" | Should Be "root\sms\site_XXX"
            $TestObject | select -ExpandProperty "ComputerName" | Should Be "TestProviderServer"
        } 

        It "Use specified values" {
            $TestObject = New-CMInstance -Class "TestClass" 

            $TestObject | select -ExpandProperty "Class" | Should Be "TestClass"
            $TestObject | select -ExpandProperty "List" | Should Be $True
            $TestObject | select -ExpandProperty "CreateInstanceCalled" | Should Be $true
            
        }

    }

}

Describe "Invoke-CMMethod" {
    Context "No connection" {
        Mock Invoke-WmiMethod {throw}

        It "Throw if no connection has been established" {
            {Invoke-WMIMethod -Class "TestClass" -Name "TestMethod"} | Should Throw
        }
    }

    Context "Connection established" {

        Mock Invoke-WmiMethod { 
            [PSCustomObject]@{ReturnValue = 0; Class = $Class; Name = $Name; ArgumentList = $ArgumentList; Namespace = $Namespace; ComputerName = $ComputerName}
        }

        $script:CMProviderServer = "TestProviderServer"
        $script:CMSiteCode = "XXX"
        $script:CMNamespace = "root\sms\Site_XXX"

        It "Throw if no class or method is supplied" {
            {Invoke-CMMethod -Class "" -Name "TestMethod"} | Should Throw
            {Invoke-CmMethod -Class "TestClass" -Name ""} | Should Throw
        } 

        It "Use Provider connection" {
            $TestMethod = Invoke-CMMethod -Class "TestClass" -Name "TestMethod" 
            
            $TestMethod | select -ExpandProperty "Namespace" | Should Be "root\sms\site_XXX"
            $TestMethod | select -ExpandProperty "ComputerName" | Should Be "TestProviderServer"
        }

        It "Use specified values for WMI Method invocation" {
            $TestMethod = Invoke-CMMethod -Class "TestClass" -Name "TestMethod" 
            
            $TestMethod | select -ExpandProperty "Class" | Should Be "TestClass"
            $TestMethod | select -ExpandProperty "Name" | Should Be "TestMethod"
            $TestMethod | select -ExpandProperty "ArgumentList" | Should Be $null

            Invoke-CMMethod -Class "TestClass" -Name "TestMethod" -ArgumentList @("TestArgument") | select -ExpandProperty "ArgumentList" | Should Be @("TestArgument")
        }

    }

}

Describe "Get-DriverPackage" {
    
    Mock Get-CMObject { 
        [PSCustomObject]@{Class = $Class; Filter = $Filter}
    }
    
    It "Throw if no name is supplied" {
        {Get-DriverPackage -Name ""} | Should Throw
    } 

    It "Return Driver Package by Name" {
        $TestPackage = Get-DriverPackage -Name "TestPackage" 
        
        $TestPackage | select -ExpandProperty "Class" | Should Be "SMS_DriverPackage"
        $TestPackage | select -ExpandProperty "Filter" | Should Not Be ""
        $TestPackage | select -ExpandProperty "Filter" | Should Be "Name = 'TestPackage'"
    }
}

Describe "New-DriverPackage" {
    
    Mock Set-WmiInstance { 
        [PSCustomObject]@{Class = $Class; Arguments = $Arguments}
    }
    
    It "Throw if no name is supplied" {
        {New-DriverPackage -Name ""} | Should Throw
    } 

    It "Use supplied values" {
        $TestPackage = New-DriverPackage -Name "TestPackage" -SourcePath "\\CMServer01\Packages$\TestPackage" -Description "Just a Test package"  

        $TestPackage | select -ExpandProperty "Class" | Should Be "SMS_DriverPackage"
        $TestPackage | select -ExpandProperty "Arguments" | %{$_.Name} | Should Be "TestPackage"
        $TestPackage | select -ExpandProperty "Arguments" | %{$_.PkgSourcePath} | Should Be "\\CMServer01\Packages$\TestPackage"
        $TestPackage | select -ExpandProperty "Arguments" | %{$_.Description} | Should Be "Just a Test package"
        $TestPackage | select -ExpandProperty "Arguments" | %{$_.PkgSourceFlag} | Should Be 2

    }
}

Describe "Get-Driver" {
    
    Mock Get-CMObject { 
        [PSCustomObject]@{Class = $Class; Filter = $Filter}
    }
    
    It "Throw if no name is supplied" {
        {Get-Driver -Name ""} | Should Throw
    } 

    It "Return Drivers by Driver Package" {
        $TestDriver = Get-Driver -DriverPackage @{Name = "TestPackage"; PackageID = "TST00001"}
        
        $TestDriver | select -ExpandProperty "Class" | Should Be "SMS_Driver"
        $TestDriver | select -ExpandProperty "Filter" | Should Not Be ""
        $TestDriver | select -ExpandProperty "Filter" | Should Be "CI_ID IN (SELECT CTC.CI_ID FROM SMS_CIToContent AS CTC JOIN SMS_PackageToContent AS PTC ON CTC.ContentID=PTC.ContentID JOIN SMS_DriverPackage AS Pkg ON PTC.PackageID=Pkg.PackageID WHERE Pkg.PackageID='TST00001')"
    }
}

Describe "Get-Category" {
    
    Mock Get-CMObject { 
        [PSCustomObject]@{Class = $Class; Filter = $Filter}
    }
    
    It "Throw if no name is supplied" {
        {Get-Category -Name ""} | Should Throw
    } 

    It "Return Category by Name" {
        $TestCategory = Get-Category -Name "TestCategory"
        
        $TestCategory | select -ExpandProperty "Class" | Should Be "SMS_CategoryInstance"
        $TestCategory | select -ExpandProperty "Filter" | Should Not Be ""
        $TestCategory | select -ExpandProperty "Filter" | Should Be "LocalizedCategoryInstanceName = 'TestCategory'"
    }
}

Describe "New-Category" {
    
    Mock New-CMInstance {
        [PSCustomObject]@{CategoryInstanceName = ""; LocaleID = 0}
    }

    Mock Set-WmiInstance { 
        [PSCustomObject]@{Class = $Class; Arguments = $Arguments}
    }
    
    It "Throw if no name is supplied or type is not correct" {
        {New-Category -Name "" -TypeName "Locale"} | Should Throw
        {New-Category -Name "TestCategory" -TypeName "FailType"} | Should Throw
    } 

    It "Use supplied values" {
        $TestCategory = New-Category -Name "TestCategory" -TypeName "Locale"

        $TestCategory | select -ExpandProperty "Class" | Should Be "SMS_CategoryInstance"
        $TestCategory | select -ExpandProperty "Arguments" | %{$_.LocalizedInformation} | %{$_.CategoryInstanceName} |  Should Be "TestCategory"
        $TestCategory | select -ExpandProperty "Arguments" | %{$_.LocalizedInformation} | %{$_.LocaleID} |  Should Be "1033"
        $TestCategory | select -ExpandProperty "Arguments" | %{$_.CategoryInstance_UniqueID} | Should match "Locale"
        $TestCategory | select -ExpandProperty "Arguments" | %{$_.SourceSite} | Should Be "XXX"

        $TestCategory = New-Category -Name "TestCategory" -TypeName "DriverCategories" -LocaleID 42
        $TestCategory | select -ExpandProperty "Arguments" | %{$_.LocalizedInformation} | %{$_.LocaleID} |  Should Be "42"
    }

}

Describe "New-DriverCategory" {
    
    Mock New-CMInstance {
        [PSCustomObject]@{CategoryInstanceName = ""; LocaleID = 0}
    }

    Mock Set-WmiInstance { 
        [PSCustomObject]@{Class = $Class; Arguments = $Arguments}
    }
    
    It "Throw if no name is supplied or type is not correct" {
        {New-Category -Name "" -TypeName "DriverCategories"} | Should Throw
        {New-Category -Name "TestCategory" -TypeName "FailType"} | Should Throw
    } 

    It "Use supplied values" {
        $TestCategory = New-DriverCategory -Name "TestCategory"

        $TestCategory | select -ExpandProperty "Class" | Should Be "SMS_CategoryInstance"
        $TestCategory | select -ExpandProperty "Arguments" | %{$_.LocalizedInformation} | %{$_.CategoryInstanceName} |  Should Be "TestCategory"
        $TestCategory | select -ExpandProperty "Arguments" | %{$_.LocalizedInformation} | %{$_.LocaleID} |  Should Be "1033"
        $TestCategory | select -ExpandProperty "Arguments" | %{$_.CategoryInstance_UniqueID} | Should match "DriverCategories"
        $TestCategory | select -ExpandProperty "Arguments" | %{$_.SourceSite} | Should Be "XXX"

        $TestCategory = New-Category -Name "TestCategory" -TypeName "DriverCategories" -LocaleID 42
        $TestCategory | select -ExpandProperty "Arguments" | %{$_.LocalizedInformation} | %{$_.LocaleID} |  Should Be "42"
    }

}

Describe "Get-Folder" {
    
    Mock Get-CMObject { 
        [PSCustomObject]@{Class = $Class; Filter = $Filter}
    }
    
    It "Throw if no name is supplied" {
        {Get-Folder -Name ""} | Should Throw
    } 

    It "Return Folder by Name" {
        $TestFolder = Get-Folder -Name "TestFolder"
        
        $TestFolder | select -ExpandProperty "Class" | Should Be "SMS_ObjectContainerNode"
        $TestFolder | select -ExpandProperty "Filter" | Should Not Be ""
        $TestFolder | select -ExpandProperty "Filter" | Should Be "Name = 'TestFolder'"
    }
}

Describe "New-Folder" {
    
    Mock Set-WmiInstance { 
        [PSCustomObject]@{Class = $Class; Arguments = $Arguments}
    }
    
    It "Throw if no name is supplied or type is not correct" {
        {New-Folder -Name "" -Type "DriverPackage"} | Should Throw
        {New-Folder -Name "TestFolder" -Type "FailType"} | Should Throw
    } 

    It "Use supplied values" {
        $TestFolder = New-Folder -Name "TestFolder" -Type "DriverPackage"  

        $TestFolder | select -ExpandProperty "Class" | Should Be "SMS_ObjectContainerNode"
        $TestFolder | select -ExpandProperty "Arguments" | %{$_.Name} | Should Be "TestFolder"
        $TestFolder | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 23
        $TestFolder | select -ExpandProperty "Arguments" | %{$_.ParentContainerNodeid} | Should Be 0
        New-Folder -Name "TestFolder" -Type "DriverPackage" -ParentFolderID 1| select -ExpandProperty "Arguments" | %{$_.ParentContainerNodeid} | Should Be 1
    }

    It "Use correct object type" {

        New-Folder -Name "TestFolder" -Type "Package" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 2
        New-Folder -Name "TestFolder" -Type "Advertisement" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 3
        New-Folder -Name "TestFolder" -Type "Query" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 7
        New-Folder -Name "TestFolder" -Type "Report" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 8
        New-Folder -Name "TestFolder" -Type "MeteredProductRule" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 9
        New-Folder -Name "TestFolder" -Type "ConfigurationItem" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 11
        New-Folder -Name "TestFolder" -Type "OperatingSystemInstallPackage" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 14
        New-Folder -Name "TestFolder" -Type "StateMigration" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 17
        New-Folder -Name "TestFolder" -Type "ImagePackage" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 18
        New-Folder -Name "TestFolder" -Type "BootImagePackage" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 19
        New-Folder -Name "TestFolder" -Type "TaskSequencePackage" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 20
        New-Folder -Name "TestFolder" -Type "DeviceSettingPackage" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 21
        New-Folder -Name "TestFolder" -Type "DriverPackage" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 23
        New-Folder -Name "TestFolder" -Type "Driver" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 25
        New-Folder -Name "TestFolder" -Type "SoftwareUpdate" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 1011
        New-Folder -Name "TestFolder" -Type "BaselineConfigurationItem" | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 2011
    }
}
