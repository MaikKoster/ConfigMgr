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




