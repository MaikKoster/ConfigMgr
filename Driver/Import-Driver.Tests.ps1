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
            [PSCustomObject]@{Class = $Class; Query = $Query}
        }

        It "Throw if no class is supplied" {
            {Get-CMObject -Class ""} | Should Throw
        } 

        It "Use specified values for WMI Query" {
            Get-CMObject -Class "TestClass" | select -ExpandProperty "Class" | Should Be "TestClass"
            Get-CMObject -Class "TestClass" | select -ExpandProperty "Query" | Should Be ""
            Get-CMObject -Class "TestClass" -Filter "Name = 'TestFilter'" | select -ExpandProperty "Query" | Should Be "SELECT * FROM TestClass WHERE Name = 'TestFilter'"
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
            [PSCustomObject]@{ReturnValue = 0; Class = $Class; Name = $Name; ArgumentList = $ArgumentList}
        }

        $script:CMProviderServer = "TestProviderServer"
        $script:CMSiteCode = "XXX"
        $script:CMNamespace = "root\sms\Site_XXX"

        It "Throw if no class or method is supplied" {
            {Invoke-CMMethod -Class "" -Name "TestMethod"} | Should Throw
            {Invoke-CmMethod -Class "TestClass" -Name ""} | Should Throw
        } 

        It "Use specified values for WMI Method invocation" {
            Invoke-CMMethod -Class "TestClass" -Name "TestMethod" | select -ExpandProperty "Class" | Should Be "TestClass"
            Invoke-CMMethod -Class "TestClass" -Name "TestMethod" | select -ExpandProperty "Name" | Should Be "TestMethod"
            Invoke-CMMethod -Class "TestClass" -Name "TestMethod" | select -ExpandProperty "ArgumentList" | Should Be $null
            Invoke-CMMethod -Class "TestClass" -Name "TestMethod" -ArgumentList @("TestArgument") | select -ExpandProperty "ArgumentList" | Should Be @("TestArgument")
        }

    }

}
