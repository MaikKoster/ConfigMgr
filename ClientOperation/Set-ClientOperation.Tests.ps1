###############################################################################
# 
# PESTER tests for Set-ClientOperation.ps1
# 
# Author:  Maik Koster
# Date:    07.12.2015 
# History:
#          07.12.2015 - Initial setup of tests
#
###############################################################################

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here\Set-ClientOperation.ps1" -Cancel -All

Describe "Script with parameter" {

    Mock New-CMConnection { 
        [PSCustomObject]@{ProviderServer = $ProviderServer; SiteCode = $SiteCode; Credentials = $Credential}
    }

    Context "Start" {

        Mock New-CMConnection {}
        Mock Start-ClientOperation { [PSCustomObject]@{Operation = $Operation; CollectionID = $CollectionID; ResourceID = $ResourceID} }
        Mock Get-DeviceCollection {}
        Mock Get-DeviceCollection { [PSCustomObject]@{CollectionID = "TST00001"} } -ParameterFilter { $CollectionName -eq "TestCollection"}
        Mock Get-Computer {}
        Mock Get-Computer { [PSCustomObject]@{ResourceID = "987654"} } -ParameterFilter { $Name -eq "TestComputer"}
        Mock Get-Computer { [PSCustomObject]@{ResourceID = "876543"} } -ParameterFilter { $Name -eq "TestComputer2"}

       
        It "Throw if no CollectionID/CollectionName and ResourceID/ResourceName is supplied" {
            {Main -Start FullScan} | Should Throw
        }

        It "Only CollectionID specified" {
            Main -Start FullScan -CollectionID "TST00001" | select -ExpandProperty CollectionID | Should Be "TST00001"
            Main -Start FullScan -CollectionID "TST00001" | select -ExpandProperty ResourceID | Should Be $null
        }

        It "Only CollectionName specified" {
            Main -Start FullScan -CollectionName "TestCollection" | select -ExpandProperty CollectionID | Should Be "TST00001"
            Main -Start FullScan -CollectionName "TestCollection" | select -ExpandProperty ResourceID | Should Be $null
        }

        It "Throw if CollectionName cannot be resolved" {
            {Main -start FullScan -CollectionName "WrongTestCollection"} | Should Throw
        }

        It "No Collection specified but ResourceID" {
            Main -Start FullScan -ResourceID 123456 | select -ExpandProperty CollectionID | Should Be "SMS00001"
            Main -Start FullScan -ResourceID 123456 | select -ExpandProperty ResourceID | Should Be 123456
        }

        It "No Collection specified but ResourceName" {
            Main -Start FullScan -ResourceName "TestComputer" | select -ExpandProperty CollectionID | Should Be "SMS00001"
            Main -Start FullScan -ResourceName "TestComputer" | select -ExpandProperty ResourceID | Should Be 987654
        }

        It "Collection ID and ResourceID" {
            Main -Start FullScan -CollectionID "TST00001" -ResourceID 123456 | select -ExpandProperty CollectionID | Should Be "TST00001"
            Main -Start FullScan -CollectionID "TST00001" -ResourceID 123456 | select -ExpandProperty ResourceID | Should Be 123456
        }

        It "Collection ID and multiple ResourceID" {
            Main -Start FullScan -CollectionID "TST00001" -ResourceID @(123456,234567) | select -ExpandProperty CollectionID | Should Be "TST00001"
            $ResourceID = Main -Start FullScan -CollectionID "TST00001" -ResourceID @(123456,234567) | select -ExpandProperty ResourceID
            $ResourceID | select -First 1 | Should Be 123456
            $ResourceID | select -Skip 1 -First 1 | Should Be 234567
        }

        It "Collection ID and ResourceName" {
            Main -Start FullScan -CollectionID "TST00001" -ResourceName "TestComputer" | select -ExpandProperty CollectionID | Should Be "TST00001"
            Main -Start FullScan -CollectionID "TST00001" -ResourceName "TestComputer" | select -ExpandProperty ResourceID | Should Be 987654
        }

        It "Collection ID and multiple ResourceNames" {
            Main -Start FullScan -CollectionID "TST00001" -ResourceName @("TestComputer", "TestComputer2") | select -ExpandProperty CollectionID | Should Be "TST00001"
            $ResourceID = Main -Start FullScan -CollectionID "TST00001" -ResourceName "TestComputer", "TestComputer2"| select -ExpandProperty ResourceID
            $ResourceID | select -First 1 | Should Be 987654
            $ResourceID | select -Skip 1 -First 1 | Should Be 876543
        }

        It "Valid operation" {
            Main -Start FullScan -CollectionID "TST00001" | select -ExpandProperty Operation | Should Be "FullScan"
            Main -Start QuickScan -CollectionID "TST00001" | select -ExpandProperty Operation | Should Be "QuickScan"
            Main -Start DownloadDefinition -CollectionID "TST00001" | select -ExpandProperty Operation | Should Be "DownloadDefinition"
            Main -Start EvaluateSoftwareUpdates -CollectionID "TST00001" | select -ExpandProperty Operation | Should Be "EvaluateSoftwareUpdates"
            Main -Start RequestComputerPolicy -CollectionID "TST00001" | select -ExpandProperty Operation | Should Be "RequestComputerPolicy"
            Main -Start RequestUserPolicy -CollectionID "TST00001" | select -ExpandProperty Operation | Should Be "RequestUserPolicy"
        }

        it "Throw on invalid operation" {
            {Main -Start Fail} | Should Throw
        }

    }

    Context "Cancel" {

        Mock New-CMConnection {}
        Mock Cancel-ClientOperation { [PSCustomObject]@{ReturnValue = $OperationID} }
        Mock Get-ClientOperation { [PSCustomObject]@{ID = 42} }

        It "Warning if no OperationID available" {
            (Main -Cancel 3>&1) | Should Be "No Client Operation available."
        }        

        It "Use supplied OperationID" {
            Main -Cancel -OperationID 42 | select -ExpandProperty ReturnValue | Should Be 42
        }

        It "Get All Client Operations" {
            Main -Cancel -All | select -ExpandProperty ReturnValue | Should Be 42
        }
    }

    Context "Delete" {

        Mock New-CMConnection {}
        Mock Delete-ClientOperation { [PSCustomObject]@{ReturnValue = $OperationID} }
        Mock Get-ClientOperation { [PSCustomObject]@{ID = 42} }
        Mock Get-ClientOperation { [PSCustomObject]@{ID = 43} } -ParameterFilter {$Expired.IsPresent}

        It "Warning if no OperationID available" {
            (Main -Delete 3>&1) | Should Be "No Client Operation available."
        }        

        It "Use supplied OperationID" {
            Main -Delete -OperationID 42 | select -ExpandProperty ReturnValue | Should Be 42
        }

        It "Get All Client Operations" {
            Main -Delete -All | select -ExpandProperty ReturnValue | Should Be 42
        }

        It "Get Expired Client Operations" {
            Main -Delete -Expired | select -ExpandProperty ReturnValue | Should Be 43
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

Describe "Get-DeviceCollection" {
    
    Mock Get-CMObject { 
        [PSCustomObject]@{Class = $Class; Filter = $Filter}
    }
    
    It "Throw if no name is supplied" {
        {Get-DeviceCollection -Name ""} | Should Throw
    } 

    It "Return Device Collection by Name" {
        Get-DeviceCollection -Name "TestCollection" | select -ExpandProperty "Class" | Should Be "SMS_Collection"
        Get-DeviceCollection -Name "TestCollection" | select -ExpandProperty "Filter" | Should Not Be ""
        Get-DeviceCollection -Name "TestCollection" | select -ExpandProperty "Filter" | Should Be "((Name = 'TestCollection') AND (CollectionType = 2))"
    }
}

Describe "Get-Computer" {
    
    Mock Get-CMObject { [PSCustomObject]@{Class = $Class; Filter = $Filter} }
    
    It "Throw if no name is supplied" {
        {Get-Computer -Name ""} | Should Throw
    } 

    It "Return computer by name" {
        Get-Computer -Name "TestComputer" | select -ExpandProperty "Class" | Should Be "SMS_R_System"
        Get-Computer -Name "TestComputer" | select -ExpandProperty "Filter" | Should Not Be ""
        Get-Computer -Name "TestComputer" | select -ExpandProperty "Filter" | Should Be "((Name = 'TestComputer') OR (NetbiosName = 'TestComputer'))"
    }

    It "Return list of computers by multiple names" {
        $Computers = Get-Computer -Name "TestComputer1", "TestComputer2" | select -ExpandProperty "Filter" 

        ($Computers).Count | Should be 2
        $Computers |  select -First 1 | Should Be "((Name = 'TestComputer1') OR (NetbiosName = 'TestComputer1'))"
        $Computers |  select -Skip 1 -First 1 | Should Be "((Name = 'TestComputer2') OR (NetbiosName = 'TestComputer2'))"
    }

    It "Return Computer by name using the pipepline" {
        "TestComputer" | Get-Computer | select -ExpandProperty "Class" | Should Be "SMS_R_System"
        "TestComputer" | Get-Computer | select -ExpandProperty "Filter" | Should Not Be ""
        "TestComputer" | Get-Computer | select -ExpandProperty "Filter" | Should Be "((Name = 'TestComputer') OR (NetbiosName = 'TestComputer'))"
    }

    It "Return list of computers by multiple names using the pipeline" {
        $Computers = @("TestComputer1", "TestComputer2") | Get-Computer | select -ExpandProperty "Filter" 

        ($Computers).Count | Should be 2
        $Computers |  select -First 1 | Should Be "((Name = 'TestComputer1') OR (NetbiosName = 'TestComputer1'))"
        $Computers |  select -Skip 1 -First 1 | Should Be "((Name = 'TestComputer2') OR (NetbiosName = 'TestComputer2'))"
    }
}

Describe "Get-ClientOperation" {
    
    Mock Get-CMObject { [PSCustomObject]@{Class = $Class; Filter = $Filter} }
    
    It "Return Client Operations" {
        Get-ClientOperation | select -ExpandProperty "Class" | Should Be "SMS_ClientOperation"
        Get-ClientOperation | select -ExpandProperty "Filter" | Should Be ""
        Get-ClientOperation -Expired | select -ExpandProperty "Filter" | Should Be "((State = 0) OR (State = 2))"
    }
}

Describe "Start-ClientOperation" {
    
    Mock Invoke-CMMethod { [PSCustomObject]@{ReturnValue = 0; Class = $Class; Name = $Name; ArgumentList = $ArgumentList} }

    It "Throw if Operation is invalid or no CollectionID supplied" {
        
        {Start-ClientOperation -Operation "" -CollectionID "XXX00001" } | Should Throw
        {Start-ClientOperation -Operation FullScan -CollectionID "" } | Should Throw

    }
      
    It "Invoke correct WMI Method" {
        Start-ClientOperation -Operation FullScan -CollectionID "XXX00001" | select -ExpandProperty "Class" | Should Be "SMS_ClientOperation"
        Start-ClientOperation -Operation FullScan -CollectionID "XXX00001" | select -ExpandProperty "Name" | Should Be "InitiateClientOperation"
    }

    It "Use supplied CollectionID" {
        Start-ClientOperation -Operation FullScan -CollectionID "XXX00001" | select -ExpandProperty "ArgumentList" | select -first 1 |  Should Be "XXX00001"
    }

    It "Use supplied ResourceID" {
        Start-ClientOperation -Operation FullScan -CollectionID "XXX00001" -ResourceID 42| select -ExpandProperty "ArgumentList" | select -Skip 1 -First 1 | Should Be "42"

        $ResourceIDs = Start-ClientOperation -Operation FullScan -CollectionID "XXX00001" -ResourceID 42,43,44| select -ExpandProperty "ArgumentList" | select -skip 1 -First 1 
        ($ResourceIDs).Count | Should Be 3
        $ResourceIDs |  select -First 1 | Should Be "42"
        $ResourceIDs |  select -Skip 1 -First 1 | Should Be "43"
        $ResourceIDs |  select -Skip 2 -First 1 | Should Be "44"
    }

    It "Use correct Operation" {
        Start-ClientOperation -Operation FullScan -CollectionID "XXX00001" | select -ExpandProperty "ArgumentList" | select -skip 1 -First 1 | Should Be "1"
        Start-ClientOperation -Operation QuickScan -CollectionID "XXX00001" | select -ExpandProperty "ArgumentList" | select -skip 1 -First 1 | Should Be "2"
        Start-ClientOperation -Operation DownloadDefinition -CollectionID "XXX00001" | select -ExpandProperty "ArgumentList" | select -skip 1 -First 1 | Should Be "3"
        Start-ClientOperation -Operation EvaluateSoftwareUpdates -CollectionID "XXX00001" | select -ExpandProperty "ArgumentList" | select -skip 1 -First 1 | Should Be "4"
        Start-ClientOperation -Operation RequestComputerPolicy -CollectionID "XXX00001" | select -ExpandProperty "ArgumentList" | select -skip 1 -First 1 | Should Be "8"
        Start-ClientOperation -Operation RequestUserPolicy -CollectionID "XXX00001" | select -ExpandProperty "ArgumentList" | select -skip 1 -First 1 | Should Be "9"
    }

}

Describe "Cancel-ClientOperation" {
    
    Mock Invoke-CMMethod { [PSCustomObject]@{ReturnValue = 0; Class = $Class; Name = $Name; ArgumentList = $ArgumentList} }
      
    It "Invoke correct WMI Method" {
        Cancel-ClientOperation -OperationID 42 | select -ExpandProperty "Class" | Should Be "SMS_ClientOperation"
        Cancel-ClientOperation -OperationID 42| select -ExpandProperty "Name" | Should Be "CancelClientOperation"
    }

    It "Use correct OperationID" {
        Cancel-ClientOperation -OperationID 42| select -ExpandProperty "ArgumentList" | Should Be 42
        Cancel-ClientOperation -OperationID 42,43,44 | select -ExpandProperty "ArgumentList" -First 1 | Should Be 42
        Cancel-ClientOperation -OperationID 42,43,44 | select -ExpandProperty "ArgumentList" -Last 1 | Should Be 44
    }

    It "Invoke correct WMI Method by pipeline" {
        42 | Cancel-ClientOperation | select -ExpandProperty "Class" | Should Be "SMS_ClientOperation"
        42 | Cancel-ClientOperation | select -ExpandProperty "Name" | Should Be "CancelClientOperation"
    }

    It "Use correct OperationID by pipeline" {
        42 | Cancel-ClientOperation | select -ExpandProperty "ArgumentList" | Should Be @(42)
        @(42,43,44) | Cancel-ClientOperation | select -ExpandProperty "ArgumentList" -First 1 | Should Be 42
        @(42,43,44) | Cancel-ClientOperation | select -ExpandProperty "ArgumentList" -Last 1 | Should Be 44
    }
}

Describe "Delete-ClientOperation" {
    
    Mock Invoke-CMMethod { [PSCustomObject]@{ReturnValue = 0; Class = $Class; Name = $Name; ArgumentList = $ArgumentList} }
      
    It "Invoke correct WMI Method" {
        Delete-ClientOperation -OperationID 42 | select -ExpandProperty "Class" | Should Be "SMS_ClientOperation"
        Delete-ClientOperation -OperationID 42| select -ExpandProperty "Name" | Should Be "DeleteClientOperation"
    }

    It "Use correct OperationID" {
        Delete-ClientOperation -OperationID 42| select -ExpandProperty "ArgumentList" | Should Be 42
        Delete-ClientOperation -OperationID 42,43,44 | select -ExpandProperty "ArgumentList" -First 1 | Should Be 42
        Delete-ClientOperation -OperationID 42,43,44 | select -ExpandProperty "ArgumentList" -Last 1 | Should Be 44
    }

    It "Invoke correct WMI Method by pipeline" {
        42 | Delete-ClientOperation | select -ExpandProperty "Class" | Should Be "SMS_ClientOperation"
        42 | Delete-ClientOperation | select -ExpandProperty "Name" | Should Be "DeleteClientOperation"
    }

    It "Use correct OperationID by pipeline" {
        42 | Delete-ClientOperation | select -ExpandProperty "ArgumentList" | Should Be 42
        @(42,43,44) | Delete-ClientOperation | select -ExpandProperty "ArgumentList" -First 1 | Should Be 42
        @(42,43,44) | Delete-ClientOperation | select -ExpandProperty "ArgumentList" -Last 1 | Should Be 44
    }
}