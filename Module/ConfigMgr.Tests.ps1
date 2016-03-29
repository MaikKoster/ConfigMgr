###############################################################################
# 
# PESTER tests for ConfigMgr.psm1
# 
# Author:  Maik Koster
# Date:    14.03.2016
# History:
#          14.03.2016 - Initial setup of tests
#
###############################################################################

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
#. "$here\ConfigMgr.psm1"

# Ensure Module is unloaded first before reloading

Get-Module ConfigMgr | Remove-Module

Import-Module ConfigMgr

InModuleScope ConfigMgr {

    #################################
    #region  ConfigMgr Connection   #
    #################################

    Describe "New-CMConnection" {
        # Ensure global module variables are cleared
        $global:CMProviderServer = ""
        $global:CMSiteCode = ""
        $global:CMNamespace = ""

        Context "No server available" {

            It "Throw an exception if server cannot be contacted" {
                Mock Get-CMSession {}
        
                { New-CMConnection } | Should Throw
            }
        }

        Context "Server available" {

            Mock Get-CimInstance { 
                [PSCustomObject]@{NamespacePath = "\\$ComputerName\root\sms\Site_TST"; Machine = "$($CimSession.ComputerName)"; SiteCode = "TST"} 
            } 

            Mock Get-CimInstance { 
                [PSCustomObject]@{NamespacePath = "\\TestLocalServer\root\sms\Site_TST"; Machine = "TestLocalServer"; SiteCode = "TST"} 
        
            } -ParameterFilter {$CimSession.ComputerName -eq "."}
        
            Mock Get-CimInstance {
                [PSCustomObject]@{NamespacePath = "\\$ComputerName\root\sms\Site_ZZZ"; Machine = "$($CimSession.ComputerName)"; SiteCode = "ZZZ"}

            } -ParameterFilter {$Filter -like "*SiteCode =*"}

            It "Use local computername if no server name was supplied." {
                New-CMConnection
            
                $global:CMProviderServer | Should Be "$Env:ComputerName"
            }

            It "Evaluate SiteCode if no SiteCode is supplied." {
                New-CMConnection
            
                $global:CMSiteCode | Should Be "TST"
            }

            It "Use supplied server name." {
                New-CMConnection -ProviderServerName "localhost"
            
                $global:CMProviderServer | Should Be "localhost"
            }

            It "Use supplied site code." {
                New-CMConnection -ProviderServerName "localhost" -SiteCode "ZZZ"
            
                $global:CMSiteCode | Should Be "ZZZ"
            }
        }
    }

    Describe "Test-CMConnection" {
        $global:CMProviderServer = ""
        $global:CMSiteCode = ""
        $global:CMNamespace = ""

        It "Throw if connection cannot be created" {
            Mock New-CMConnection { Throw }

            {Test-CMConnection}  | Should Throw
        }

        It "Connect to local Site Server" {
            Mock New-CMConnection {}
            Test-CMConnection
        }
    }

    Describe "Get-CMSession" {

        Context "DCOM" {
            Mock Test-WSMan {}

            It "Throw an exception if server cannot be contacted" {
                Get-CimSession | Remove-CimSession
        
                Mock New-CimSession { throw } -ParameterFilter {$ComputerName -eq "DoesNotExists"}
                {Get-CMSession "DoesNotExists" -ErrorAction Stop} | Should Throw
            }

            It "Create Dcom sesssion to local computer if WSMAN fails" {
                Get-CimSession | Remove-CimSession

                $Result = Get-CMSession 
                $Result | Should Not Be $null
                $Result.ComputerName | Should be $env:COMPUTERNAME
                $Result.Protocol | Should Be "Dcom"
            }
        } 

        Context "WSMAN" {
            It "Create WSMAN Session to local computer on default" {
                Get-CimSession | Remove-CimSession

                $Result = Get-CMSession 
                $Result | Should Not Be $null
                $Result.ComputerName | Should be $env:COMPUTERNAME
                $Result.Protocol | Should Be "WSMAN"
            }

            It "Create WSMAN Session to specified computer" {
                Get-CimSession | Remove-CimSession

                $Result = Get-CMSession localhost
                $Result | Should Not Be $null
                $Result.ComputerName | Should be "localhost"
                $Result.Protocol | Should Be "WSMAN"
            }

            It "Return existing session" {
                Mock New-CimSession {}
        
                $Result = Get-CMSession localhost
                $Result | Should Not Be $null
                $Result.ComputerName | Should be "localhost"
                $Result.Protocol | Should Be "WSMAN"
            }
        }
    }
    #################################
    #endregion ConfigMgr Connection #
    #################################

    ########################
    #region  CIM Methods   #
    ########################
    Describe "Get-CMInstance" {
        Context "No connection" {
            Mock Get-CimInstance {throw}

            It "Throw if no connection has been established" {
                {Get-CMInstance -Class "TestMe"} | Should Throw
            }
        }

        Context "Connection established" {
            Get-CimSession | Remove-CimSession
            $global:CMProviderServer = "TestProviderServer"
            $global:CMSiteCode = "TST"
            $global:CMNamespace = "root\sms\Site_TST"
            $global:CMSession = New-CimSession $env:COMPUTERNAME

            Mock Get-CimInstance { 
                [PSCustomObject]@{ClassName = $ClassName; Filter = $Filter; Namespace = $Namespace; ComputerName = $($CimSession.ComputerName)}
            }

            It "Throw if no class is supplied" {
                {Get-CMInstance -ClassName ""} | Should Throw
            }
        
            It "Use Provider connection" {
                Get-CMInstance -ClassName "TestClass" | select -ExpandProperty "Namespace" | Should Be "root\sms\site_TST"
                Get-CMInstance -ClassName "TestClass" | select -ExpandProperty "ComputerName" | Should Be "$Env:ComputerName"
            } 

            It "Use specified values for WMI Query" {
                Get-CMInstance -ClassName "TestClass" | select -ExpandProperty "ClassName" | Should Be "TestClass"
                Get-CMInstance -ClassName "TestClass" | select -ExpandProperty "Filter" | Should Be ""
                Get-CMInstance -ClassName "TestClass" -Filter "Name = 'TestFilter'" | select -ExpandProperty "Filter" | Should Be "Name = 'TestFilter'"
            }
        }
    }

    Describe "New-CMInstance" {
        Context "No connection" {
            Mock Test-CMConnection {throw}

            It "Throw if no connection has been established" {
                {New-CMObject -ClassName "TestClass" -Arguments @{TestArg=1} } | Should Throw
            }
        }

        Context "Connection established" {
            Get-CimSession | Remove-CimSession

            $global:CMProviderServer = "TestProviderServer"
            $global:CMSiteCode = "TST"
            $global:CMNamespace = "root\sms\Site_TST"
            $global:CMSession = New-CimSession $env:COMPUTERNAME

            Mock New-CimInstance { [PSCustomObject]@{Class = $ClassName; Arguments = $Property; Namespace = $Namespace; ComputerName = $($CimSession.ComputerName)} }

            It "Throw if class or arguments are missing" {
                {New-CMInstance -ClassName "" -Arguments @{TestArg=1}} | Should Throw
                {New-CMInstance -ClassName "TestClass" -Arguments $null } | Should Throw
            }
        
            It "Use Provider connection" {
                $TestObject = New-CMInstance -ClassName "TestClass" -Arguments @{TestArg=1} 
            
                $TestObject | select -ExpandProperty "Namespace" | Should Be "root\sms\site_TST"
                $TestObject | select -ExpandProperty "ComputerName" | Should Be "$Env:ComputerName"
            } 

            It "Use specified values" {
                $TestObject = New-CMInstance -ClassName "TestClass" -Arguments @{TestArg=1} 

                $TestObject | select -ExpandProperty "Class" | Should Be "TestClass"
                $TestObject | select -ExpandProperty "Arguments" | %{$_.TestArg} |  Should Be 1
            }

            It "Create embedded class" {
                Mock New-CMWMIInstance {@{Class = $ClassName; Embedded = $true}}

                $TestObject = New-CMInstance -ClassName "TestClass" -Arguments @{TestArg=1} -EnforceWMI

                $TestObject.Class | Should Be "TestClass"
                $TestObject.Embedded | Should Be $true
                $TestObject.TestArg | Should Be 1
            }
        }
    }

    Describe "Set-CMInstance" {
        Context "No connection" {
            Mock Test-CMConnection {throw}

            It "Throw if no connection has been established" {
                {Get-CMInstance -ClassName "TestMe"} | Should Throw
            }
        }

        Context "Connection established" {
            Get-CimSession | Remove-CimSession
            $global:CMProviderServer = "TestProviderServer"
            $global:CMSiteCode = "TST"
            $global:CMNamespace = "root\sms\Site_TST"
            $global:CMSession = New-CimSession $env:COMPUTERNAME

            Mock Set-CimInstance { 
                $InputObject
                #[PSCustomObject]@{ClassName = $($InputObject.ClassName); Filter = $($InputObject.Filter)}
            }

            Mock Get-CMInstance {
                Get-CimInstance "Win32_Process" -Filter $Filter | Select -First 1
            }  -ParameterFilter {$ClassName -eq "TestClass"}

            Mock Invoke-CimCommand { Get-CimInstance "Win32_ComputerSystem" }

            

            It "Throw if no classname and filter is supplied" {
                {Set-CMInstance -ClassName "" -filter "Name='Test'"} | Should Throw
                {Set-CMInstance -ClassName "TestClass" -filter ""} | Should Throw
            }

            It "Throw if no ClassInstance is supplied" {
                {Set-CMInstance -ClassInstance $null } | Should Throw
            }
        
            It "Resolve Class name and filter to object" {
                $Result = Set-CMInstance -ClassName "TestClass" -Filter "Name='explorer.exe'"

                $Result.CreationClassName | Should Be "Win32_Process"
                $Result.Name | Should Be "explorer.exe"
            } 

            #It "Use specified values for WMI Query" {
            #    Set-CMInstance -ClassName "TestClass" | select -ExpandProperty "ClassName" | Should Be "TestClass"
            #    Set-CMInstance -ClassName "TestClass" | select -ExpandProperty "Filter" | Should Be ""
            #    Set-CMInstance -ClassName "TestClass" -Filter "Name = 'TestFilter'" | select -ExpandProperty "Filter" | Should Be "Name = 'TestFilter'"
            #}
        }
    }

    Describe "New-CMWMIInstance" {

        Context "No connection" {
            Mock Test-CMConnection {throw}

            It "Throw if no connection has been established" {
                {New-WMIInstance -ClassName "TestClass" -Arguments @{TestArg=1} } | Should Throw
            }
        }

        Context "Connection established" {
            Get-CimSession | Remove-CimSession

            $global:CMProviderServer = "TestProviderServer"
            $global:CMSiteCode = "TST"
            $global:CMNamespace = "root\sms\Site_TST"
            $global:CMCredential = [System.Management.Automation.PSCredential]::Empty
            $global:CMSession = New-CimSession $env:COMPUTERNAME

            Mock Get-WmiObject { 
                $Result = [PSCustomObject]@{List = $List; ClassName = $ClassName; Namespace = $Namespace; ComputerName = $ComputerName} 
                $Result | Add-Member -MemberType ScriptMethod -Name CreateInstance -Value {
                        [PSCustomObject]@{List = $this.List; ClassName = $this.ClassName; Namespace = $this.Namespace; ComputerName = $this.ComputerName; CreateInstanceCalled=$True}
                    }  
                $Result  
            }

            It "Throw if class is missing" {
                {New-CMWMIInstance -ClassName ""} | Should Throw
            }
        
            It "Use Provider connection" {
                $TestObject = New-CMWMIInstance -ClassName "TestClass" 
            
                $TestObject | select -ExpandProperty "Namespace" | Should Be "root\sms\site_TST"
                $TestObject | select -ExpandProperty "ComputerName" | Should Be "TestProviderServer"
            } 

            It "Use specified values" {
                $TestObject = New-CMWMIInstance -ClassName "TestClass" 

                $TestObject | select -ExpandProperty "Classname" | Should Be "TestClass"
                $TestObject | select -ExpandProperty "List" | Should Be $True
                $TestObject | select -ExpandProperty "CreateInstanceCalled" | Should Be $true
            }
        }
    }


    Describe "Invoke-CMMethod" {
        Context "No connection" {
            Mock Test-CMConnection {throw}

            It "Throw if no connection has been established" {
                {Invoke-CMMethod -ClassName "TestClass" -Name "TestMethod"} | Should Throw
            }
        }

        Context "Connection established" {

            Mock Invoke-CimMethod { 
                [PSCustomObject]@{ReturnValue = 0; Class = $ClassName; Name = $MethodName; Arguments = $Arguments; Namespace = $Namespace; ComputerName = $($CimSession.ComputerName)}
            }

            Get-CimSession | Remove-CimSession
            $global:CMProviderServer = "TestProviderServer"
            $global:CMSiteCode = "TST"
            $global:CMNamespace = "root\sms\Site_TST"
            $global:CMSession = New-CimSession $env:COMPUTERNAME

            It "Throw if no class or method is supplied" {
                {Invoke-CMMethod -ClassName "" -MethodName "TestMethod"} | Should Throw
                {Invoke-CmMethod -ClassName "TestClass" -MethodName ""} | Should Throw
                {Invoke-CmMethod -ClassInstance $null -MethodName "TestMethod"} | Should Throw
            } 

            It "Use Provider connection" {
                $TestMethod = Invoke-CMMethod -ClassName "TestClass" -MethodName "TestMethod" 
            
                $TestMethod | select -ExpandProperty "Namespace" | Should Be "root\sms\site_TST"
                $TestMethod | select -ExpandProperty "ComputerName" | Should Be "$Env:ComputerName"
            }

            It "Use specified values for static WMI Method invocation" {
                $TestMethod = Invoke-CMMethod -ClassName "TestClass" -MethodName "TestMethod" -Arguments @{TestArgument="Test"}
            
                $TestMethod.Class | Should Be "TestClass"
                $TestMethod.Name  | Should Be "TestMethod"
                $TestMethod.Arguments.TestArgument | Should Be "Test"

                Invoke-CMMethod -ClassName "TestClass" -MethodName "TestMethod"  | select -ExpandProperty "Arguments" | Should Be $null
            }

            #It "Use specified values for static WMI Method invocation" {
            #    $TestClassInstance = [PSCustomObject]{Name="TestClass"}
            #    $TestClassInstance | Add-Member -MemberType ScriptMethod -Name "RunMe" -Value {[PSCustomObject]@{ReturnValue=0;}}
            #    $TestMethod = Invoke-CMMethod -ClassName "TestClass" -Name "TestMethod" 
            #    
            #    $TestMethod.Class | Should Be "TestClass"
            #    $TestMethod.Name  | Should Be "TestMethod"
            #    $TestMethod.Arguments | Should Be $null
    #
            #    Invoke-CMMethod -ClassName "TestClass" -Name "TestMethod" -Arguments @("TestArgument") | select -ExpandProperty "Arguments" | Should Be @("TestArgument")
            #}
        }
    }
    ########################
    #endregion CIM Methods #
    ########################


    #########################
    #region  SMS_R_System   #
    #########################
    Describe "Get-Device" {
    
        Mock Get-CMInstance { [PSCustomObject]@{ClassName = $ClassName; Filter = $Filter} }
    
        It "Throw if no name is supplied" {
            {Get-Device -Name ""} | Should Throw
        } 

        It "Return computer by name" {
            Get-Device -Name "TestComputer" | select -ExpandProperty "ClassName" | Should Be "SMS_R_System"
            Get-Device -Name "TestComputer" | select -ExpandProperty "Filter" | Should Not Be ""
            Get-Device -Name "TestComputer" | select -ExpandProperty "Filter" | Should Be "((Name = 'TestComputer') OR (NetbiosName = 'TestComputer'))"
        }

        It "Return list of computers by multiple names" {
            $Computers = Get-Device -Name "TestComputer1", "TestComputer2" | select -ExpandProperty "Filter" 

            $Computers | Should Be "(((Name = 'TestComputer1') OR (Name = 'TestComputer2')) OR ((NetbiosName = 'TestComputer1') OR (NetbiosName = 'TestComputer2')))"
        }

        It "Return Computer by name using the pipepline" {
            "TestComputer" | Get-Device | select -ExpandProperty "ClassName" | Should Be "SMS_R_System"
            "TestComputer" | Get-Device | select -ExpandProperty "Filter" | Should Not Be ""
            "TestComputer" | Get-Device | select -ExpandProperty "Filter" | Should Be "((Name = 'TestComputer') OR (NetbiosName = 'TestComputer'))"
        }

        It "Return list of computers by multiple names using the pipeline" {
            $Computers = @("TestComputer1", "TestComputer2") | Get-Device | select -ExpandProperty "Filter" 

            $Computers.Count | Should be 2
            $Computers[0] | Should Be "((Name = 'TestComputer1') OR (NetbiosName = 'TestComputer1'))"
            $Computers[1] | Should Be "((Name = 'TestComputer2') OR (NetbiosName = 'TestComputer2'))"
        }
    }
    #########################
    #endregion SMS_R_System #
    #########################


    ###########################
    #region  SMS_Collection   #
    ###########################
    Describe "Get-Collection" {
    
        Mock Get-CMInstance { 
            [PSCustomObject]@{ClassName = $ClassName; Filter = $Filter}
        }
    
        It "Throw if no name is supplied" {
            {Get-Collection -Name ""} | Should Throw
        } 

        It "Throw if wrong type is supplied" {
            {Get-Collection -Type CustomType -Name "TestCollection"} | Should Throw
        } 

        It "Return Collection by ID" {
            $Collection = Get-Collection -ID "TST00001" 
            
            $Collection | select -ExpandProperty "ClassName" | Should Be "SMS_Collection"
            $Collection.Filter.Contains("(CollectionID = 'TST00001'") | Should Be $true
        }

        It "Return Collection by Name" {
            $Collection = Get-Collection -Name "TestCollection" 
            
            $Collection | select -ExpandProperty "ClassName" | Should Be "SMS_Collection"
            $Collection.Filter.Contains("(Name = 'TestCollection'") | Should Be $true

            $Collection = Get-Collection -Type Any -Name "TestCollection" 
            
            $Collection | select -ExpandProperty "ClassName" | Should Be "SMS_Collection"
            $Collection.Filter.Contains("(Name = 'TestCollection'") | Should Be $true
        }

        It "Search Collection by Name" {
            $Collection = Get-Collection -Name "%Collection" -Search
            
            $Collection | select -ExpandProperty "ClassName" | Should Be "SMS_Collection"
            $Collection.Filter.Contains("(Name LIKE '%Collection'") | Should Be $true

            $Collection = Get-Collection -Type Any -Name "%Collection" -Search
            
            $Collection | select -ExpandProperty "ClassName" | Should Be "SMS_Collection"
            $Collection.Filter.Contains("(Name LIKE '%Collection'") | Should Be $true
        }

        It "Search Collection by custom filter " {
            $Collection = Get-Collection -Filter "(Name LIKE '%Collection')"
            
            $Collection | select -ExpandProperty "ClassName" | Should Be "SMS_Collection"
            $Collection.Filter.Contains("(Name LIKE '%Collection'") | Should Be $true
        }

        It "Return Device Collection by Name" {
            $Collection = Get-Collection -Type Device -Name "TestCollection" 
            
            $Collection | select -ExpandProperty "ClassName" | Should Be "SMS_Collection"
            $Collection.Filter.Contains("(CollectionType = 2)") | Should Be $true
            $Collection.Filter.Contains("(Name = 'TestCollection'") | Should Be $true
        }

        It "Return User Collection by Name" {
            $Collection = Get-Collection -Type User -Name "TestCollection" 
            
            $Collection | select -ExpandProperty "ClassName" | Should Be "SMS_Collection"
            $Collection.Filter.Contains("(CollectionType = 1)") | Should Be $true
            $Collection.Filter.Contains("(Name = 'TestCollection'") | Should Be $true
        }

        It "Return Other Collection by Name" {
            $Collection = Get-Collection -Type Other -Name "TestCollection" 
            
            $Collection | select -ExpandProperty "ClassName" | Should Be "SMS_Collection"
            $Collection.Filter.Contains("(CollectionType = 0)") | Should Be $true
            $Collection.Filter.Contains("(Name = 'TestCollection'") | Should Be $true
        }
    }
    ########################
    #region SMS_Collection #
    ########################


    #################################
    #region  SMS_CategoryInstance   #
    #################################
    Describe "Get-Category" {
    
        Mock Get-CMInstance { [PSCustomObject]@{ClassName = $ClassName; Filter = $Filter} }
    
        It "Throw if no name is supplied" {
            {Get-Category -Name ""} | Should Throw
        } 

        It "Return Category by ID" {
            $TestCategory = Get-Category -ID 42 
            $TestCategory | select -ExpandProperty "ClassName" | Should Be "SMS_CategoryInstance"
            $TestCategory | select -ExpandProperty "Filter" | Should Be "((CategoryInstanceID = 42))"
        }

        It "Return Category by name" {
            $TestCategory = Get-Category -Name "TestCategory" 
            $TestCategory | select -ExpandProperty "ClassName" | Should Be "SMS_CategoryInstance"
            $TestCategory | select -ExpandProperty "Filter" | Should Be "((LocalizedCategoryInstanceName = 'TestCategory'))"
        }

        It "Return list of Categories by multiple names" {
            $Categories = Get-Category -Name "TestCategory1", "TestCategory2" | select -ExpandProperty "Filter" 
            $Categories | Should Be "(((LocalizedCategoryInstanceName = 'TestCategory1') OR (LocalizedCategoryInstanceName = 'TestCategory2')))"
        }

        It "Return Category by name using the pipepline" {
            $Category = "TestCategory" | Get-Category 
            $Category | select -ExpandProperty "ClassName" | Should Be "SMS_CategoryInstance"
            $Category | select -ExpandProperty "Filter" | Should Be "((LocalizedCategoryInstanceName = 'TestCategory'))"
        }

        It "Return list of Categories by multiple names using the pipeline" {
            $Categories = @("TestCategory1", "TestCategory2") | Get-Category | select -ExpandProperty "Filter" 

            $Categories.Count | Should be 2
            $Categories[0] | Should Be "((LocalizedCategoryInstanceName = 'TestCategory1'))"
            $Categories[1] | Should Be "((LocalizedCategoryInstanceName = 'TestCategory2'))"
        }

        It "Return Categories by ParentID" {
            $TestCategory = Get-Category -Name "%" -Search -ParentID 42 
            $TestCategory | select -ExpandProperty "ClassName" | Should Be "SMS_CategoryInstance"
            $TestCategory | select -ExpandProperty "Filter" | Should Be "((LocalizedCategoryInstanceName LIKE '%') AND (ParentCategoryInstanceID = 42))"
        }

        It "Return Categories by Type" {
            $TestCategory = Get-Category -Name "%" -Search -TypeName DriverCategories
            $TestCategory | select -ExpandProperty "ClassName" | Should Be "SMS_CategoryInstance"
            $TestCategory | select -ExpandProperty "Filter" | Should Be "((LocalizedCategoryInstanceName LIKE '%') AND (CategoryTypeName = 'DriverCategories'))"
        }

        It "Return Categories by custom filter" {
            $TestCategory = Get-Category -Filter "LocalizedCategoryInstanceName LIKE '%'"
            $TestCategory | select -ExpandProperty "ClassName" | Should Be "SMS_CategoryInstance"
            $TestCategory | select -ExpandProperty "Filter" | Should Be "(LocalizedCategoryInstanceName LIKE '%')"
        }
    }

    Describe "New-Category" {
        It "Tests are implemented"{
            $False | Should be $True
        }
    }

    Describe "Remove-Category" {
        It "Tests are implemented" {
            $False | Should be $True
        }
    }
    #################################
    #endregion SMS_CategoryInstance #
    #################################


    ####################################
    #region  SMS_ObjectContainerNode   #
    ####################################
    Describe "Move-CMObject" {

        Context "No connection" {
            Mock Invoke-WmiMethod {throw}

            It "Throw if no connection has been established" {
                {Move-CMObject -Class "TestClass" -MethodName "TestMethod"} | Should Throw
            }
        }

        Context "Connection established" {
            Mock Test-CMConnection { $true }
            Mock Get-CMInstance { [PSCustomObject]@{Name = "TestFolder1"; ObjectType = 2} }
            Mock Get-CMInstance { [PSCustomObject]@{Name = "TestFolder1"; ObjectType = 3} } -ParameterFilter {$Filter -eq "ContainerNodeID = 3"}
            Mock Get-CMInstance { $null } -ParameterFilter {$Filter -eq "ContainerNodeID = 42"}
            Mock Invoke-CMMethod { [PSCustomObject]@{ReturnValue = 1; Arguments = $Arguments} }
            Mock Invoke-CMMethod { [PSCustomObject]@{ReturnValue = 0; Arguments = $Arguments} } -ParameterFilter {(Compare-Object $Arguments @{InstanceKeys=@(5);ContainerNodeID=1;TargetContainerNodeID=4;ObjectType=2} -PassThru) -eq $null}
            Mock Invoke-CMMethod { [PSCustomObject]@{ReturnValue = 0; Arguments = $Arguments} } -ParameterFilter {(Compare-Object $Arguments @{InstanceKeys=@(1);ContainerNodeID=3;TargetContainerNodeID=0;ObjectType=3} -PassThru) -eq $null}
    
            $global:CMProviderServer = "TestProviderServer"
            $global:CMSiteCode = "TST"
            $global:CMNamespace = "root\sms\Site_TST"

            It "Return False if Folder is not available and write Warning" {
                $Result = (Move-CMObject -SourceFolderID 42 -TargetFolderID 2 -ObjectID "1" 3>&1)

                $Result[0] | Should Match "Unable to move object 1. SourceFolder 42 cannot be retrieved.*"
                $Result[1] | Should Be $false

                $Result = (Move-CMObject -SourceFolderID 2 -TargetFolderID 42 -ObjectID "1" 3>&1)
                $Result[0] |Should Match "Unable to move object 1. TargetFolder 42 cannot be retrieved.*"
            } 

            It "Return False if folder types don't match and write warning" {
                $Result =(Move-CMObject -SourceFolderID 2 -TargetFolderID 3 -ObjectID "1" 3>&1) 
            
                $Result[0] | Should Match "Unable to move object 1. SourceFolder 2 ObjectType 2 and TargetFolder 3 ObjectType 3 don't match."
                $Result[1] | Should Be $false
            }

            It "Use supplied values" {
                Move-CMObject -SourceFolderID 1 -TargetFolderID 4 -ObjectID "5" | Should Be $true
            }

            It "Handle Root" {
                Move-CMObject -SourceFolderID 3 -TargetFolderID 0 -ObjectID "1" | Should Be $true
            }
        }
    }


    Describe "Get-Folder" {
    
        Mock Get-CMInstance { 
            [PSCustomObject]@{Class = $ClassName; Filter = $Filter}
        }
    
        It "Throw if no name is supplied" {
            {Get-Folder -Name "" -Type Package} | Should Throw
        } 

        It "Throw on wrong Type" {
            {Get-Folder -Name "Test" -Type FailType} | Should Throw
        }

        It "Return Folder by Name with no Parent Folder specified" {
            $TestFolder = Get-Folder -Name "TestFolder" -Type Package
        
            $TestFolder | select -ExpandProperty "Class" | Should Be "SMS_ObjectContainerNode"
            $TestFolder | select -ExpandProperty "Filter" | Should Match "(Name = 'TestFolder')"
            $TestFolder | select -ExpandProperty "Filter" | Should Match "(ObjectType=2)"
            $TestFolder | select -ExpandProperty "Filter" | Should Not Match "ParentContainerNodeID = "
        }

        It "Return Folder by Name with with specific parent folder" {
            $TestFolder = Get-Folder -Name "TestFolder" -Type Package -ParentFolderID 0
        
            $TestFolder | select -ExpandProperty "Class" | Should Be "SMS_ObjectContainerNode"
            $TestFolder | select -ExpandProperty "Filter" | Should Match "(Name = 'TestFolder')"
            $TestFolder | select -ExpandProperty "Filter" | Should Match "(ObjectType=2)"
            $TestFolder | select -ExpandProperty "Filter" | Should Match "(ParentContainerNodeID = 0)"
        }

        It "Return Folder by ID" {
            $TestFolder = Get-Folder -ID 1
        
            $TestFolder | select -ExpandProperty "Class" | Should Be "SMS_ObjectContainerNode"
            $TestFolder | select -ExpandProperty "Filter" | Should Be "ContainerNodeID = 1"
        }

        It "Return custom Root object on ID 0" {
            $TestFolder = Get-Folder -ID 0
        
            $TestFolder | select -ExpandProperty "Name" | Should Be "Root"
            $TestFolder | select -ExpandProperty "ContainerNodeID" | Should Be "0"
            $TestFolder | select -ExpandProperty "ObjectType" | Should Be "0"
        }

        It "Use correct object type" {
            Get-Folder -Name "TestFolder" -Type Package | select -ExpandProperty "Filter" | Should Match "(ObjectType=2)"
            Get-Folder -Name "TestFolder" -Type Advertisement | select -ExpandProperty "Filter" | Should Match "(ObjectType=3)"
            Get-Folder -Name "TestFolder" -Type Query | select -ExpandProperty "Filter" | Should Match "(ObjectType=7)"
            Get-Folder -Name "TestFolder" -Type Report | select -ExpandProperty "Filter" | Should Match "(ObjectType=8)"
            Get-Folder -Name "TestFolder" -Type MeteredProductRule | select -ExpandProperty "Filter" | Should Match "(ObjectType=9)"
            Get-Folder -Name "TestFolder" -Type ConfigurationItem | select -ExpandProperty "Filter" | Should Match "(ObjectType=11)"
            Get-Folder -Name "TestFolder" -Type OSInstallPackage | select -ExpandProperty "Filter" | Should Match "(ObjectType=14)"
            Get-Folder -Name "TestFolder" -Type StateMigration | select -ExpandProperty "Filter" | Should Match "(ObjectType=17)"
            Get-Folder -Name "TestFolder" -Type ImagePackage | select -ExpandProperty "Filter" | Should Match "(ObjectType=18)"
            Get-Folder -Name "TestFolder" -Type BootImagePackage | select -ExpandProperty "Filter" | Should Match "(ObjectType=19)"
            Get-Folder -Name "TestFolder" -Type TaskSequencePackage | select -ExpandProperty "Filter" | Should Match "(ObjectType=20)"
            Get-Folder -Name "TestFolder" -Type DeviceSettingPackage | select -ExpandProperty "Filter" | Should Match "(ObjectType=21)"
            Get-Folder -Name "TestFolder" -Type DriverPackage | select -ExpandProperty "Filter" | Should Match "(ObjectType=23)"
            Get-Folder -Name "TestFolder" -Type Driver | select -ExpandProperty "Filter" | Should Match "(ObjectType=25)"
            Get-Folder -Name "TestFolder" -Type SoftwareUpdate | select -ExpandProperty "Filter" | Should Match "(ObjectType=1011)"
            Get-Folder -Name "TestFolder" -Type ConfigurationBaseline | select -ExpandProperty "Filter" | Should Match "(ObjectType=2011)"
            Get-Folder -Name "TestFolder" -Type DeviceCollection | select -ExpandProperty "Filter" | Should Match "(ObjectType=5000)"
            Get-Folder -Name "TestFolder" -Type UserCollectioN | select -ExpandProperty "Filter" | Should Match "(ObjectType=5001)"
        }
    }

    Describe "New-Folder" {

        Get-CimSession | Remove-CimSession
        $global:CMProviderServer = "TestProviderServer"
        $global:CMSiteCode = "TST"
        $global:CMNamespace = "root\sms\Site_TST"
        $global:CMSession = New-CimSession $env:COMPUTERNAME

        Mock New-CimInstance { 
            [PSCustomObject]@{Class = $ClassName; Arguments = $Arguments}
        }
    
        It "Throw if no name is supplied" {
            {New-Folder -Name "" -Type Package} | Should Throw
        } 

        It "Throw on wrong Type" {
            {New-Folder -Name "TestFolder" -Type FailType} | Should Throw
        }

        It "Use supplied values" {
            $TestFolder = New-Folder -Name "TestFolder" -Type "DriverPackage" -PassThru 

            $TestFolder | select -ExpandProperty "Class" | Should Be "SMS_ObjectContainerNode"
            $TestFolder | select -ExpandProperty "Arguments" | %{$_.Name} | Should Be "TestFolder"
            $TestFolder | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 23
            $TestFolder | select -ExpandProperty "Arguments" | %{$_.ParentContainerNodeid} | Should Be 0
            New-Folder -Name "TestFolder" -Type "DriverPackage" -ParentFolderID 1 -PassThru| select -ExpandProperty "Arguments" | %{$_.ParentContainerNodeid} | Should Be 1
        }

        It "Use correct object type" {
            New-Folder -Name "TestFolder" -PassThru -Type Package | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 2
            New-Folder -Name "TestFolder" -PassThru -Type Advertisement | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 3
            New-Folder -Name "TestFolder" -PassThru -Type Query | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 7
            New-Folder -Name "TestFolder" -PassThru -Type Report | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 8
            New-Folder -Name "TestFolder" -PassThru -Type MeteredProductRule | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 9
            New-Folder -Name "TestFolder" -PassThru -Type ConfigurationItem | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 11
            New-Folder -Name "TestFolder" -PassThru -Type OSInstallPackage | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 14
            New-Folder -Name "TestFolder" -PassThru -Type StateMigration | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 17
            New-Folder -Name "TestFolder" -PassThru -Type ImagePackage | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 18
            New-Folder -Name "TestFolder" -PassThru -Type BootImagePackage | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 19
            New-Folder -Name "TestFolder" -PassThru -Type TaskSequencePackage | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 20
            New-Folder -Name "TestFolder" -PassThru -Type DeviceSettingPackage | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 21
            New-Folder -Name "TestFolder" -PassThru -Type DriverPackage | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 23
            New-Folder -Name "TestFolder" -PassThru -Type Driver | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 25
            New-Folder -Name "TestFolder" -PassThru -Type SoftwareUpdate | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 1011
            New-Folder -Name "TestFolder" -PassThru -Type ConfigurationBaseline | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 2011
            New-Folder -Name "TestFolder" -PassThru -Type DeviceCollection | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 5000
            New-Folder -Name "TestFolder" -PassThru -Type UserCollection | select -ExpandProperty "Arguments" | %{$_.ObjectType} | Should Be 5001
        }

    }
    ####################################
    #endregion SMS_ObjectContainerNode #
    ####################################


    ####################################
    #region  SMS_TaskSequencePackage   #
    ####################################
    Describe Get-TaskSequencePackage {

        Get-CimSession | Remove-CimSession
        $global:CMProviderServer = "TestProviderServer"
        $global:CMSiteCode = "TST"
        $global:CMNamespace = "root\sms\Site_TST"
        $global:CMSession = New-CimSession $env:COMPUTERNAME

        Mock Get-CimInstance { 
            [PSCustomObject]@{Class = $ClassName; Filter = $Filter}
        }
        
        It "Throw if no or empty ID is supplied" {
            {Get-TaskSequencePackage -ID $null} | Should Throw
            {Get-TaskSequencePackage -ID ""} | Should Throw
        }

        It "Throw if no or empty name is supplied" {
            {Get-TaskSequencePackage -Name $null} | Should Throw
            {Get-TaskSequencePackage -Name ""} | Should Throw
        }

        It "Get TaskSequencePackage by ID" {
            $TestPackage = Get-TaskSequencePackage -ID "TST000001" 

            $TestPackage | select -ExpandProperty "Class" | Should Be "SMS_TaskSequencePackage"
            $TestPackage | select -ExpandProperty "Filter" | Should Be "PackageID='TST000001'"
        } 

        It "Get TaskSequencePackage by Name" {
            $TestPackage = Get-TaskSequencePackage -Name "Test TaskSequence" 

            $TestPackage | select -ExpandProperty "Class" | Should Be "SMS_TaskSequencePackage"
            $TestPackage | select -ExpandProperty "Filter" | Should Be "Name='Test TaskSequence'"
        } 
    }

    Describe New-TaskSequencePackage {

        Get-CimSession | Remove-CimSession
        $global:CMProviderServer = "TestProviderServer"
        $global:CMSiteCode = "TST"
        $global:CMNamespace = "root\sms\Site_TST"
        $global:CMSession = New-CimSession $env:COMPUTERNAME
        #-Class "SMS_TaskSequencePackage" -Arguments @{Name=$Name;Description=$Description}
        Mock New-CimInstance { 
            [PSCustomObject]@{Class = $ClassName; Arguments = $Arguments}
        }

        It "Throw if no or empty name is supplied" {
            {New-TaskSequencePackage -Name $null} | Should Throw
            {New-TaskSequencePackage -Name ""} | Should Throw
        }

        It "Create new TaskSequencePackage" {
            $TestPackage = New-TaskSequencePackage -Name "New TaskSequencePackage" -Description "Test Description" -TaskSequence @{Name="Test"}

            $TestPackage | select -ExpandProperty "Class" | Should Be "SMS_TaskSequencePackage"
            $TestPackage | select -ExpandProperty "Arguments" | %{$_.Name} | Should Be "New TaskSequencePackage"
            $TestPackage | select -ExpandProperty "Arguments" | %{$_.Description} | Should Be "Test Description"
        } 

    }
    ####################################
    #endregion SMS_TaskSequencePackage #
    ####################################

    #############################
    #region  SMS_TaskSequence   #
    #############################
    Describe Get-TaskSequence {

        Get-CimSession | Remove-CimSession
        $global:CMProviderServer = "TestProviderServer"
        $global:CMSiteCode = "TST"
        $global:CMNamespace = "root\sms\Site_TST"
        $global:CMSession = New-CimSession $env:COMPUTERNAME

        Mock Invoke-CimMethod { 
            [PSCustomObject]@{ClassName = $ClassName; MethodName = $MethodName; ReturnValue = 0; TaskSequence = $Arguments}
        }

        Mock Get-CimInstance { 
            [PSCustomObject]@{ClassName = $ClassName; Filter = $Filter}
        }
        
        It "Throw if no or empty ID is supplied" {
            {Get-TaskSequence -ID $null} | Should Throw
            {Get-TaskSequence -ID ""} | Should Throw
        }

        It "Throw if no or empty name is supplied" {
            {Get-TaskSequence -Name $null} | Should Throw
            {Get-TaskSequence -Name ""} | Should Throw
        }

        It "Get TaskSequence by ID" {
            $TestSequence = Get-TaskSequence -ID "TST000001" 

            $TestSequence |  %{$_.TaskSequencePackage.ClassName} | Should Be "SMS_TaskSequencePackage"
        } 

        It "Get TaskSequence by Name" {
            $TestSequence = Get-TaskSequence -Name "Test TaskSequence" 

            $TestSequence |  %{$_.TaskSequencePackage.ClassName} | Should Be "SMS_TaskSequencePackage"
        } 
    }
    #############################
    #endregion SMS_TaskSequence #
    #############################

    ################################
    #region  SMS_ClientOperation   #
    ################################
    Describe "Get-ClientOperation" {
    
        Mock Get-CMInstance { [PSCustomObject]@{ClassName = $ClassName; Filter = $Filter} }
    
        It "Return Client Operations" {
            Get-ClientOperation | select -ExpandProperty "ClassName" | Should Be "SMS_ClientOperation"
            Get-ClientOperation | select -ExpandProperty "Filter" | Should Be ""
            Get-ClientOperation -Expired | select -ExpandProperty "Filter" | Should Be "((State = 0) OR (State = 2))"
        }
    }

    Describe "Start-ClientOperation" {    
        Mock Invoke-CMMethod { [PSCustomObject]@{ReturnValue = 0; ClassName = $ClassName; MethodName = $MethodName; Arguments = $Arguments} }

        It "Throw if Operation is invalid or no CollectionID supplied" {        
            {Start-ClientOperation -Operation "" -CollectionID "XXX00001" } | Should Throw
            {Start-ClientOperation -Operation FullScan -CollectionID "" } | Should Throw
        }
      
        It "Invoke correct WMI Method" {
            Start-ClientOperation -Operation FullScan -CollectionID "XXX00001" | select -ExpandProperty "ClassName" | Should Be "SMS_ClientOperation"
            Start-ClientOperation -Operation FullScan -CollectionID "XXX00001" | select -ExpandProperty "MethodName" | Should Be "InitiateClientOperation"
        }

        It "Use supplied CollectionID" {
            Start-ClientOperation -Operation FullScan -CollectionID "XXX00001" | select -ExpandProperty "Arguments" | %{$_.TargetCollectionID} |  Should Be "XXX00001"
        }

        It "Use supplied ResourceID" {
            Start-ClientOperation -Operation FullScan -CollectionID "XXX00001" -ResourceID 42| select -ExpandProperty "Arguments" | %{$_.TargetResourceIDs}  | Should Be "42"

            $ResourceIDs = Start-ClientOperation -Operation FullScan -CollectionID "XXX00001" -ResourceID 42,43,44| select -ExpandProperty "Arguments" | %{$_.TargetResourceIDs} 
            $ResourceIDs.Count | Should Be 3
            $ResourceIDs | Should Be ("42", "43", "44")
        }

        It "Use correct Operation" {
            Start-ClientOperation -Operation FullScan -CollectionID "XXX00001" | select -ExpandProperty "Arguments" |  %{$_.Type} | Should Be "1"
            Start-ClientOperation -Operation QuickScan -CollectionID "XXX00001" | select -ExpandProperty "Arguments" | %{$_.Type} | Should Be "2"
            Start-ClientOperation -Operation DownloadDefinition -CollectionID "XXX00001" | select -ExpandProperty "Arguments" | %{$_.Type} | Should Be "3"
            Start-ClientOperation -Operation EvaluateSoftwareUpdates -CollectionID "XXX00001" | select -ExpandProperty "Arguments" | %{$_.Type} | Should Be "4"
            Start-ClientOperation -Operation RequestComputerPolicy -CollectionID "XXX00001" | select -ExpandProperty "Arguments" | %{$_.Type} | Should Be "8"
            Start-ClientOperation -Operation RequestUserPolicy -CollectionID "XXX00001" | select -ExpandProperty "Arguments" | %{$_.Type} | Should Be "9"
        }
    }

    Describe "Stop-ClientOperation" {
    
        Mock Invoke-CMMethod { [PSCustomObject]@{ReturnValue = 0; ClassName = $ClassName; MethodName = $MethodName; Arguments = $Arguments} }
      
        It "Invoke correct WMI Method" {
            Stop-ClientOperation -OperationID 42 | select -ExpandProperty "ClassName" | Should Be "SMS_ClientOperation"
            Stop-ClientOperation -OperationID 42| select -ExpandProperty "MethodName" | Should Be "CancelClientOperation"
        }

        It "Use correct OperationID" {
            Stop-ClientOperation -OperationID 42| select -ExpandProperty "Arguments" | %{$_.OperationID} | Should Be 42
            Stop-ClientOperation -OperationID 42,43,44 | select -ExpandProperty "Arguments" -First 1 | %{$_.OperationID} | Should Be 42
            Stop-ClientOperation -OperationID 42,43,44 | select -ExpandProperty "Arguments" -Last 1 | %{$_.OperationID} | Should Be 44
        }

        It "Invoke correct WMI Method by pipeline" {
            42 | Stop-ClientOperation | select -ExpandProperty "ClassName" | Should Be "SMS_ClientOperation"
            42 | Stop-ClientOperation | select -ExpandProperty "MethodName" | Should Be "CancelClientOperation"
        }

        It "Use correct OperationID by pipeline" {
            42 | Stop-ClientOperation | select -ExpandProperty "Arguments" | %{$_.OperationID} | Should Be 42
            @(42,43,44) | Stop-ClientOperation | select -ExpandProperty "Arguments" -First 1 | %{$_.OperationID} | Should Be 42
            @(42,43,44) | Stop-ClientOperation | select -ExpandProperty "Arguments" -Last 1 | %{$_.OperationID} | Should Be 44
        }
    }

    Describe "Remove-ClientOperation" {
    
        Mock Invoke-CMMethod { [PSCustomObject]@{ReturnValue = 0; ClassName = $ClassName; MethodName = $MethodName; Arguments = $Arguments} }
      
        It "Invoke correct WMI Method" {
            Remove-ClientOperation -OperationID 42 | select -ExpandProperty "ClassName" | Should Be "SMS_ClientOperation"
            Remove-ClientOperation -OperationID 42| select -ExpandProperty "MethodName" | Should Be "DeleteClientOperation"
        }

        It "Use correct OperationID" {
            Remove-ClientOperation -OperationID 42| select -ExpandProperty "Arguments" | %{$_.OperationID} | Should Be 42
            Remove-ClientOperation -OperationID 42,43,44 | select -ExpandProperty "Arguments" -First 1 | %{$_.OperationID} | Should Be 42
            Remove-ClientOperation -OperationID 42,43,44 | select -ExpandProperty "Arguments" -Last 1 | %{$_.OperationID} | Should Be 44
        }

        It "Invoke correct WMI Method by pipeline" {
            42 | Remove-ClientOperation | select -ExpandProperty "ClassName" | Should Be "SMS_ClientOperation"
            42 | Remove-ClientOperation | select -ExpandProperty "MethodName" | Should Be "DeleteClientOperation"
        }

        It "Use correct OperationID by pipeline" {
            42 | Remove-ClientOperation | select -ExpandProperty "Arguments" | %{$_.OperationID} | Should Be 42
            @(42,43,44) | Remove-ClientOperation | select -ExpandProperty "Arguments" -First 1 | %{$_.OperationID} | Should Be 42
            @(42,43,44) | Remove-ClientOperation | select -ExpandProperty "Arguments" -Last 1 | %{$_.OperationID} | Should Be 44
        }
    }
    ################################
    #endregion SMS_ClientOperation #
    ################################

}