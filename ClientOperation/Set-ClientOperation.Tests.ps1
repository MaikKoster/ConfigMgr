###############################################################################
# 
# PESTER tests for Set-ClientOperation.ps1
# 
# Author:  Maik Koster
# Date:    25.03.2016
# History:
#          07.12.2015 - Initial setup of tests
#          25.03.2016 - Updated Tests for use with ConfigMgr Module
#
###############################################################################

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here\Set-ClientOperation.ps1" -Cancel -All

Describe "Script with parameter" {

    Mock New-CMConnection { [PSCustomObject]@{ProviderServer = $ProviderServer; SiteCode = $SiteCode; Credentials = $Credential} }
    Mock Invoke-CMMethod { [PSCustomObject]@{Methodname=$MethodName; Arguments = $Arguments} }
    Mock Get-CimInstance {}
    Mock Get-CimInstance { [PSCustomObject]@{ID = 1} } -ParameterFilter {$ClassName -eq "SMS_ClientOperation"}
    Mock Get-CimInstance { [PSCustomObject]@{ID = 2} } -ParameterFilter {($ClassName -eq "SMS_ClientOperation") -and ($Filter -Like "*State = 2*")}

    Context "Start" {

        Mock New-CMConnection {}
        Mock Get-CimInstance { [PSCustomObject]@{CollectionID = "TST00001"} } -ParameterFilter { $Filter -Like "*TestCollection*"}
        Mock Get-CimInstance { [PSCustomObject]@{ResourceID = "987654"} } -ParameterFilter { $Filter -Like "*FirstTest*"}
        Mock Get-CimInstance { [PSCustomObject]@{ResourceID = "876543"} } -ParameterFilter { $Filter -Like "*SecondTest*"}
        Mock Get-CimInstance { [PSCustomObject]@{ResourceID = @("987654","876543")} } -ParameterFilter { $Filter -Like "*FirstTest*SecondTest*"}
       
        It "Throw if no CollectionID/CollectionName and ResourceID/ResourceName is supplied" {
            {Main -Start FullScan} | Should Throw
        }

        It "Only CollectionID specified" {
            $Result = Main -Start FullScan -CollectionID "TST00001"
            $Result.Arguments["TargetCollectionID"] | Should Be "TST00001"
            $Result.Arguments["TargetResourceIDs"] | Should Be $null
        }

        It "Only CollectionName specified" {
            $Result = Main -Start FullScan -CollectionName "TestCollection" 
            $Result.Arguments["TargetCollectionID"] | Should Be "TST00001"
            $Result.Arguments["TargetResourceIDs"] | Should Be $null
        }

        It "Throw if CollectionName cannot be resolved" {
            {Main -start FullScan -CollectionName "FailCollection"} | Should Throw
        }

        It "No Collection specified but ResourceID" {
            $Result = Main -Start FullScan -ResourceID 123456 
            $Result.Arguments["TargetCollectionID"] | Should Be "SMS00001"
            $Result.Arguments["TargetResourceIDs"] | Should Be 123456
        }

        It "No Collection specified but ResourceName" {
            $Result = Main -Start FullScan -ResourceName "FirstTest" 
            $Result.Arguments["TargetCollectionID"] | Should Be "SMS00001"
            $Result.Arguments["TargetResourceIDs"] | Should Be 987654
        }

        It "Collection ID and ResourceID" {
            $Result = Main -Start FullScan -CollectionID "TST00001" -ResourceID 123456 
            $Result.Arguments["TargetCollectionID"] | Should Be "TST00001"
            $Result.Arguments["TargetResourceIDs"] | Should Be 123456
        }

        It "Collection ID and multiple ResourceID" {
            $Result = Main -Start FullScan -CollectionID "TST00001" -ResourceID @(123456,234567) 
            $Result.Arguments["TargetCollectionID"] | Should Be "TST00001"
            $ResourceID = $Result.Arguments["TargetResourceIDs"]
            $ResourceID[0] | Should Be 123456
            $ResourceID[1] | Should Be 234567
        }

        It "Collection ID and ResourceName" {
            $Result = Main -Start FullScan -CollectionID "TST00001" -ResourceName "FirstTest" 
            $Result.Arguments["TargetCollectionID"] | Should Be "TST00001"
            $Result.Arguments["TargetResourceIDs"] | Should Be 987654
        }

        It "Collection ID and multiple ResourceNames" {
            $Result = Main -Start FullScan -CollectionID "TST00001" -ResourceName @("FirstTest", "SecondTest") 
            $Result.Arguments["TargetCollectionID"] | Should Be "TST00001"
            $ResourceID = $Result.Arguments["TargetResourceIDs"]
            $ResourceID[0] | Should Be 987654
            $ResourceID[1] | Should Be 876543
        }

        It "Valid operation" {
            $Result = Main -Start FullScan -CollectionID "TST00001"
            $Result.Methodname | Should Be "InitiateClientOperation"
            $Result.Arguments["Type"] | Should Be 1 
            $Result = Main -Start QuickScan -CollectionID "TST00001" 
            $Result.Arguments["Type"] | Should Be 2
            $Result = Main -Start DownloadDefinition -CollectionID "TST00001" 
            $Result.Arguments["Type"] | Should Be 3
            $Result = Main -Start EvaluateSoftwareUpdates -CollectionID "TST00001" 
            $Result.Arguments["Type"] | Should Be 4
            $Result = Main -Start RequestComputerPolicy -CollectionID "TST00001" 
            $Result.Arguments["Type"] | Should Be 8
            $Result = Main -Start RequestUserPolicy -CollectionID "TST00001" 
            $Result.Arguments["Type"] | Should Be 9
        }

        it "Throw on invalid operation" {
            {Main -Start Fail} | Should Throw
        }
    }

    Context "Cancel" {

        Mock New-CMConnection {}

        It "Warning if no OperationID available" {
            (Main -Cancel 3>&1) | Should Be "No Client Operation available."
        }        

        It "Use correct WMI Method" {
            $Result = Main -Cancel -OperationID 42 
            $Result.MethodName = "CancelClientOperation"
        }
        It "Use supplied OperationID" {
            $Result = Main -Cancel -OperationID 42 
            $Result.Arguments["OperationID"] | Should Be 42
        }

        It "Cancel All Client Operations" {
            $Result = Main -Cancel -All 
            $Result.Arguments["OperationID"] | Should Be 1
        }
    }

    Context "Delete" {

        Mock New-CMConnection {}

        It "Warning if no OperationID available" {
            (Main -Delete 3>&1) | Should Be "No Client Operation available."
        }        

        It "Use supplied OperationID" {
            $Result = Main -Delete -OperationID 42 
            $Result.Arguments["OperationID"] |  Should Be 42
        }

        It "Get All Client Operations" {
            $Result = Main -Delete -All
            $Result.Arguments["OperationID"]| Should Be 1
        }

        It "Get Expired Client Operations" {
            $Result = Main -Delete -Expired 
            $Result.Arguments["OperationID"]| Should Be 2
        }
    }
}