###############################################################################
# 
# PESTER tests for Import-TaskSequence.ps1
# 
# Author:  Maik Koster
# Date:    15.03.2016
# History:
#          15.03.2016 - Initial setup of tests
#
###############################################################################

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here\Import-TaskSequence.ps1" -Path $MyInvocation.MyCommand.Path -ID "TST000001" 

Describe "Start-Import" {

    It "Needs to be implemented" {
        $False | Should Be $True
    }
}