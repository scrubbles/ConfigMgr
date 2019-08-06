<#
.SYNOPSIS
    Name: New-Application_AD-Group_CM-Colleciton.ps1
    The purpose of this script is to easily create AD Based Configuration Manager collecitons.
    
.DESCRIPTION
    Script should do the following, depending on switch
        Check for existing AD groups with same name and report warning if already exists
        Check for existing CM Collection with same name and report warning if already exists
        Create the AD Group and CM Colleciton
    Be sure to replace default values in parameters to your specifications.

.PARAMETER ApplicationName
Specify the name of the application. This uses the ConventionalName parameter set to name the AD Group and CM Collection according to convention specified in script (default values for the variables)

.PARAMETER BaseADGroupName
Used in the ConventionalName parameter set. Specify your naming convention for AD Groups here by setting the default value for the variable. Use this switch to override.

.PARAMETER BaseCMCollectionName
Used in the ConventionalName parameter set. Specify your naming convention for CM Collections here by setting the default value for the variable. Use this switch to override.

.PARAMETER CMCollectionName
Mnaually specifies a name for the CM Colleciton

.PARAMETER ADGroupName
Mnaually specifies a name for the AD Group

.PARAMETER OU
Path to Active Directory OU where group should be created. Manually specified in script to ensure consistency. Use this switch to override and pick a new OU.

.PARAMETER Description
Description for AD Group. Manually specified in script to ensure consistency. Use this switch to add more info for description.

.NOTES
Release Date: 2019.04.15
Update Date: 
Author: scrubbles_lc

You should have your target OU added to Hierarchy discovery. Otherwise be sure to manually go add the new group to to the Active Directory Group Discovery method in CM > Administration >  Hierarchy Configuration > Discovery Methods.

.EXAMPLE
$AppName = 'AdobeFlashPlayer'    
New-Application_AD-Group_CM-Colleciton.ps1 -AppName $AppName

Creates an AD group called "SCCM.Applications.AdobeFlashPlayer" and CM User Collection called "Applications.AdobeFlashPlayer" using the naming convention specified in the script parameters

.EXAMPLE
New-Application_AD-Group_CM-Colleciton.ps1 -CMName 'Manual CM Name' -ADName 'Manual AD Name'

Creates an AD group and CM Collection with manually specified names, not using the predefined convention.
#>

[CmdletBinding()]

PARAM ( 
    [Parameter(
        Mandatory=$False,
        HelpMessage="SCCM Site Server Name"
        )]
    [Alias('SiteServer')]
    [String]$ProviderMachineName = "CM01",

    [Parameter(
        Mandatory=$False,
        HelpMessage="SCCM Site Code"
        )]
    [String]$SiteCode = "PS1",

    [Parameter(
        Mandatory=$False,
        HelpMessage="OU Location to create the AD Group")]
    [Alias('Path')]
    [string]$OU = "OU=Groups,DC=Company,DC=com",

    [Parameter(
        Mandatory=$False,
        HelpMessage="Specify Limiting collection for new CM collection")]
    [string]$LimitingCollection = "All Users and User Groups",

    [Parameter(
        Mandatory=$False,
        HelpMessage="Description. Short comment about purpose"
        )]
    [Alias('descr')]
    [String]$Description,

#parameters for manual naming    
    [Parameter(
        ParameterSetName="ManualName",
        Mandatory=$True,
        HelpMessage="Manuall specify the name of the CM Collection.")]
    [Alias('CMName')]
    [string]$ManualCMCollectionName,

    [Parameter(
        ParameterSetName="ManualName",
        Mandatory=$True,
        HelpMessage="Manually specify the name of the AD Group.")]
    [Alias('ADName')]
    [string]$ManualADGroupName,

#parameters for enforcing naming convention
    [Parameter(
        ParameterSetName="ConventionalName",
        Mandatory=$False,
        HelpMessage="Specify the base name of the CM Collection according to your naming convention.")]
    [string]$BaseCMCollectionName='Applications.',

    [Parameter(
        ParameterSetName="ConventionalName",
        Mandatory=$False,
        HelpMessage="Specify the base name of the AD Group according to your naming convention.")]
    [string]$BaseADGroupName='SCCM.Applications.',

    [Parameter(
        ParameterSetName="ConventionalName",
        Mandatory=$True,
        HelpMessage="Specify the name of the applicaiton. Appended to basenames.")]
    [Alias('AppName')]
    [string]$ApplicationName
)
#----------------[ Declarations ]------------------------------------------------------

# Set any initial values
$DomainName = (Get-ADDomain).Name


if ($PsCmdlet.ParameterSetName -eq 'ConventionalName') {
    $CMCollectionName = "$BaseCMCollectionName" + "$ApplicationName"
    $ADGroupName = "$BaseADGroupName" + "$ApplicationName"
} elseif ($PsCmdlet.ParameterSetName -eq 'ManualName') {
    $CMCollectionName = $ManualCMCollectionName
    $ADGroupName = $ManualADGroupName
}

if ([string]::IsNullOrEmpty($Description)) {
    $ADDescription = "Group for SCCM $CMcollectionName, provides access to $ApplicationName. This group created by script."
} else {
    $ADDescription = $Description + " - Group for SCCM $CMcollectionName, provides access to $ApplicationName. This group created by script."
}

#----------------[ Functions ]---------------------------------------------------------

function  Connect-CM {
    param ()
    #
    # Press 'F5' to run this script. Running this script will load the ConfigurationManager
    # module for Windows PowerShell and will connect to the site.
    #
    # This script was auto-generated at '10/17/2018 2:16:30 PM'.

    # Uncomment the line below if running in an environment where script signing is 
    # required.
    #Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

    # Site configuration
    $ProviderMachineName = $SiteServer # SMS Provider machine name

    # Customizations
    $initParams = @{}
    #$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
    #$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

    # Do not change anything below this line

    # Import the ConfigurationManager.psd1 module 
    if($null -eq (Get-Module ConfigurationManager)) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
    }

    # Connect to the site's drive if it is not already present
    if($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams
}

#----------------[ Main Execution ]----------------------------------------------------

# Script Execution goes here

Push-Location
Import-Module ActiveDirectory
connect-cm

$ADGroupProps = @{
    Name = $ADGroupName;
    DisplayName = $ADGroupName;
    Description = $ADDescription;
    Path = $OU;
    GroupCategory = "Security";
    GroupScope = "Universal";
}

New-ADGroup @ADGroupProps -PassThru

#wait for AD Group Discovery

$CMResource = Get-CMResource -fast | Where-Object {$_.name -eq "$DomainName\$ADGroupName"}
$StopWatch = New-Object -TypeName System.Diagnostics.Stopwatch 
if ($null -eq $CMResource) {do {
    $StopWatch.Start()
    #does
    Write-host "Waiting 30 seconds for AD Group Discovery, then rechecking."
    start-sleep -Seconds 30
    $CMResource = Get-CMResource -fast | Where-Object {$_.name -eq "$DomainName\$ADGroupName"}
    write-host "Elapsed Time $($StopWatch.Elapsed.ToString())"
} until ($null -ne $CMResource)}

$StopWatch.Stop()
Write-host "Found $($CMResource.Name). creating collection and direct membership rule."
write-host "Total Time $($StopWatch.Elapsed.ToString())"

#Create CM Collection

$CMCollection = New-CMCollection -CollectionType User -Name $CMCollectionName -LimitingCollectionName $LimitingCollection
Add-CMUserCollectionDirectMembershipRule -CollectionId $CMCOllection.CollectionID -ResourceId $CMResource.ResourceID -PassThru

#Remove update schedule
$WMIColl = Get-WmiObject -ComputerName $ProviderMachineName -Namespace "ROOT\SMS\site_$SiteCode" -Class "SMS_Collection" -Filter "CollectionID = '$($CMCollection.CollectionID)'"
$WMIColl.RequestRefresh()
Set-CMCollection -CollectionId $CMCollection.CollectionID -RefreshType None  

#reset shell location 
Pop-Location

