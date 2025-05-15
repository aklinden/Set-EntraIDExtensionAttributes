<#
    .SYNOPSIS
        Creates Extensionattribute properties based on the location of the object in Active Directory tree, compares them with the existing extensionattributes in Entra ID, and updates them if there's a mismatch. 
    .DESCRIPTION
        After gathering all Active Directory Computer objects, Get-Extensionattributes will generate $Office, $Department, and $Devicetype using the distinguishedName property of each device. Next, Get-MgDevice finds any 
        objects with the same name. The ID property for each returned result is then used to pull the existing extensionattributes in EntraID so they can be compared with the AD objects. Any differences between the Entra
        ID objects and the Active Directory objects will trigger an update to be sent to the Entra ID object's extensionattribute.         
    .NOTES
        Author: Andy Linden
        Date: 11/27/2024
#>


function Get-Extensionattributes {

    [cmdletbinding()]
    param (
        # Specify the user you'd like to get info for.
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelinebypropertyname = $true)]
        [Alias('CN', 'Name')]
        [string[]] $computername     
    )
    


    process {
        foreach ($computer in $computername) {

            $ADObject = get-adcomputer -filter "Name -like '$computer'" -Properties description
            #Determine Office based on OU path
            $Office = switch -Regex ($ADObject.distinguishedName) {
                "DEN" { "DEN"; break }
                "LAX" { "LAX"; break }
                "NYC" { "NYC"; break }
                "PDX" { "PDX"; break }
                Default { "Unknown" }
            }# Office



            #Determine department based on OU path
            $dept = switch -Regex ($ADObject.distinguishedName) {
                "Accounting Computers" { "Accounting"; break }
                "IT Computers" { "IT"; break }
                "Marketing Computers" { "Marketing"; break }
                "Support Computers" { "Support"; break }
                "Travel Computers" { "Travel"; break }
                "Servers" { "IT"; break }
                Default { "Unknown" }
            }# Department



            #determine the machine type based on OU path
            $devicetype = switch -Regex ($ADObject.distinguishedName) {
                "Desktop Servers" { "Server"; break }
                "Servers" { "Server"; break }
                "Desktop" { "Desktop"; break }
                "Laptop" { "Laptop"; break }
                Default { "Unknown" }
            }# Machine type

            [PSCustomObject]@{
                'ComputerName' = $ADObject.Name
                'Description'  = $ADObject.description
                'Office'       = $Office
                'Department'   = $dept
                'Machinetype'  = $devicetype
            }#properties
            
        }#foreach ADObject
        
    }#process

}# function

function Compare-EntraIDExtensionAttributes {
    [CmdletBinding()]
    param (
        [string]$AzureID,
        [string]$Office,
        [string]$Department,
        [string]$Devicetype,
        [string]$Infrastructure,
        [string]$computername,
        [PSCustomObject]$mgdeviceObject
    )

    # Initialize parameters
    $params = @{
        extensionAttributes = @{}
    }

    # Compare and prepare updates for extension attributes
    if ($mgdeviceObject.ExtensionAttribute1 -ne $Office) {
        $params.extensionAttributes.extensionAttribute1 = $Office
    }

    if($mgdeviceObject.extensionAttribute2 -ne $Department) {
        $params.extensionAttributes.extensionAttribute2 = $Department
    }

    if ($mgdeviceObject.extensionAttribute3 -ne $Devicetype) {
        $params.extensionAttributes.extensionAttribute3 = $Devicetype
    }

    if ($mgdeviceObject.extensionAttribute4 -ne $Infrastructure) {
        $params.extensionAttributes.extensionAttribute4 = $Infrastructure
    }

    # Only proceed if there are attributes to update
    if ($params.extensionAttributes.Count -gt 0) {
        Write-Verbose "Updating attributes for: $computerName"

        try {
            # Send the updates to Entra ID
            Update-MgDevice -DeviceId $AzureID -BodyParameter $params
            Write-Verbose "Successfully updated attributes for: $computerName"
        } catch {
            Write-Verbose "Error occurred while trying to update: $computerName - $_"
        }
    }
}

### LOGGING BLOCK ### 
$date = Get-Date -Format 'MM-dd HH-mm'
$logpath = "$home\documents\"
$TranscriptFile = "SetEntraIDExtensionAttributesTranscript_$date.log"

$SetIntuneExtensionTranscript = Join-Path -Path $logpath -ChildPath $TranscriptFile

#Begin transcript
Start-Transcript -Path $SetIntuneExtensionTranscript

Write-Host "Started processing at [$([DateTime]::Now)]."

#Remove older logs
$OldFiles = Get-ChildItem -Path $logpath -Filter "SetEntraIDExtensionAttributes*" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) }

#remove files older than 7 days
foreach ($file in $OldFiles) {
    try {
        Remove-Item -Path $file.FullName -Force
        Write-Host "LOGGING: Removing: $($file.FullName)"
    } catch {
        Write-host "LOGGING: Error deleting file: $($file.FullName) - $_"
    }
}

#Connect to Graph API
Connect-mgGraph

# Get on-premises computer list
$ADComputerList = Get-ADComputer -Filter * -Properties Description

Write-host "Attempting to update the following computers: "

# Iterate through each computer in the list
foreach($computer in $ADComputerList){
    # Build variables
    $ADName = $computer.Name
    $ADGUID = ($computer.objectguid).guid
    $Azureobject = Get-MgDevice -Filter "startswith(displayName, '$ADName')" -ConsistencyLevel eventual
    
    # Extensionattributes
    $Attributes = Get-Extensionattributes -computername $ADName
    $Office = $Attributes.Office
    $Department = $Attributes.Department
    $Devicetype = $Attributes.Machinetype

    foreach($match in $Azureobject){

        $AzureID = $match.id
        $AzureDeviceID = $match.deviceID

        if ($AzureDeviceID -match $ADGUID){
            $Infrastructure = "On-Premises"
        }
        else{
            $Infrastructure = "Cloud"
        } 
        
        $filter = "$($AzureID)?`$select=id,deviceid,displayname,extensionAttributes,trustType"
        $uri = "https://graph.microsoft.com/v1.0/devices/$filter"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri

        if ($response) {

            if($response.trustType -eq "ServerAd"){
                $Infrastructure = "On-Premises"
            }
            elseif ($response.trustType -eq "AzureAd") {
                $Infrastructure = "Cloud"
            }

            # Create a new PSCustomObject to hold the device objects EntraID extension attributes
            $mgdeviceObject = [PSCustomObject][ordered]@{
            'displayName' = $response.displayName
            'deviceID' = $response.deviceId
            'ExtensionAttribute1' = $response.extensionAttributes.extensionattribute1
            'ExtensionAttribute2' = $response.extensionAttributes.extensionattribute2
            'ExtensionAttribute3' = $response.extensionAttributes.extensionattribute3
            'ExtensionAttribute4' = $response.extensionAttributes.extensionattribute4
            }
         

            $compareParams = @{
                AzureID =        $AzureID
                Office =         $Office
                Department =     $Department
                Devicetype =     $Devicetype
                Infrastructure = $Infrastructure
                Computername =   $ADName
                mgdeviceObject = $mgdeviceObject
            }

            # Call the function to update extension attributes
            Compare-EntraIDExtensionAttributes @compareParams
        }#If $response
    }#foreach $computer in $AzureObject
}#foreach $computer in $ADComputerList

Stop-Transcript
