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
    
    Begin {
        $results = @()
    }


    process {
        foreach ($computer in $computername) {

            $ADObject = get-adcomputer -filter "Name -like '$computer'" -Properties description
            #Determine Office based on OU path
            switch -Regex ($ADObject.distinguishedName) {
                "DEN" { $Office = "DEN"; break }
                "LAX" { $Office = "LAX"; break }
                "NYC" { $Office = "NYC"; break }
                "PDX" { $Office = "PDX"; break }
                Default { $Office = "Unknown" }
            }# Office



            #Determine department based on OU path
            switch -Regex ($ADObject.distinguishedName) {
                "Accounting Computers" { $dept = "Accounting"; break }
                "IT Computers" { $dept = "IT"; break }
                "Marketing Computers" { $dept = "Marketing"; break }
                "Support Computers" { $dept = "Support"; break }
                "Travel Computers" { $dept = "Travel"; break }
                "Servers" { $dept = "IT"; break }
                
                Default { $dept = "Unknown" }
            }# Department



            #determine the machine type based on OU path
            switch -Regex ($ADObject.distinguishedName) {
                "Desktop Servers" { $devicetype = "Server"; break }
                "Servers" { $devicetype = "Server"; break }
                "Desktop" { $devicetype = "Desktop"; break }
                "Laptop" { $devicetype = "Laptop"; break }
                
                Default { $devicetype = "Unknown" }
            }# Machine type

            $properties = @{
                'ComputerName' = $ADObject.Name
                'Description'  = $ADObject.description
                'Office'       = $Office
                'Department'   = $dept
                'Machinetype'  = $devicetype
            }#properties
            
        }#foreach ADObject
        
        $results += New-Object -Type psobject -Property $properties
    }#process

    end {
        $results
    }#end 

}# function

function Set-EntraIDExtensionAttributes {
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

    if ($mgdeviceObject.extensionAttribute2 -ne $Department) {
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
        Add-Content -Path $SetIntuneExtensionLogFile -Value "Updating attributes for: $computerName"

        try {
            # Send the updates to Entra ID
            Update-MgDevice -DeviceId $AzureID -BodyParameter $params
        } catch {
            Add-Content -Path $SetIntuneExtensionLogFile -Value "Error occurred while trying to update: $computerName - $_"
        }
    }
}

#Generate variables for logging
$date = Get-Date -Format 'MM-dd HH-mm'
$logpath = "$home\documents\SetEntraIDExtensionAttributes"
$LogFile = "Set-EntraIDExtensionAttributes_$date.log"
$TranscriptFile = "Set-EntraIDExtensionAttributesTranscript_$date.log"

$SetIntuneExtensionLogFile = Join-Path -Path $logpath -ChildPath $LogFile
$SetIntuneExtensionTranscript = Join-Path -Path $logpath -ChildPath $TranscriptFile

#Begin transcript
Start-Transcript -Path $SetIntuneExtensionTranscript

if((Test-Path -Path $SetIntuneExtensionLogFile) -eq $false){
    New-Item -Path $SetIntuneExtensionLogFile -ItemType File
}#create log file if not found

Add-Content -path $SetIntuneExtensionLogFile -Value "Started processing at [$([DateTime]::Now)]."

#Remove older logs
$OldFiles = Get-ChildItem -Path $logpath -Filter "Set-EntraIDExtensionAttributes*" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) }

#remove files older than 7 days
foreach ($file in $OldFiles) {
    try {
        Remove-Item -Path $file.FullName -Force
        Add-content -path $SetIntuneExtensionLogFile -Value "Removing: $($file.FullName)"
    } catch {
        Add-content -path $SetIntuneExtensionLogFile -Value "Error deleting file: $($file.FullName) - $_"
    }
}

#Connect to Graph API
Connect-mgGraph

# Get on-premises computer list
$ADComputerList = Get-ADComputer -Filter * -Properties Description

Add-Content -Path $SetIntuneExtensionLogFile -Value "Attempting to update the following computers: "

# Iterate through each computer in the list
foreach ($computer in $ADComputerList) {
    # Build variables
    $ADName = $computer.Name
    $ADGUID = ($computer.objectguid).guid
    $Azureobject = Get-MgDevice -Filter "startswith(displayName, '$ADName')"
    
    # Extensionattributes
    $Attributes = Get-Extensionattributes -computername $computer.name
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
        
        # Variables for checking existing extension attributes
        $uri = "https://graph.microsoft.com/v1.0/devices/$($AzureID)?$select=id,displayname,extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionattribute4"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri

        if ($response) {
            # Initialize a hashtable to hold the extension attributes
            $mgdevice = @{}
        
            # Loop through the hashtable to find the extension attributes
            foreach ($key in $response.Keys) {
                if (($key -match "extensionAttribute*") -or ($key -match "displayName") -or ($key -match "deviceId")) {
                    # Add the key and value to the extensionAttributes hashtable
                    $mgdevice[$key] = $response[$key]
                }
            }
            # Create a new PSCustomObject to hold the device objects EntraID extension attributes
            $mgdeviceObject = New-Object PSObject -Property @{
                'displayName' = $mgdevice.displayName
                'deviceID' = $mgdevice.deviceId
                'ExtensionAttribute1' = $mgdevice.values.extensionattribute1
                'ExtensionAttribute2' = $mgdevice.values.extensionattribute2
                'ExtensionAttribute3' = $mgdevice.values.extensionattribute3
                'ExtensionAttribute4' = $mgdevice.values.extensionattribute4
            }
         

            # Call the function to update extension attributes
            Set-EntraIDExtensionAttributes -AzureID $AzureID `
                                            -Office $Office `
                                            -Department $Department `
                                            -Devicetype $Devicetype `
                                            -Infrastructure $Infrastructure `
                                            -computername $match.displayName `
                                            -mgdeviceObject $mgdeviceObject
        }#If $response
    }#foreach $computer in $AzureObject
}#foreach $computer in $ADComputerList

Stop-Transcript
