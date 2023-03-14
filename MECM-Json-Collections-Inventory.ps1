<#
.SYNOPSIS
This script retrieves information about MECM (Microsoft Endpoint Configuration Manager) collections and saves it in a JSON file. 
The collections are sorted by folders and each folder contains information about the collections it contains. 
The JSON file is created or overwritten if it already exists.

.PARAMETER SiteCode
The MECM site code.

.PARAMETER ServerName
The name of the MECM server.

.PARAMETER JsonPath
The path to the JSON file to create or overwrite.

.EXAMPLE
powershell.exe -ExecutionPolicy bypass -file "MECM-Json-Collections-Inventory.ps1" -SiteCode "AAA" -ServerName "myserver.aaa.local" -JsonPath "C:\users\%USERNAME%\documents\export.json"

.AUTHOR Camille POIROT
https://www.linkedin.com/in/camille-poirot-85919621b/
#>


[CmdletBinding()]
Param (

    [Parameter(Mandatory = $True)]
    [String]$SiteCode ,

    [Parameter(Mandatory = $true)]
    [String]$ServerName,

    [Parameter(Mandatory = $true)]
    [string]$JsonPath # exemple : "C:\Users\$env:USERNAME\Documents\test.json"
)

# MECM connexion
try {

        # Customizations
        $initParams = @{}

        # Import the ConfigurationManager.psd1 module 
        if((Get-Module ConfigurationManager) -eq $null) {
            Import-Module "C:\ProgramData\*\MECM\BIN\ConfigurationManager.psd1" @initParams 
        }

        # Connect to the site's drive if it is not already present
        if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ServerName @initParams
        }

        # Set the current location to be the site code.
        Set-Location "$($SiteCode):\" @initParams


} catch { exit.code(1) }


[string]$Namespace = "root/SMS/site_$SiteCode"

if(Test-Path $JsonPath ){ Remove-Item -path $JsonPath -Force -Recurse ; New-Item -Path $JsonPath}
else{ New-Item -Path $JsonPath }


# Retourne les Path des collections USER (1) ou Device (2)
function Get-FoldersPathFromCollections{

                    param( [string]$CollectionType) # 1 USER - 2 DEVICE
                    $SCCMCollectionQuery ="select ObjectPath from SMS_Collection where CollectionType=$CollectionType"
                    $Collections_ObjectPath = (Get-WmiObject -Namespace $Namespace -Query $SCCMCollectionQuery -ComputerName $ServerName ).ObjectPath
                    return $Collections_ObjectPath
}#Get-FoldersPathFromCollections

function Get-CollectionsInFolderID ( $FolderID ){

        $SCCMCollectionQuery ="select CollectionID,Name from SMS_Collection where CollectionID is in(select InstanceKey from SMS_ObjectContainerItem `
                                where ContainerNodeID='$FolderID')"
        $CollectionInfos =  Get-WmiObject -Namespace $Namespace -Query $SCCMCollectionQuery -ComputerName $ServerName
        return $CollectionInfos


}#Get-CollectionInFolderID

function Get-CollectionINFOS{

        param(
        [string]$CollectionID,
        [string]$CollectionType
        )

        $return = Get-CMCollection -CollectionType $CollectionType -Id $CollectionID | select Name,CollectionID,CollectionRules,MemberCount,LimitToCollectionName,LimitToCollectionID,RefreshSchedule,Comment
        try {$deployement = (Get-CMDeployment -CollectionName $return.Name | select ApplicationName).ApplicationName } catch{}
        if($deployement){ $return | Add-Member -MemberType NoteProperty -Name "Deployment" -Value $deployement }
        return $return

}#Get-CollectionINFOS

function GetFolderContainerNodeID{

        param(
            [Parameter(Mandatory = $True)]
            [ValidateSet('User', 'Device')]
            [String]$Type,
            [Parameter(Mandatory = $True)]
            $folderName
            )
                       
            $ObjectTypeName = "SMS_Collection_$Type"
            $query = "SELECT * FROM SMS_ObjectContainerNode WHERE Name='$folderName' and ObjectTypeName='$ObjectTypeName'"
            $folder = Get-WmiObject -Namespace $Namespace -Query $Query -ComputerName $ServerName
            $ContainerNodeID = $folder.ContainerNodeID
            return $ContainerNodeID

}#GetFolderContainerNodeID

function Set-TreeInJsonFile{
            param(
                [Parameter(Mandatory = $True)]
                [ValidateSet('User', 'Device')]
                [String]$Type = 'User',
                [Parameter(Mandatory = $True)]
                $allObjectPath,
                [Parameter(Mandatory = $True)]
                $JsonPath
                )

                $Collections = New-Object -TypeName PSObject
                $Collections = Get-Content -Path "$JsonPath" | ConvertFrom-Json 
                [string]$CollectionType = $Type+"_Collections" ; echo $CollectionType

                if(-not($Collections)){ break }

                # Trier la liste
                $comparer = [System.Collections.Comparer]::Default
                $allObjectPath.Sort($comparer)
                $allObjectPath | ForEach-Object {
        
                       # NewEntryInJsonByPath($_)
                        [string]$path = $_
                        [string]$path = $path.Substring(1).replace("/",".") 
  
                        if(-not($path.Split(".")[1])){
        
                               try { 
                                        [string]$Name = $path
                                        $Collections.$CollectionType | Add-Member -MemberType NoteProperty  -Name $Name -Value (New-Object -TypeName PSObject) -Force -ErrorAction SilentlyContinue
                                        $Collections.$CollectionType.$Name | Add-Member -MemberType NoteProperty  -Name "ObjectType" -Value "Folder" -Force -ErrorAction SilentlyContinue
                                        [string]$ContainerNodeID = GetFolderContainerNodeID -Type $type -folderName $name
                                        $Collections.$CollectionType.$Name | Add-Member -MemberType NoteProperty  -Name "ContainerNodeID" -Value "$ContainerNodeID" -Force -ErrorAction SilentlyContinue

                                        $IDCollectionsInFolder = Get-CollectionsInFolderID($ContainerNodeID) 
                                        $IDCollectionsInFolder | % {

                                                                    $val = Get-CollectionINFOS -CollectionID $_.CollectionID -CollectionType $type
                                                                    $CollectionName
                                                                    $val
                                                                    [string]$CollectionName = $_.Name
                                                                    echo $Collections.$CollectionType.$Name
                                                                    $Collections.$CollectionType.$Name |  Add-Member -MemberType NoteProperty -Name "$CollectionName" -Value $val

                                        }

                                    }catch{}
                                }
                        else{
                               try {
                                        [string]$name = ($path.Split(".")[1] )#.replace(" ", "_")
                                        [string]$Path = $path#.replace(" ", "_") 
                                        [string]$Path = $path.Split(".")[0]
                                        $Collections.$CollectionType.$Path | Add-Member -MemberType NoteProperty  -Name "$name" -Value (New-Object -TypeName PSObject) -Force -ErrorAction SilentlyContinue
                                        $Collections.$CollectionType.$Path.$Name | Add-Member -MemberType NoteProperty  -Name "ObjectType" -Value "Folder" -Force -ErrorAction SilentlyContinue
                                        [string]$ContainerNodeID = GetFolderContainerNodeID -Type $type -folderName $name
                                        $Collections.$CollectionType.$Path.$Name | Add-Member -MemberType NoteProperty  -Name "ContainerNodeID" -Value "$ContainerNodeID" -Force -ErrorAction SilentlyContinue

                                        $IDCollectionsInFolder = Get-CollectionsInFolderID($ContainerNodeID) 
                                        $IDCollectionsInFolder | % {

                                                                    $val = Get-CollectionINFOS -CollectionID $_.CollectionID -CollectionType $type
                                                                    $val
                                                                    [string]$CollectionName = $_.Name
                                                                    $Collections.$CollectionType.$Path.$Name |  Add-Member -MemberType NoteProperty -Name "$CollectionName" -Value $val

                                        }

                                    }catch{}
                            
                        }

                }
                $Collections |  ConvertTo-Json -Depth 5 | Out-File -FilePath $JsonPath
}#


$json = New-Object -TypeName PSObject
$json | Add-Member -MemberType NoteProperty  -Name "User_Collections" -Value (New-Object -TypeName PSObject)
$json.User_Collections | Add-Member -MemberType NoteProperty  -Name "ObjectType" -Value "Folder" -Force -ErrorAction SilentlyContinue
$json | Add-Member -MemberType NoteProperty  -Name "Device_Collections" -Value (New-Object -TypeName PSObject)
$json.Device_Collections | Add-Member -MemberType NoteProperty  -Name "ObjectType" -Value "Folder" -Force -ErrorAction SilentlyContinue


$ObjectPath = "/"
$SCCMCollectionQuery ="select * from SMS_Collection where ObjectPath = '$ObjectPath' and CollectionType='1'"
$CollectionsUser = Get-WmiObject -Namespace $Namespace -Query $SCCMCollectionQuery -ComputerName $ServerName 

$ObjectPath = "/"
$SCCMCollectionQuery ="select * from SMS_Collection where ObjectPath = '$ObjectPath' and CollectionType='2'"
$CollectionsDevice = Get-WmiObject -Namespace $Namespace -Query $SCCMCollectionQuery -ComputerName $ServerName 

$CollectionsUser |  %{
                    $_.CollectionID
                    $CollectionsInfos = Get-CollectionINFOS -CollectionID $_.CollectionID -CollectionType "User"
                    $CollectionsInfos
                    [string]$CollectionName = $CollectionsInfos.Name
                    $CollectionName
                    try { $json.User_Collections | Add-Member -MemberType NoteProperty  -Name "$CollectionName" -Value $CollectionsInfos -ErrorAction SilentlyContinue } catch {}

}
$CollectionsDevice |  %{
                    $_.CollectionID
                    $CollectionsInfos = Get-CollectionINFOS -CollectionID $_.CollectionID -CollectionType "Device"
                    $CollectionsInfos
                    [string]$CollectionName = $CollectionsInfos.Name
                    $CollectionName
                    try { $json.Device_Collections | Add-Member -MemberType NoteProperty  -Name "$CollectionName" -Value $CollectionsInfos -ErrorAction SilentlyContinue } catch {}

}

$json | ConvertTo-Json | Out-File -FilePath $JsonPath

$allObjectPath = New-Object System.Collections.ArrayList
$Collections_ObjectPath = Get-FoldersPathFromCollections("1") | %{ if($_ -notin $allObjectPath ){ $allObjectPath.add($_) >> $Null }}
Set-TreeInJsonFile -type User -allObjectPath $allObjectPath  -JsonPath $JsonPath

$allObjectPath = New-Object System.Collections.ArrayList
$Collections_ObjectPath = Get-FoldersPathFromCollections("2") | %{ if($_ -notin $allObjectPath ){ $allObjectPath.add($_) >> $Null }}
Set-TreeInJsonFile -type Device -allObjectPath $allObjectPath -JsonPath $JsonPath

