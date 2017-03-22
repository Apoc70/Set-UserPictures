<# 
    .SYNOPSIS 
    This script 

    Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Version 1.0, 2017-xx-xx

    Please send ideas, comments and suggestions to support@granikos.eu 

    .LINK 
    http://scripts.granikos.eu


    .DESCRIPTION 
    The script 


    .NOTES 
    Requirements 
    - 

    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0      Initial release 

    This PowerShell script has been developed using ISESteroids - www.powertheshell.com 


    .PARAMETER 

    .EXAMPLE
     

#>
[CmdletBinding()]
param(
  [string]$PictureSource="D:\UserPictures\SOURCE",
  [string]$TargetPathAD = 'D:\UserPictures\AD',
  [string]$TargetPathExchange = 'D:\UserPictures\Exchange',
  [string]$TargetPathIntranet = 'D:\UserPictures\Intranet',
  [switch]$Exchange,
  [switch]$Intranet,
  [switch]$ActiveDirectory,
  [switch]$SaveUserStatus,
  [switch]$MoveToProcessedFolder
)


# IMPORT GLOBAL MODULE AND SET INITIAL VARIABLES

# Import-Module GlobalFunctions
Import-Module BDRFunctions

$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write('Script started')

$FileFilter = '*.jpg'
$ImageResizeTool = 'ResizeImage.exe'
$UserStatusXml = 'UserStatus.xml'

### BEGIN Variables -----------------------------------------------------------

<# For update ProfilPicture_ChangeCounter in DB
  For future development
  $SQLServer = 'mcsmdeSQL18'
  $SQLDB = 'inhouse'
  $SQLUser = 'mcsm\srv-ADPictureImport'
  $SQLPassword = 'xXxXxXxXxXx'
  $ConnectionString = "Server=$SQLServer; Uid=$SQLUser; Pwd=$SQLPassword; Database=$SQLDB; Trusted_Connection=True;"
#>

### END Variables -------------------------------------------------------------

function Set-UserStatus { 
param(
  [string]$User = ''
)
  if($User -ne '') { 
    $XmlPath = Join-Path -Path $ScriptDir -ChildPath $UserStatusXml
    [xml]$xml=[xml](Get-Content -Path $XmlPath)

    $UserNode = $xml.Data.Users.User | Where-Object {$_.Name -eq $User}
    $TimeStamp = Get-Date -Format s

    if($UserNode -eq $null) {
      # Append new node
      $newNode = $xml.CreateElement("User")
      $newNode.SetAttribute("Name",$User.ToUpper())
      $newNode.SetAttribute("LastUpdated",$TimeStamp)
      $xml.SelectSingleNode('/Data/Users').AppendChild($newNode) | Out-Null
    }
    else {
      # Update node
      $UserNode.LastUpdated = $TimeStamp.ToString()
    }

    $xml.Save($XmlPath) | Out-Null
  }
}

function Convert-ToTargetPicture {
[CmdletBinding()]
param(
  [string]$SourcePath = '',
  [string]$TargetPath = ''
)
  if(($SourcePath -ne '') -and ($TargetPath -ne '')) { 
    
    $cmd = Join-Path -Path $TargetPath -ChildPath $ImageResizeTool
    $logger.Write(('Executing {0} for source {1}' -f $cmd, $SourcePath))

    & $cmd $PictureSource $TargetPath
  }
}

function Move-ToProcessedFolder {
[CmdletBinding()]
param(
  [string]$SourcePath = ''
)

}

function Set-ExchangePhoto { 
[CmdletBinding()]
param(
  [string]$SourcePath = ''
)
  if($SourcePath -ne '') {

    $ExchangePictures = Get-ChildItem -Path $SourcePath -Filter $FileFilter

    if(($ExchangePictures | Measure-Object).Count -gt 0) {
    
      foreach ($file in $ExchangePictures) {
        $user = $null
        try{

          $user = Get-ADUser -Identity $file.BaseName 

        }
        catch{}

        if($user -ne $null) { 

          $Photo = ([System.IO.File]::ReadAllBytes($file.FullName))

          $logger.Write(('Set EXCH UserPhoto for {0}' -f $file.BaseName))

          Set-UserPhoto -Identity $file.BaseName -PictureData $Photo -Confirm:$false -ErrorAction SilentlyContinue | Out-Null

        }
        else {
          $logger.Write(('No AD user found for {0}' -f $file.BaseName))
        }
      }    

    }
    else {
      $logger.Write("Exchange path $($SourcePath) is empty!")
    }
  }
}

function Set-ActiveDirectoryThumbnail {
[CmdletBinding()]
param(
  [string]$SourcePath = ''
)
  if($SourcePath -ne '') { 
    $AdPictures = Get-ChildItem -Path $SourcePath -Filter $FileFilter

    if(($AdPictures | Measure-Object).Count -gt 0) {

      foreach ($file in $AdPictures) {
        $user = $null
        try{

          $user = Get-ADUser -Identity $file.BaseName 

        }
        catch{}

        if($user -ne $null) { 

          if($file.length -lt 10KB) {
            # file size is less then 10KB
            $Photo = ([System.IO.File]::ReadAllBytes($file.FullName))
            $logger.Write(('Set thumbnailPhoto for {0}' -f $file.BaseName))

            # Set-ADUser -Identity $file.BaseName -Replace @{thumbnailPhoto=$Photo}

            if($SaveUserStatus) { 
              Set-UserStatus -User $file.BaseName
            }
          }
          else {
            $logger.Write("File size for $($file.BaseName) is too large!")
          }
        }
        else {
          $logger.Write(('No AD user found for {0}' -f $file.BaseName))
        }
      }
    }
    else {
      $logger.Write("AD path $($SourcePath) is empty!")
    }
  }
}


### BEGIN Main ----------------------------------------------------------------

if(Test-Path -Path (Join-Path -Path $PictureSource -ChildPath $FileFilter) ) {

  # Fetch file information
  $Pictures = Get-ChildItem -Path $PictureSource -Filter $FileFilter
  
  $logger.Write(('Found {0} file(s)' -f ($Pictures | Measure-Object).Count))

  if($Exchange) {
    # Convert images for Exchange and push to Exchange
    Convert-ToTargetPicture -SourcePath $PictureSource -TargetPath $TargetPathExchange  

    Set-ExchangePhoto -SourcePath $TargetPathExchange
  }
  elseif($Intranet) {
    # Convert images for Intranet, convert only
    Convert-ToTargetPicture -SourcePath $PictureSource -TargetPath $TargetPathIntranet
  }
  elseif($ActiveDirectory) {
    # Convert images for Active Directory thumbnail
    Convert-ToTargetPicture -SourcePath $PictureSource -TargetPath $TargetPathAD

    Set-ActiveDirectoryThumbnail -SourcePath $TargetPathAD 

  }
}
else {
  # Ooops, source directory seems to be empty
  $logger.Write('Pictures source directory is empty!')
}


$logger.Write('Script finished')
### END Main ------------------------------------------------------------------