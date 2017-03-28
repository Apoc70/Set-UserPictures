<# 
    .SYNOPSIS 
    This script fetches images from a source folder, resizes images for use with Exchange, 
    Active Directory and Intranet. Resized images are are written to Exchange and Active Directory.

    Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Version 1.0, 2017-03-23

    Please send ideas, comments and suggestions to support@granikos.eu 

    .LINK 
    http://scripts.granikos.eu

    .DESCRIPTION 
    The script parses all images provided in a dedicated source folder. The images are resized and cropped 
    for use with the following targets:
    - Exchange Server user photo
    - Active Directory thumbnailPhoto attribute
    - Local intranet use in your local infrastructure

    Source images must be named using the respective user logon name. 
    Example: MYDOMAIN\JohnDoe --> JohnDoe.jpg

    Preferably, the images are stored in jpg format.

    Optionally, processed image files are moved from the respective folder (Exchange, AD, Intranet) to a
    processed folder.

    .NOTES 
    Requirements 
    - GlobalFunctions PowerShell module, described here: http://scripts.granikos.eu 
    - ResizeImage.exe executable, described here: 
    - Exchange Server 2013+ Management Shell (EMS) for storing user photos in on-premises mailboxes
    - Exchange Online Management Shell for storing user photos in cloud mailboxes 
    - Write access to thumbnailPhoto attribute in Active Directory 

    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0      Initial release 

    This PowerShell script has been developed using ISESteroids - www.powertheshell.com 

  
    .PARAMETER PictureSource
    Absolute path to source images
    Filenames must match the logon name of the user

    .PARAMETER TargetPathAD
    Absolute path to store images resized for Active Directory

    .PARAMETER TargetPathExchange
    Absolute path to store images resized for Exchange

    .PARAMETER TargetPathIntranet
    Absolute path to store images resized for Intranet

    .PARAMETER Exchange
    Switch to create resized images for Exchange and store images in users mailbox
    Requires the image tool to be available in TargetPathExchange

    .PARAMETER ActiveDirectory
    Switch to create resized images for Active Directory and store images in users thumbnailPhoto attribute
    Requires the image tool to be available in TargetPathAD

    .PARAMETER Intranet
    Switch to create resized images for Intranet
    Requires the image tool to be available in TargetPathIntranet

    .PARAMETER SaveUserStatus
    Switch to save a last modified status in a local Xml file. Currently in development.

    .PARAMETER MoveAction
    Optional action to move processed images to a dedicated sub folder.
    Possible value:
    MoveTargetToProcessed = Move Exchange, AD or Intranet pictures to a subfolder
    MoveSourceToProcessed = Move image source to a subfolder

    .EXAMPLE
    Resize photos stored in the default PictureSource folder for Exchange (648x648) and write images to user mailboxes
    .\Set-UserPictures.ps1 -Exchange   

    .EXAMPLE
    Resize photos stored on a SRV01 share for Exchange and save resized photos on a SRV02 share
    .\Set-UserPictures.ps1 -Exchange -PictureSource '\\SRV01\HRShare\Photos' -TargetPathExchange '\\SRV02\ExScripts\Photos'

    .EXAMPLE
    Resize photos stored in the default PictureSource folder for Active Directory (96x96) and write images to user thumbnailPhoto attribute
    .\Set-UserPictures.ps1 -ActiveDirectory

    .EXAMPLE
    Resize photos stored in the default PictureSource folder for Intranet (150x150)
    .\Set-UserPictures.ps1 -Intranet
#>
[CmdletBinding()]
param(
  [string]$PictureSource='D:\UserPhotos\SOURCE',
  [string]$TargetPathAD = 'D:\UserPhotos\AD',
  [string]$TargetPathExchange = 'D:\UserPhotos\Exchange',
  [string]$TargetPathIntranet = 'D:\UserPhotos\Intranet',
  [switch]$Exchange,
  [switch]$Intranet,
  [switch]$ActiveDirectory,
  [switch]$SaveUserStatus,
  [ValidateSet('MoveTargetToProcessed','MoveSourceToProcessed')]
  [string]$MoveAction
)


# IMPORT GLOBAL MODULE AND SET INITIAL VARIABLES
try { 
  Import-Module -Name GlobalFunctions
}
catch {
  # Loading GlobalFUnction module failed. 
  Write-Warning -Message 'GlobalFunctions module could not be loaded!'
  Write-Warning -Message 'Please see: http://scripts.granikos.eu'
  exit 99
}

$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write('Script started')

# File filter for resizing and importing images
$FileFilter = '*.jpg'

# Tool used for image resizing
# It is assumed that the ImageResizeTool is located in EACH target directory
$ImageResizeTool = 'ResizeImage.exe'

# Tracking of user status (last modified), open issue #2
$UserStatusXml = 'UserStatus.xml'

# folder name for processed images
$ProcessedFolderName = 'processed'

### BEGIN Variables -----------------------------------------------------------

<# 
    For update ProfilPicture_ChangeCounter in DB?
    https://github.com/Apoc70/Set-UserPictures/issues/2

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
      $newNode = $xml.CreateElement('User')
      $newNode.SetAttribute('Name',$User.ToUpper())
      $newNode.SetAttribute('LastUpdated',$TimeStamp)
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
    [string]$SourcePath = '',
    [string]$Filename =''
  )
  $ProcessedPath = Join-Path -Path $SourcePath -ChildPath $ProcessedFolderName
  if(!(Test-Path -Path $ProcessedPath)) { 
    # Create processed target folder first
    $null = New-Item -Path $ProcessedPath -ItemType Directory
    $logger.Write(('Directory created: {0}' -f ($ProcessedPath)))
  }

  # if SourcePath AND Filename are set, move a single file
  if(($SourcePath -ne '') -and ($Filename -ne '')){
    # move file to processed folder
    $null = Move-Item -Path (Join-Path -Path $SourcePath -ChildPath $Filename) -Destination (Join-Path -Path $ProcessedPath -ChildPath $Filename) -Force
    $logger.Write(('Moved {0} from {1} to {2}' -f $Filename, $SourcePath, $ProcessedPath))
  }
  elseif($SourcePath -ne '') {
    # Move full directory to processed folder
    $null = Move-Item -Path (Join-Path -Path $SourcePath -ChildPath '*') -Destination (Join-Path -Path $ProcessedPath -ChildPath '*') -Force
    $logger.Write(('Moved all files from {0} to {1}' -f $SourcePath, $ProcessedPath))
  }
}

function Set-ExchangePhoto { 
  [CmdletBinding()]
  param(
    [string]$SourcePath = ''
  )
  if($SourcePath -ne '') {

    # fetch all files from Exchange directory
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

          Set-UserPhoto -Identity $user.UserPrincipalName -PictureData $Photo -Confirm:$false -ErrorAction SilentlyContinue | Out-Null

          if($MoveToProcessedFolder) {
            Move-ToProcessedFolder -SourcePath $SourcePath -Filename $file.Name
          }

        }
        else {
          $logger.Write(('No AD user found for {0}' -f $file.BaseName), 2)
        }
      }    
    }
    else {
      # Source path is empty
      $logger.Write(('Exchange path {0} is empty!' -f $SourcePath))
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

            if($MoveToProcessedFolder) {
              Move-ToProcessedFolder -SourcePath $SourcePath -Filename $file.Name
            }

            if($SaveUserStatus) { 
              Set-UserStatus -User $file.BaseName
            }
          }
          else {
            # File size, too large
            # Open issue #1
            $logger.Write(('File size for {0} is too large!' -f $file.BaseName), 1)
          }
        }
        else {
          # Ooops, we haven't found an AD user object
          $logger.Write(('No AD user found for {0}' -f $file.BaseName), 2)
        }
      }
    }
    else {
      # Source path is empty
      $logger.Write(('AD path {0} is empty!' -f $SourcePath))
    }
  }
}


### BEGIN Main ----------------------------------------------------------------

if(Test-Path -Path (Join-Path -Path $PictureSource -ChildPath $FileFilter) ) {

  $MoveToProcessedFolder = $false
  $MoveSourceToProcessedFolder = $false

  switch($MoveAction) {
    'MoveTargetToProcessed' { $MoveToProcessedFolder = $true }
    'MoveSourceToProcessed' { $MoveSourceToProcessedFolder = $true }
  }  

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

  # Do we need to move source images to processed?
  if($MoveSourceToProcessedFolder) {
    Move-ToProcessedFolder -SourcePath $PictureSource
  }

}
else {
  # Ooops, source directory seems to be empty
  $logger.Write('Pictures source directory is empty!')
}

$logger.Write('Script finished')
### END Main ------------------------------------------------------------------