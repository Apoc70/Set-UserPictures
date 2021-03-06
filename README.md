# Set-UserPictures.ps1

This script fetches images from a source folder, resizes images for use with Exchange, Active Directory and Intranet. Resized images are are written to Exchange and Active Directory.

## Description

The script parses all images provided in a dedicated source folder. The images are resized and cropped for use with the following targets:

* Exchange Server user photo
* Active Directory thumbnailPhoto attribute
* Local intranet use in your local infrastructure

Source images must be named using the respective user logon name. 

Example: MYDOMAIN\JohnDoe --> JohnDoe.jpg

Preferably, the images are stored in jpg format.

Optionally, processed image files are moved from the respective folder (Exchange, AD, Intranet) to a processed folder.

## Requirements

* GlobalFunctions PowerShell module, described here: [http://scripts.granikos.eu](http://scripts.granikos.eu)
* ResizeImage.exe executable, described here: [https://www.granikos.eu/en/justcantgetenough/PostId/307/add-resized-user-photos-automatically](https://www.granikos.eu/en/justcantgetenough/PostId/307/add-resized-user-photos-automatically)
* Exchange Server 2013+ Management Shell (EMS) for storing user photos in on-premises mailboxes
* Exchange Online Management Shell for storing user photos in cloud mailboxes
* Write access to thumbnailPhoto attribute in Active Directory

## Parameters

### PictureSource

Absolute path to source images
Filenames must match the logon name of the user

### TargetPathAD

Absolute path to store images resized for Active Directory

### TargetPathExchange

Absolute path to store images resized for Exchange

### TargetPathIntranet

Absolute path to store images resized for Intranet

### ExchangeOnPrem

Switch to create resized images for Exchange On-Premesis and store images in users mailbox
Requires the image tool to be available in TargetPathExchange

### ExchangeOnline

Switch to create resized images for Exchange Online and store images in users mailbox
Requires the image tool to be available in TargetPathExchange

### ExchangeOnPremisesFqdn

Name the on-premises Exchange Server Fqdn to connect to using remote PowerShell

### ActiveDirectory

Switch to create resized images for Active Directory and store images in users thumbnailPhoto attribute
Requires the image tool to be available in TargetPathAD

### Intranet

Switch to create resized images for Intranet
Requires the image tool to be available in TargetPathIntranet

### SaveUserStatus

Switch to save a last modified status in a local Xml file. Currently in development.

### MoveAction

Optional action to move processed images to a dedicated sub folder.

Possible values:

* MoveTargetToProcessed = Move Exchange, AD or Intranet pictures to a subfolder
* MoveSourceToProcessed = Move image source to a subfolder

## Examples

``` PowerShell
.\Set-UserPictures.ps1 -ExchangeOnPrem
```

Resize photos stored in the default PictureSource folder for Exchange (648x648) and write images to user mailboxes

``` PowerShell
.\Set-UserPictures.ps1 -ExchangeOnline -PictureSource '\\SRV01\HRShare\Photos' -TargetPathExchange '\\SRV02\ExScripts\Photos'
```

Resize photos stored on a SRV01 share for Exchange Online and save resized photos on a SRV02 share

``` PowerShell
.\Set-UserPictures.ps1 -ActiveDirectory
```

Resize photos stored in the default PictureSource folder for Active Directory (96x96) and write images to user thumbnailPhoto attribute

``` PowerShell
.\Set-UserPictures.ps1 -Intranet
```

Resize photos stored in the default PictureSource folder for Intranet (150x150)

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)