### .ps1 
### PowerShell-Script to resize pictures and import them into Active Directory
###
### Version 1.5:
###  - Update 1.5: Write Status in Scopeland-DB to re-import contacts in Exchange Mailboxes
###
### (c) 2014
###
### Requirements:  
###  - SQL Server Permissions
###  - Quest ActiveRoles AD Management SnapIn




### BEGIN Variables -----------------------------------------------------------

$logfile = "C:\Scripts\Profilbilder\Logs\$(get-date -format yyyy-MM-dd).log"

$rootpath="\\mcsmdeinet02\Fotos"
$imagesroot=get-childitem $rootpath *.jpg
$imageserrorpath="\\mcsmdeinet02\Profilbilder\Fehler"

$imagesproceededad="\\mcsmdeinet02\Profilbilder\AD"
$adtemppath="C:\Scripts\Profilbilder\AD\TEMP"

$imagesproceededexchange="C:\Scripts\Profilbilder\Exchange"
$exchangetemppath="\\mcsmdeinet02\Profilbilder\Exchange"

$imagesproceededintranet="C:\Scripts\Profilbilder\Intranet"
$intranettemppath="\\mcsmdeinet02\Profilbilder\Intranet"

# For update ProfilPicture_ChangeCounter in DB
$SQLServer = "mcsmdeSQL18"
$SQLDB = "inhouse"
$SQLUser = "mcsm\srv-ADPictureImport"
$SQLPassword = "xxxxx"
$ConnectionString = "Server=$SQLServer; Uid=$SQLUser; Pwd=$SQLPassword; Database=$SQLDB; Trusted_Connection=True;"

### END Variables -------------------------------------------------------------





### BEGIN Functions -----------------------------------------------------------

# Logging
Function Log
{
   Param ([string]$logstring)
   Add-content $logfile -value $logstring
}

### END Functions -------------------------------------------------------------





### BEGIN Main ----------------------------------------------------------------

# START (only if JPG exists in $rootpath)
if (Test-Path $rootpath\*.jpg)
{ 
  Write-Host " "
  Log " "
  Write-Host "New pictures found...              "
  Log "$(get-date -format hh:mm:ss) New pictures found..."
  Write-Host "============================================"
  Write-Host $imagesroot
  Log "$(get-date -format hh:mm:ss) $imagesroot"
  Write-Host " "
    

  
# Copy and convert images
  Write-Host " "
  Write-Host "Starting to convert images...      "
  Write-Host "============================================"
  
  Write-Host "Converting to Intranet Format (150 x 150)"
  Log "$(get-date -format hh:mm:ss) Converting to Intranet Format (150 x 150)"
  C:\Scripts\Profilbilder\Intranet\TEMP\icr.exe $rootpath $intranettemppath
  Write-Host " "
	
  Write-Host "Converting to AD Format (96 x 96)"
  Log "$(get-date -format hh:mm:ss) Converting to AD Format (96 x 96)"
  C:\Scripts\Profilbilder\AD\TEMP\icr.exe $rootpath $adtemppath
  Write-Host " "
	
  Write-Host "Converting to Exchange Format (648 x 648)"
  Log "$(get-date -format hh:mm:ss) Converting to Exchange Format (648 x 648)"
  C:\Scripts\Profilbilder\Exchange\TEMP\icr.exe $rootpath $exchangetemppath
  Write-Host " "
   
  

# Active Directory: Replace thumbnailPhoto in user objects with pictures from specific path
# Filenames = sAMAccountName
  $imagesactivedirectorytemp=get-childitem $adtemppath *.jpg
  Write-Host " "
  Write-Host "Starting Active Directory Import..."
  Log "$(get-date -format hh:mm:ss) Starting Active Directory Import..."
  Write-Host "============================================"
  # For every file in $imagesactivedirectorytemp
  foreach ($file in $imagesactivedirectorytemp) {
	$photo = [byte[]](Get-Content $file.FullName -Encoding byte)
	Write-Host "Importing thumbnailPhoto for User:" $file.BaseName
	Log "$(get-date -format hh:mm:ss) Importing thumbnailPhoto for User: $file"
	# Check if file is > 10 KB
	if ($file.length -gt 10KB)
	{
        Write-Host "File $file exceeds size of 10 KB, moving file to $imageserrorpath"
		Log "$(get-date -format hh:mm:ss) File $file exceeds size of 10 KB, moving file to $imageserrorpath"
		Move-Item $adtemppath\$file $imageserrorpath -force
		C:\Scripts\Profilbilder\Scripts\blat.exe EmailcontentFileSize.txt -t Profilbilder@mcsm.de -s $file -server relay.mcsm.de -f Bildimportskript@mcsm.de
    }

	else
	{
		# Check if user is in Active Directory, import and move file
		# add-pssnapin quest.activeroles.admanagement
		if ( (Get-PSSnapin -Name Quest.Activeroles.ADManagement -ErrorAction SilentlyContinue) -eq $null )
		{
			Add-PsSnapin Quest.Activeroles.ADManagement -ErrorAction SilentlyContinue
			
			if ( (Get-PSSnapin -Name Quest.Activeroles.ADManagement -ErrorAction SilentlyContinue) -eq $null )
			{
				Write-Host "Quest.Activeroles.ADManagement could NOT be loaded!" -ForegroundColor Red
				Write-Host "Verify that the Quest AD Management is installed on this computer!" -ForegroundColor Red
			}
		}
		# $user = Get-ADUser $file.BaseName | select SamAccountName
		$user = Get-QADUser $file.BaseName | select SamAccountName

		if ($user.SamAccountName -eq $file.Basename) 
		{
			#Set-ADUser $file.BaseName -Replace @{thumbnailPhoto=$photo}
			Set-QADUser $file.BaseName -ObjectAttributes @{thumbnailPhoto=$photo}
			Write-Host "Moving imported picture $file to" $imagesproceededad
			Log "$(get-date -format hh:mm:ss) Moving imported picture $file to $imagesproceededad"
			Move-Item $adtemppath\$file $imagesproceededad -force
			# Clean-Up Rootpath
			Write-Host "Removing from rootpath:" $file
			Log "$(get-date -format hh:mm:ss) Removing from rootpath: $file"
			Remove-Item $rootpath\$file
            
            # Update ProfilPicture_ChangeCounter in DB
            Write-Host
            Write-Host "Updating ProfilePicture_ChangeCounter in DB."
            Log "$(get-date -format hh:mm:ss) Updating ProfilePicture_ChangeCounter in DB."
            $Connection = New-Object System.Data.SqlClient.SqlConnection
            $Connection.ConnectionString = $ConnectionString
            $Connection.Open()
            $SQLCmd = New-Object System.Data.SQLClient.SQLCommand
            $TargetUser = $file.BaseName.ToLower()
            $SQLQuery = "UPDATE [inhouse].[dbo].[VIEW_OutlookSync_Mitarbeiter_Update] SET Profilepicture_Changecounter = Profilepicture_Changecounter + 1 WHERE lower(Username) = '$TargetUser'"
            $SQLCmd.CommandText = $SQLQuery
            $SQLCmd.Connection = $Connection
            $Result = $SQLCmd.ExecuteNonQuery()
            $Connection.Close()
		}

		else 
        {
			Write-Host "User for $file does not exist in AD, moving file to $imageserrorpath"
			Log "$(get-date -format hh:mm:ss) User for $file does not exist in AD, moving file to $imageserrorpath"
			Move-Item $rootpath\$file $imageserrorpath -force
			Remove-Item $adtemppath\$file
			#C:\Scripts\Profilbilder\Scripts\blat.exe EmailcontentADuser.txt -t Profilbilder@mcsm.de -s $file -server relay.mcsm.de -f Bildimportskript@mcsm.de
			C:\Scripts\Profilbilder\Scripts\blat.exe EmailcontentADuser.txt -t Profilbilder@mcsm.de -s $file -server smtp.mcsm.de -f alert-ADPictureImport@mcsm.de -u srv-smtp-ADPictureImport@mcsm.de -pw fPFsKBDSz9dr2W8w67qk
		} 

	}

  }
  
    
# END of global if
}


else
{
    Write-Host " "
    Write-Host "No new pictures found in" $rootpath
    Log "$(get-date -format hh:mm:ss) No new pictures found in $rootpath"
    Write-Host " "
}

### END Main ------------------------------------------------------------------