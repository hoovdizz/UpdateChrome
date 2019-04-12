#Created to check and download latest Chrome
#Checks Live Version vs what is deployed, then downloads and email notification if newer version is found
#Creation date : 4-2-2019
#Creator: Alix N Hoover




#Variables to your software Share
$SCCMSource = '\\chsccm02\Software\Google\Chrome'

#Variables-Mail
$MailServer = "CHMAIL01"
$recip = "NetworkEngineers@lyco.org"
$sender = "Powershell@lyco.org"
$subject = "Chrome Update"
#where you put your schedule task
$ServerName = "CHSCCM02" 
#where you put your documentation on deploying files
$doc='onenote:///L:\Documentation\$Lycoming%20OneNote\SCCM.one#Deploying%20Chrome%20Update&section-id={B839BB76-97C7-4482-A067-D2408C04CB35}&page-id={BB768083-B904-48AF-BB7D-32DA865C8BAE}&end'



#weblinks
$uricheck = 'http://feeds.feedburner.com/GoogleChromeReleases'
$URIX86 = 'https://dl.google.com/edgedl/chrome/install/GoogleChromeStandaloneEnterprise.msi'
$URI = 'https://dl.google.com/edgedl/chrome/install/GoogleChromeStandaloneEnterprise64.msi'


#Check Current Version

           [xml]$strReleaseFeed = Invoke-webRequest $uricheck -UseBasicParsing
           [string]$versioncheck = ($strReleaseFeed.feed.entry | Where-object{$_.title.'#text' -match 'Stable'}).content | Select-Object{$_.'#text'} | Where-Object{$_ -match 'Windows'} | ForEach{[version](($_ | Select-string -allmatches '(\d{1,4}\.){3}(\d{1,4})').matches | select-object -first 1 -expandProperty Value)} | Sort-Object -Descending | Select-Object -first 1
            $versioncheck
            $checkfolder = "$SCCMSource\$versioncheck"
IF (!(test-path $checkfolder)) {


#Variables-64
$OutFile = 'GoogleChromeStandaloneEnterprise64.msi'
$OutFile = "$SCCMSource\$OutFile"

#Variables-86
$OutFileX86 = 'GoogleChromeStandaloneEnterprise.msi'
$OutFileX86 = "$SCCMSource\$OutFileX86"

# Download Chrome from the web
Write-Output "Downloading $URI to $OutFile"
$start_time = Get-Date
Invoke-WebRequest -Uri $URI -OutFile $OutFile
Write-Output "Download completed in: $((Get-Date).Subtract($start_time).Seconds) second(s)"

# Download Chromex86 from the web
Write-Output "Downloading $URIX86 to $OutFileX86"
$start_time = Get-Date
Invoke-WebRequest -Uri $URIX86 -OutFile $OutFileX86
Write-Output "Download completed in: $((Get-Date).Subtract($start_time).Seconds) second(s)"





# Get file metadata
$a = 0 
$objShell = New-Object -ComObject Shell.Application 
$objFolder = $objShell.namespace((Get-Item $OutFile).DirectoryName) 

foreach ($File in $objFolder.items()) {
    IF ($file.path -eq $outfile) {
        $FileMetaData = New-Object PSOBJECT 
        for ($a ; $a  -le 266; $a++) {  
         if($objFolder.getDetailsOf($File, $a)) { 
             $hash += @{$($objFolder.getDetailsOf($objFolder.items, $a)) = $($objFolder.getDetailsOf($File, $a)) }
            $FileMetaData | Add-Member $hash 
            $hash.clear()  
           } #end if 
       } #end for  
    }
}


# Get file metadata x86
$a = 0 
$objShellX86 = New-Object -ComObject Shell.Application 
$objFolderX86 = $objShellX86.namespace((Get-Item $OutFileX86).DirectoryName) 

foreach ($FileX86 in $objFolderX86.items()) {
    IF ($fileX86.path -eq $outfileX86) {
        $FileMetaDataX86 = New-Object PSOBJECT 
        for ($aX86 ; $aX86  -le 266; $aX86++) {  
         if($objFolderX86.getDetailsOf($FileX86, $aX86)) { 
             $hashX86 += @{$($objFolderX86.getDetailsOf($objFolderX86.items, $aX86)) = $($objFolderX86.getDetailsOf($FileX86, $aX86)) }
            $FileMetaDataX86 | Add-Member $hashX86 
            $hashX86.clear()  
           } #end if 
       } #end for  
    }
}


# Move the downloaded file to the appropriate location
$ChromeVersion = $FileMetaData.Comments.split(' ')[0]
$Filename = $((get-item $OutFile).name)
$destinationfolder = "$SCCMSource\$chromeversion"

Write-Output "Downloaded version: $ChromeVersion"
Write-Output "Destination folder is $destinationfolder"


[System.IO.Directory]::CreateDirectory($destinationfolder)
Write-Output "Creating $destinationfolder"
[System.IO.File]::Move($OutFile,"$destinationfolder\$Filename")
Write-Output "Moving $OutFile to $destinationfolder"




# Move the downloaded file to the appropriate location -x86
$ChromeVersionX86 = $FileMetaDataX86.Comments.split(' ')[0]
$FilenameX86 = $((get-item $OutFileX86).name)
$destinationfolderX86 = "$SCCMSource\$chromeversionX86"
[System.IO.File]::Move($OutFileX86,"$destinationfolder\$FilenameX86")
Write-Output "Moving $OutFileX86 to $destinationfolderX86"

#sendemail

$body ="<html></body> <BR> Chrome Version <p style='color:#FF0000'>  $ChromeVersion </p> is ready for deployment Via SCCM  <BR>"

$body+= "<a href=$doc>Here are Directions</a> "
$body+="<BR> this is a Scheduled task on $ServerName"
 
Send-MailMessage -From $sender -To $recip -Subject $subject -Body ( $Body | out-string ) -BodyAsHtml -SmtpServer $MailServer

}

   ELSE { Write-Output "$checkfolder already exists"}