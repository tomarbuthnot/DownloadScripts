<#  
.SYNOPSIS  
   	Downloads Teched MP4's and PPTX from channel 9 site, Based on input CSV

.DESCRIPTION  
    Downloads Teched MP4's and PPTX from channel 9 site, Based on input CSV.
    CSV format Code,Name, Where Code is session code and Name is the name you want to call the download file.

.NOTES  
    Modified for Teched 2013 NA by Tom Arbuthnot lyncdup.com
    Original provided by blog.SCOMfaq.ch / Stefan Roth
    credit http://blog.scomfaq.ch/2012/06/13/teched-2012-orlando-download-sessions-offline-viewing/
    Credit: Pat Richard for New-Download Function http://www.ehloworld.com

    Microsoft handily use a simple format of URL/SessionID.MP4/PPTX, so in future techeds this method will also likely
    work
    
    Use completely at your own risk

.LINK  
	www.lyncdup.com

.EXAMPLE
    .\Download-TechEdSessions.ps1 -SessionCSV .\LyncSessionsTechedNA2013.csv -DownloadTargetDirectory C:\Downloads

.INPUTS
    None. You cannot pipe objects to this script.
#>

Param(
        [parameter(mandatory=$true)]
        [String]$SessionCSV,

        [parameter(mandatory=$true)]
        [String]$DownloadTargetDirectory
        
        )
Write-host " "
Write-host " "
Write-host " "
Write-host " "
Write-host " "
Write-host " "
Write-host " "
Write-host " "
Write-Host "###############################################################################################"  -ForegroundColor Yellow
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "                    Download-TechEdSessions by Tom Arbuthnot http://lyncdup.com                "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "            Credit: Original 2012 downloader by http://blog.scomfaq.ch (c) 2012 Stefan Roth   "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "                 Credit: Pat Richard for New-Download Function http://www.ehloworld.com/       "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "###############################################################################################"  -ForegroundColor Yellow

function Set-ModuleStatus { 
	[CmdletBinding(SupportsShouldProcess = $True)]
	param	(
		[parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true, HelpMessage = "No module name specified!")] 
		[string]$name
	)
	if(!(Get-Module -name "$name")) { 
		if(Get-Module -ListAvailable | ? {$_.name -eq "$name"}) { 
			Import-Module -Name "$name" 
			# module was imported
			return $true
		} else {
			# module was not available
			return $false
		}
	}else {
		# module was already imported
		# Write-Host "$name module already imported"
		return $true
	}
} # end function Set-ModuleStatus

function New-FileDownload {
	[CmdletBinding(SupportsShouldProcess = $True)]
	param(
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
		[ValidateNotNullOrEmpty()]
		[string]$SourceFile,
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
		[string]$DestFolder,
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
		[string]$DestFile
	)
	[bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
	# I should switch this to using a param block and pipelining from property name - just for consistency
	# I should clean up the display text to be consistent with other functions
	$error.clear()
	if (!($DestFolder)){$DestFolder = $TargetFolder}
	Set-ModuleStatus -name BitsTransfer
	if (!($DestFile)){[string] $DestFile = $SourceFile.Substring($SourceFile.LastIndexOf("/") + 1)}
	if (Test-Path $DestFolder){
		Write-Host "Folder: `"$DestFolder`" exists."
	} else{
		Write-Host "Folder: `"$DestFolder`" does not exist, creating..." -NoNewline
		New-Item $DestFolder -type Directory | Out-Null
		Write-Host "Done! " -ForegroundColor Green
	}
	if (Test-Path "$DestFolder\$DestFile"){
		Write-Host "File: $DestFile exists."
	}else{
		if ($HasInternetAccess){
			Write-Host "File: $DestFile does not exist, downloading..." -NoNewLine
			Start-BitsTransfer -Source "$SourceFile" -Destination "$DestFolder\$DestFile"
			Write-Host "Done! " -ForegroundColor Green
		}else{
			Write-Host "Internet access not detected. Please resolve and try again." -ForegroundColor red
		}
	}
} # end function New-FileDownload

$sessions = Import-csv $SessionCSV

#Downloading the files from here. this changes depending on year/event
$url="http://video.ch9.ms/sessions/teched/na/2013/"

$TotalDownloads =  $($sessions.count) * 2
$DownloadTargetDirectory =

Foreach($session in $sessions){

$i = $i+1


try {

            # Download MP4

            write-host -ForegroundColor Yellow ("Downloading Session MP4..." + $($session.Code) + " Number: " + $i + " out of " + $TotalDownloads)
            $MP4ToDownload=$url+"$($session.Code)"+".mp4"

            $DestinationFileMP4 = "$($session.Code)" + " " + "$($session.Name)" + ".mp4"

            New-FileDownload -SourceFile $MP4ToDownload -DestFolder $DownloadTargetDirectory -DestFile $DestinationFileMP4

            # Download PPTX (some not available)
            
            write-host -ForegroundColor Yellow ("Downloading Session PPTX..." + $($session.Code) + " Number: " + $i + " out of " + $TotalDownloads)
            $PPTXToDownload=$url+"$($session.Code)"+".pptx"

            $DestinationFilePPTX = "$($session.Code)" + " " + "$($session.Name)" + ".pptx"

            New-FileDownload -SourceFile $PPTXToDownload -DestFolder $DownloadTargetDirectory -DestFile $DestinationFilePPTX
        
     } 
     
catch {
            
            $errorfile = $scriptdir + "\sessions_notavailable.txt";
            write-host -foregroundColor Red ("Session not available please check $errorfile")
            if(!(test-path -path $errorfile)) { new-item $errorfile -type file};
            $session | out-file $errorfile -Append

        }



}

write-host -ForegroundColor Green "Finished!"

