
Write-host " "
Write-Host "###############################################################################################"  -ForegroundColor Yellow
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "            Download LyncConf 2014 Videos and PPTX by Tom Arbuthnot http://lyncdup.com         "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "                 Thanks to everyone involved with LyncConf 2014 for a great show               "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "             Credit: Pat Richard for New-Download Function http://www.ehloworld.com/           "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "###############################################################################################"  -ForegroundColor Yellow


$DownloadTargetFolder = Read-Host "Please enter download target folder, e.g. C:\Tom\Downloads"

$SessionsToDownload = Read-Host "Sessions are divided by level, please choose 100,200,300,400 or All"

$TestPath = Test-Path $DownloadTargetFolder
        If ($TestPath -ne $true)
            {
            Write-Error "Your input for download path could not be verified" -ErrorAction Stop
            }


$TestPath2 = Test-Path .\LynnConf_100_URLs.txt
        If ($TestPath2 -ne $true)
            {
            Write-Error "This script must be run from the script directory" -ErrorAction Stop
            }

switch ($SessionsToDownload) 
        { 
        100 {
            $URLList = (Get-Content .\LynnConf_100_URLs.txt)
            # PowerShell ignores multiple backslashes so we don't need to worry about them
            $DownloadTargetFolder = $DownloadTargetFolder + "\lyncconf2014\100\"  
            } 
        200 {
            $URLList = (Get-Content .\LynnConf_200_URLs.txt)
            $DownloadTargetFolder = $DownloadTargetFolder + "\lyncconf2014\200\"
            }
        300 {
            $URLList = (Get-Content .\LynnConf_300_URLs.txt)
            $DownloadTargetFolder = $DownloadTargetFolder + "\lyncconf2014\300\"
            } 
        400 {
            $URLList = (Get-Content .\LynnConf_400_URLs.txt)
            $DownloadTargetFolder = $DownloadTargetFolder + "\lyncconf2014\400\"
            } 
        All {
            $URLList = (Get-Content .\LynnConf_100_URLs.txt)
            $URLList += (Get-Content .\LynnConf_200_URLs.txt)
            $URLList += (Get-Content .\LynnConf_300_URLs.txt)
            $URLList += (Get-Content .\LynnConf_400_URLs.txt)
            $DownloadTargetFolder = $DownloadTargetFolder + "\lyncconf2014\"
            }  
        default {
                Write-Error "Your input could not be determined, please enter 100,200,300,400 or All" -ErrorAction Stop
                }
        }



# Credit Pat Richard

# 0.4 added rety loop for 400 fail, seem to get this often on some machines

# Credit Pat Richard
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
		# return $true
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

	if (!($DestFolder)){$DestFolder = $TargetFolder}
	# Write-Host "Checking for BitsModule"
    Set-ModuleStatus -name BitsTransfer
	
    if (!($DestFile)){[string] $DestFile = $SourceFile.Substring($SourceFile.LastIndexOf("/") + 1)}
	if (Test-Path $DestFolder){
		Write-Verbose "Folder: `"$DestFolder`" exists."
	} else{
		Write-Host "Folder: `"$DestFolder`" does not exist, creating..."
		New-Item $DestFolder -type Directory | Out-Null
		Write-Host "Done! " -ForegroundColor Green
	}
	if (Test-Path "$DestFolder\$DestFile"){
		Write-Host -ForegroundColor Yellow "File: $DestFile already exists."
        #write finish result to global
        $Global:NewFileDownloadResult = $?
	}else{
		if ($HasInternetAccess){
			Write-Host "File: $DestFile does not exist, downloading..." 
			
            Try {
                # Forcing the error output to  a custom variable, as it was the only way to catch the non-terminating error

                # clear down error
                $bitserror = $null
                Start-BitsTransfer -Source "$SourceFile" -Destination "$DestFolder\$DestFile" -RetryInterval 60 -RetryTimeout 600 -ErrorVariable BitsError -ErrorAction Continue
                # Write-Host "Done! " -ForegroundColor Green

                # This sends the result variable of the last command run to the global scope
                # $? is true if command ran successfully and false if it didn't
                
                # Show-ErrorDetails $bitserror
                
                # Write if this was successful or not to a global variable
                $Global:NewFileDownloadResult = $?
                $Global:NewFileDownloadError = $bitserror
                
                $bitserror

                If ($BitsError)
                            {
                                # loop three times
                                    While ($loop -lt "2" -and $Global:NewFileDownloadResult -eq $false) 
                                    {
                                    $loop = $loop + 1
                                    Write-Host "Download Retry attempt $loop of 2"
                                    $bitserror = $null
                                    Write-Host "Trying for Session: $url"
                                    Write-Host "Trying URL: $SourceFile"
                                    Start-BitsTransfer -Source "$SourceFile" -Destination "$DestFolder\$DestFile" -RetryInterval 60 -RetryTimeout 600 -ErrorVariable BitsError -ErrorAction Continue 
                                    $Global:NewFileDownloadResult = $?
                                    Start-Sleep -Seconds 5
                                    
                                    
                                    
                                    }
                             

                            } # Close If error 400


                }
	       catch
                {
                Write-Host "Hit Generic Catch on New-FileDownload"
                Write-Host $bitserror
                
                }
		


			
		}else{
			Write-Host "Internet access not detected. Please resolve and try again." -ForegroundColor red
		}
	}
} # end function New-FileDownload


write-host "Start IE, allow user to login so that we can grab the Dynamic URLs for the downloads" -ForegroundColor Green
write-host "The liveID redirects make it difficult to login in pure PowerShell" -ForegroundColor Green
write-host "Please leave Internet Explorer running as long as you are downloading, PowerShell will control the session" -ForegroundColor Green
write-host " "

$url = "mylync.lyncconf.com"
$ie = New-Object -comobject InternetExplorer.Application 
$ie.visible = $true 
$ie.silent = $true 
$ie.Navigate( $url )
while( $ie.busy){Start-Sleep 1} 


                            ###################################################################################
		                    # Visual Sleep Timer
			                $SleepTime = 45
                            Write-Host "This script has launched an IE window, Please sign in with a Valid Microsoft Account" -ForegroundColor Green
                            Write-Host "You have $SleepTime Seconds to Sign In" -ForegroundColor Green
			                $SleepTime..0 | Foreach-Object {
							Write-Host "$_, " -ForegroundColor Green -NoNewLine
							Start-Sleep -seconds 1
							}
			                Write-Host " " # Write Blank Line to stop next verbose overlapping with countdown
			                ###################################################################################


# Get List of URLS from Text file which must be in same directory as script

# Total Sessions to download
$totalcount = $URLList.Count
# Clear previous counter
$i = $null

Foreach ($url in $URLList)
        {
        
        # Running Count
        $i = $i+1
        Write-Host " "
        Write-host "Setting up download $i of $totalcount"
        Write-Host $url
        # Naviate to Page

        # Build the Download URL
        $DownloadURL = $url + "#download"

        $ie.Navigate( $DownloadURL )
        while( $ie.busy){Start-Sleep 1} 

        #Validate a Proper URL has been hit
         If ($ie.Document.title -like "*MyLync 2014 - Sign In*")
            {
            write-host "This Script Requires you sign into the Microsoft Account you used for LyncConf 2014"
            Write-Host "These Videos and PowerPoints can only be downloaded by people who attended LyncConf 2014 at this time"
            }
         else
            {

        
        $mp4link = $ie.Document.links | where { $_.hostname -eq "hubb.blob.core.windows.net" -and $_.pathname -like "*.mp4" }
        
        $pptxlink = $ie.Document.links | where { $_.hostname -eq "hubb.blob.core.windows.net" -and $_.pathname -like "*.pptx" }         

        $mp4filename = "$($mp4link.outerText)"

        # Fix name by removing non valid characters
        $mp4filename = $mp4filename.Replace("\"," ")
                $mp4filename = $mp4filename.Replace("/"," ")
                $mp4filename = $mp4filename.Replace(":"," ")
                $mp4filename = $mp4filename.Replace("`*"," ")
                $mp4filename = $mp4filename.Replace("`?"," ")
                $mp4filename = $mp4filename.Replace("|"," ")
                $mp4filename = $mp4filename.Replace('"',"")
                $mp4filename = $mp4filename.Replace(' .mp4',".mp4")

        $pptxfilename = "$($pptxlink.outerText)"

        # Fix name by removing non valid characters
        $pptxfilename = $pptxfilename.Replace("\"," ")
                $pptxfilename = $pptxfilename.Replace("/"," ")
                $pptxfilename = $pptxfilename.Replace(":"," ")
                $pptxfilename = $pptxfilename.Replace("`*"," ")
                $pptxfilename = $pptxfilename.Replace("`?"," ")
                $pptxfilename = $pptxfilename.Replace("|"," ")
                $pptxfilename = $pptxfilename.Replace('"',"")
                $pptxfilename = $pptxfilename.Replace(' .pptx',".pptx")

        write-host " "
        Write-Host "MP4 file name is `"$mp4filename`""
        write-host " "
        Write-host "MP4 link is $($mp4link.href)"
        write-host " "
        Write-Host "PPTX file name is `"$pptxfilename`""
        write-host " "
        write-host "PPTX link is $($pptxlink.href)"
        write-host " "

        # Download the files
        New-FileDownload -SourceFile $($mp4link.href) -DestFolder $DownloadTargetFolder -DestFile "$mp4filename"

        New-FileDownload -SourceFile $($pptxlink.href) -DestFolder $DownloadTargetFolder -DestFile "$pptxfilename"
            
            } # Close Else
        } # Close ForEach
