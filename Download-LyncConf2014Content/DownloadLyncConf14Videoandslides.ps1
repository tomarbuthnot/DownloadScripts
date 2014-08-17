# Originally published at https://gist.github.com/nzthiago/5736907
# I Customized it for MEC 2014 and now for LyncConf 2014 with slides (find more info about my script on www.msdigest.net)
# Original SharePoint script by:  Vlad Catrinescu (http://absolute-sharepoint.com/2014/03/ultimate-script-download-sharepoint-conference-2014-videos-slides.html)
# If you like it, leave me a comment
# If you don't like it, complain to Github. :)

Write-host " "
Write-Host "###############################################################################################"  -ForegroundColor Yellow
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "            Download Microsoft Lync Conference 2014 Videos and PPTX                            "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "         Based on the SharePoint Conference download script by Vlad Catrinescu                 "
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "             Download this script and updates from www.msdigest.net                            "  
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "###############################################################################################"  -ForegroundColor Yellow

[Environment]::CurrentDirectory=(Get-Location -PSProvider FileSystem).ProviderPath 
$rss = (new-object net.webclient)

# Grab the RSS feed for the MP4 downloads

# Lync Conference 2014 Videos
$a = ([xml]$rss.downloadstring("http://channel9.msdn.com/Events/Lync-Conference/Lync-Conference-2014/RSS/mp4high")) 
$b = ([xml]$rss.downloadstring("http://channel9.msdn.com/Events/Lync-Conference/Lync-Conference-2014/RSS/slides")) 

#other qualities for the videos only. Choose the one you want!
# $a = ([xml]$rss.downloadstring("http://channel9.msdn.com/Events/Lync-Conference/Lync-Conference-2014/RSS/mp4")) 
#$a = ([xml]$rss.downloadstring("http://channel9.msdn.com/Events/Lync-Conference/Lync-Conference-2014/RSS/mp3")) 

#Preferably enter something not too long to not have filename problems! cut and paste them afterwards
$downloadlocation = "C:\LyncConf2014"

	if (-not (Test-Path $downloadlocation)) { 
		Write-Host "Folder $fpath dosen't exist. Creating it..."  
		New-Item $downloadlocation -type directory 
	}
set-location $downloadlocation

#Download all the slides	
$b.rss.channel.item | foreach{   
	$code = $_.comments.split("/") | select -last 1	   
	
	# Grab the URL for the PPTX file
	$urlpptx = New-Object System.Uri($_.enclosure.url)  
    $filepptx = $code + "-" + $_.creator + " - " + $_.title.Replace(":", "-").Replace("?", "").Replace("/", "-").Replace("<", "").Replace("|", "").Replace('"',"").Replace("*","")
	$filepptx = $filepptx.substring(0, [System.Math]::Min(120, $filepptx.Length))
	$filepptx = $filepptx.trim()
	$filepptx = $filepptx + ".pptx" 
	if ($code -ne "")
	{
		 $folder = $code + " - " + $_.title.Replace(":", "-").Replace("?", "").Replace("/", "-").Replace("<", "").Replace("|", "").Replace('"',"").Replace("*","")
		 $folder = $folder.substring(0, [System.Math]::Min(100, $folder.Length))
		 $folder = $folder.trim()
	}
	else
	{
		$folder = "NoCodeSessions"
	}
	
	if (-not (Test-Path $folder)) { 
		Write-Host "Folder $folder dosen't exist. Creating it..."  
		New-Item $folder -type directory 
	}
		
	
	#text description from session . Thank you VaperWare
	$OutFile = New-Item -type file "$($downloadlocation)\$($Folder)\$($Code.trim()).txt" -Force  
    $Category = "" ; $Content = ""
    $_.category | foreach {$Category += $_ + ","}
    $Content = $_.title.trim() + "`r`n" + $_.creator + "`r`n" + $_.summary.trim() + "`r`n" + "`r`n" + $Category.Substring(0,$Category.Length -1)
   add-content $OutFile $Content

	# Make sure the PowerPoint file doesn't already exist
	if (!(test-path "$downloadlocation\$folder\$filepptx"))     
	{ 	
		# Echo out the  file that's being downloaded
		$filepptx
		$wc = (New-Object System.Net.WebClient)  

		# Download the MP4 file
		$wc.DownloadFile($urlpptx, "$downloadlocation\$filepptx")
		mv $filepptx $folder 

	}
	}

#download all the mp4

# Walk through each item in the feed 
$a.rss.channel.item | foreach{   
	$code = $_.comments.split("/") | select -last 1	   
	
	# Grab the URL for the MP4 file
	$url = New-Object System.Uri($_.enclosure.url)  
	
	# Create the local file name for the MP4 download
	$file = $code + "-" + $_.creator + "-" + $_.title.Replace(":", "-").Replace("?", "").Replace("/", "-").Replace("<", "").Replace("|", "").Replace('"',"").Replace("*","")
	$file = $file.substring(0, [System.Math]::Min(120, $file.Length))
	$file = $file.trim()
	$file = $file + ".mp4"  
	
	if ($code -ne "")
	{
		 $folder = $code + " - " + $_.title.Replace(":", "-").Replace("?", "").Replace("/", "-").Replace("<", "").Replace("|", "").Replace('"',"").Replace("*","")
		 $folder = $folder.substring(0, [System.Math]::Min(100, $folder.Length))
		 $folder = $folder.trim()
	}
	else
	{
		$folder = "NoCodeSessions"
	}
	
	if (-not (Test-Path $folder)) { 
		Write-Host "Folder $folder) dosen't exist. Creating it..."  
		New-Item $folder -type directory 
	}
	
	
	
	# Make sure the MP4 file doesn't already exist

	if (!(test-path "$folder\$file"))     
	{ 	
		# Echo out the  file that's being downloaded
		$file
		$wc = (New-Object System.Net.WebClient)  

		# Download the MP4 file
		$wc.DownloadFile($url, "$downloadlocation\$file")
		mv $file $folder
	}
	
		
	}