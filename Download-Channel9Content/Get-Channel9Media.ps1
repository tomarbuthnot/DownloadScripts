<#  
.SYNOPSIS  
   	Downloads  MP4's and PPTX from channel 9 site, Based on input CSV

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
    .\Get-Channel9Media.ps1 -SessionCSV .\InputCSV.csv -DownloadTargetDirectory C:\Downloads

.INPUTS
    None. You cannot pipe objects to this script.
#>


Write-Host "Hi"
Write-Host " "
Write-Host " It looks like you've tried to run this as a script but it's actually a module (a package of scripts)"
write-Host " "
write-Host " Don't panic, just Run Import-Module followed by the folder path to the module.psm1 file"
write-Host " You only need the top file, it will auto load any `"Helper Modules`" in sub folders"
write-host " e.g. Import-Module c:\mydownloads\GetChannel9MediaPSModule\GetChannel9MediaPSModule.psm1"
write-Host " One the module is loaded you can use the cmdlet in the normal way"


function Get-Channel9Media {

	[CmdletBinding(SupportsShouldProcess = $True)]
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
Write-Host "                    Get-Channel9Media by Tom Arbuthnot http://lyncdup.com                "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "            Credit: Original 2012 downloader by http://blog.scomfaq.ch (c) 2012 Stefan Roth   "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "                 Credit: Pat Richard for New-Download Function http://www.ehloworld.com/       "  -ForegroundColor Green
Write-Host "                                                                                               "  -ForegroundColor Yellow
Write-Host "###############################################################################################"  -ForegroundColor Yellow



Begin {
    	

$sessions = Import-csv $SessionCSV

$TotalDownloads =  $($sessions.count)

           }


Process {
    Foreach($session in $sessions){ #ForEach Loop 1

            $i = $i+1


        try { # Try 1


                $name = ($session.Name)
                Write-Host "$code $name"

                $name = $name.Replace("\"," ")
                $name = $name.Replace("/"," ")
                $name = $name.Replace(":"," ")
                $name = $name.Replace("`*"," ")
                $name = $name.Replace("`?"," ")
                $name = $name.Replace("|"," ")

            If ($DownloadMP3Only)
                {

               
                $MP4URL = $($session.URL)

                $MP3URL = $MP4URL.Replace("_high.mp4",".mp3")


                $DestinationFileMP3 = "$($name)" + ".mp3"

                Write-Host "MP3 URL is $MP3URL"

                write-host -ForegroundColor Yellow ("Downloading Session MP3..." + $($DestinationFileMP3) + " Number: " + $i + " out of " + $TotalDownloads)

                New-FileDownload -SourceFile $MP3URL -DestFolder $DownloadTargetDirectory -DestFile $DestinationFileMP3



                } # close download MP3 only block


            else 
                {

                # Download MP4

                write-host -ForegroundColor Yellow ("Downloading Session MP4..." + $($name) + " Number: " + $i + " out of " + $TotalDownloads)
                $MP4ToDownload = $($session.URL)

                $DestinationFileMP4 = "$($name)" + ".mp4"

                New-FileDownload -SourceFile $MP4ToDownload -DestFolder $DownloadTargetDirectory -DestFile $DestinationFileMP4

                } # close else block


    } # close Try1
     
    catch {
            
            $errorfile = $scriptdir + "\sessions_notavailable.txt";
            write-host -foregroundColor Red ("Session not available please check $errorfile")
            if(!(test-path -path $errorfile)) { new-item $errorfile -type file};
            $session | out-file $errorfile -Append

        } # Close Catch


            } # Close Foreach Loop1


write-host -ForegroundColor Green "Finished!"

} # Close Process Block

} # Close Function

# SIG # Begin signature block
# MIIfawYJKoZIhvcNAQcCoIIfXDCCH1gCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUId6xAJ0hZlWMlVvgx+tW4Wmb
# WoOgghqdMIIGajCCBVKgAwIBAgIQA5/t7ct5W43tMgyJGfA2iTANBgkqhkiG9w0B
# AQUFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVk
# IElEIENBLTEwHhcNMTMwNTIxMDAwMDAwWhcNMTQwNjA0MDAwMDAwWjBHMQswCQYD
# VQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxJTAjBgNVBAMTHERpZ2lDZXJ0IFRp
# bWVzdGFtcCBSZXNwb25kZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC6aUqBTW+lFBaqis1nvku/xmmPWBzgeegenVgmmNpc1Hyj+dsrjBI2w/z5ZAax
# u8KomAoXDeGV60C065ZtmL+mj3nPvIqSe22cGAZR2KUYUzIBJxlh6IRB38bw6Mr+
# d61f2J57jGBvhVxGvWvnD4DO5wPDfDHPt2VVxvvgmQjkc1r7l9rQTL60tsYPfyaS
# qbj8OO605DqkSNBM6qlGJ1vPkhGTnBan/tKtHyLFHqzBce+8StsBCUTfmBwtZ7qo
# igMzyVG19wJNCaRN/oBexddFw30IqgEzzDPYTzAW5P8iMi7rfjvw+R4y65Ul0vL+
# bVSEutXl1NHdG6+9WXuUhTABAgMBAAGjggM1MIIDMTAOBgNVHQ8BAf8EBAMCB4Aw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCCAb8GA1UdIASC
# AbYwggGyMIIBoQYJYIZIAYb9bAcBMIIBkjAoBggrBgEFBQcCARYcaHR0cHM6Ly93
# d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkA
# IAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUA
# IABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAA
# bwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEA
# bgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIA
# ZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkA
# bABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUA
# ZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjALBglg
# hkgBhv1sAxUwHwYDVR0jBBgwFoAUFQASKxOYspkH7R7for5XDStnAs0wHQYDVR0O
# BBYEFGMvyd95knu1I8q74aTuM37j4p36MH0GA1UdHwR2MHQwOKA2oDSGMmh0dHA6
# Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3JsMDig
# NqA0hjJodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURD
# QS0xLmNybDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3Nw
# LmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2VydHMuZGlnaWNl
# cnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcnQwDQYJKoZIhvcNAQEFBQAD
# ggEBAKt0vUAATHYVJVc90xwD/31FyEUSZucoZWDY3zuz+g3BrDOP9IG5YfGd+5hV
# 195HQ7qAPfFIzD9nMFYfzvTQTIS9h6SexeEPqAZd0C9uXtwZ6PCH6uBOrz1sII5z
# b37WhxjghtOa/J7qjHLpQQ+4cbU4LPgpstUcop0b7F8quNw3IOHLu/DQbGyls8uf
# SvZU4yY0PS64wSsct/bDPf7RLR5Q9JTI+P3uc9tJtRv09f+lkME5FBvY7XEbapj7
# +kCaRKkpDlVeeLi3pIPDcAHwZkDlrnk04StNA6Et5ttUYhjt1QmLoqrWDMhPGr6Z
# JXhpmYnUWYne34jw02dedKWdpkQwggajMIIFi6ADAgECAhAPqEkGFdcAoL4hdv3F
# 7G29MA0GCSqGSIb3DQEBBQUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdp
# Q2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0Rp
# Z2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMTAyMTExMjAwMDBaFw0yNjAy
# MTAxMjAwMDBaMG8xCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xLjAsBgNVBAMTJURpZ2lDZXJ0IEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBLTEwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQCcfPmgjwrKiUtTmjzsGSJ/DMv3SETQPyJumk/6zt/G0ySR/6hS
# k+dy+PFGhpTFqxf0eH/Ler6QJhx8Uy/lg+e7agUozKAXEUsYIPO3vfLcy7iGQEUf
# T/k5mNM7629ppFwBLrFm6aa43Abero1i/kQngqkDw/7mJguTSXHlOG1O/oBcZ3e1
# 1W9mZJRru4hJaNjR9H4hwebFHsnglrgJlflLnq7MMb1qWkKnxAVHfWAr2aFdvftW
# k+8b/HL53z4y/d0qLDJG2l5jvNC4y0wQNfxQX6xDRHz+hERQtIwqPXQM9HqLckvg
# VrUTtmPpP05JI+cGFvAlqwH4KEHmx9RkO12rAgMBAAGjggNDMIIDPzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMwggHDBgNVHSAEggG6MIIBtjCC
# AbIGCGCGSAGG/WwDMIIBpDA6BggrBgEFBQcCARYuaHR0cDovL3d3dy5kaWdpY2Vy
# dC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6C
# AVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBp
# AGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABh
# AG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBD
# AFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5
# ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABs
# AGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABv
# AHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBj
# AGUALjASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsGAQUFBwEBBG0wazAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsNC5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMB0GA1UdDgQW
# BBR7aM4pqsAXvkl64eU/1qf3RY81MjAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYun
# pyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEAe3IdZP+IyDrBt+nnqcSHu9uUkteQ
# WTP6K4feqFuAJT8Tj5uDG3xDxOaM3zk+wxXssNo7ISV7JMFyXbhHkYETRvqcP2pR
# ON60Jcvwq9/FKAFUeRBGJNE4DyahYZBNur0o5j/xxKqb9to1U0/J8j3TbNwj7aqg
# TWcJ8zqAPTz7NkyQ53ak3fI6v1Y1L6JMZejg1NrRx8iRai0jTzc7GZQY1NWcEDzV
# sRwZ/4/Ia5ue+K6cmZZ40c2cURVbQiZyWo0KSiOSQOiG3iLCkzrUm2im3yl/Brk8
# Dr2fxIacgkdCcTKGCZlyCXlLnXFp9UH/fzl3ZPGEjb6LHrJ9aKOlkLEM/zCCBrMw
# ggWboAMCAQICEAeD459qgkM8IIqAW+ceb8cwDQYJKoZIhvcNAQEFBQAwbzELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEuMCwGA1UEAxMlRGlnaUNlcnQgQXNzdXJlZCBJRCBDb2RlIFNp
# Z25pbmcgQ0EtMTAeFw0xMzA3MDEwMDAwMDBaFw0xNDA3MDkxMjAwMDBaMH8xCzAJ
# BgNVBAYTAkdCMRYwFAYDVQQIEw1IZXJ0Zm9yZHNoaXJlMRIwEAYDVQQHEwlTdGV2
# ZW5hZ2UxITAfBgNVBAoTGFRob21hcyBDaGFybGVzIEFyYnV0aG5vdDEhMB8GA1UE
# AxMYVGhvbWFzIENoYXJsZXMgQXJidXRobm90MIIBIjANBgkqhkiG9w0BAQEFAAOC
# AQ8AMIIBCgKCAQEAscTUpaPp4CHYTnhSINweZybDzXuiE7ixnnMyKtjkIUnHgb8g
# h8SDoked4he02tj1BsfaAG+gIn+P6swtSIbQbyPDafSY4DJYRwEOBrntIPMjo4fI
# ozef46QLkPrBBJLSUfl73RK56CzlvEt9kVNIVZVEqweGgK0oVV69oQohu3ZLzATk
# V2yzYEKyaB/tFwiuPitlzGDc5nVQO0PAI8idyvIzSP8EL7zWP9A9Mnh+7at9JUQo
# QsdHz0k34wB2twltYsdmX1KCxUYNYbXnBBSJVgH4DUzPpnBP7dK3v+ImBYYW9l1w
# j4DuwRZllIjIkyt7z+wJgR47VtEiapVttLVFfwIDAQABo4IDOTCCAzUwHwYDVR0j
# BBgwFoAUe2jOKarAF75JeuHlP9an90WPNTIwHQYDVR0OBBYEFDhIVxDUOkhgJprW
# 8lsCAwRQ7v9RMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzBz
# BgNVHR8EbDBqMDOgMaAvhi1odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vYXNzdXJl
# ZC1jcy0yMDExYS5jcmwwM6AxoC+GLWh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9h
# c3N1cmVkLWNzLTIwMTFhLmNybDCCAcQGA1UdIASCAbswggG3MIIBswYJYIZIAYb9
# bAMBMIIBpDA6BggrBgEFBQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3Ns
# LWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkA
# IAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUA
# IABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAA
# bwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEA
# bgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIA
# ZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkA
# bABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUA
# ZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjCBggYI
# KwYBBQUHAQEEdjB0MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wTAYIKwYBBQUHMAKGQGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRENvZGVTaWduaW5nQ0EtMS5jcnQwDAYDVR0TAQH/BAIwADAN
# BgkqhkiG9w0BAQUFAAOCAQEAeA/3K+BEtjNJPZNfUTJQtpUJsIeyQgZTj3Ywyh8/
# XzDPRXr6TXEnSiqX47I+HmAkE6cB4o6KjV3LPHLpQdhMISYUeC8jnwqknPL3fSzO
# kJYKho4hQF6DLbOhh34QoHQfWkDgUP3duS3dPjweU//ngUGRJ6IMv9DwfDVqvtpx
# nHQY4B8c4trZCNSHMkpbVmn1LadZfYVK8zkei1B5AOnRRWoiA93t3sUPXoaJFmQ9
# 1JahD9tBxL6SlDMb2fp48x7UekfHQUUHvwwFh3csOk+zhedmXD2lAa4CgvtPiWyh
# Q0wf5IDSKj6NGC0tTMFdQRkLzLQRJgFfFO14yZvOZSTFPzCCBs0wggW1oAMCAQIC
# EAb9+QOWA63qAArrPye7uhswDQYJKoZIhvcNAQEFBQAwZTELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTA2MTEx
# MDAwMDAwMFoXDTIxMTExMDAwMDAwMFowYjELMAkGA1UEBhMCVVMxFTATBgNVBAoT
# DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UE
# AxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0xMIIBIjANBgkqhkiG9w0BAQEFAAOC
# AQ8AMIIBCgKCAQEA6IItmfnKwkKVpYBzQHDSnlZUXKnE0kEGj8kz/E1FkVyBn+0s
# nPgWWd+etSQVwpi5tHdJ3InECtqvy15r7a2wcTHrzzpADEZNk+yLejYIA6sMNP4Y
# SYL+x8cxSIB8HqIPkg5QycaH6zY/2DDD/6b3+6LNb3Mj/qxWBZDwMiEWicZwiPkF
# l32jx0PdAug7Pe2xQaPtP77blUjE7h6z8rwMK5nQxl0SQoHhg26Ccz8mSxSQrllm
# CsSNvtLOBq6thG9IhJtPQLnxTPKvmPv2zkBdXPao8S+v7Iki8msYZbHBc63X8djP
# Hgp0XEK4aH631XcKJ1Z8D2KkPzIUYJX9BwSiCQIDAQABo4IDejCCA3YwDgYDVR0P
# AQH/BAQDAgGGMDsGA1UdJQQ0MDIGCCsGAQUFBwMBBggrBgEFBQcDAgYIKwYBBQUH
# AwMGCCsGAQUFBwMEBggrBgEFBQcDCDCCAdIGA1UdIASCAckwggHFMIIBtAYKYIZI
# AYb9bAABBDCCAaQwOgYIKwYBBQUHAgEWLmh0dHA6Ly93d3cuZGlnaWNlcnQuY29t
# L3NzbC1jcHMtcmVwb3NpdG9yeS5odG0wggFkBggrBgEFBQcCAjCCAVYeggFSAEEA
# bgB5ACAAdQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEA
# dABlACAAYwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMA
# ZQAgAG8AZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMA
# IABhAG4AZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEA
# ZwByAGUAZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEA
# YgBpAGwAaQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEA
# dABlAGQAIABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4w
# CwYJYIZIAYb9bAMVMBIGA1UdEwEB/wQIMAYBAf8CAQAweQYIKwYBBQUHAQEEbTBr
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUH
# MAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJ
# RFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6
# Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmww
# HQYDVR0OBBYEFBUAEisTmLKZB+0e36K+Vw0rZwLNMB8GA1UdIwQYMBaAFEXroq/0
# ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBBQUAA4IBAQBGUD7Jtygkpzgdtlsp
# r1LPUukxR6tWXHvVDQtBs+/sdR90OPKyXGGinJXDUOSCuSPRujqGcq04eKx1XRcX
# NHJHhZRW0eu7NoR3zCSl8wQZVann4+erYs37iy2QwsDStZS9Xk+xBdIOPRqpFFum
# hjFiqKgz5Js5p8T1zh14dpQlc+Qqq8+cdkvtX8JLFuRLcEwAiR78xXm8TBJX/l/h
# HrwCXaj++wc4Tw3GXZG5D2dFzdaD7eeSDY2xaYxP+1ngIw/Sqq4AfO6cQg7Pkdcn
# txbuD8O9fAqg7iwIVYUiuOsYGk38KiGtSTGDR5V3cdyxG0tLHBCcdxTBnU8vWpUI
# KRAmMYIEODCCBDQCAQEwgYMwbzELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEuMCwGA1UEAxMlRGln
# aUNlcnQgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EtMQIQB4Pjn2qCQzwgioBb
# 5x5vxzAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUopK3PPoKu1eYXd6ARt71tmwx958wDQYJKoZI
# hvcNAQEBBQAEggEAXXsnIfsIvO5+32Uzpem6qBnsxVc/yrPi0wSojKEA7SLdqhFV
# KqBd1M2CIQPNZdOtF/7IYKTZ0/hNKBdbXteTidrm+gynNGcF//8+6FUN8Wsl3NCI
# 3dKAdAlTve7+6G0ZWuuLoCc/6NAX5396cbeVoo5AKCV2ilngz03hJmULraFNEihu
# v7au/H58qdsKjK7EmkzvHdC4v19XrxUEcPQZz23WvauOVpXiC9iJW65TdUJcPbAU
# vi9aI3V9VlSclyev7trOmT82ODsmj5nD+W7VQeg2NicpQbYlcJBUXbeM5AuzJQZg
# oQBRGao5u8zjLMI5bwhrlnTyJ+CMqO/wTCRYC6GCAg8wggILBgkqhkiG9w0BCQYx
# ggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IEFzc3VyZWQgSUQgQ0EtMQIQA5/t7ct5W43tMgyJGfA2iTAJBgUrDgMCGgUAoF0w
# GAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTMwNzAy
# MDkzMzQ0WjAjBgkqhkiG9w0BCQQxFgQUYWICQp33QyGoUWXWJlKi8SEcigswDQYJ
# KoZIhvcNAQEBBQAEggEAkz3bW9IFzwqWD16/+6JfNHAfQVIFFbuQw0QV0kvrkxSc
# XBEyKLL2pQAbM7DBag+4044vyDzFk3PHzCFVBdfDsFdsWHAvcymsNPy7jMKtzYth
# BYprOyfIVIMUCqkgAzqUBY3c84AsvqMxAX2xhHgMw0wPHZAgoOtOhmmpSU/afCO9
# QpeKgL1AYifvljxgo9anSI9y0pbZQR9UOO/0+zQ6U6WEJeZ2zjkyMO2Iv0m8s24X
# kh3ujOsVHw/g88UWYQ7ObW54pQ6OFUoIYhVQYfBEVyJr2VUHBUC51fnMmReP4RtQ
# NZWs50Uzvk3VRm+k1aixJgVH/yKba+CQ8QdF7YWdug==
# SIG # End signature block
