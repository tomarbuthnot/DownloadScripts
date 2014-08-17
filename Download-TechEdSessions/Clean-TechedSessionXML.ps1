[xml]$sessions = Get-Content C:\temp\TechedEU\Sessions.xml

$list = $sessions.feed.entry

$OutputCollection= @()

foreach ($session in $list)
{
$code = ($session.content.properties.Code)
$name = ($session.title.InnerText)
Write-Host “$code $name”

$name = $name.Replace(“\”,” “)
$name = $name.Replace(“/”,” “)
$name = $name.Replace(“:”,” “)
$name = $name.Replace(“`*”,” “)
$name = $name.Replace(“`?”,” “)
$name = $name.Replace(“|”,” “)

$output = New-Object -TypeName PSobject

$output | add-member NoteProperty “Code” -value $($code)
$output | add-member NoteProperty “Name” -value “$($name)”
$OutputCollection += $output

}
$OutputCollection | Sort-Object Code | Export-Csv C:\temp\TechedEU\Teched_EU_2013SessionList.csv -NoTypeInformation