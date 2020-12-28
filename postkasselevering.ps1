param(
    $postnr,
    $outfile
)
<#
.SYNOPSIS
    Generates ICS file for post delivery dates. (Posten Norway)
.DESCRIPTION
    Generates ICS file for post delivery dates based on post code. (Posten Norway)
.EXAMPLE
    ./postkasselevering.ps1 -postnr 7010
#>
function New-ICSevent {
    [CmdletBinding()]
    [OutputType([String])]
    param(
        [Parameter(Mandatory=$true)][string]$Location,
        [Parameter(Mandatory=$true)][string]$Subject,
        [Parameter(Mandatory=$true)][datetime]$Start,
        [Parameter(Mandatory=$true)][datetime]$End,
        [Parameter(Mandatory=$true)][string]$Description,
        [Parameter(Mandatory=$true)][ValidateSet('Private', 'Public', 'Confidential')][string]$Visibility = 'Public',
        [Parameter(Mandatory=$true)][ValidateSet('Free', 'Busy')]$ShowAs = 'Busy'
    )

    begin {
    }

    process {
        $icsDateFormat = "yyyyMMddTHHmmssZ"

        $Subject = "Postkasselevering"
        $EventDescription = "Postbudet kommer."

        Write-Output "BEGIN:VEVENT"
        Write-Output "UID:$([guid]::NewGuid())"
        Write-Output "CREATED:$((Get-Date).ToUniversalTime().ToString($icsDateFormat))"
        Write-Output "DTSTAMP:$((Get-Date).ToUniversalTime().ToString($icsDateFormat))"
        Write-Output "LAST-MODIFIED:$((Get-Date).ToUniversalTime().ToString($icsDateFormat))"
        Write-Output "CLASS:$Visibility"
        Write-Output "CATEGORIES:$($Category -join ',')"
        Write-Output "SEQUENCE:0"
        Write-Output "DTSTART:$($Start.ToUniversalTime().ToString($icsDateFormat))"
        Write-Output "DTEND:$($End.ToUniversalTime().ToString($icsDateFormat))"
        Write-Output "DESCRIPTION:$Description"
        Write-Output "SUMMARY:$Subject"
        Write-Output "LOCATION:$Location"
        Write-Output "TRANSP:$(if($ShowAs -eq 'Free') {'TRANSPARENT'})"
        Write-Output "END:VEVENT"
    }

    end {
    }
}


$res = Invoke-WebRequest -Uri "https://www.posten.no/levering-av-post-2020/_/component/main/1/leftRegion/1?postCode=$postnr" `
-Headers @{
"User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36"
  "x-requested-with"="XMLHttpRequest"
  "Accept"="*/*"
  "Sec-Fetch-Site"="same-origin"
  "Sec-Fetch-Mode"="cors"
  "Sec-Fetch-Dest"="empty"
  "Referer"="https://www.posten.no/levering-av-post-2020"
  "Accept-Encoding"="gzip, deflate, br"
  "Accept-Language"="en-NO,en;q=0.9,nb-NO;q=0.8,nb;q=0.7,en-US;q=0.6,no;q=0.5"
  "dnt"="1"
  "sec-gpc"="1"
} `
-ContentType "application/json"

$ics = if ($res.StatusCode -eq 200) {

    # RegEx
    [regex]$regex = " (\d{1,2})\. ([a-z]+)$"

    # Months
    $month = @{
        januar = 1
        februar = 2
        marsj = 3
        april = 4
        mai = 5
        juni = 6
        juli = 7
        august = 8
        september = 9
        oktober = 10
        november = 11
        desember = 12
    }

    $lastday = Get-Date
    $year = $lastday.Year

    $deliverydays = ($res.Content | ConvertFrom-Json).nextDeliveryDays

    if ($deliverydays.count -ne 0) {             
           Write-Output "BEGIN:VCALENDAR"
           Write-Output "VERSION:2.0"
           Write-Output "METHOD:PUBLISH"
           Write-Output "PRODID:-//TurboSnute//Postkasselevering//EN"
           foreach ($stringDate in $deliverydays) {
             $match = $regex.Match($stringDate)
             if (($match.Groups[1].success) -and ($match.Groups[2].success)) {
               $deliveryday = $match.Groups[1].Value
               $deliveryMonth = $month[$match.Groups[2].Value]
               #write-host -foregroundcolor green $deliveryday
               #write-host -foregroundcolor red $($lastday.month)
               # Check if last day of year.
               if ($deliveryMonth -lt $lastday.Month) {
                   $year++
               }
               $deliveryDate = Get-Date "$year-$deliveryMonth-$deliveryday"
               #"$year-$deliveryMonth-$deliveryday"
               $lastday = $deliveryDate
               $start = $deliveryDate
               $end = ($deliveryDate.AddDays(1)).AddHours(-4)
               New-ICSevent -Location $postnr -Subject "Postlevering for $postCode" -Description "Postbudet lever post i dag" -Start $start -End $end -Visibility 'Public' -ShowAs 'Free'
             } #end if
           } #end foreach
           Write-Output "END:VCALENDAR"
    } #end if
}

$ics -replace "`n", "`r`n" | out-file -path $outfile
