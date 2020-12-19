<#
.SYNOPSIS
    Short description
.DESCRIPTION
    Long description
.EXAMPLE
    Example of how to use this cmdlet
.EXAMPLE
    Another example of how to use this cmdlet
#>
function New-ICS {
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

        $postnr = 7066
        $Subject = "Postkasselevering"
        $EventDescription = "Postbudet kommer."

@"
BEGIN:VEVENT
UID: $([guid]::NewGuid())
CREATED: $((Get-Date).ToUniversalTime().ToString($icsDateFormat))
DTSTAMP: $((Get-Date).ToUniversalTime().ToString($icsDateFormat))
LAST-MODIFIED: $((Get-Date).ToUniversalTime().ToString($icsDateFormat))
CLASS:$Visibility
CATEGORIES:$($Category -join ',')
SEQUENCE:0
DTSTART: $($Start.ToUniversalTime().ToString($icsDateFormat))
DTEND: $($End.ToUniversalTime().ToString($icsDateFormat))
DESCRIPTION: $Description
SUMMARY: $Subject
LOCATION: $Location
TRANSP:$(if($ShowAs -eq 'Free') {'TRANSPARENT'})
END:VEVENT
"@

    }
    
    end {
    }
}


$postCode = 7066


$res = Invoke-WebRequest -Uri "https://www.posten.no/levering-av-post-2020/_/component/main/1/leftRegion/1?postCode=$postCode" `
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

if ($res.StatusCode -eq 200) {

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
    
    $ical = foreach ($stringDate in $deliverydays) {
        $match = $regex.Match($stringDate)

        if (($match.Groups[1].success) -and ($match.Groups[2].success)) {
            $deliveryday = $match.Groups[1].Value
            $deliveryMonth = $month[$match.Groups[2].Value]

            # Check if last day of year.
            if (($lastday.Month -eq 12) -and ($lastday.day -eq 31) -and ($deliveryday -ne 31)) {
                $year++
            }

           $deliveryDate = Get-Date "$year-$deliveryMonth-$deliveryday"
           #"$year-$deliveryMonth-$deliveryday"
           $lastday = $deliveryday
           $start = $deliveryDate
           $end = $deliveryDate.AddDays(1)
           Write-Output "BEGIN:VCALENDAR"
           Write-Output "VERSION:2.0"
           Write-Output "METHOD:PUBLISH"
           Write-Output "PRODID:-//JHP//We love PowerShell!//EN"
           New-ICSevent -Location $postCode -Subject "Postlevering for $postCode" -Description "Postbudet lever post i dag" -Start $start -End $end -Visibility 'Public' -ShowAs 'Free'
           Write-Output "END:VCALENDAR"
        }
    }
}

$ical