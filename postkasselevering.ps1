param(
    $postal_code,
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

#
# Find API-key
#
$postal_code = "7010"

$url = "https://www.posten.no/levering-av-post"
$regex_pattern = 'data-react4xp-ref="parts_mailbox-delivery__\S*"\s*type="application\/json">([^<]+)<\/script>'
$rx=[regex]$regex_pattern
$result = Invoke-WebRequest -Uri $url
$apiKey = ($rx.Match($result.RawContent).Groups[1].Value | ConvertFrom-Json).props.apikey

#
# Get Delivery Days
#
$headers = @{
    "kp-api-token"="$apiKey"
    "referer"="https://www.posten.no/levering-av-post"
}

$deliveryDays = Invoke-RestMethod -Uri "https://www.posten.no/levering-av-post/_/service/no.posten.website/delivery-days?postalCode=$postal_code" -Headers $headers
$deliveryDates = $deliveryDays.delivery_dates

$ics = if ($deliveryDates.count -gt 0) {
    Write-Output "BEGIN:VCALENDAR"
    Write-Output "VERSION:2.0"
    Write-Output "METHOD:PUBLISH"
    Write-Output "PRODID:-//TurboSnute//Postkasselevering//EN"

    foreach ($date in $deliveryDates) {
        $post_date = (Get-Date $date).AddHours(11)
        New-ICSevent -Location $postal_code -Subject "Postlevering for $postal_code" -Description "Postbudet lever post i dag" -Start $post_date -End $post_date.AddHours(1) -Visibility 'Public' -ShowAs 'Free'
    }
    $ics -replace "`n", "`r`n" | out-file -path $outfile
}