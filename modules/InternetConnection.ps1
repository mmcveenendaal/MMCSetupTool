function Test-Internet {
    $status = Test-NetConnection
    return $status.PingSucceeded
}

function Connect-Wifi {
    # get the profile
    $wlanprofile = "assets/mmc_guest.xml"

    # install the profile
    netsh wlan add profile filename=$wlanprofile | Out-Null
    netsh wlan connect name="MMC_Guest" | Out-Null

    function Get-Connection {
        return Get-NetRoute | Where-Object DestinationPrefix -eq '0.0.0.0/0' | Get-NetIPInterface | Where-Object ConnectionState -eq 'Connected'
    }

    $connected = $false
    $count = 0

    while (!$connected) {
        if ($count++ -eq 5) {
            return $false
        }

        $connected = Get-Connection
        Start-Sleep -Seconds 3
    }

    return $true
}
