
# Update the details below
$controller = "controller_url"
$username = "username"
$password = "password"
$port = 8443

# Force PowerShell to use TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Disables certificate verification - primarily for self-signed certs
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

# Function to create session and login
Function Login-Unifi ($controller, $username, $password, $port) {
    $credential = @{
        username = $username
        password = $password
    }
    $uri = 'https://{0}:{1}/api/login' -f $controller, $port
    $response = Invoke-RestMethod -Uri $uri -Body (ConvertTo-Json $credential) -Method Post -SessionVariable session
    return $session
}

# Function to get devices
Function Get-UnifiDevices ($session, $controller, $port, $site_id) {
    $uri = 'https://{0}:{1}/api/s/{2}/stat/device' -f $controller, $port, $site_id
    $response = Invoke-RestMethod -Uri $uri -Method Get -WebSession $session
    return $response.data
}

# Function to get sites
Function Get-UnifiSites ($session, $controller, $port) {
    $uri = 'https://{0}:{1}/api/self/sites' -f $controller, $port
    $response = Invoke-RestMethod -Uri $uri -Method Get -WebSession $session
    return $response.data
}

#Function to get clients
Function Get-UnifiClients ($session, $controller, $port, $site_id) {
    $uri = 'https://{0}:{1}/api/s/{2}/stat/sta' -f $controller, $port, $site_id
    $response = Invoke-RestMethod -Uri $uri -Method Get -WebSession $session
    return $response.data
}


$session = Login-Unifi -controller $controller -username $username -password $password -port $port
$sites = Get-UnifiSites -session $session -controller $controller -port $port

# Empty array to hold our device data
$deviceData = @()

foreach ($site in $sites) {
    $devices = Get-UnifiDevices -session $session -controller $controller -port $port -site_id $site.name
    foreach ($device in $devices) {
        $deviceData += New-Object PSObject -Property @{
            "Site Name" = $site.desc
            "Model" = "Ubiquiti " + $device.model
            "Firmware Version" = $device.version
            "MAC Address" = $device.mac
            "IP Address" = $device.ip
        }
    }
}

# Empty array to hold our client data
$clientData = @()

foreach ($site in $sites) {
    ...
    $clients = Get-UnifiClients -session $session -controller $controller -port $port -site_id $site.name
    foreach ($client in $clients) {
        $clientData += New-Object PSObject -Property @{
            "Site Name" = $site.desc
            "Client Name" = $client.name
            "IP Address" = $client.ip
            "MAC Address" = $client.mac
            "Last Seen" = $client.last_seen
            "Device" = "Ubiquiti " + $client.device
        }
    }
}
# Output the data to a CSV file
$deviceData | Export-Csv -Path $env:USERPROFILE\Desktop\unifi_devices.csv -NoTypeInformation
$clientData | Export-Csv -Path $env:USERPROFILE\Desktop\clients.csv -NoTypeInformation
