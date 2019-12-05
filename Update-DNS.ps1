Param (
    [String]$KeyPath = "C:\Windows\System32\drivers\etc\windows-update-client.txt",
    [String]$NSScriptPath = "$env:Temp\nsupdate.txt",
    [String]$NSUpdatePath = "$env:SystemRoot\System32"
)

begin {
    #Gather status of system IP Addresses, DNS Servers, and domains
    $IPAddresses = Get-NetIPAddress | Where-Object -FilterScript { ($_.InterfaceAlias -like "Ethernet*" -or $_.InterfaceAlias -like "WiFi*") -and $_.IPAddress -notlike "fe*"}
    $DNSServers = Get-DnsClientServerAddress | Where-Object -FilterScript { $_.InterfaceAlias -like "Ethernet*" -or $_.InterfaceAlias -like "WiFi*"}
    $DNSClient = Get-DnsClient | Where-Object -FilterScript { $_.InterfaceAlias -like "Ethernet*" -or $_.InterfaceAlias -like "WiFi*"}
}

process {
    [array]$RequestOutput = @()
    #Parse network status into simplified objects
    foreach ( $if in $IPAddresses ) {
        $requesthash = @{
            IPAddress = @{Address = $if.IPAddress;AddressFamily = $if.AddressFamily}
            Zone = $DNSClient | Where-Object -FilterScript { $_.InterfaceAlias -eq $if.InterfaceAlias } | Select-Object -ExpandProperty "ConnectionSpecificSuffix" -First 1
            Servers = $DnsServers | Where-Object -FilterScript { $_.InterfaceAlias -eq $if.InterfaceAlias } | Select-Object -ExpandProperty "ServerAddresses"
        }
        $RequestObj = New-Object -TypeName psobject -Property $requesthash
        $RequestOutput += $RequestObj 
        
    }

    #Condense zones from multiple interfaces
    [array]$UniqueZones = ($RequestOutput.Zone|Sort-Object -Unique)
    #Combine IPv6 and IPv4 addresses into a single object property for each zone
    [array]$CombinedOutput = @()
    for ($i=0;$i -lt $UniqueZones.count;$i++) {
        $Combinedhash = @{
            Addresses = $RequestOutput | Where-Object -FilterScript {$_.Zone -eq $UniqueZones[$i]} | Select-Object -ExpandProperty "IPAddress"
            Servers = $RequestOutput | Where-Object -FilterScript {$_.Zone -eq $UniqueZones[$i]} | Select-Object -ExpandProperty "Servers" | Sort-Object -Unique
            Zone = $UniqueZones[$i]
        }
        $CombinedObj = New-Object -TypeName psobject -Property $Combinedhash
        $CombinedOutput += $CombinedObj 
    }

    $CombinedOutput
    foreach ( $o in $CombinedOutput ) {
        foreach ( $s in $o.Servers ) {
            $CurrentRecords = Resolve-DnsName $env:COMPUTERNAME`.$($o.Zone) -Server $s -Type "A_AAAA" -DnsOnly -DnssecOK -QuickTimeout -ErrorAction "SilentlyContinue" | Select-Object -ExpandProperty "IPAddress" -ErrorAction "SilentlyContinue"
            if ( $CurrentRecords ) {
                $CurrentState = Compare-Object $IPAddresses.IPAddress $CurrentRecords -ErrorAction "SilentlyContinue"
            } else {
                $CurrentState = $true
            }

            if ( $CurrentState ) {
                $script += "server $s
"
                foreach ( $a in $o.Addresses ) {
                    if ( $a.AddressFamily -eq "IPv4" ) {
                        $PTR = $a.Address -replace '^(\d+)\.(\d+)\.\d+\.(\d+)$','$3.$2.$1.in-addr.arpa.'
                    } else {
                        $PTR = (([char[]][BitConverter]::ToString(([IPAddress]$a.Address).GetAddressBytes())-ne'-')[31..0]-join".")+'.ip6.arpa.'
                    }
                    $script += "update delete $env:COMPUTERNAME.$($o.Zone). $(if($a.AddressFamily -eq "IPv4"){"A"}else{"AAAA"})

update add $env:COMPUTERNAME.$($o.Zone). 60 $(if($a.AddressFamily -eq "IPv4"){"A"}else{"AAAA"}) $($a.Address)

update delete $PTR PTR

update add $PTR 60 PTR $env:COMPUTERNAME.$($o.Zone).


"
                }
            }

        }
    }
}

end {
    $script| Out-File -FilePath $NSScriptPath -Encoding "ascii" -Force
    Start-Process -FilePath (Join-Path -Path $NSUpdatePath -ChildPath "nsupdate.exe") -ArgumentList "-d -k `"$KeyPath`" `"$NSScriptPath`"" -Wait -NoNewWindow -RedirectStandardError "$env:TEMP\nsstderr" -RedirectStandardOutput "$env:TEMP\nsstdout" -WorkingDirectory $NSUpdatePath | Out-Null
    
}
