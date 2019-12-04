$StartupPath = "$([Environment]::GetFolderPath('MyDocuments'))\Startup"
<#
#AppName
Start-GracefulProcess -Path "$env:HOMEDRIVE$env:HOMEPATH" -Executable "$env:SystemRoot\system32\mmc.exe" -ArgumentList "$env:SystemRoot\system32\gpmc.msc" -AltExecutable "null.exe"

#>

function Wait-DiskActivity {
    param (
        [Int]$IopsLimit
    )
    
    do {
        $CurrentIops = Get-Counter -SampleInterval 5 -Counter "\Process(_total)\IO Data Operations/sec" | Select-Object -ExpandProperty "CounterSamples" | Select-Object -ExpandProperty "CookedValue"
        Write-Host -Object "waiting for IOPS: $CurrentIops"
    } while ( $CurrentIops -gt $IopsLimit )
}

function Start-GracefulProcess {
    param (
        [int]$IopsLimit = 1300,
        [String]$Shortcut,
        [String]$LaunchType,
        [String]$AlternateExecutable,
        [String]$WindowTitleIncludes
    )
Write-Host $LaunchType
    switch ( $LaunchType ) {
        #Executable Only
        "eo" { 
            $sh = New-Object -ComObject "WScript.Shell"
            $Executable = $sh.CreateShortcut($Shortcut).TargetPath | Select-String -AllMatches -Pattern "[a-zA-Z0-9 ]{1}\\([A-Za-z0-9. ]*)\.[a-zA-Z0-9 ]{3}" | Foreach-Object {$_.Matches.Groups.Value[1]}
            $Processes = Get-Process -Name $Executable -ErrorAction "SilentlyContinue"
            if ( $Processes ) {
                $Running = $true
            }
        }
        #Executable and WindowTitle
        "et" {
            $sh = New-Object -ComObject "WScript.Shell"
            $Executable = $sh.CreateShortcut($Shortcut).TargetPath | Select-String -AllMatches -Pattern "[a-zA-Z0-9 ]{1}\\([A-Za-z0-9. ]*)\.[a-zA-Z0-9 ]{3}" | Foreach-Object {$_.Matches.Groups.Value[1]}
            $Processes = Get-Process -Name $Executable -ErrorAction "SilentlyContinue"
            ForEach ( $w in $Processes.MainWindowTitle ) {        
                if ( $w -like "*${WindowTitleIncludes}*" ) {
                    $Running = $true
                }
            }
        }
        #WindowTitle Only
        "to" {
            $Processes = Get-Process -ErrorAction "SilentlyContinue"
            ForEach ( $w in $Processes.MainWindowTitle ) {        
                if ( $w -like "*${WindowTitleIncludes}*" ) {
                    $Running = $true
                }
            }
        }
        #Alternate Executable
        "ae" {
            $Processes = Get-Process -Name $AlternateExecutable -ErrorAction "SilentlyContinue"
            if ( $Processes ) {
                $Running = $true
            }
        }
    }

    if ( !$Running ) {
        Wait-DiskActivity -IopsLimit $IopsLimit
        Write-Host "Launching $Shortcut"
        Invoke-Item -Path $Shortcut
    } else {
        Write-Host "Running $Shortcut"
    }

}

################
# Start standard logon applications
#################

$LaunchTasks = Get-ChildItem -Path $StartupPath -Filter "*.lnk" | Select-Object -ExpandProperty "FullName" | Sort-Object
$LaunchTasks.count
ForEach ( $lt in $LaunchTasks ) {
    $PatternMatch = Split-Path $lt -Leaf | Select-String -AllMatches -Pattern '^([A-Za-z0-9]*)-([A-Za-z]*)-([A-Za-z0-9 ]*)-*([A-Za-z0-9 ]*)' | Foreach-Object {$_.Matches.Groups}

    $LaunchOrder = $PatternMatch[1].Value
    $DetectionMethod = $PatternMatch[2].Value
    $ApplicationTitle = $PatternMatch[3].Value
        $AlternateExecutable = $PatternMatch[4].Value
    switch ( $DetectionMethod ) {
        #Executable Only
        "eo" { 
                Start-GracefulProcess -Shortcut $lt -LaunchType "eo"
            }
        #Executable and WindowTitle
        "et" {
                Start-GracefulProcess -Shortcut $lt -WindowTitleIncludes $ApplicationTitle -LaunchType 'et'
            }
        #WindowTitle Only
        "to" {
            Start-GracefulProcess -Shortcut $lt -WindowTitleIncludes $ApplicationTitle -LaunchType 'to'
        }
        #Alternate Executable
        "ae" {
            Start-GracefulProcess -Shortcut $lt -WindowTitleIncludes $ApplicationTitle -AlternateExecutable $AlternateExecutable -LaunchType 'ae'
        }
    }

}
