$StartupPath = "$([Environment]::GetFolderPath('MyDocuments'))\Startup"
<#
Sequentially launches shortcuts in the $StartupPath following a filename format while waiting for IOPS to settle before proceeding
to the next application.
Besure to disable/delete the auto start feature of any applications included.

<Launch Order>-<Launch Type>-<Window Title Includes>-<(Optional) Alternate Executable>

Launch Order:
    An alphanumeric indication of which application to start first. It is suggested to specify applications that trigger
    high IOPS useage during launch are specified with alpha indications to ensure they are last. Initally setting files in increments
    of 10 (i.e. 010,020,030,etc...) to make re-ordering less tedious. A new application could be specified with 005,015,025,etc...
    without renaming additional shortcuts.
        010, 020, 030, 032, 049, A, AAA, Z, ZZZ are examples of valid order indications.

Launch Type:
    Indicates how the application is detected before launching in one of 4 abbreviations
    
    eo - Executable only. If the executable in the shortcut is running, the application will be detected and not launched, otherwise the
    shortcut will be invoked.
    et - Executable and Title. If the executable in the shortcut is running and a window title which matches specified in the "Window 
    Title Includes" segment is found, the application will be detected and not launched, otherwise the shortcut will be invoked.
    to - Title Only. If any window of any application is found with the title which matches specified in the "Window Title Includes"
    segment is found, the application will be detected and not launched, otherwise the shortcut will be invoked.
    ae - Alternate Executable. If a program is found running specified in the otherwise optional "Optional Alternate Executable" segment
    the application will be detected and not launched, otherwise the shortcut will be invoked. This is useful when the shortcut does not
    include the executable that eventually will be run and the title can change depending on the application's context.
    
Window Title Includes:
    Window titles are matched by wild card "*Window Title*" matches windows with the title "My Window Title Example" and "Window Title
    Example". This has no effect if Launch Types eo or ae are specified but should be included for identifying the program in the
    startup folder by a user.
    
Alternate Executable:
    Specifies an executable to detect application status which may not be called by the shortcut. This is optional for all Launch Types
    other than ae.
    
    
Examples:
010-eo-Outlook
020-eo-Word
030-et-Active Directory Users and Computers
040-et-DNS
050-ae-Microsoft Teams-Teams
Z-eo-OneDrive
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
