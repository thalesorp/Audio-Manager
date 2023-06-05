<#
.SYNOPSIS
Cmdlet that provides some tools to programmatically manage audio on Windows systems.

.DESCRIPTION
TBD

.PARAMETER NextOutput
Change the audio output to the next available output.

.PARAMETER SetVolume
Change the volume of currently in use output.

.INPUTS
TBD

.EXAMPLE
.\Audio-Manager.ps1 -NextOutput
Change the audio output to the next available output.

.NOTES
Author: Thales Pinto
Version: 0.4.0
Licence: This code is licensed under the MIT license.
#>

using module .\OutputList.psm1

[CmdletBinding(DefaultParameterSetName="None")]

param (
    [Parameter(Mandatory=$false, ParameterSetName="NextOutput")]
    [Alias("Next")][switch]$NextOutput,

    [Parameter(Mandatory=$false, ParameterSetName="SetOutput",
        HelpMessage="Enter the nickname (previously setted in Profile file) of output to be setted as playback device.")]
    [Alias("Set")][string]$SetOutput,

    [Parameter(Mandatory=$false, ParameterSetName="SetOutput")]
    [ValidateRange(0,100)][int]$Volume,

    [Parameter(Mandatory=$false, ParameterSetName="SetVolume")]
    [ValidateRange(0,100)][int]$SetVolume,

    [Parameter(Mandatory=$false, ParameterSetName="ListOutput")]
    [Alias("List")][switch]$ListOutput

)

begin {

    <#
    .SYNOPSIS
    Install the modules needed to Audio Manager works properly.
    #>
    function Resolve-Dependencies {
        $dependencies = "BurntToast", "AudioDeviceCmdlets", "PSMenu"

        forEach ($dependency in $dependencies) {
            if ((Get-Module -ListAvailable | Where-Object {$_.Name -eq $dependency}) -eq $null) {
                Start-Process -wait powershell -Verb runAs -ArgumentList "Write-Host `"Install-Module -Name $dependency`" ; Install-Module -Name $dependency"
            }
        }
    }

    <#
    .SYNOPSIS
    Initialize all resources: program settings, outputs and profile settings.
    #>
    function Initialize-Resources {
        $global:settings = @{}
        $global:settings.NotificationImagePath = "."
        $global:outputs = Get-AvailableOutputs
        Initialize-Profile
    }

    <#
    .SYNOPSIS
    Create a default profile JSON. If it already exists, read the information on it.
    #>
    function Initialize-Profile {
        $profilePath = Join-Path $PSScriptRoot -ChildPath "Profile.json"

        if ((Test-Path -Path $profilePath -PathType Leaf) -eq $false) {
            $outputsCustomObject = [PSCustomObject] @{
                NotificationImagePath = $global:settings.NotificationImagePath
                Outputs = [Ordered]@{}
            }

            ForEach ($output in $global:outputs.GetEnumerator()) {
                $outputValues = [PSCustomObject] @{
                    SystemName = $output.Value.SystemName
                    Nickname = $output.Value.Nickname
                    DefaultVolume = $output.Value.DefaultVolume
                    Enabled = $True
                }
                $outputsCustomObject.Outputs.Add($output.Name, $outputValues)
            }

            $outputsCustomObject | ConvertTo-Json | Out-File $profilePath
        } else {
            $profileContent = Get-Content $profilePath -Raw | ConvertFrom-Json

            if (Test-Path -Path $profileContent.NotificationImagePath -PathType Leaf) {
                $global:settings.NotificationImagePath = $profileContent.NotificationImagePath
            }
            ForEach ($output in $profileContent.Outputs.PSObject.Properties) {
                $OutputJsonData = [PSCustomObject] @{
                    ID = $output.Name
                    SystemName = $output.Value.SystemName
                    Nickname = $output.Value.Nickname
                    DefaultVolume = $output.Value.DefaultVolume
                    Enabled = $output.Value.Enabled
                }
                $global:outputs.Update($OutputJsonData)
            }
        }
    }

    <#
    .SYNOPSIS
    Send notification with message.
    #>
    function Notification {
        param (
            [string]$Message
        )
        New-BurntToastNotification -AppLogo $global:settings.NotificationImagePath -Text "Audio Manager", $Message
    }

    <#
    .SYNOPSIS
    Builds the output list containing all playback devices currently available.
    #>
    function Get-AvailableOutputs {
        $audioDevices = Get-AudioDevice -List | Where-Object {$_.Type -eq "Playback"}

        if ($audioDevices -eq $null) {
            Notification "No audio output available"
            Write-Error "No audio output available"
            exit
        }

        $outputs = [OutputList]::new()
        ForEach ($audioDevice in $audioDevices) {
            $outputs.Add($audioDevice)
        }

        return $outputs
    }

    <#
    .SYNOPSIS
    Change the audio device output based on the "ID" parameter. Send a notification about the change.
    #>
    function Set-Output {
        param (
            [Parameter(Mandatory=$true, ParameterSetName="SetByNickname")][String]$Nickname,
            [Parameter(Mandatory=$true, ParameterSetName="SetByID")][String]$ID,
            [ValidateRange(0,100)][Int]$Volume
        )

        if ($PSCmdlet.ParameterSetName -eq "SetByNickname") {
            $ID = $global:outputs.GetOutputID($Nickname)
        }

        $audioDevice = Get-AudioDevice -ID $ID

        if ($audioDevice -eq $null) {
            # TODO: Deal with this possible error
            return $null
        }

        if ($audioDevice.Default) {
            return $false
        }

        $audioDevice | Set-AudioDevice | Out-Null

        if (-not ($PSBoundParameters.ContainsKey("Volume"))) {
            $Volume = $global:outputs.outputs[$ID].DefaultVolume
        }

        Set-AudioDevice -PlaybackVolume $Volume

        Notification "Output: $($global:outputs.outputs[$ID].Nickname)`nVolume: $Volume%"
        return $true
    }

    <#
    .SYNOPSIS
    Change the volume of currently in use output.
    #>
    function Set-OutputVolume {
        param (
            [ValidateRange(0,100)][Int]$Volume
        )

        $defaultOutputID = (Get-AudioDevice -List | Where-Object {$_.Type -eq "Playback" -and $_.Default -eq $true}).ID

        Set-AudioDevice -PlaybackVolume $Volume

        Notification "Output: $($global:outputs.outputs[$defaultOutputID].Nickname)`nVolume: $Volume%"
    }

    <#
    .SYNOPSIS
    Set the next audio output available as default playback device.
    #>
    function Step-Output {
        $defaultOutputID = (Get-AudioDevice -List | Where-Object {$_.Type -eq "Playback" -and $_.Default -eq $true}).ID

        $outputIndex = $($global:outputs.outputs.keys).indexOf($defaultOutputID)

        if ($outputIndex -eq -1) {
            Notification "Unknown audio output currently used as default"
            Write-Error "Unknown audio output currently used as default"
            Exit
        }

        # Reorder the indexes in a new list, starting with the next device, so that it will go through all devices once (in the worst case)
        $reorganizedIndexes =  0..($global:outputs.outputs.count - 1) | %{ ($_ + $outputIndex + 1) % $global:outputs.outputs.count }

        # Cycling throught the list of output devices; if it's the last item, go to the beginning
        forEach ($i in $reorganizedIndexes) {
            $outputID = $global:outputs.outputs.Keys.Where({ $global:outputs.outputs[$PSItem] -eq $global:outputs.outputs[$i]; }, [System.Management.Automation.WhereOperatorSelectionMode]::First)

            if (Set-Output -ID $outputID) {
                return
            }
        }

        Notification "Unable to fetch audio output"
    }

    <#
    .SYNOPSIS
    Get a list of outputs containing all your propreties.
    #>
    function Get-OutputList {
        $(ForEach ($output in $global:outputs.GetEnumerator()) {
            $outputsCustomObject = [PSCustomObject] @{
                ID = $output.Name
                SystemName = $output.Value.SystemName
                DefaultVolume = $output.Value.DefaultVolume
                Nickname = $output.Value.Nickname
            }
            $outputsCustomObject
        }) | Format-List
    }

    <#
    .SYNOPSIS
    Print the interactive menu.
    #>
    function Show-MainMenu {
        $Option = Show-Menu @("Next output", "Set output", "Set volume", "List outputs", "Exit")
        switch ($Option) {
            "Next output" {
                Step-Output
                break
            }
            "Set output" {
                $OutputOptions = @($(Get-MenuSeparator))
                ForEach ($output in $global:outputs.GetEnumerator()) {
                    $OutputOptions += $output.Value.Nickname
                }
                $Option = Show-Menu $OutputOptions
                Set-Output -Nickname $Option | Out-Null
                break
            }
            "Set volume" {
                [int]$Volume = Read-Host "Volume level (0-100)"
                Set-OutputVolume $Volume
                break
            }
            "List outputs" {
                Get-OutputList
                break
            }
            "Exit" {
                Exit
            }
        }
    }
}

process {
    Initialize-Resources

    switch ($PSCmdlet.ParameterSetName) {
        "None" {
            Show-MainMenu
        }
        "NextOutput" {
            Step-Output
        }
        "SetOutput" {
            if ($PSBoundParameters.ContainsKey("Volume")) {
                Set-Output -Nickname $SetOutput -Volume $Volume | Out-Null
            } else {
                Set-Output -Nickname $SetOutput | Out-Null
            }
        }
        "SetVolume" {
            Set-OutputVolume $SetVolume
        }
        "ListOutput" {
            Get-OutputList
        }
    }
}
