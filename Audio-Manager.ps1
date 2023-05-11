<#
.SYNOPSIS
Cmdlet that provides some tools to programmatically manage audio on Windows systems.

.DESCRIPTION
TBD

.PARAMETER NextOutput
Change the audio output to the next available output.

.INPUTS
TBD

.EXAMPLE
.\Audio-Manager.ps1 -NextOutput
Change the audio output to the next available output.

.NOTES
Author: Thales Pinto
Version: 0.1.0
Licence: This code is licensed under the MIT license.
#>

using module .\PlaybackAudioDevice.psm1

[CmdletBinding()]

param (
    [Parameter(Mandatory=$false, ParameterSetName="NextOutput")]
    [Alias("Next")]
    [switch]
    $NextOutput
)

begin {

    <#
    .SYNOPSIS
    Install the modules needed to Audio Manager works properly.
    #>
    function Resolve-Dependencies {
        $dependencies = "BurntToast", "AudioDeviceCmdlets"

        forEach ($dependency in $dependencies) {
            if ((Get-Module -ListAvailable | Where-Object {$_.Name -eq $dependency}) -eq $null) {
                Start-Process -wait powershell -Verb runAs -ArgumentList "Write-Host `"Install-Module -Name $dependency`" ; Install-Module -Name $dependency"
            }
        }
    }

    <#
    .SYNOPSIS
    Create a default profile JSON. If it already exists, read the information on it.
    #>
    function Initialize-Profile {
        $root = $PSScriptRoot
        $profileJson = "Profile.json"
        $profileJsonPath = Join-Path $root -ChildPath $profileJson

        if ((Test-Path -Path $profileJsonPath -PathType Leaf) -eq $false) {
            New-Item $profileJsonPath -ItemType File | Out-Null

            $outputs = [PSCustomObject] @{
                NotificationImagePath = ""
                Outputs = @()
            }

            ForEach ($output in $global:outputs) {
                $outputs.Outputs += @{
                    "Index" = $output.Index;
                    "Default" = $output.Default;
                    "Name" = $output.Name;
                    "Nickname" = $output.Nickname;
                    "ID" = $output.ID;
                    "DefaultVolume" = $output.DefaultVolume;
                    "Enabled" = $true
                }
            }

            $outputs | ConvertTo-Json | Out-File $profileJsonPath
        } else {
            $(Get-Content $profileJsonPath -Raw | ConvertFrom-Json).outputs | ForEach-Object {
                Update-AudioDevice -ID $_.ID -DefaultVolume $_.DefaultVolume -Nickname $_.Nickname -Enabled $_.Enabled
            }
            $NotificationImagePath = $(Get-Content $profileJsonPath -Raw | ConvertFrom-Json).NotificationImagePath
            if ($NotificationImagePath -ne "") {
                $global:settings.NotificationImagePath = $NotificationImagePath
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
    Create a list of PlaybackAudioDevice, containing all available playback devices.
    #>
    function Get-AudioDevices {
        $audioDevices = Get-AudioDevice -List | Where-Object {$_.Type -eq "Playback"}

        if ($audioDevices -eq $null) {
            Notification "No audio output available"
            Write-Error "No audio output available"
            exit
        }

        $outputs = @()
        ForEach ($device in $audioDevices) {
            $outputs += [PlaybackAudioDevice]::new($device)
        }
        return $outputs
    }

    <#
    .SYNOPSIS
    Remove a device by index, name or ID from the list.
    #>
    function Remove-AudioDevice {
        param (
            [Parameter(Mandatory=$true, ParameterSetName="Index")][int]$Index,
            [Parameter(Mandatory=$true, ParameterSetName="Name")][string]$Name,
            [Parameter(Mandatory=$true, ParameterSetName="ID")][string]$ID
        )

        switch ($PSBoundParameters.Keys) {
            "Index" { $global:outputs = $global:outputs | Where-Object {$_.Index -ne $Index} ; return }
            "Name" { $global:outputs = $global:outputs | Where-Object {$_.Name -NotLike "*$Name*"} ; return }
            "ID" { $global:outputs = $global:outputs | Where-Object {$_.ID -NotLike "*$ID*"} ; return }
        }
    }

    <#
    .SYNOPSIS
    Update device propreties.
    #>
    function Update-AudioDevice {
        param (
            [Parameter(Mandatory=$true)][String]$ID,
            [bool]$Enabled,
            [ValidateRange(0,100)][Int]$DefaultVolume,
            [String]$Nickname
        )

        $output = $global:outputs | Where-Object {$_.ID -Like "*$ID*"}

        if ($output -eq $null) {
            #Notification "Unable to find audio output"
            return
        }

        if (($PSBoundParameters.ContainsKey("Enabled")) -and ($Enabled -eq $false)) {
            $global:outputs = $global:outputs | Where-Object {$_.ID -NotLike "*$($ID)*"}
            return
        }

        $outputIndex = [array]::indexof($global:outputs, $output)

        if ($PSBoundParameters.ContainsKey("Nickname")) {
            $global:outputs[$outputIndex].Nickname = $Nickname
        }

        if ($PSBoundParameters.ContainsKey("DefaultVolume")) {
            $global:outputs[$outputIndex].DefaultVolume = $DefaultVolume
        }
    }

    <#
    .SYNOPSIS
    Change the audio device output based on the "index" parameter. Send a notification about the change.
    #>
    function Set-AudioOutput {
        param (
            [int]$index,
            [ValidateRange(0,100)][Int]$Volume
        )

        $audioDevice = Get-AudioDevice -ID $global:outputs[$index].ID

        if ($audioDevice -eq $null) {
            # TODO: Deal with this possible error
            return $null
        }

        $audioDevice | Set-AudioDevice | Out-Null

        if (-not ($PSBoundParameters.ContainsKey("Volume"))) {
            $Volume = $global:outputs[$index].DefaultVolume
        }

        Set-AudioDevice -PlaybackVolume $Volume

        Notification "Output: $($global:outputs[$index].Nickname)`nVolume: $Volume%"
        return $true
    }

    <#
    .SYNOPSIS
    Set the next audio output available as default playback device.
    #>
    function Step-AudioOutput {
        $defaultOutput = $global:outputs | Where-Object {$_.Default -eq $true}
        $outputIndex = [array]::indexof($global:outputs, $defaultOutput)

        # Reorder the indexes in a new list, starting with the next device, so that it will go through all devices once (in the worst case)
        $reorganizedIndexes =  0..($global:outputs.count - 1) | %{ ($_ + $outputIndex + 1) % $global:outputs.count }

        # Cycling throught the list of output devices; if it's the last item, go to the beginning
        foreach ($i in $reorganizedIndexes) {
            if (Set-AudioOutput $i) {
                return
            }
        }

        Notification "Unable to fetch audio output"
    }

}

process {
    $global:settings = @{}
    $global:settings.NotificationImagePath = "."
    $global:outputs = Get-AudioDevices
    Initialize-Profile
    if ($PSBoundParameters.ContainsKey("NextOutput")) {
        Step-AudioOutput
    }
}
