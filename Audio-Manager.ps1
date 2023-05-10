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
    [Parameter(Mandatory=$true, ParameterSetName="NextOutput")]
    [Alias("Next")]
    [switch]
    $NextOutput
)

begin {

    $global:outputs = @()

    <#
    .SYNOPSIS
    Send notification with message.
    #>
    function Notification {
        param (
            [string]$Message
        )

        New-BurntToastNotification -AppLogo "." -Text "Audio Manager", $Message
    }

    <#
    .SYNOPSIS
    Create a list of PlaybackAudioDevice, containing all available playback devices.
    #>
    function Get-AudioDevices {
        $AudioDevices = Get-AudioDevice -List | Where-Object {$_.Type -eq "Playback"}
        ForEach ($device in $AudioDevices) {
            $global:outputs += [PlaybackAudioDevice]::new($device)
        }
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
            [Parameter(Mandatory=$true, ParameterSetName="Index")][int]$Index,
            [Parameter(Mandatory=$true, ParameterSetName="Name")][string]$Name,
            [Parameter(Mandatory=$true, ParameterSetName="ID")][string]$ID,
            [ValidateRange(0,100)][Int]$DefaultVolume
        )

        switch ($PSBoundParameters.Keys) {
            "Index" { $output = $global:outputs | Where-Object {$_.Index -eq $Index} ; break }
            "Name" { $output = $global:outputs | Where-Object {$_.Name -Like "*$Name*"} ; break }
            "ID" { $output = $global:outputs | Where-Object {$_.ID -Like "*$ID*"} ; break }
        }

        if ($output -eq $null) {
            #Notification "Unable to find audio output"
            return
        }

        $output.DefaultVolume = $DefaultVolume

        $outputIndex = [array]::indexof($global:outputs, $output)
        $global:outputs[$outputIndex] = $output
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

        Notification "Output: $($global:outputs[$index].Name)`nVolume: $Volume%"
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
    Get-AudioDevices

    if ($PSBoundParameters.ContainsKey("NextOutput")) {
        Step-AudioOutput
    }
}
