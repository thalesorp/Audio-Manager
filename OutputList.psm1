class OutputList {
    $Outputs

    OutputList() {
        $this.Outputs = [Ordered]@{}
    }

    [void]Add([System.Object]$AudioDevice) {
        $this.Outputs.add($AudioDevice.ID, [Output]::new($AudioDevice))
    }

    [void]Update([System.Object]$OutputJsonData) {
        if (-not $this.Outputs.Contains($OutputJsonData.ID)) {
            return
        }
        if ($OutputJsonData.Enabled -eq $false) {
            $this.Outputs.Remove($OutputJsonData.ID)
            return
        }
        $this.Outputs[$OutputJsonData.ID].SystemName = $OutputJsonData.SystemName
        $this.Outputs[$OutputJsonData.ID].Nickname = $OutputJsonData.Nickname
        $this.Outputs[$OutputJsonData.ID].DefaultVolume = $OutputJsonData.DefaultVolume
    }

    [string] GetOutputID([string]$nickname) {
        return ($this.Outputs.GetEnumerator() | Where-Object {$_.Value.Nickname -eq $nickname}).Key
    }

    [System.Collections.IEnumerator] GetEnumerator() {
        return $this.Outputs.GetEnumerator()
    }

    [object[]] Values() {
        return $this.Outputs.Values
    }

}

class Output {
    [String]$SystemName
    [int]$DefaultVolume
    [String]$Nickname

    Output([System.Object]$AudioDevice) {
        $this.SystemName = $AudioDevice.Name
        $this.Nickname = $AudioDevice.Name
        $this.DefaultVolume = 20
    }

    Output([Output]$Output) {
        $this.Nickname = $Output.Nickname
        $this.DefaultVolume = $Output.DefaultVolume
    }

}
