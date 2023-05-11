class PlaybackAudioDevice {
    [int]$Index
    [Bool]$Default
    [String]$Name
    [String]$ID
    [int]$DefaultVolume
    [String]$Nickname
    [Bool]$Enabled

    PlaybackAudioDevice([System.Object]$AudioDevice) {
        $this.Index = $AudioDevice.Index
        $this.Default = $AudioDevice.Default
        $this.Name = $AudioDevice.Name
        $this.ID = $AudioDevice.ID
        $this.DefaultVolume = 20
        $this.Nickname = $AudioDevice.Name
        $this.Enabled = $True
    }
}
