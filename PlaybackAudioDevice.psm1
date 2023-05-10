class PlaybackAudioDevice {
    [int]$Index
    [Bool]$Default
    [String]$Name
    [String]$ID
    [int]$DefaultVolume

    PlaybackAudioDevice([System.Object]$AudioDevice) {
        $this.Index = $AudioDevice.Index
        $this.Default = $AudioDevice.Default
        $this.Name = $AudioDevice.Name
        $this.ID = $AudioDevice.ID
        $this.DefaultVolume = 20
    }
}
