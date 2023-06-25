# Audio Manager

Cmdlet that provides some tools to programmatically manage audio on Windows systems.



## 🔧 Dependencies

Before running this script, ensure that the following modules are installed:
- [AudioDeviceCmdlets](https://github.com/frgnca/AudioDeviceCmdlets)
- [BurntToast](https://github.com/Windos/BurntToast)
- [PSMenu](https://github.com/Sebazzz/PSMenu)



## 💡 Usage

- **Next Output**: Change the audio output to the next available output.

```powershell
.\Audio-Manager.ps1 -NextOutput
```

- **Set Output**: Change the audio output to a specific device. You can provide the nickname of the output device, previously set in the profile file.

```powershell
.\Audio-Manager.ps1 -SetOutput "Nickname"
```

- **Set Volume**: Change the volume of the currently in-use output. Specify the desired volume level as a percentage (0-100).

```powershell
.\Audio-Manager.ps1 -SetVolume 80
```

- **List Outputs**: List all available audio output devices and their properties.

```powershell
.\Audio-Manager.ps1 -ListOutput
```



## 📋 Profile

The script uses a profile file named `Profile.json` to store information about output devices and their settings. If the profile file does not exist, the script creates a default profile with all available output devices enabled. You can modify the profile file manually to customize the settings.



## 📢 Notifications

The script sends notifications using the BurntToast module. Notifications will appear with the title "Audio Manager" and provide relevant information about the audio output or volume changes.



## 📃 License

This code is licensed under the MIT license. See the file LICENSE in the project root for full license information.
