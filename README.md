# GmaShellInfo
A shell extension for Windows Explorer which displays the metadata of .gma addon archive files used by Garry's Mod.

### Before
![Before](https://github.com/user-attachments/assets/51ef3617-ce57-4d2f-995b-ed4982eb805c)

### After
![After](https://github.com/user-attachments/assets/138087e9-4038-4ab6-abb0-ac2b1a740983)

The following shell properties are provided from the .gma metadata:
| .gma Metadata | Shell Property Mapping | Notes |
| --- | --- | --- |
| Name | Title | |
| Author | Authors | Very few .gma files contain this data |
| Description | Description | Typically, only old(er) format .gma files contain this data |
| Type | Categories | Windows already uses property "Type" for file type (i.e. extension) |
| Tags | Tags |

<br/>

# Setup
GmaShellInfo is compatible with Windows Vista and newer.

GmaShellInfo requires the Microsoft Visual C++ 2010 runtime. This is included in the GmaShellInfo installer and will automatically be installed, if needed.

### Installation
- Go to the [releases page](https://github.com/TiberiumFusion/GmaShellInfo/releases) and select a release, such as the [latest](https://github.com/TiberiumFusion/GmaShellInfo/releases/latest).
- Dowload the installer which matches [your system's architecture](https://www.howtogeek.com/21726/how-do-i-know-if-im-running-32-bit-or-64-bit-windows-answers/).
  - On 64-bit systems, download the GmaShellInfo_1.0.0_**x64**.exe installer. Most systems today are 64-bit.
  - On 32-bit systems, download the GmaShellInfo_1.0.0_**x86**.exe installer.
- Run the installer.
- GmaShellInfo will be activated as soon as the installer completes.

#### IMPORTANT
Shell extensions are architecture-specific.
- The 32-bit version of GmaShellInfo only works in 32-bit Windows Explorer, and the 64-bit version only works in 64-bit Windows Explorer.
- The 32-bit extension will *not* work in 64-bit Windows Explorer. However, if you are running 32-bit applications on your 64-bit system, you may need to additionally install the 32-bit version of GmaShellInfo in order for the .gma properties to appear in the various file browse/open/save/etc dialogs created by the 32-bit application.

### Usage Hints
Once installed, open Explorer and navigate to a folder containing .gma files.
- In the details view, right-click the column headers to add additional columns </br> ![ExportToGif](https://github.com/user-attachments/assets/ab62f71a-98a0-4ff8-aebc-db9775098cdc)
- Selecting a file will show all available properties in the properties panel... <br/> ![File properties panel](https://github.com/user-attachments/assets/cf7ae1ae-562f-4f5d-94c8-a10689551a65)
- ...and in the full property sheet <br/> ![Property sheet](https://github.com/user-attachments/assets/eedbf529-e253-4207-bd3e-8a1e77adf820)

### Troubleshooting
If you do not see the metadata of your .gma files:
- Reboot, or log off and log back on. This will restart Windows Explorer and the COM server it uses to host shell extensions.
- Check your .gma files in a hex editor. Some .gma files may be missing some or all metadata.
- Ensure you installed the correct version of GmaShellInfo for your system's architecture.

<br/>

# Building
This repository consists of a single Visual Studio 2017 solution containing all relevant projects. Requirements are listed below.
- `GmaShellPropertyHandler` is the Property Handler shell extension that constitutes GmaShellInfo. It is set up to build with the Visual Studio 2010 (v100) MSVC toolset and the Windows 7 SDK.
- `Installer.Msi.GmaShellInfo` and `Installer.Bundle` are [WiX](http://wixtoolset.org/) v3 projects that produce the release installers. The [WiX v3 toolset](https://github.com/wixtoolset/wix/releases) and Visual Studio [extension](https://marketplace.visualstudio.com/items?itemName=WixToolset.WiXToolset) must be installed to build these two projects.
- `WixCaShellAssocNotify` is a custom action for the WiX installer projects. Building it requires the v100 MSVC toolset, the Windows 7 SDK, and the WiX v3 toolset.

`Installer.Bundle` embeds the v100 SP1 MSVC redistributable installers to run during the GmaShellInfo installation. These vcredist installers are **not** present in this repository. The build will fail when these files are missing.
1. [Acquire them](https://github.com/abbodi1406/vcredist/blob/master/source_links/README.md#microsoft-visual-c-2010-redistributables---v10) from Microsoft.
2. Place the x86 installer at `Installer.Bundle\EmbedPrereqInstallers\vcredist_x86_2010_SP1.exe`
3. Place the x64 installer at `Installer.Bundle\EmbedPrereqInstallers\vcredist_x64_2010_SP1.exe`


<br />

# Credits
- `GmaShellPropertyHandler` uses code and is adapted from Microsoft's public sample **[PlaylistPropertyHandler](https://github.com/microsoft/Windows-classic-samples/tree/main/Samples/Win7Samples/winui/shell/appshellintegration/PlaylistPropertyHandler)**.
- `GmaShellPropertyHandler` uses the **[cJSON](https://github.com/DaveGamble/cJSON)** library.
