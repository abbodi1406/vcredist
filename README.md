# VisualCppRedist AIO

## Overview:

- AIO Repack for latest Microsoft Visual C++ Redistributable Runtimes, without the original setup bloat payload.

- Built upon VBCRedist_AIO_x86_x64.exe by **@ricktendo64**

- The process is handled by a windows command script, which runs hidden in the background by default.

- Before installation, the script will check and remove existing non-compliant Visual C++ Runtimes, including the original EXE or MSI setups, or older MSI packages versions.

- The uninstallation option/script will remove any detected VC++ runtimes (except UCRT).

- Custom AIO v35 packs for Windows XP XP SP3 / NT 5.2 support all the features and switches of recent AIO pack.  
old general AIO v35 pack itself has partial support.

- You can extract the installer file with 7-zip or WinRar to a short path, and run Installer.cmd as administrator

## Contents:

- Visual C++ Redistributables (x86/x64)  
2005: 8.0.50727.6229  
2008: 9.0.30729.7523  
2010: 10.0.40219.473  
2012: 11.0.61135.400  
2013: 12.0.40664.0  
2022: Latest  
2026: Insider

- Visual Studio 2010 Tools for Office Runtime (x86/x64)  
10.0.60922

- Legacy Runtimes (x86)  
Visual C++ 2002: 7.0.9975.0  
Visual C++ 2003: 7.10.6119.0  
Visual Basic Runtimes  

- Universal CRT:  
a complementary part of VC++ 2022 redist.  
inbox component for Windows 10/11.  
delivered as an update for Windows Vista/7/8/8.1, either in Monthly Quality Rollup, KB3118401, or KB2999226.  
installed with VC++ 2019 redist for Windows XP.  
this repack will install KB3118401 if UCRT is not available.  

- VC++ 2022 runtimes are binary compatible with VC++ 2015-2017-2019 and cover all VS 2015-2017-2019-2022 programs.

- VC++ 2026 runtimes for Windows 10/11 are binary compatible with all VC++ 2015-2017-2019-2022.

- Note:  
by design, Microsoft Windows Installer creates restore point for each msi package, if System Restore is active.

## Windows 7/8/8.1 Notice:

* VC++ 2022 version 14.44.35211.0 = [VisualCppRedist_AIO v0.92.0](https://github.com/abbodi1406/vcredist/releases/tag/v0.92.0) will the last version compatible with Windows 7/8/8.1

newer servicing versions of the same 14.44.35xxx.x range might get released in future. The AIO repack will be updated accordingly by then.

## Windows Vista Notice:

* VC++ 2022 version 14.32.31332.0 = [VisualCppRedist_AIO v0.61.0](https://github.com/abbodi1406/vcredist/releases/tag/v0.61.0) is the last version compatible with Windows Vista

## Windows XP Notice:

* VC++ 2019 version 14.28.29213.0 = [VisualCppRedist_AIO v0.35.0](https://github.com/abbodi1406/vcredist/releases/tag/v0.35.0) is the last version compatible with Windows XP

## Credits:

- [@ricktendo64](https://forums.mydigitallife.net/members/28038/) / MDL forums - repacks.net - wincert.net  
VBCRedist_AIO_x86_x64.exe creator,  modded MSI installers

- [@burfadel](https://forums.mydigitallife.net/members/84828/) / MDL forums - @thatguy91 / guru3D Forums  
original installation script

- Visual Basic and Visual C++ are registered trademarks of Microsoft Corporation.

## Unattended switches:

- For command-line options and examples, run:  
`VisualCppRedist_AIO_x86_x64.exe /?`

```
Usage:  
VisualCppRedist_AIO_x86_x64.exe [switches]

All switches are optional, case-sensitive.

/y  
Passive mode, shows progress. *All* Runtime packages are installed.

/ai  
Quiet mode, no output shown. *All* Runtime packages are installed.

/aiA  
Quiet mode. *All* Runtime packages are installed, and hide ARP entries.

/ai5  
Quiet mode. *Only* 2005 package is installed.

/ai8  
Quiet mode. *Only* 2008 package is installed.

/aiX  
Quiet mode. *Only* 2010 package is installed.

/ai2  
Quiet mode. *Only* 2012 package is installed.

/ai3  
Quiet mode. *Only* 2013 package is installed.

/ai9  
Quiet mode. *Only* 2022 package is installed.

/aiT  
Quiet mode. *Only* VSTOR 2010 package is installed.

/aiE  
Quiet mode. *Only* Extra VB/C package is installed. 
 
/aiB  
Quiet mode. *Only* Extra VB package is installed.

/aiC  
Quiet mode. *Only* Extra VC package is installed.

/aiV  
Quiet mode. *Only* VC++ packages are installed.

/aiM  
Manual Install mode, shows installation script with prompt.

/aiR  
Auto Uninstall mode, remove all detected runtimes.

/aiD  
Debug mode, create VCpp_debug.log without installing/uninstalling any package.

/aiP  
Manual Hide or Show Runtimes entries in Add/Remove Programs panel.

/ai1  
Update mode. Only already installed packages are updated.

/aiF  
Repair mode. Only already installed packages are reinstalled or updated.

/gm2  
Optional switch to disable extraction dialog for all other switches.

/sfxlang:  
Set the program display language, if possible. Example: /sfxlang:1031

/h | /?  
Display this help.
```
```
Examples:

Automatically install all packages and display progress:  
VisualCppRedist_AIO_x86_x64.exe /y

Silently install all packages and display no progress:  
VisualCppRedist_AIO_x86_x64.exe /ai /gm2

Silently install 2022 package:  
VisualCppRedist_AIO_x86_x64.exe /ai9

Silently install 2010/2012/2013/2022 packages:  
VisualCppRedist_AIO_x86_x64.exe /aiX239

Silently install VSTOR and Extra VB/C packages:  
VisualCppRedist_AIO_x86_x64.exe /aiTE

Silently install all packages and hide ARP entries:  
VisualCppRedist_AIO_x86_x64.exe /aiA /gm2
```

- **/y** gives the same default behavior, but without the beginning prompt and the finish message  

- only **/sfxlang** and **/gm2** can be specified with other switches  
if other switches are specified together, only the last one will have an effect. Example, this will only install Extra VB/C package:  
`/ai5 /ai8 /aiT /aiE`

- to install separate packages together, combine their latest switch character after **/ai** , Example:  
`/ai58X239E`

- **/sfxlang** must be the first switch to have effect. Example:  
`/sfxlang:1031 /aiV`

## Download

- Latest release zip file:  
https://kutt.it/vcppredist
- Latest release exe file:  
https://kutt.it/vcpp
- All releases:  
https://github.com/abbodi1406/vcredist/releases  
https://tiny.cc/vcredist
