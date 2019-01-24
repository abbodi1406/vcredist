# VisualCppRedist AIO

## Overview:

- AIO Repack for latest Microsoft Visual C++ Redistributable Runtimes, without the original setup bloat payload.

- Built upon VBCRedist_AIO_x86_x64.exe by **@ricktendo64**

- The process is handled by a windows command script, which runs hidden in the background by default.

- Before installation, the script will check and remove existing non-compliant Visual C++ Runtimes, including the original EXE or MSI setups, or older MSI packages versions.

- The uninstallation option/script will remove any detected VC++ runtimes (except UCRT).

- Windows XP support is partial, the pack will install and detect latest runtimes versions, but it will not check and remove non-compliant versions.

- You can extract the installer file with 7-zip or WinRar to a short path, and run Installer.cmd as administrator

## Contents:

- Visual C++ Redistributables (x86/x64)  
2005: 8.0.50727.6229  
2008: 9.0.30729.7523  
2010: 10.0.40219.473  
2012: 11.0.61135.400  
2013: 12.0.40664.0  
2017: 14.20.27305.0

- Visual Studio 2010 Tools for Office Runtime (x86/x64)  
10.0.60833.0

- Legacy Runtimes (x86)  
Visual C++ 2002: 7.0.9975.0  
Visual C++ 2003: 7.10.6119.0  
Visual Basic Runtimes  

- Universal CRT:  
complementary part of VC++ 2017 redist.  
inbox component for Windows 10.  
delivered as an update for Windows Vista/7/8/8.1, either in Monthly Quality Rollup, KB3118401, or KB2999226.  
installed with VC++ 2017 redist for Windows XP.  
this repack will install KB3118401 if UCRT is not available.  

- VC++ 2017 runtimes are binary compatible with VC++ 2015 and cover both VS 2015/2017 applications.

## Credits:

- [@ricktendo64](https://forums.mydigitallife.net/members/28038/) / MDL forums - repacks.net - wincert.net  
VBCRedist_AIO_x86_x64.exe creator,  modded MSI installers

- [@burfadel](https://forums.mydigitallife.net/members/84828/) / MDL forums - @thatguy91 / guru3D Forums  
original installation script

- Visual Basic and Visual C++ are registered trademarks of Microsoft Corporation.

## Unattended switches:

- For command-line options and examples, run:  
`VisualCppRedist_AIO_x86_x64.exe /?`

- **/y** give the same default behavior, but without the begin prompt and finnish message  

- only **/sfxlang** and **/gm2** can be specified with other switches  
if other switches specified together, only the latest will have effect. Example, this will only install Extra VB/C package:  
`/ai5 /ai8 /aiO /aiE`

- **/sfxlang** most be first switch to have effect. Example:  
`/sfxlang:1031 /aiV`

- VSTOR switch **/aiO** is letter O not zero 0  

## [Download](https://tiny.cc/vcredist)

