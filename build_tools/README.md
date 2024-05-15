# Build Tools

* Sample tools and description for preparing VisualCppRedist AIO's msi packages.

## Requirements

- VBScript files to modify and slim msi files (created by dumpydooby, modded by ricktendo64).
- WiSumInf.vbs to update msi summary information stream (part of Windows SDK Windows Installer utility scripts).
- [WiX Toolset v3](https://github.com/wixtoolset/wix3/releases/) to extract VC++ 2012 and later Bootstrappers, and build msi files for legacy VB/C runtimes.
- [7zSfxMod](https://github.com/chrislake/7zsfxmm) - [7z SFX Modified Module](http://forum.oszone.net/showthread.php?t=51547) to build the AIO executable installer.

## General Steps

- Place the required files (original exe) per version in its folder.
- Open Command Prompt as administrator in the same folder location.
- Extract the original VC++ redistributables.
- Optionally, remove all the extracted files except msi and cab files (and msp file for VC++ 2010).
- Run the vbs script to slim the msi database.
- Create administrative installation for the modded msi to get rid of the internal unneeded files, and/or reduce the overall 7z AIO archive.

## WiX Tip

- If not already set, add WiX binaries folder to **PATH** environment variable for easier usage

Example, global system path:  
`setx PATH "W:\GitHub\dotNetFx4xW7\BIN;%PATH%" /M`

Example, per cmd session:  
`set "PATH=W:\GitHub\dotNetFx4xW7\BIN;%PATH%"`

- Supported compression levels **dcl** for light.exe command:  
`none, low, mszip, medium, high`

## VC++ 2005

- Extract
```
start /w vcredist_x64.exe /Q /C /T:"%cd%\vc64"
start /w vcredist_x86.exe /Q /C /T:"%cd%\vc86"
```
- Modify
```
cscript vc08.vbs vc64\vcredist.msi
cscript vc08.vbs vc86\vcredist.msi
```
- Admin install
```
start /w msiexec.exe /a vc64\vcredist.msi /quiet TARGETDIR="%cd%\2005\x64"
start /w msiexec.exe /a vc86\vcredist.msi /quiet TARGETDIR="%cd%\2005\x86"
rmdir /s /q vc64\ vc86\
```

## VC++ 2008

- Extract
```
start /w vcredist_x64.exe /quiet /extract:"%cd%\vc64"
start /w vcredist_x86.exe /quiet /extract:"%cd%\vc86"
```
- Modify
```
cscript vc09.vbs vc64\vc_red.msi
cscript vc09.vbs vc86\vc_red.msi
```
- Admin install
```
start /w msiexec.exe /a vc64\vc_red.msi /quiet TARGETDIR="%cd%\2008\x64"
start /w msiexec.exe /a vc86\vc_red.msi /quiet TARGETDIR="%cd%\2008\x86"
rmdir /s /q vc64\ vc86\
```

## VC++ 2010

- Extract
```
start /w vcredist_x64.exe /quiet /extract:"%cd%\tmp"
robocopy /NJH /NJS tmp\ vc10\x64\ *.cab *.msi *.msp
rmdir /s /q tmp\

start /w vcredist_x86.exe /quiet /extract:"%cd%\tmp"
robocopy /NJH /NJS tmp\ vc10\x86\ *.cab *.msi *.msp
rmdir /s /q tmp\
```
- Modify
```
cscript vc10.vbs vc10\x64\vc_red.msi
cscript vc10.vbs vc10\x86\vc_red.msi
```
- Admin install
```
for /f "tokens=2* delims== " %a in ('cscript WiSumInf.vbs vc10\x64\vc_red.msi ^| findstr /i Subject') do set name="%b"
for /f "tokens=2* delims== " %a in ('cscript WiSumInf.vbs vc10\x64\vc_red.msi ^| findstr /i Comments') do set desc="%b"
start /w msiexec.exe /a vc10\x64\vc_red.msi /quiet TARGETDIR="%cd%\vc10\z64"
start /w msiexec.exe /a vc10\z64\vc_red.msi /quiet TARGETDIR="%cd%\2010\x64" PATCH="%cd%\vc10\x64\msp_kb2890375.msp"
cscript WiSumInf.vbs vc10\z64\vc_red.msi Subject=%name% Comments=%desc%
move /y vc10\z64\vc_red.msi 2010\x64\

for /f "tokens=2* delims== " %a in ('cscript WiSumInf.vbs vc10\x86\vc_red.msi ^| findstr /i Subject') do set name="%b"
for /f "tokens=2* delims== " %a in ('cscript WiSumInf.vbs vc10\x86\vc_red.msi ^| findstr /i Comments') do set desc="%b"
start /w msiexec.exe /a vc10\x86\vc_red.msi /quiet TARGETDIR="%cd%\vc10\z86"
start /w msiexec.exe /a vc10\z86\vc_red.msi /quiet TARGETDIR="%cd%\2010\x86" PATCH="%cd%\vc10\x86\msp_kb2890375.msp"
cscript WiSumInf.vbs vc10\z86\vc_red.msi Subject=%name% Comments=%desc%
move /y vc10\z86\vc_red.msi 2010\x86\

rmdir /s /q vc10\
```

## VC++ 2012

- Extract
```
dark.exe vcredist_x64.exe -x "%cd%\vc64"
dark.exe vcredist_x86.exe -x "%cd%\vc86"
```
- Modify
```
cscript vc11.vbs vc64\AttachedContainer\packages\vcRuntimeMinimum_amd64\vc_runtimeMinimum_x64.msi
cscript vc11.vbs vc64\AttachedContainer\packages\vcRuntimeAdditional_amd64\vc_runtimeAdditional_x64.msi

cscript vc11.vbs vc86\AttachedContainer\packages\vcRuntimeMinimum_x86\vc_runtimeMinimum_x86.msi
cscript vc11.vbs vc86\AttachedContainer\packages\vcRuntimeAdditional_x86\vc_runtimeAdditional_x86.msi
```
- Admin install
```
start /w msiexec.exe /a vc64\AttachedContainer\packages\vcRuntimeMinimum_amd64\vc_runtimeMinimum_x64.msi /quiet TARGETDIR="%cd%\2012\x64"
start /w msiexec.exe /a vc64\AttachedContainer\packages\vcRuntimeAdditional_amd64\vc_runtimeAdditional_x64.msi /quiet TARGETDIR="%cd%\2012\x64"

start /w msiexec.exe /a vc86\AttachedContainer\packages\vcRuntimeMinimum_x86\vc_runtimeMinimum_x86.msi /quiet TARGETDIR="%cd%\2012\x86"
start /w msiexec.exe /a vc86\AttachedContainer\packages\vcRuntimeAdditional_x86\vc_runtimeAdditional_x86.msi /quiet TARGETDIR="%cd%\2012\x86"

rmdir /s /q vc64\ vc86\
```

## VC++ 2013

- Extract
```
dark.exe vcredist_x64.exe -x "%cd%\vc64"
dark.exe vcredist_x86.exe -x "%cd%\vc86"
```
- Modify
```
cscript vc12.vbs vc64\AttachedContainer\packages\vcRuntimeMinimum_amd64\vc_runtimeMinimum_x64.msi
cscript vc12.vbs vc64\AttachedContainer\packages\vcRuntimeAdditional_amd64\vc_runtimeAdditional_x64.msi

cscript vc12.vbs vc86\AttachedContainer\packages\vcRuntimeMinimum_x86\vc_runtimeMinimum_x86.msi
cscript vc12.vbs vc86\AttachedContainer\packages\vcRuntimeAdditional_x86\vc_runtimeAdditional_x86.msi
```
- Admin install
```
start /w msiexec.exe /a vc64\AttachedContainer\packages\vcRuntimeMinimum_amd64\vc_runtimeMinimum_x64.msi /quiet TARGETDIR="%cd%\2013\x64"
start /w msiexec.exe /a vc64\AttachedContainer\packages\vcRuntimeAdditional_amd64\vc_runtimeAdditional_x64.msi /quiet TARGETDIR="%cd%\2013\x64"

start /w msiexec.exe /a vc86\AttachedContainer\packages\vcRuntimeMinimum_x86\vc_runtimeMinimum_x86.msi /quiet TARGETDIR="%cd%\2013\x86"
start /w msiexec.exe /a vc86\AttachedContainer\packages\vcRuntimeAdditional_x86\vc_runtimeAdditional_x86.msi /quiet TARGETDIR="%cd%\2013\x86"

rmdir /s /q vc64\ vc86\
```

## VC++ 2015-2022

- Extract
```
dark.exe VC_redist.x64.exe -x "%cd%\vc64"
dark.exe VC_redist.x86.exe -x "%cd%\vc86"
```
- Modify
```
cscript vc14.vbs vc64\AttachedContainer\packages\vcRuntimeMinimum_amd64\vc_runtimeMinimum_x64.msi
cscript vc14.vbs vc64\AttachedContainer\packages\vcRuntimeAdditional_amd64\vc_runtimeAdditional_x64.msi

cscript vc14.vbs vc86\AttachedContainer\packages\vcRuntimeMinimum_x86\vc_runtimeMinimum_x86.msi
cscript vc14.vbs vc86\AttachedContainer\packages\vcRuntimeAdditional_x86\vc_runtimeAdditional_x86.msi
```
- Admin install
```
start /w msiexec.exe /a vc64\AttachedContainer\packages\vcRuntimeMinimum_amd64\vc_runtimeMinimum_x64.msi /quiet TARGETDIR="%cd%\2022\x64"
start /w msiexec.exe /a vc64\AttachedContainer\packages\vcRuntimeAdditional_amd64\vc_runtimeAdditional_x64.msi /quiet TARGETDIR="%cd%\2022\x64"

start /w msiexec.exe /a vc86\AttachedContainer\packages\vcRuntimeMinimum_x86\vc_runtimeMinimum_x86.msi /quiet TARGETDIR="%cd%\2022\x86"
start /w msiexec.exe /a vc86\AttachedContainer\packages\vcRuntimeAdditional_x86\vc_runtimeAdditional_x86.msi /quiet TARGETDIR="%cd%\2022\x86"

rmdir /s /q vc64\ vc86\
```

## VSTOR 2010

- Extract
```
start /w vstor_redist.exe /quiet /extract:"%cd%\tmp"
start /w tmp\vstor40\vstor40_x64.exe /quiet /extract:"%cd%\vc64"
start /w tmp\vstor40\vstor40_x86.exe /quiet /extract:"%cd%\vc86"
rmdir /s /q tmp\
```
- Modify
```
cscript vstor40.vbs vc64\vstor40_x64.msi
cscript vstor40.vbs vc86\vstor40_x86.msi
```
- Admin install
```
start /w msiexec.exe /a vc86\vstor40_x86.msi /quiet TARGETDIR="%cd%\vstor"
start /w msiexec.exe /a vc64\vstor40_x64.msi /quiet TARGETDIR="%cd%\vstor"
rmdir /s /q vc64\ vc86\
```

## Legacy VB/C++

- Extract **VBCRun.7z**
- Build VB/C++ MSI
```
candle.exe vbcrun.wxs
light.exe vbcrun.wixobj -spdb -sice:ICE09 -dcl:none
```
- Build VC++ MSI
```
candle.exe vcrun.wxs
light.exe vcrun.wixobj -spdb -sice:ICE09 -dcl:none
```
- Build VB MSI
```
candle.exe vbrun.wxs
light.exe vbrun.wixobj -spdb -sice:ICE09 -dcl:none
```
- Admin install
```
start /w msiexec.exe /a vbrun.msi /quiet TARGETDIR="%cd%\vbc"
start /w msiexec.exe /a vcrun.msi /quiet TARGETDIR="%cd%\vbc"
start /w msiexec.exe /a vbcrun.msi /quiet TARGETDIR="%cd%\vbc"
```

## Universal C Runtime / UCRT

- Download MSU files
```
http://download.windowsupdate.com/c/msdownload/update/software/updt/2016/02/windows8.1-kb3118401-x64_2d9f2a496d7a35dc5e68b541b7218ecf00a68108.msu
http://download.windowsupdate.com/c/msdownload/update/software/updt/2016/02/windows8.1-kb3118401-x86_a746ed4d040c315daca0b5b886d832ebec7b40f5.msu

http://download.windowsupdate.com/c/msdownload/update/software/updt/2016/02/windows8-rt-kb3118401-x64_704ddb69e2e8073d06f1b0905673c248f0d23d56.msu
Windows8-RT-KB3118401-x86.msu / extract from WindowsUCRT.zip: https://www.microsoft.com/en-us/download/details.aspx?id=50410

http://download.windowsupdate.com/d/msdownload/update/software/updt/2016/02/windows6.1-kb3118401-x64_99153d75ee4d103a429464cdd9c63ef4e4957140.msu
http://download.windowsupdate.com/c/msdownload/update/software/updt/2016/02/windows6.1-kb3118401-x86_db0267a39805ae9e98f037a5f6ada5b34fa7bdb2.msu

http://download.windowsupdate.com/c/msdownload/update/software/updt/2018/06/windows6.0-kb4132941-x64_20144f9f3a533aff2406761c0363b6a44108e358.msu
http://download.windowsupdate.com/c/msdownload/update/software/updt/2018/06/windows6.0-kb4132941-x86_448f787762f5a23499d669c4e584073e17303474.msu
```

- Group the msu files next to `UCRT.cmd` and run the script

## VisualCppRedist_AIO

- Move and group the created administrative installation directories into `_AIO` folder:
```
2005
2008
2010
2012
2013
2022
ucrt
vbc
vstor
```

along with these needed files:
```
7zSfx_x86_x64.cmd
7zSfx_x86only.cmd
7zSfxConfig.txt
7zSfxMod.sfx
ARP.cmd
Installer.cmd
Uninstaller.cmd
```

- To update `Installer.cmd` script with new runtime versions:  
run `MSIProductCode.vbs` against new msi files to obtain new ProductCode  
edit the script and go around line 180  
update the files version for `_verXX` variables  
update the following `code` variables

- Use a resource editor to update `File Version` and `Product Version` for **7zSfxMod.sfx** according to latest VC++ 14 version

- Run `7zSfx_x86_x64.cmd` and/or `7zSfx_x86only.cmd` scripts to create the AIO installers

the scripts are configured to use `7z.exe` installed at `"%ProgramFiles%\7-Zip"`  
if you have a different path or use a portable 7-Zip, adjust the 2nd line path accordingly

the switch `-bso0` require 7-Zip 15.01 or later

### Example Work Folder Tree

<details><summary>Spoiler</summary>


```
|   
|---build_tools
|   |   
|   |   README.md
|   |   
|   |---_AIO
|   |       7zSfxConfig.txt
|   |       7zSfxMod.sfx
|   |       7zSfx_x86only.cmd
|   |       7zSfx_x86_x64.cmd
|   |       MSIProductCode.vbs
|   |   
|   |---_m08
|   |       vc08.vbs
|   |       vcredist_x64.exe
|   |       vcredist_x86.exe
|   |       
|   |---_m09
|   |       vc09.vbs
|   |       vcredist_x64.exe
|   |       vcredist_x86.exe
|   |       
|   |---_m10
|   |       vc10.vbs
|   |       vcredist_x64.exe
|   |       vcredist_x86.exe
|   |       WiSumInf.vbs
|   |       
|   |---_m11
|   |       vc11.vbs
|   |       vcredist_x64.exe
|   |       vcredist_x86.exe
|   |       
|   |---_m12
|   |       vc12.vbs
|   |       vcredist_x64.exe
|   |       vcredist_x86.exe
|   |       
|   |---_m14
|   |       vc14.vbs
|   |       VC_redist.x64.exe
|   |       VC_redist.x86.exe
|   |       
|   |---_ucrt
|   |       UCRT.cmd
|   |       Windows6.0-KB4132941-x64.msu
|   |       Windows6.0-KB4132941-x86.msu
|   |       Windows6.1-KB3118401-x64.msu
|   |       Windows6.1-KB3118401-x86.msu
|   |       Windows8-RT-KB3118401-x64.msu
|   |       Windows8-RT-KB3118401-x86.msu
|   |       Windows8.1-KB3118401-x64.msu
|   |       Windows8.1-KB3118401-x86.msu
|   |       
|   |---_vbc
|   |       VBCRun.7z
|   |       
|   |---_vstor
|           vstor40.vbs
|           vstor_redist.exe
```
</details>
