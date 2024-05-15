@echo off
cd /d "%~dp0"
if exist "ucrt\*.mum" if exist "ucrt\*ucrt*.manifest" (
echo Notice: ucrt directory is already found
echo remove it to create fresh one
pause
exit /b
)
if not exist "*KB3118401*.msu" (
echo Error: ucrt msu files are not found
pause
exit /b
)
call :Work 1>nul 2>nul
del /f /q *.cab 1>nul 2>nul
echo Finished
pause
exit /b

:Work
if not exist ucrt\ mkdir .\ucrt
expand.exe -f:*Windows*.cab *KB3118401*.msu .\ucrt
expand.exe -f:*Windows*.cab *KB4132941*.msu .\ucrt
cd ucrt\

>nul expand.exe -f:* Windows8.1-KB3118401-x64.cab .\
ren update.mum 9600-x64.mum
ren update.cat 9600-x64.cat
>nul expand.exe -f:* Windows8.1-KB3118401-x86.cab .\
ren update.mum 9600-x86.mum
ren update.cat 9600-x86.cat

>nul expand.exe -f:* Windows8-RT-KB3118401-x64.cab .\
ren update-bf.mum 9200-x64.mum
ren update-bf.cat 9200-x64.cat
>nul expand.exe -f:* Windows8-RT-KB3118401-x86.cab .\
ren update-bf.mum 9200-x86.mum
ren update-bf.cat 9200-x86.cat
del /f /q update.* *6.2.9200.1*.manifest *kb3118401~*6.2*.* *kb3118401_rtm~*6.2*.*
for /f %%# in ('dir /b /ad *6.2.9200.1*') do rmdir /s /q %%#\

>nul expand.exe -f:* Windows6.1-KB3118401-x64.cab .\
ren update-bf.mum 7601-x64.mum
ren update-bf.cat 7601-x64.cat
>nul expand.exe -f:* Windows6.1-KB3118401-x86.cab .\
ren update-bf.mum 7601-x86.mum
ren update-bf.cat 7601-x86.cat
del /f /q update.* *6.1.7601.1*.manifest *kb3118401~*6.1*.* *kb3118401_sp1~*6.1*.*
for /f %%# in ('dir /b /ad *6.1.7601.1*') do rmdir /s /q %%#\

>nul expand.exe -f:* Windows6.0-KB4132941-x64.cab .\
ren update.mum 6002-x64.mum
ren update.cat 6002-x64.cat
>nul expand.exe -f:* Windows6.0-KB4132941-x86.cab .\
ren update.mum 6002-x86.mum
ren update.cat 6002-x86.cat

if exist "6002-*.mum" exit /b
>nul expand.exe -f:* Windows6.0-KB3118401-x64.cab .\
ren update-bf.mum 6002-x64.mum
ren update-bf.cat 6002-x64.cat
>nul expand.exe -f:* Windows6.0-KB3118401-x86.cab .\
ren update-bf.mum 6002-x86.mum
ren update-bf.cat 6002-x86.cat
del /f /q update.* *6.0.6002.1*.manifest *kb3118401~*6.0*.* *kb3118401_client~*.* *kb3118401_server~*.* *kb3118401_sc~*.* *kb3118401_client_2~*.* *kb3118401_server_1~*.* *kb3118401_sc_1~*.*
for /f %%# in ('dir /b /ad *6.0.6002.1*') do rmdir /s /q %%#\

exit /b
