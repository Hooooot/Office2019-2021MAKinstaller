@echo off
fltmc > nul
if "%errorlevel%" NEQ "0" (goto UACPrompt) else (goto UACAdmin)
:UACPrompt
start mshta vbscript:createobject("shell.application").shellexecute("%~0","%~1 %~2 %~3 %~4 %~5 %~6 %~7 %~8 %~9",,"runas",1)(window.close)&exit
exit /B
:UACAdmin

set CurrentPath=%~dp0
cd /d "%CurrentPath%"

title Office 2019/2021 ProPlus MAK edition installer

@REM default setting

@REM 2019   2021
set edition=2021
@REM 1: will install, 0: not install
set excel=1
set powerPoint=1
set word=1
set outlook=0
set access=0
set groove=0
set lync=0
set oneDrive=0
set oneNote=0
set publisher=0
set visioPro=0
set projectPro=0
@REM zh-cn  en-us
set language=zh-cn



set ret=
goto main

:download_file
@REM argus: 
@REM   %1: URL
@REM   %2: file full path.
bitsadmin /transfer %~nx2 /download /priority high "%1"  "%2"

goto:eof


:reverse_var
    if "%1"=="0"  ( if "%edition%"=="2021" ( set edition=2019) else ( set edition=2021) )
    if "%1"=="1"  ( if "%excel%"=="0"      ( set excel=1) else ( set excel=0) )
    if "%1"=="2"  ( if "%powerPoint%"=="0" ( set powerPoint=1) else ( set powerPoint=0) )
    if "%1"=="3"  ( if "%word%"=="0"       ( set word=1) else ( set word=0) )
    if "%1"=="4"  ( if "%outlook%"=="0"    ( set outlook=1) else ( set outlook=0) )
    if "%1"=="5"  ( if "%access%"=="0"     ( set access=1) else ( set access=0) )
    if "%1"=="6"  ( if "%groove%"=="0"     ( set groove=1) else ( set groove=0) )
    if "%1"=="7"  ( if "%lync%"=="0"       ( set lync=1) else ( set lync=0) )
    if "%1"=="8"  ( if "%oneDrive%"=="0"   ( set oneDrive=1) else ( set oneDrive=0) )
    if "%1"=="9"  ( if "%oneNote%"=="0"    ( set oneNote=1) else ( set oneNote=0) )
    if "%1"=="10" ( if "%publisher%"=="0"  ( set publisher=1) else ( set publisher=0) )
    if "%1"=="11" ( if "%visioPro%"=="0"   ( set visioPro=1) else ( set visioPro=0) )
    if "%1"=="12" ( if "%projectPro%"=="0" ( set projectPro=1) else ( set projectPro=0) )
goto:eof

:choose_functions
@REM let user choose functions.
    cls
    call :show_current_functions
    :loop
        echo.

        set /p choosed_function=Input the index to choose your function: 
        set choosed_function=%choosed_function: =%

        if "%choosed_function%"=="15" ( goto :break )
        if "%choosed_function%"=="20" ( goto :break )
        if "%choosed_function%"=="21" ( goto :break )
        if "%choosed_function%"=="22" ( goto :break )
        if "%choosed_function%"=="23" ( goto :break )

        call :reverse_var %choosed_function%
        cls
        call :show_current_functions
    goto :loop
    :break
    set ret=%choosed_function%
goto:eof

:show_current_functions
                             echo Index     Function        Install(Y/N)
    if "%edition%"=="2021" ( echo   0.      Edition           2021) else ( echo   0.      Edition           2019)
    if "%excel%"=="0"      ( echo   1.      Excel               N ) else ( echo   1.      Excel               Y )
    if "%powerPoint%"=="0" ( echo   2.      PowerPoint          N ) else ( echo   2.      PowerPoint          Y )
    if "%word%"=="0"       ( echo   3.      Word                N ) else ( echo   3.      Word                Y )
    if "%outlook%"=="0"    ( echo   4.      Outlook             N ) else ( echo   4.      Outlook             Y )
    if "%access%"=="0"     ( echo   5.      Access              N ) else ( echo   5.      Access              Y )
    if "%groove%"=="0"     ( echo   6.      Groove              N ) else ( echo   6.      Groove              Y )
    if "%edition%"=="2021" (
    if "%lync%"=="0"       ( echo   7.      Teams               N ) else ( echo   7.      Teams               Y )
    ) else (
    if "%lync%"=="0"       ( echo   7.      Skype               N ) else ( echo   7.      Skype               Y )
    )
    if "%oneDrive%"=="0"   ( echo   8.      OneDrive            N ) else ( echo   8.      OneDrive            Y )
    if "%oneNote%"=="0"    ( echo   9.      OneNote             N ) else ( echo   9.      OneNote             Y )
    if "%publisher%"=="0"  ( echo   10.     Publisher           N ) else ( echo   10.     Publisher           Y )
    if "%visioPro%"=="0"   ( echo   11.     VisioPro            N ) else ( echo   11.     VisioPro            Y )
    if "%projectPro%"=="0" ( echo   12.     ProjectPro          N ) else ( echo   12.     ProjectPro          Y )
                             echo.
                             echo   15.     Exit  
                             echo.
                             echo   20      Start installation ^(Manual^)
                             echo   21      Start installation ^(Auto^)
                             echo   22      KMS active
                             echo   23      Start installation with KMS active
goto:eof

:generate_config
@REM Generate config file.
    set config_file_name=configuration-Office%edition%Enterprise.xml

    echo ^<Configuration^> > %config_file_name%
    echo. >> %config_file_name%
    echo  ^<Add OfficeClientEdition="64" Channel="PerpetualVL%edition%"^> >> %config_file_name%
    echo    ^<Product ID="ProPlus%edition%Volume"^> >> %config_file_name%
    echo      ^<Language ID="%language%" /^> >> %config_file_name%

    if "%excel%"=="0"      ( echo      ^<ExcludeApp ID="Excel" /^> >> %config_file_name% )
    if "%powerPoint%"=="0" ( echo      ^<ExcludeApp ID="PowerPoint" /^> >> %config_file_name% )
    if "%word%"=="0"       ( echo      ^<ExcludeApp ID="Word" /^> >> %config_file_name% )
    if "%outlook%"=="0"    ( echo      ^<ExcludeApp ID="Outlook" /^> >> %config_file_name% )
    if "%access%"=="0"     ( echo      ^<ExcludeApp ID="Access" /^> >> %config_file_name% )
    if "%groove%"=="0"     ( echo      ^<ExcludeApp ID="Groove" /^> >> %config_file_name% )
    if "%lync%"=="0"       ( echo      ^<ExcludeApp ID="Lync" /^> >> %config_file_name% )
    if "%oneDrive%"=="0"   ( echo      ^<ExcludeApp ID="OneDrive" /^> >> %config_file_name% )
    if "%oneNote%"=="0"    ( echo      ^<ExcludeApp ID="OneNote" /^> >> %config_file_name% )
    if "%publisher%"=="0"  ( echo      ^<ExcludeApp ID="Publisher" /^> >> %config_file_name% )
    echo    ^</Product^> >> %config_file_name%
    echo. >> %config_file_name%
    if "%visioPro%"=="1" (
        echo    ^<Product ID="VisioPro%edition%Volume"^> >> %config_file_name%
        echo      ^<Language ID="%language%" /^> >> %config_file_name%
        echo    ^</Product^> >> %config_file_name%
    )
    if "%projectPro%"=="1" (
        echo    ^<Product ID="ProjectPro%edition%Volume"^> >> %config_file_name%
        echo      ^<Language ID="%language%" /^> >> %config_file_name%
        echo    ^</Product^> >> %config_file_name%
    )
    echo  ^</Add^> >> %config_file_name%
    echo. >> %config_file_name%
    echo ^</Configuration^> >> %config_file_name%
goto:eof


:install_office
if exist "C:\Program Files\Microsoft Office\root" (
    echo You have already installed office, please uninstall it firstly.
    pause
    exit /B
)

if exist "C:\Program Files (x86)\Microsoft Office\root" (
    echo You have already installed office, please uninstall it firstly.
    pause
    exit /B
)
start /wait setup.exe /configure configuration-Office%edition%Enterprise.xml
goto:eof

:release_deployment_tool
    powershell.exe -ExecutionPolicy Unrestricted ^
-Command ^"^& { (New-Object System.Net.WebClient).DownloadString('https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117') -match ^
'url^=^(^?^<url^>https://.*/^(^?^<fileName^>officedeploymenttool.*exe))' ^> $null; $filePath='%CurrentPath%'+$Matches.fileName; Invoke-WebRequest -Uri $Matches.url -OutFile $filePath; ^
start $filePath -ArgumentList /extract:%1, /quiet -Wait; del $filePath }^"
goto:eof


@REM
:kms_active
@REM https://docs.microsoft.com/zh-cn/DeployOffice/vlactivation/gvlks?redirectedfrom=MSDN
    if not exist "C:\Program Files\Microsoft Office\Office16" (
        echo You did not install the office.
        goto:eof
    )
    if exist "C:\Program Files\Microsoft Office\root\Licenses16\ProPlus2021VL_KMS_Client_AE-ppd.xrm-ms" (
        set edition=2021
    ) else (
        set edition=2019
    )
    if "%edition%"=="2019" (
        set kms_key=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
    ) else (
        set kms_key=FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
    )

    cd /d "C:\Program Files\Microsoft Office\Office16"
    cscript ospp.vbs /inpkey:%kms_key%

    if "%edition%"=="2019" (
        set kms_key_visio=9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
    ) else (
        set kms_key_visio=KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4
    )
    cscript ospp.vbs /inpkey:%kms_key_visio%


    if "%edition%"=="2019" (
        set kms_key_project=B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
    ) else (
        set kms_key_project=FTNWT-C6WBT-8HMGF-K9PRX-QV9H8
    )
    cscript ospp.vbs /inpkey:%kms_key_project%
    
    cscript ospp.vbs /sethst:kms.03k.org
    cscript ospp.vbs /act
    @REM cscript ospp.vbs /dstatus
goto:eof


@REM main 
:main
set extractFolder=office_setup
set setupPath="%CurrentPath%%extractFolder%"


if not exist %setupPath% ( 
    @REM https://www.microsoft.com/en-us/download/details.aspx?id=49117
    echo Downloading deployment tool, please waiting...^(around 5 MiB^)
    call :release_deployment_tool %extractFolder%
)
if not exist %setupPath% (
    echo Download failed.
    echo Please install officedeploymenttool, press any key to exit.
    pause >nul
    exit /B
)
cd %setupPath%
call :choose_functions
set user_choice=%ret%
if "%user_choice%"=="15" (
    exit /B
)
if "%user_choice%"=="22" (
    call :kms_active
    pause
    exit /B
)
call :generate_config
echo Start downloading the newest office.
echo Downloading, please wait...^(around 2.1 GiB, start time %date:~0,10% %time%^)
start /wait setup.exe /download configuration-Office%edition%Enterprise.xml
timeout /T 2 /NOBREAK
cls
echo Download successfully!
if "%user_choice%"=="20" (
    echo Press any key to install the office.
    pause > nul
) else (
    echo Start installing the office.
)
call :install_office
if "%user_choice%"=="23" (
    echo Install successfully, press any key to set kms active.
    pause 
    call :kms_active
    echo Press any key to exit.
) else (
    echo Install successfully, press any key to exit.
)
pause
exit /B
