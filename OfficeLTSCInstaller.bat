<# :
    @echo off
    setlocal enabledelayedexpansion
    set arg="%~f0"
    for %%x in (%*) do set arg=!arg! /, "%%x"
    start /b powershell /nologo /noprofile /command ^
"Start-Process powershell -Verb RunAs '/nologo /noprofile /command ^
"""^&{ $ScriptFullPath="""""""""%~f0"""""""""; $ScriptName="""""""""%~xn0"""""""""; $ScriptPath="""""""""%~dp0"""""""""; Set-Location $ScriptPath; ^
    icm ([scriptblock]::Create((gc $ScriptFullPath -Raw))) -ArgumentList ("""""""""!arg!""""""""" -split """"""""" /, """""""""); }""" '"
    endlocal
    exit /B
#>

# write your powershell command here
$PSDefaultParameterValues['*:Encoding'] = 'utf8'
# Write-Host $ScriptFullPath
# Write-Host $ScriptPath
# Write-Host $ScriptName
# Write-Output $args

function GetLanguageString($stringID) {
    if ((Get-Culture).LCID -eq 2052) { # GBK
        switch ($stringID) {
            "inputOfficeVersion" { "请选择需要安装的Office版本序号" }
            "officeVersionIndexList" { "序号                版本" }
            "appInstallListHeader" {"序号           名称                    安装"}
            "Yes" { "是" }
            "No" { "否" }
            "StartInstall" { "开始安装" }
            "InputFunctionIndex" { "请输入要选择的功能序号" }
            "invalidIndex" { "无效的序号!" }
            "invalidInput" { "无效输入：" }
            "getExeUrl" { "正在获取Setup.exe链接......" }
            "showExeUrl" { "Setup.exe链接：" }
            "downloadExe" { "正在下载Setup.exe（大约4MiB）" }
            "downloadingOffice" { "正在下载Office..." }
            "installingOffice" { "正在安装Office..." }
            "defaultStr" { "（默认）" }
            Default { "Invalid language string!" }
        }
    } else {
        switch ($stringID) {
            "inputOfficeVersion" { "Please select the Office version index that needs to be installed" }
            "officeVersionIndexList" { "Index                Version" }
            "appInstallListHeader" { "Index           Name                     Install" }
            "Yes" { "Yes" }
            "No" { "No" }
            "StartInstall" { "Start Install" }
            "InputFunctionIndex" { "Input the index to select function" }
            "invalidIndex" { "Invalid index!" }
            "invalidInput" { "Invalid input: " }
            "getExeUrl" { "Getting Setup.exe url..." }
            "showExeUrl" { "Setup.exe url: " }
            "downloadExe" { "Downloading Setup.exe (around 4MiB)." }
            "downloadingOffice" { "Downloading Office..." }
            "installingOffice" { "Installing Office..." }
            "defaultStr" { "(Default)" }
            Default { "Invalid language string!" }
        }
    }
}

function ReleaseODT($extractFolder) {
    Write-Host $(GetLanguageString "getExeUrl") -NoNewLine
    $webContent = (New-Object System.Net.WebClient).DownloadString('https://www.microsoft.com/en-us/download/details.aspx?id=49117')
    $webContent -match '"url":"(?<url>https://.*(?<fileName>officedeploymenttool.*exe))",' > $null
    $filePath = $CurrentPath + "\" +$Matches.fileName
    Write-Host ""
    Write-Host "$(GetLanguageString "showExeUrl")$($Matches.url)"
    Write-Host $(GetLanguageString "downloadExe")
    Invoke-WebRequest -Uri $Matches.url -OutFile $filePath
    Start-Process $filePath -ArgumentList "/extract:$extractFolder /quiet" -Wait
    Remove-Item "$extractFolder*.xml"
    Remove-Item $filePath
}


$configCommon = @"
<Configuration>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
</Configuration>
"@

# https://config.office.com/deploymentsettings
$configMap = @(
    @{
        "Name" = "Office LTSC Professional Plus 2024";
        "Channel" = "PerpetualVL2024";
        "Product" = @(
            @{"ID" = "ProPlus2024Volume"; "PIDKEY" = "XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB";    "Install" = $true };
            @{"ID" = "VisioPro2024Volume"; "PIDKEY" = "B7TN8-FJ8V3-7QYCP-HQPMV-YY89G";   "Install" = $false; "Name" = "Visio Pro 2024  " };
            @{"ID" = "ProjectPro2024Volume"; "PIDKEY" = "FQQ23-N4YCY-73HQ3-FM9WC-76HF4"; "Install" = $false; "Name" = "Project Pro 2024" }
        );
        "LanguageID" = "MatchOS";
        "ExcludeApp" = @(
            @{ "Name" = "Word              "; "Install" = $true;   "ID" = "Word"       };
            @{ "Name" = "Excel             "; "Install" = $true;   "ID" = "Excel"      };
            @{ "Name" = "PowerPoint        "; "Install" = $true;   "ID" = "PowerPoint" };
            @{ "Name" = "Outlook           "; "Install" = $false;  "ID" = "Outlook"    };
            @{ "Name" = "OneNote           "; "Install" = $false;  "ID" = "OneNote"    };
            @{ "Name" = "Access            "; "Install" = $false;  "ID" = "Access"     };
            @{ "Name" = "Skype for Business"; "Install" = $false;  "ID" = "Lync"       };
            @{ "Name" = "OneDrive Desktop  "; "Install" = $false;  "ID" = "OneDrive"   };
            @{ "Name" = "Publisher         "; "Install" = $false;  "ID" = "Publisher"  }
        )
    };
    @{
        "Name" = "Office LTSC Professional Plus 2021";
        "Channel" = "PerpetualVL2021";
        "Product" = @(
            @{"ID" = "ProPlus2021Volume"; "PIDKEY" = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH";    "Install" = $true };
            @{"ID" = "VisioPro2021Volume"; "PIDKEY" = "KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4";   "Install" = $false; "Name" = "Visio Pro 2021  " };
            @{"ID" = "ProjectPro2021Volume"; "PIDKEY" = "FTNWT-C6WBT-8HMGF-K9PRX-QV9H8"; "Install" = $false; "Name" = "Project Pro 2021" }
        );
        "LanguageID" = "MatchOS";
        "ExcludeApp" = @(
            @{ "Name" = "Word              "; "Install" = $true;   "ID" = "Word"       };
            @{ "Name" = "Excel             "; "Install" = $true;   "ID" = "Excel"      };
            @{ "Name" = "PowerPoint        "; "Install" = $true;   "ID" = "PowerPoint" };
            @{ "Name" = "Outlook           "; "Install" = $false;  "ID" = "Outlook"    };
            @{ "Name" = "OneNote           "; "Install" = $false;  "ID" = "OneNote"    };
            @{ "Name" = "Access            "; "Install" = $false;  "ID" = "Access"     };
            @{ "Name" = "Skype for Business"; "Install" = $false;  "ID" = "Lync"       };
            @{ "Name" = "OneDrive Desktop  "; "Install" = $false;  "ID" = "OneDrive"   };
            @{ "Name" = "Publisher         "; "Install" = $false;  "ID" = "Publisher"  }
        )
    };
    @{
        "Name" = "Office Professional Plus 2019";
        "Channel" = "PerpetualVL2019";
        "Product" = @(
            @{"ID" = "ProPlus2019Volume"; "PIDKEY" = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP";    "Install" = $true };
            @{"ID" = "VisioPro2019Volume"; "PIDKEY" = "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB";   "Install" = $false; "Name" = "Visio Pro 2019  " };
            @{"ID" = "ProjectPro2019Volume"; "PIDKEY" = "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B"; "Install" = $false; "Name" = "Project Pro 2019" }
        );
        "LanguageID" = "MatchOS";
        "ExcludeApp" = @(
            @{ "Name" = "Word              "; "Install" = $true;   "ID" = "Word"       };
            @{ "Name" = "Excel             "; "Install" = $true;   "ID" = "Excel"      };
            @{ "Name" = "PowerPoint        "; "Install" = $true;   "ID" = "PowerPoint" };
            @{ "Name" = "Outlook           "; "Install" = $false;  "ID" = "Outlook"    };
            @{ "Name" = "OneNote           "; "Install" = $false;  "ID" = "OneNote"    };
            @{ "Name" = "Access            "; "Install" = $false;  "ID" = "Access"     };
            @{ "Name" = "OneDrive (Groove) "; "Install" = $false;  "ID" = "Groove"     };
            @{ "Name" = "Skype for Business"; "Install" = $false;  "ID" = "Lync"       };
            @{ "Name" = "OneDrive Desktop  "; "Install" = $false;  "ID" = "OneDrive"   };
            @{ "Name" = "Publisher         "; "Install" = $false;  "ID" = "Publisher"  }
        )
    }
)

function SwitchOfficeVersion {
    [uint32]$in = 65536
    do {
        Write-Host $(GetLanguageString "officeVersionIndexList")
        for ($i = 0; $i -lt $configMap.Count; $i++) {
            if ($i -eq 0) {
                Write-Host "  $i        $($configMap[$i]["Name"]) $(GetLanguageString "defaultStr")"
            } else {
                Write-Host "  $i        $($configMap[$i]["Name"])"
            }
        }
        Write-Host ""
        $answer = Read-Host $(GetLanguageString "inputOfficeVersion")
        Clear-Host
        $in = 65536
        try {
            $in = [uint32]$answer
            if ($in -ge $configMap.Count) {
                throw [System.IndexOutOfRangeException]
            }
        }
        catch {
            Write-Host "$(GetLanguageString "invalidInput") : $answer"
            $in = 65536
        }
    } while ($in -ge $configMap.Count)
    return $in
}

function GetUserChoice($officeVersionIndex) {
    $config = $configMap[$officeVersionIndex]
    $apps = $config["ExcludeApp"]
    Write-Host $(GetLanguageString "appInstallListHeader")
    $index = 1
    foreach ($app in $apps) {
        $s = " " + $index.ToString() + "`t`t" + $app["Name"] + "`t"
        if ($app["Install"]) {
            $s = $s + $(GetLanguageString "Yes")
        } else {
            $s = $s + $(GetLanguageString "No")
        }
        Write-Host $s
        $index++
    }
    for ($i = 1; $i -lt $config["Product"].Count; $i++) {
        $s = " " + $index.ToString() + "`t`t" + $config["Product"][$i]["Name"] + "`t"
        if ($app["Install"]) {
            $s = $s + $(GetLanguageString "Yes")
        } else {
            $s = $s + $(GetLanguageString "No")
        }
        Write-Host $s
        $index++
    }
    Write-Host `n
    $installString = GetLanguageString "StartInstall"
    $s = " " + $index.ToString() + "`t`t" + $installString + "`t"
    Write-Host $s
    $answer = Read-Host $(GetLanguageString "InputFunctionIndex")
    Clear-Host
    $in = -1
    try {
        $in = [int]$answer - 1
        if ($in -gt $index -or $in -lt 0) {
            throw [System.IndexOutOfRangeException]
        }
    }
    catch {
        Write-Host "$(GetLanguageString "invalidInput")$answer"
        $in = -1
    }
    return $in
}

function SwitchApp($officeVersionIndex) {
    $installIndex = $configMap[$officeVersionIndex]["ExcludeApp"].Count + $configMap[$officeVersionIndex]["Product"].Count - 1
    do {
        $in = GetUserChoice $officeVersionIndex
        if ($in -lt 0) {
            continue
        } elseif ($in -lt $configMap[$officeVersionIndex]["ExcludeApp"].Count) {
            $configMap[$officeVersionIndex]["ExcludeApp"][$in]["Install"] = -not $configMap[$officeVersionIndex]["ExcludeApp"][$in]["Install"]
        } elseif ($in -lt ($configMap[$officeVersionIndex]["Product"].Count - 1 + $configMap[$officeVersionIndex]["ExcludeApp"].Count)) {
            $idx = $in - $configMap[$officeVersionIndex]["ExcludeApp"].Count
            $configMap[$officeVersionIndex]["Product"][$idx]["Install"] = -not $configMap[$officeVersionIndex]["Product"][$idx]["Install"]
        } elseif ($in -eq $installIndex) {
            break;
        }
    } while ($true)
}

function GenterateConfig($officeVersionIndex, $targetPath) {
    $config = $configMap[$officeVersionIndex]
    $xml = New-Object -TypeName xml
    $xml.LoadXml($configCommon)
    $addElement = $xml.CreateElement("Add")
    $addElement.SetAttribute("OfficeClientEdition", "64")
    $addElement.SetAttribute("Channel", $config["Channel"])

    foreach($product in $config["Product"]) {
        if ($product["Install"]) {
            $productElement = $xml.CreateElement("Product");
            $productElement.SetAttribute("ID", $product["ID"])
            $productElement.SetAttribute("PIDKEY", $product["PIDKEY"])
        
            $langElement = $xml.CreateElement("Language");
            $langElement.SetAttribute("ID", $config["LanguageID"])
            $productElement.AppendChild($langElement)
        
            foreach($app in $config["ExcludeApp"]) {
                if (-not $app["Install"]) {
                    $excludeAppElement = $xml.CreateElement("ExcludeApp")
                    $excludeAppElement.SetAttribute("ID", $app["ID"])
                    $productElement.AppendChild($excludeAppElement)
                }
            }
            $addElement.AppendChild($productElement)
        }
    }

    $configurationNode = $xml.SelectSingleNode("Configuration")
    $first = $configurationNode.FirstChild
    $configurationNode.InsertBefore($addElement, $first)
    $xml.Save($targetPath)
}

function DownloadOffice($exeFolder, $configPath) {
    Write-Host $(GetLanguageString "downloadingOffice")
    Set-Location $exeFolder
    Start-Process -FilePath "$exeFolder\setup.exe" -ArgumentList @("/download", "$configPath") -Wait
    Set-Location $ScriptPath
}

function InstallOffice($exeFolder, $configPath) {
    Write-Host $(GetLanguageString "installingOffice")
    Set-Location $exeFolder
    Start-Process -FilePath "$exeFolder\setup.exe" -ArgumentList @("/configure", "$configPath") -Wait
    Set-Location $ScriptPath
}

function Main {
    ReleaseODT "$ScriptPath\OfficeSetup\"
    $officeVersion = SwitchOfficeVersion
    SwitchApp $officeVersion
    GenterateConfig $officeVersion "$ScriptPath\OfficeSetup\configuration.xml"
    DownloadOffice "$ScriptPath\OfficeSetup" "$ScriptPath\OfficeSetup\configuration.xml"
    InstallOffice  "$ScriptPath\OfficeSetup" "$ScriptPath\OfficeSetup\configuration.xml"
}

Main
pause
