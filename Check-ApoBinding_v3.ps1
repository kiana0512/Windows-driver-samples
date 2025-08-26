<#
================================================================================
 Check-ApoBinding_v2.ps1  —  USB Audio EFX APO end-to-end diagnostics
 Author: you
 Notes:
  - Pure ASCII output. Rich comments in Chinese (ASCII punctuation).
  - Requires Administrator for writing MMDevices FxProperties and device toggling.
  - Supports:
      1) Installed Extension packages check
      2) Device instance present and FX\0 binding
      3) MMDevices mapping (strict by InstanceId)
      4) Endpoint-wide FxProperties search (Render/Capture)
      5) COM/CLSID and DLL presence
      6) Set default render endpoint to your device (IPolicyConfig)
      7) Clear Disable_SysFx = 0 (enable enhancements)
      8) Play system WAVs to trigger audiodg load
      9) Show Audio Effects related Windows events
     10) Parse APO DLL PE header (arch/subsystem/timestamp)
     11) Scan setupapi.dev.log for install/AddReg clues

 Usage examples:
   powershell.exe -ExecutionPolicy Bypass -File .\Check-ApoBinding_v2.ps1
   powershell.exe -ExecutionPolicy Bypass -File .\Check-ApoBinding_v2.ps1 `
     -SearchFx -SetDefaultToTarget -ClearDisableSysFx -PlayTest -ShowEvents -ScanSetupLog
================================================================================
#>

[CmdletBinding()]
param(
  # === Parameters (edit as needed) ===
  [string]$HwPrefix = 'USB\VID_0A67&PID_30A2&MI_00',
  [string]$ApoDll   = 'MyCompanyEfxApo.dll',
  [string]$ApoClsid = '{8E3E0B71-5B8A-45C9-9B3D-3A2E5B418A10}',
  [string]$ExtId    = '{E6F0C0C8-2A0D-4B5D-9B6E-6B3D7B2C9D11}',
  [switch]$CycleDevice,         # Disable/Enable matching device instances
  [switch]$RestartAudio,        # Restart Windows Audio service (audiosrv)
  [switch]$PlayTest,            # Play system WAVs to trigger audiodg load
  [switch]$ShowEvents,          # Show recent Audio/Effects events
  # New capabilities:
  [switch]$SearchFx,            # Search all endpoints (Render/Capture) for our CLSID
  [switch]$ClearDisableSysFx,   # Force Disable_SysFx = 0 on endpoints with our CLSID
  [switch]$SetDefaultToTarget,  # Set default render endpoint to our device
  [switch]$ParsePE,             # Parse APO DLL PE header
  [switch]$ScanSetupLog         # Scan setupapi.dev.log for HwPrefix/ExtId
)

# ------------------------------------------------------------------------------
# Auto-elevate to Administrator if needed (so that registry writes succeed)
# ------------------------------------------------------------------------------
$wi = [Security.Principal.WindowsIdentity]::GetCurrent()
$wp = New-Object Security.Principal.WindowsPrincipal($wi)
if(-not $wp.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)){
  $argsLine = @()
  if($PSBoundParameters.ContainsKey('HwPrefix')){$argsLine += "-HwPrefix `"$HwPrefix`""}
  if($PSBoundParameters.ContainsKey('ApoDll')){$argsLine += "-ApoDll `"$ApoDll`""}
  if($PSBoundParameters.ContainsKey('ApoClsid')){$argsLine += "-ApoClsid `"$ApoClsid`""}
  if($PSBoundParameters.ContainsKey('ExtId')){$argsLine += "-ExtId `"$ExtId`""}
  foreach($s in 'CycleDevice','RestartAudio','PlayTest','ShowEvents','SearchFx','ClearDisableSysFx','SetDefaultToTarget','ParsePE','ScanSetupLog'){
    if($PSBoundParameters[$s]){ $argsLine += "-$s" }
  }
  $arg = "-ExecutionPolicy Bypass -File `"$PSCommandPath`" " + ($argsLine -join ' ')
  Start-Process -FilePath "powershell.exe" -ArgumentList $arg -Verb RunAs
  return
}

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------------------
# Utilities
# ------------------------------------------------------------------------------
function W([string]$s,[string]$color='Gray'){ Write-Host $s -ForegroundColor $color }
function Section($t){ W "`n== $t ==" 'Cyan' }
function SubSection($t){ W "`n== $t ==" 'DarkCyan' }

# Property keys
# PKEY_Device_InstanceId: note the correct GUID tail is 9F5B
$PKEY_Instance = '{78C34FC8-104A-4D11-9F5B-700F2848BCA5},256'
# EndpointEffect CLSID
$PKEY_FX_EndpointEffectClsid = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},7'
# Optional vendor/custom list
$PKEY_FX_Extra                = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},17'
# Disable system effects (1 disables enhancements)
$PKEY_Disable_SysFx           = '{1DA5D803-D492-4EDD-8C23-E0C0FFEE7F0E},5'

function Get-DevRegPath([string]$iid){
  # Build Enum root: HKLM\SYSTEM\CCS\Enum\<parent>\<leaf>
  $m = [regex]::Match($iid,'^(.*)\\([^\\]+)$')
  if($m.Success){
    return "HKLM:\SYSTEM\CurrentControlSet\Enum\$($m.Groups[1].Value)\$($m.Groups[2].Value)"
  }
  return $null
}

function Restart-AudioEngine {
  W "Restarting Windows Audio service..." 'DarkYellow'
  try { Stop-Process -Name audiodg -Force -ErrorAction SilentlyContinue } catch {}
  & net stop audiosrv  | Out-Null
  Start-Sleep -Seconds 1
  & net start audiosrv | Out-Null
}

function Cycle-MatchingDevices([string]$prefix){
  $targets = Get-PnpDevice -PresentOnly | Where-Object InstanceId -like "$prefix*"
  if(-not $targets){ W "No present instances found: $prefix*" 'Red'; return }
  $targets | Disable-PnpDevice -Confirm:$false
  Start-Sleep -Seconds 1
  $targets | Enable-PnpDevice -Confirm:$false
}

function Play-TestSound {
  # Play built-in WAVs to trigger audiodg/EFX load; no Add-Type needed.
  $candidates = @(
    "$env:WINDIR\Media\Windows Background.wav",
    "$env:WINDIR\Media\Windows Notify Calendar.wav",
    "$env:WINDIR\Media\Windows Notify System Generic.wav"
  ) | Where-Object { Test-Path $_ }

  if(-not $candidates){
    W "No system WAV files found. Skipping playback." 'Yellow'
    return
  }

  foreach($wav in $candidates){
    try{
      $sp = New-Object System.Media.SoundPlayer($wav)
      $sp.Load()
      W "Playing: $wav" 'Gray'
      $sp.PlaySync()
    }catch{
      W "Sound playback failed: $($_.Exception.Message)" 'Yellow'
    }
  }
}

function Show-AudioEvents {
  # Show recent effect and audio logs; filter to the most relevant keywords.
  W "Recent events from Audio-Effects-Manager:" 'Gray'
  try{
    Get-WinEvent -LogName "Microsoft-Windows-Audio-Effects-Manager/Operational" -MaxEvents 80 |
      Where-Object { $_.Message -match ($ApoClsid -replace '[{}]') -or $_.Message -match 'Effect|APO|EFX|Load|Initialize|Raw' } |
      Select TimeCreated, Id, LevelDisplayName, Message | Format-List
  }catch{
    W "Cannot read Microsoft-Windows-Audio-Effects-Manager/Operational." 'Yellow'
  }

  W "`nRecent events from Microsoft-Windows-Audio:" 'Gray'
  try{
    Get-WinEvent -LogName "Microsoft-Windows-Audio/Operational" -MaxEvents 80 |
      Where-Object { $_.Message -match 'audiodg|stream|endpoint|fx' } |
      Select TimeCreated, Id, LevelDisplayName, Message | Format-List
  }catch{
    W "Cannot read Microsoft-Windows-Audio/Operational." 'Yellow'
  }
}

function Find-EndpointsByFxClsid([string]$clsid){
  # Search all Render and Capture endpoints for our CLSID in FxProperties
  $roots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture'
  )
  $found = @()
  foreach($root in $roots){
    if(-not (Test-Path $root)){ continue }
    Get-ChildItem "$root\*" | ForEach-Object {
      $props = Join-Path $_.PSPath 'Properties'
      $fxp   = Join-Path $_.PSPath 'FxProperties'
      if(Test-Path $fxp){
        $pp = Get-ItemProperty $fxp
        $fxcls = $pp.$PKEY_FX_EndpointEffectClsid
        if($fxcls -and ($fxcls.Trim('{}').ToLower() -eq $clsid.Trim('{}').ToLower())){
          $name = (Get-ItemProperty $props -EA SilentlyContinue).'{a45c254e-df1c-4efd-8020-67d146a850e0},2'
          $iid  = (Get-ItemProperty $props -EA SilentlyContinue).$PKEY_Instance
          $dis  = $pp.$PKEY_Disable_SysFx
          $found += [pscustomobject]@{
            Kind          = (Split-Path $root -Leaf)
            EndpointGUID  = $_.PSChildName
            Name          = $name
            InstanceId    = $iid
            Disable_SysFx = $(if($dis -ne $null){$dis}else{'(unset)'})
            FxRoot        = $fxp
          }
        }
      }
    }
  }
  return $found
}

function Clear-DisableSysFx-OnEndpoints($endpoints){
  # Set Disable_SysFx = 0 on given endpoints (requires Admin)
  foreach($e in $endpoints){
    try{
      Set-ItemProperty -Path $e.FxRoot -Name $PKEY_Disable_SysFx -Type DWord -Value 0 -ErrorAction Stop
      W "Set Disable_SysFx=0 on $($e.EndpointGUID) [$($e.Kind)]" 'Green'
    }catch{
      W "Failed to set Disable_SysFx on $($e.EndpointGUID): $($_.Exception.Message)" 'Yellow'
    }
  }
}

function Set-DefaultRenderToInstancePrefix([string]$prefix){
  # Choose a Render endpoint by InstanceId match; fallback to Fx CLSID match.
  $candidates = @()
  $renderRoot = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render'
  if(Test-Path $renderRoot){
    Get-ChildItem "$renderRoot\*" | ForEach-Object {
      $props = Join-Path $_.PSPath 'Properties'
      $iid   = (Get-ItemProperty $props -EA SilentlyContinue).$PKEY_Instance
      if($iid -and $iid -like "$prefix*"){
        $name = (Get-ItemProperty $props -EA SilentlyContinue).'{a45c254e-df1c-4efd-8020-67d146a850e0},2'
        $candidates += [pscustomobject]@{ EndpointGUID=$_.PSChildName; Name=$name; Match='Instance' }
      }
    }
  }
  if(-not $candidates){
    $eps = Find-EndpointsByFxClsid -clsid $ApoClsid
    $candidates = $eps | Where-Object { $_.Kind -eq 'Render' } |
      Select @{n='EndpointGUID';e={$_.EndpointGUID}}, @{n='Name';e={$_.Name}}, @{n='Match';e={'FxClsid'}}
  }
  if(-not $candidates){ W "No render endpoints matching HwPrefix or CLSID." 'Yellow'; return $null }

  $target = $candidates | Select -First 1
  W ("Setting default render endpoint to {0} ({1})..." -f $target.EndpointGUID,$target.Name) 'DarkYellow'

  # C# interop: support both IPolicyConfig and IPolicyConfigVista, and both CoClasses
  $code = @"
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")]
class CPolicyConfigClient {}

[ComImport, Guid("294935CE-F637-4E7C-A41B-AB255460B862")]
class CPolicyConfigVista {}

[Guid("F8679F50-850A-41CF-9C72-430F290290C8"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IPolicyConfig {
    int Unused1();
    int Unused2();
    int Unused3();
    int Unused4();
    int Unused5();
    int Unused6();
    int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string devID, int role);
}

[Guid("568B9108-44BF-40B4-9006-86AFE5B5A620"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IPolicyConfigVista {
    int Unused1();
    int Unused2();
    int Unused3();
    int Unused4();
    int Unused5();
    int Unused6();
    int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string devID, int role);
}

public static class PolicyConfigHelper {
   public static void SetDefault(string devId){
     object obj = null;
     try { obj = new CPolicyConfigClient(); } catch {}
     if (obj == null) { obj = new CPolicyConfigVista(); }

     var pc  = obj as IPolicyConfig;
     var pcV = obj as IPolicyConfigVista;

     if (pc != null) {
        pc.SetDefaultEndpoint(devId, 0); // Console
        pc.SetDefaultEndpoint(devId, 1); // Multimedia
        pc.SetDefaultEndpoint(devId, 2); // Communications
        Marshal.ReleaseComObject(pc);
     } else if (pcV != null) {
        pcV.SetDefaultEndpoint(devId, 0);
        pcV.SetDefaultEndpoint(devId, 1);
        pcV.SetDefaultEndpoint(devId, 2);
        Marshal.ReleaseComObject(pcV);
     } else {
        Marshal.ReleaseComObject(obj);
        throw new COMException("No supported IPolicyConfig interface (Vista/Win10).", unchecked((int)0x80004002));
     }
   }
}
"@
  try{
    $asm = Add-Type -TypeDefinition $code -Language CSharp -PassThru -ErrorAction Stop
    [PolicyConfigHelper]::SetDefault($target.EndpointGUID)  # Endpoint ID is the GUID key
    W "Default render endpoint set." 'Green'
    return $target.EndpointGUID
  }catch{
    W "Failed to set default endpoint: $($_.Exception.Message)" 'Yellow'
    return $null
  }
}

function Get-PEInfo([string]$path){
  # Parse PE header for arch/subsystem/timestamp; quick mismatch detector.
  if(-not (Test-Path $path)){ throw "File not found: $path" }
  $fs = [System.IO.File]::Open($path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
  try{
    $br = New-Object System.IO.BinaryReader($fs)
    if($br.ReadUInt16() -ne 0x5A4D){ throw "Not an MZ file." }
    $fs.Position = 0x3C
    $peOff = $br.ReadUInt32()
    $fs.Position = $peOff
    if($br.ReadUInt32() -ne 0x00004550){ throw "Invalid PE signature." }
    $machine = $br.ReadUInt16()
    $numSec  = $br.ReadUInt16()
    $time    = [DateTime]::UnixEpoch.AddSeconds($br.ReadUInt32()).ToLocalTime()
    $fs.Position += 12
    $optSize = $br.ReadUInt16()
    $charcs  = $br.ReadUInt16()
    $isPE32Plus = $false
    $subsystem = 0
    if($optSize -gt 0){
      $magic = $br.ReadUInt16()
      $isPE32Plus = ($magic -eq 0x20B)
      if(-not $isPE32Plus){ $fs.Position += 66 } else { $fs.Position += 82 }
      $subsystem = $br.ReadUInt16()
    }
    [pscustomobject]@{
      MachineHex = ('0x{0:X4}' -f $machine)
      Machine    = $(switch($machine){ 0x8664{'AMD64'}; 0x014C{'I386'}; default{"Unknown"} })
      PE32Plus   = $isPE32Plus
      Subsystem  = $subsystem
      Timestamp  = $time
    }
  } finally {
    $fs.Dispose()
  }
}

function Scan-SetupApiLogForClues([string]$prefix,[string]$extId){
  # Scan setupapi.dev.log for device and ExtensionId related records.
  $log = "$env:WINDIR\inf\setupapi.dev.log"
  if(-not (Test-Path $log)){ W "setupapi.dev.log not found." 'Yellow'; return }
  $pat = @(
    [regex]::Escape($prefix),
    [regex]::Escape($extId.Trim('{}')),
    'Class\s*:\s*Extension',
    'Apo|FX|Device Parameters\\FX\\0|AddReg',
    'MyCompanyUsbApoExt|mycompanyusbapoext\.inf|oem\d+\.inf'
  ) -join '|'
  Select-String -Path $log -Pattern $pat -Context 0,4 | ForEach-Object {
    $_.ToString()
  }
}

# ------------------------------------------------------------------------------
# Main
# ------------------------------------------------------------------------------
Section '1) Installed Class=Extension driver packages'
try{
  cmd /c 'dism /online /get-drivers /format:table' |
    Select-String -Pattern 'oem\d+\.inf|Class\s*:\s*Extension|Original.*\.inf|Version\s*:' -Context 0,3
}catch{
  W "Failed to read driver list: $($_.Exception.Message)" 'Yellow'
}

SubSection '1.1) Packages with the same ExtensionId'
Get-ChildItem "$env:WINDIR\INF\oem*.inf" | ForEach-Object {
  $t = Get-Content $_.FullName -Raw
  if($t -match 'Class\s*=\s*Extension' -and $t -match [regex]::Escape($ExtId)){
    $dv = [regex]::Match($t,'DriverVer\s*=\s*([^\r\n]+)').Groups[1].Value.Trim()
    "{0}  DriverVer={1}" -f $_.Name,$dv
  }
}

Section '2) Present instances matching HwPrefix'
$devs = Get-PnpDevice -PresentOnly | Where-Object InstanceId -like "$HwPrefix*"
if(-not $devs){
  W "No present instances found: $HwPrefix*" 'Red'
}else{
  $devs | Select-Object Status, Class, FriendlyName, InstanceId | Format-List
}

if($CycleDevice -and $devs){
  SubSection '2.1) Disable/Enable matching instances to refresh endpoints'
  Cycle-MatchingDevices -prefix $HwPrefix
  $devs = Get-PnpDevice -PresentOnly | Where-Object InstanceId -like "$HwPrefix*"
}

Section '3) Enum...\Device Parameters\FX\0 check'
$fxFound = $false
foreach($d in $devs){
  $root = Get-DevRegPath $d.InstanceId
  if($null -eq $root){ continue }
  $fx0  = Join-Path $root 'Device Parameters\FX\0'
  if(Test-Path $fx0){
    $p = Get-ItemProperty $fx0 -ErrorAction SilentlyContinue
    $clsid = $p.$PKEY_FX_EndpointEffectClsid
    if($clsid){
      W "[OK] $fx0" 'Green'
      "    $PKEY_FX_EndpointEffectClsid = $clsid"
      $fxFound = $true
    } else {
      W "[WARN] $fx0 exists but $PKEY_FX_EndpointEffectClsid not set." 'Yellow'
    }
  } else {
    W "[MISS] $fx0" 'Red'
  }
}
if(-not $fxFound){
  W "Hint: not set implies Extension INF not applied or overridden by a higher DriverVer." 'Yellow'
}

Section '4) MMDevices mapping and FxProperties (strict mapping by InstanceId)'
$renderRoot = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render'
if(Test-Path $renderRoot){
  Get-ChildItem "$renderRoot\*" | ForEach-Object {
    $props = Join-Path $_.PSPath 'Properties'
    $iid   = (Get-ItemProperty $props -EA SilentlyContinue).$PKEY_Instance
    if($iid -and ($devs.InstanceId -contains $iid)){
      $fxp = Join-Path $_.PSPath 'FxProperties'
      W "[Endpoint] $($_.PSChildName) maps to $iid" 'Gray'
      if(Test-Path $fxp){
        $pp = Get-ItemProperty $fxp
        "  FxProperties $PKEY_FX_EndpointEffectClsid = " + $pp.$PKEY_FX_EndpointEffectClsid
        "  FxProperties $PKEY_FX_Extra = " + ($pp.$PKEY_FX_Extra -join ', ')
        $dis = $pp.$PKEY_Disable_SysFx
        "  Disable_SysFx ($PKEY_Disable_SysFx) = " + ($(if($dis -ne $null){$dis}else{'(unset)'}))
      } else {
        W "  [WARN] FxProperties missing (endpoint not rebuilt or INF not effective)." 'Yellow'
      }
    }
  }
}

if($SearchFx){
  SubSection '4b) Endpoints carrying our CLSID (Render & Capture)'
  $eps = Find-EndpointsByFxClsid -clsid $ApoClsid
  if($eps){
    $eps | Format-Table Kind, EndpointGUID, Name, InstanceId, Disable_SysFx -Auto
  }else{
    W "No endpoints contain $ApoClsid in FxProperties." 'Yellow'
  }
}

Section '5) COM/CLSID registration and DLL presence'
$clsPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$ApoClsid\InprocServer32"
if(Test-Path $clsPath){
  $def = (Get-ItemProperty $clsPath).'(default)'
  "CLSID InprocServer32 = $def"
  "ThreadingModel      = " + (Get-ItemProperty $clsPath).ThreadingModel
}else{
  W "Missing: $clsPath" 'Red'
}
$dllPath = Join-Path $env:WINDIR "System32\$ApoDll"
"DLL Exists? " + (Test-Path $dllPath) + "  => $dllPath"

if($ParsePE -and (Test-Path $dllPath)){
  SubSection '5.1) APO DLL PE header'
  try{
    $pe = Get-PEInfo -path $dllPath
    $pe | Format-List
    if(-not $pe.PE32Plus -or $pe.Machine -ne 'AMD64'){
      W "Warning: DLL is not AMD64 PE32+; audiodg.exe is 64-bit." 'Yellow'
    }
  }catch{
    W "PE parse failed: $($_.Exception.Message)" 'Yellow'
  }
}

# Quick COM activation probe (loads into powershell; not audiodg)
try {
  $t = [type]::GetTypeFromCLSID($ApoClsid)
  if($t){
    $obj = [Activator]::CreateInstance($t)
    "COM CoCreateInstance() OK: $($obj.GetType().FullName)"
  }
}catch{
  W ("COM activation failed: 0x{0:X8}  {1}" -f $_.Exception.HResult,$_.Exception.Message) 'Yellow'
}

$defaultEndpointGuid = $null
if($SetDefaultToTarget){
  SubSection '5.2) Set default render endpoint to our target'
  $defaultEndpointGuid = Set-DefaultRenderToInstancePrefix -prefix $HwPrefix
}

if($ClearDisableSysFx -or $SetDefaultToTarget){
  SubSection '5.3) Ensure enhancements (Disable_SysFx=0) for our endpoints'
  $eps2 = Find-EndpointsByFxClsid -clsid $ApoClsid
  if($eps2){
    Clear-DisableSysFx-OnEndpoints -endpoints $eps2
  }else{
    W "No endpoints found with our CLSID to clear Disable_SysFx." 'Yellow'
  }
}

if($RestartAudio){ Restart-AudioEngine }

Section '6) Is audiodg.exe running and is the DLL loaded'
$adgList = (tasklist /fi "imagename eq audiodg.exe") -join "`n"
if($adgList -notmatch 'audiodg.exe'){
  W "audiodg.exe is not running (no active audio session). Start playback or use -PlayTest." 'Yellow'
}else{
  tasklist /m $ApoDll 2>$null
}

if($PlayTest){
  SubSection '6.1) Play system WAVs to trigger load'
  Play-TestSound
  Start-Sleep -Milliseconds 300
  tasklist /fi "imagename eq audiodg.exe"
  tasklist /m $ApoDll 2>$null
}

if($ShowEvents){ Show-AudioEvents }

if($ScanSetupLog){
  Section '7) setupapi.dev.log scan for HwPrefix/ExtId'
  Scan-SetupApiLogForClues -prefix $HwPrefix -extId $ExtId
}

W "`n== END ==" 'Cyan'
