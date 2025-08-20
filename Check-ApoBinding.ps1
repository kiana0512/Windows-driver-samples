[CmdletBinding()]
param(
  # === Parameters (edit as needed) ===
  [string]$HwPrefix = 'USB\VID_0A67&PID_30A2&MI_00',
  [string]$ApoDll   = 'MyCompanyEfxApo.dll',
  [string]$ApoClsid = '{8E3E0B71-5B8A-45C9-9B3D-3A2E5B418A10}',
  [string]$ExtId    = '{E6F0C0C8-2A0D-4B5D-9B6E-6B3D7B2C9D11}',
  [switch]$CycleDevice,
  [switch]$RestartAudio,
  [switch]$PlayTest,
  [switch]$ShowEvents
)

$ErrorActionPreference = 'Stop'

function W([string]$s,[string]$color='Gray'){ Write-Host $s -ForegroundColor $color }
function Section($t){ W "`n== $t ==" 'Cyan' }
function SubSection($t){ W "`n== $t ==" 'DarkCyan' }

# Property keys
# Fix: correct PKEY (uses ...-9F5B-...)
$PKEY_Instance = '{78C34FC8-104A-4D11-9F5B-700F2848BCA5},256'
$PKEY_FX_EndpointEffectClsid = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},7'
$PKEY_FX_Extra                = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},17'
$PKEY_Disable_SysFx           = '{1DA5D803-D492-4EDD-8C23-E0C0FFEE7F0E},5'

function Require-Admin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = [Security.Principal.WindowsPrincipal]$id
  if(-not $p.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)){
    W "Note: run PowerShell as Administrator for full access to device and audio registries." 'Yellow'
  }
}

# Build Enum registry root for a given device instance
function Get-DevRegPath([string]$iid){
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
  # Play built-in WAVs to trigger audiodg/EFX load
  $candidates = @(
    "$env:WINDIR\Media\Windows Background.wav",
    "$env:WINDIR\Media\Windows Notify Calendar.wav",
    "$env:WINDIR\Media\Windows Notify System Generic.wav"
  ) | Where-Object { Test-Path $_ }

  if(-not $candidates){
    W "No system WAV files found. Skipping playback." 'Yellow'
    return
  }

  Add-Type -AssemblyName System.Media | Out-Null
  foreach($wav in $candidates){
    W "Playing: $wav" 'Gray'
    $sp = New-Object System.Media.SoundPlayer($wav)
    $sp.PlaySync()
    Start-Sleep -Milliseconds 200
  }
}

function Show-AudioEvents {
  W "Recent events from Audio-Effects-Manager:" 'Gray'
  try{
    Get-WinEvent -LogName "Microsoft-Windows-Audio-Effects-Manager/Operational" -MaxEvents 50 |
      Select TimeCreated, Id, LevelDisplayName, Message | Format-List
  }catch{
    W "Cannot read Microsoft-Windows-Audio-Effects-Manager/Operational." 'Yellow'
  }

  W "`nRecent events from Microsoft-Windows-Audio:" 'Gray'
  try{
    Get-WinEvent -LogName "Microsoft-Windows-Audio/Operational" -MaxEvents 50 |
      Select TimeCreated, Id, LevelDisplayName, Message | Format-List
  }catch{
    W "Cannot read Microsoft-Windows-Audio/Operational." 'Yellow'
  }
}

Require-Admin

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

Section '3) Enum\...\Device Parameters\FX\0 check'
$fxFound = $false
$devPaths = @()
foreach($d in $devs){
  $root = Get-DevRegPath $d.InstanceId
  if($null -eq $root){ continue }
  $fx0  = Join-Path $root 'Device Parameters\FX\0'
  $devPaths += [pscustomobject]@{ Iid=$d.InstanceId; DevRoot=$root; Fx0=$fx0 }
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
  W "Hint: If not set, the Extension INF may not have applied to this instance or was overridden by a higher DriverVer." 'Yellow'
}

Section '4) MMDevices mapping and FxProperties'
$renderRoot = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render'
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

# Optional: quick COM activation probe (only if your APO exposes a COM-visible class)
try {
  $t = [type]::GetTypeFromCLSID($ApoClsid)
  if($t){
    $obj = [Activator]::CreateInstance($t)
    "COM CoCreateInstance() OK: $($obj.GetType().FullName)"
  }
}catch{
  W ("COM activation failed: 0x{0:X8}  {1}" -f $_.Exception.HResult,$_.Exception.Message) 'Yellow'
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

W "`n== END ==" 'Cyan'
