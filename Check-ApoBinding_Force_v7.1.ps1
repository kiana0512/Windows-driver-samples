# Check-ApoBinding_Force_v7.1.ps1 — FORCE the chain: Endpoint -> audiodg.exe -> Your EFX APO DLL
# Fixes from v7: robust PE parsing (no null deref), non-fatal on PE issues, clearer logs.

[CmdletBinding()]
param(
  [string]$HwPrefix = 'USB\VID_0A67&PID_30A2&MI_00',
  [string]$ApoClsid = '{8E3E0B71-5B8A-45C9-9B3D-3A2E5B418A10}',
  [string]$ApoDll   = 'MyCompanyEfxApo.dll',
  [int]$EventCount  = 120,
  [switch]$VerboseLogs,
  [switch]$DryRun,
  [switch]$SkipPE
)

# -------------------------- Auto elevate --------------------------
$wi = [Security.Principal.WindowsIdentity]::GetCurrent()
$wp = New-Object Security.Principal.WindowsPrincipal($wi)
if(-not $wp.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)){
  $argsLine = @()
  foreach($kv in $PSBoundParameters.GetEnumerator()){
    $name=$kv.Key;$val=$kv.Value
    if($val -is [switch]){ if($val){ $argsLine += "-$name" } }
    else { $argsLine += "-$name `"$val`"" }
  }
  $arg = "-ExecutionPolicy Bypass -File `"$PSCommandPath`" " + ($argsLine -join ' ')
  Start-Process powershell.exe -Verb RunAs -ArgumentList $arg | Out-Null
  return
}
$ErrorActionPreference = 'Stop'

# -------------------------- Utilities ----------------------------
function WriteStep([string]$s){ Write-Host ("== " + $s + " ==") }
function Info([string]$s){ Write-Host $s }
function Warn([string]$s){ Write-Host ("[!] " + $s) -ForegroundColor Yellow }
function Err ([string]$s){ Write-Host ("[X] " + $s) -ForegroundColor Red }
function Vrb([string]$s){ if($VerboseLogs){ Write-Host ("[v] " + $s) -ForegroundColor DarkGray } }

$MMBase = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio'
$RenderKey = Join-Path $MMBase 'Render'

function Get-EndpointObjects {
  $list=@(); if(-not (Test-Path $RenderKey)){ return $list }
  foreach($guidKey in Get-ChildItem $RenderKey -EA SilentlyContinue){
    $propsKey = Join-Path $guidKey.PSPath 'Properties'
    $fxKey    = Join-Path $guidKey.PSPath 'FxProperties'
    $name='(unknown)';$instId='(unknown)';$disable='(unset)';$has=$false
    $p=$null; if(Test-Path $propsKey){ $p = Get-ItemProperty -Path $propsKey -EA SilentlyContinue }
    if($p){
      $n1=$p.'{a45c254e-df1c-4efd-8020-67d146a850e0},2'
      $n2=$p.'{b3f8fa53-0004-438e-9003-51a46e139bfc},14'
      if([string]::IsNullOrWhiteSpace($n1) -and $n2){ $n1=$n2 }
      if($n1){ $name=$n1.Trim() }
      $iid=$p.'{78C34FC8-104A-4D11-9F5B-700F2848BCA5},256'; if($iid){ $instId=$iid.Trim() }
      if($p.PSObject.Properties.Match('Disable_SysFx')){ $disable=[string]$p.Disable_SysFx }
    }
    $f=$null; if(Test-Path $fxKey){ $f=Get-ItemProperty -Path $fxKey -EA SilentlyContinue }
    if($f){
      if($f.PSObject.Properties.Match('Disable_SysFx')){ $disable=[string]$f.Disable_SysFx }
      foreach($np in ($f.PSObject.Properties | Where-Object {$_.MemberType -eq 'NoteProperty'})){
        try{ $v=$f.$($np.Name)
          if(($v -is [string] -and $v.Trim().ToLower() -eq $ApoClsid.ToLower()) -or ($v -is [string[]] -and ($v -contains $ApoClsid))){ $has=$true; break }
        }catch{}
      }
    }
    $list += [pscustomobject]@{ Kind='Render'; EndpointGuid=$guidKey.PSChildName; Name=$name; InstanceId=$instId; HasClsid=$has; Disable_SysFx=$disable; FxKey=$fxKey; PropsKey=$propsKey }
  }
  return $list
}

function Get-DeviceFx0ForInstance([string]$InstanceId){
  $path = "HKLM:\SYSTEM\CurrentControlSet\Enum\$InstanceId\Device Parameters\FX\0"
  if(Test-Path $path){ return Get-Item $path } else { return $null }
}

# ------------ PolicyConfig COM (guarded Add-Type to avoid duplicates) ------------
if (-not ("PC.IPolicyConfigVista" -as [type])) {
Add-Type -Language CSharp -TypeDefinition @"
using System; using System.Runtime.InteropServices;
namespace PC{
  [ComImport, Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")] public class _CPolicyConfigVistaClient{}
  [ComImport, Guid("F8679F50-850A-41CF-9C72-430F290290C8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IPolicyConfigVista{ int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string id, ERole role); }
  [ComImport, Guid("294935CE-F637-4E7C-A41B-AB255460B862"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IPolicyConfig10{
    int Unused1();int Unused2();int Unused3();int Unused4();int Unused5();int Unused6();int Unused7();
    int GetDeviceFormat(string id,int role,IntPtr ppwfx); int SetDeviceFormat(string id,IntPtr pwfxIn,IntPtr pwfxOut);
    int GetProcessingPeriod(string id,int role,IntPtr pDefault,IntPtr pMinimum); int SetProcessingPeriod(string id,IntPtr pPeriod);
    int GetShareMode(string id,IntPtr pMode); int SetShareMode(string id,IntPtr pMode);
    int GetPropertyValue(string id, ref PropertyKey key, IntPtr pv); int SetPropertyValue(string id, ref PropertyKey key, IntPtr pv);
    int SetDefaultEndpoint(string id, ERole role); int Unused8(); int Unused9(); }
  public enum ERole { eConsole=0, eMultimedia=1, eCommunications=2 }
  [StructLayout(LayoutKind.Sequential)] public struct PropertyKey { public Guid fmtid; public int pid; }
}
"@
} else { Vrb 'PolicyConfig types already loaded.' }

function Get-EndpointIdString([string]$Kind,[string]$EndpointGuid){ if($Kind -eq 'Capture'){ return "{0.0.1.00000000}.{$EndpointGuid}" } else { return "{0.0.0.00000000}.{$EndpointGuid}" } }

function Set-DefaultEndpoint([object]$Ep){
  $id = Get-EndpointIdString -Kind $Ep.Kind -EndpointGuid $Ep.EndpointGuid
  try{
    $vista = [PC.IPolicyConfigVista][Activator]::CreateInstance([type]::GetTypeFromCLSID([Guid]'870af99c-171d-4f9e-af0d-e63df40c2bc9'))
    foreach($r in 0,1,2){ $hr=$vista.SetDefaultEndpoint($id,[PC.ERole]$r); Info ("  Vista SetDefault(role=$r) hr=0x{0:X8}" -f ($hr -band 0xffffffff)) }
    Info "  Default set => $id"
  }catch{ Warn ("  Vista SetDefault failed: " + $_.Exception.Message) }
  try{
    $pc10 = [PC.IPolicyConfig10][Activator]::CreateInstance([type]::GetTypeFromCLSID([Guid]'870af99c-171d-4f9e-af0d-e63df40c2bc9'))
    foreach($r in 0,1,2){ $hr=$pc10.SetDefaultEndpoint($id,[PC.ERole]$r); Info ("  Win10 SetDefault(role=$r) hr=0x{0:X8}" -f ($hr -band 0xffffffff)) }
  }catch{ Vrb ("  Win10 SetDefault path unavailable: " + $_.Exception.Message) }
}

function Mirror-DeviceFx0-To-EndpointFx([string]$DeviceInstanceId,[object]$Endpoint){
  $fx0 = Get-DeviceFx0ForInstance $DeviceInstanceId
  if(-not $fx0){ Warn "  No device FX\\0 for $DeviceInstanceId"; return }
  $fx0Obj = Get-ItemProperty $fx0.PSPath -EA SilentlyContinue
  if(-not $fx0Obj){ Warn "  FX\\0 not readable."; return }
  $dst = $Endpoint.FxKey
  if(-not (Test-Path $dst) -and -not $DryRun){ New-Item -Path $dst -Force | Out-Null }
  foreach($p in $fx0Obj.PSObject.Properties | Where-Object {$_.MemberType -eq 'NoteProperty'}){
    $name=$p.Name; $val=$p.Value
    try{
      if(-not $DryRun){
        if($val -is [array]){ New-ItemProperty -Path $dst -Name $name -Value ([string[]]$val) -PropertyType MultiString -Force | Out-Null }
        else { New-ItemProperty -Path $dst -Name $name -Value ([string]$val) -PropertyType String -Force | Out-Null }
      }
      Info ("  Mirrored FxProperties.$name = " + (@($val) -join ', '))
    }catch{ Warn ("  Write $name failed: " + $_.Exception.Message) }
  }
}

function Force-Enable-Enhancements([object]$Endpoint){
  foreach($path in @($Endpoint.PropsKey,$Endpoint.FxKey)){
    if(-not $path){ continue }
    try{
      if(-not (Test-Path $path) -and -not $DryRun){ New-Item -Path $path -Force | Out-Null }
      if(-not $DryRun){ New-ItemProperty -Path $path -Name 'Disable_SysFx' -Value 0 -PropertyType DWord -Force | Out-Null }
      Info ("  $path : Disable_SysFx=0")
    }catch{ Warn ("  Disable_SysFx write failed at $path : " + $_.Exception.Message) }
  }
}

function Parse-PEHeader([string]$path){
  try{
    if([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path $path)){ return $null }
    $fs=[System.IO.File]::Open($path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::Read)
    try{
      $br=New-Object System.IO.BinaryReader($fs)
      $mz=[System.Text.Encoding]::ASCII.GetString($br.ReadBytes(2))
      if($mz -ne 'MZ'){ return $null }
      $fs.Seek(0x3C,0)|Out-Null; $e_lfanew=$br.ReadInt32(); if($e_lfanew -le 0){ return $null }
      $fs.Seek($e_lfanew,0)|Out-Null
      $sig=$br.ReadBytes(4); if($sig.Length -lt 4 -or $sig[0] -ne 0x50 -or $sig[1] -ne 0x45){ return $null }
      $machine=$br.ReadUInt16(); $br.ReadUInt16()|Out-Null; $time=$br.ReadUInt32(); $fs.Seek(0x10,1)|Out-Null
      $magic=$br.ReadUInt16(); $isPE32Plus=($magic -eq 0x20B); $subsysOff=if($isPE32Plus){0x5C}else{0x44}; $fs.Seek($subsysOff-2,1)|Out-Null
      $sub=$br.ReadUInt16();
      return [pscustomobject]@{ Machine=('0x{0:X4}' -f $machine); PE32Plus=$isPE32Plus; Subsystem=('0x{0:X4}' -f $sub); Time=([DateTime]::UnixEpoch).AddSeconds([double]$time).ToString('yyyy-MM-dd HH:mm:ss') }
    } finally { if($fs){ $fs.Dispose() } }
  } catch { return $null }
}

function Check-Audiodg-HasDll([string]$dll){
  $ok=$false
  try{
    $adg = Get-Process audiodg -EA SilentlyContinue
    if($adg){ Info ('audiodg.exe PID=' + $adg.Id) } else { Warn 'audiodg.exe not running (no active session).' }
    $tl = & tasklist /fi "imagename eq audiodg.exe" /m $dll 2>$null
    if($tl -and ($tl -match '^audiodg\.exe')){ $tl | ForEach-Object { Info $_ }; $ok=$true }
  }catch{ Warn 'tasklist failed.' }
  return $ok
}

# -------------------------- Force Flow ----------------------------
WriteStep '1) Verify COM/PE correctness'
# COM path
$clsKey = "Registry::HKEY_CLASSES_ROOT\\CLSID\\$ApoClsid\\InprocServer32"
$dll=$null
if(Test-Path $clsKey){
  try{ $dll = (Get-ItemProperty $clsKey -Name '(default)').'(default)'; Info "CLSID InprocServer32 = $dll" }catch{ Warn ('Cannot read InprocServer32: ' + $_.Exception.Message) }
}else{ Warn 'CLSID not found (HKCR\\CLSID).' }

if(-not $SkipPE -and $dll){
  $pe = Parse-PEHeader -path $dll
  if($pe){ Info ("PE: Machine="+$pe.Machine+", PE32Plus="+$pe.PE32Plus+", Subsystem="+$pe.Subsystem+", Time="+$pe.Time) }
  else   { Warn 'PE parse skipped/failed (non-fatal). Continue.' }
}

WriteStep '2) Discover render endpoints'
$eps = Get-EndpointObjects
$eps | Select Kind,EndpointGuid,Name,Disable_SysFx,HasClsid,InstanceId | Format-Table -AutoSize

# Choose target endpoint: prefer (HasClsid && InstanceId matches HwPrefix); else HasClsid; else HwPrefix
$target = $eps | Where-Object { $_.HasClsid -and $_.InstanceId -like "$HwPrefix*" } | Select-Object -First 1
if(-not $target){ $target = $eps | Where-Object { $_.HasClsid } | Select-Object -First 1 }
if(-not $target){ $target = $eps | Where-Object { $_.InstanceId -like "$HwPrefix*" } | Select-Object -First 1 }
if($null -eq $target){ Err 'No suitable render endpoint found. Make sure device is present and INF wrote FX\0.'; return }
Info ("Target => " + $target.Name + " [" + $target.EndpointGuid + "]")

WriteStep '3) Mirror device FX\0 -> endpoint FxProperties (to guarantee CLSID on endpoint)'
if($target.InstanceId -and $target.InstanceId -ne '(unknown)'){
  Mirror-DeviceFx0-To-EndpointFx -DeviceInstanceId $target.InstanceId -Endpoint $target
} else { Warn 'Target has no InstanceId; skip mirror.' }

WriteStep '4) Force Enable Enhancements (Disable_SysFx=0)'
Force-Enable-Enhancements -Endpoint $target

WriteStep '5) Set this endpoint as default for ALL roles (Console/Multimedia/Communications)'
if(-not $DryRun){ Set-DefaultEndpoint -Ep $target } else { Warn 'DryRun: skipped SetDefault.' }

WriteStep '6) Rebuild endpoint cache + restart audio services (optional but recommended)'
if(-not $DryRun){
  if($target.InstanceId -and $target.InstanceId -ne '(unknown)'){
    try{ Disable-PnpDevice -InstanceId $target.InstanceId -Confirm:$false -EA SilentlyContinue; Start-Sleep 1; Enable-PnpDevice -InstanceId $target.InstanceId -Confirm:$false -EA SilentlyContinue; Info 'Cycled device.' }catch{ Warn ('Cycle device failed: ' + $_.Exception.Message) }
  }
  foreach($svc in 'Audiosrv','AudioEndpointBuilder'){
    try{ Restart-Service -Name $svc -Force -EA SilentlyContinue; Info ("Restarted service: " + $svc) }catch{ Warn ("Restart service failed: " + $_.Exception.Message) }
  }
} else { Warn 'DryRun: skipped cycle/restart.' }

WriteStep '7) Play a system WAV to force audiodg load'
$wav = Join-Path $env:WINDIR 'Media\Windows Background.wav'; if(-not (Test-Path $wav)){ $wav = Join-Path $env:WINDIR 'Media\Windows Ding.wav' }
try{ $player = New-Object System.Media.SoundPlayer($wav); $player.Play(); Info ("Playing: " + (Split-Path $wav -Leaf) + " ..."); Start-Sleep 2 }catch{ Warn ("SoundPlayer failed: " + $_.Exception.Message) }

WriteStep '8) Verify audiodg has loaded your DLL (poll up to 12s)'
$ok=$false; for($i=0;$i -lt 12;$i++){ if(Check-Audiodg-HasDll -dll $ApoDll){ $ok=$true; break } Start-Sleep 1 }
if(-not $ok){ Warn "audiodg has NOT loaded $ApoDll. We'll fetch logs for root-causing." }

WriteStep '9) Recent logs (Audio Operational)'
try{ Get-WinEvent -LogName 'Microsoft-Windows-Audio/Operational' -MaxEvents $EventCount | Select-Object TimeCreated,Id,LevelDisplayName,Message | Format-List }catch{ Warn ('Cannot read Audio/Operational: ' + $_.Exception.Message) }

WriteStep '10) Notes & next steps'
Info 'If logs show RAW/exclusive path is used, EFX is bypassed. Test with shared mode first.'
Info 'If you see Initialize failed hr=..., fix APO Initialize()/format/mode and re-run.'
Info 'If DLL still not loading but endpoint has CLSID and enhancements enabled, paste steps 2/8/9 outputs for diagnosis.'

WriteStep 'Done.'
