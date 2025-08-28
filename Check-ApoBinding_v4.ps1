# Check-ApoBinding_v6.ps1 — USB Audio EFX APO end-to-end diagnostics
# ASCII-only output. Run as Administrator. Validates/fixes the chain:
# Endpoint -> audiodg.exe -> Your APO DLL
# Safe-by-default: no system-wide policy changes; only targeted, reversible tweaks.

[CmdletBinding()]
param(
    # ---- User knobs ----
    [string]$HwPrefix = 'USB\VID_0A67&PID_30A2&MI_00',
    [string]$ApoDll   = 'MyCompanyEfxApo.dll',
    [string]$ApoClsid = '{8E3E0B71-5B8A-45C9-9B3D-3A2E5B418A10}',
    [string]$ExtId    = '{E6F0C0C8-2A0D-4B5D-9B6E-6B3D7B2C9D11}',

    # Kinds filter for endpoint scan
    [ValidateSet('Render','Capture')]
    [string[]]$Kinds = @('Render','Capture'),

    # Actions
    [switch]$CycleDevice,           # Disable/Enable matched PnP instance(s) to rebuild endpoints
    [switch]$RestartAudio,          # Restart Audiosrv + AudioEndpointBuilder
    [switch]$PlayTest,              # Play a system WAV to force audiodg load path
    [switch]$ShowEvents,            # Dump Audio and Audio-Effects-Manager recent events
    [int]$EventCount = 80,

    # Enhanced capabilities
    [switch]$SearchFx,              # Search ALL endpoints for your CLSID
    [switch]$ClearDisableSysFx,     # Set Disable_SysFx = 0 on matched endpoints
    [switch]$SetDefaultToTarget,    # Switch default render endpoint to your device/CLSID carrier
    [switch]$ParsePE,               # Parse APO DLL PE header (bitness/subsystem/timestamp)
    [switch]$ScanSetupLog,          # Scan setupapi.dev.log for ExtId/HwPrefix lines
    [switch]$TestCOM,               # Only if set: CoCreate your COM CLSID (avoid loading into PowerShell by default)
    [switch]$ForceWriteEndpointFx   # Mirror device FX\0 values into Endpoint FxProperties (when cache not synced)
)

# --------------------- Auto elevate if not Administrator ---------------------
$wi = [Security.Principal.WindowsIdentity]::GetCurrent()
$wp = New-Object Security.Principal.WindowsPrincipal($wi)
if(-not $wp.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)){
    $argsLine = @()
    foreach($kv in $PSBoundParameters.GetEnumerator()){
        $name = $kv.Key; $val = $kv.Value
        if($val -is [switch]){ if($val){ $argsLine += "-$name" } }
        else { $argsLine += "-$name `"$val`"" }
    }
    $arg = "-ExecutionPolicy Bypass -File `"$PSCommandPath`" " + ($argsLine -join ' ')
    Start-Process -FilePath "powershell.exe" -ArgumentList $arg -Verb RunAs
    return
}
$ErrorActionPreference = 'Stop'

# --------------------- Console helpers ---------------------
function WriteStep([string]$s){ Write-Host ("== " + $s + " ==") }
function WriteInfo([string]$s){ Write-Host $s }
function WriteWarn([string]$s){ Write-Host ("[!] " + $s) -ForegroundColor Yellow }
function WriteErr ([string]$s){ Write-Host ("[X] " + $s) -ForegroundColor Red }

# --------------------- Registry bases ---------------------
$MMBase = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio'
$RenderKey  = Join-Path $MMBase 'Render'
$CaptureKey = Join-Path $MMBase 'Capture'

# --------------------- Endpoint enumeration (null-safe) ---------------------
function Get-EndpointObjects {
    param([string[]]$Kinds = @('Render','Capture'))
    $list = @()
    foreach($kind in $Kinds){
        $base = if($kind -eq 'Render'){ $RenderKey } else { $CaptureKey }
        if(-not (Test-Path $base)){ continue }
        foreach($guidKey in Get-ChildItem $base -ErrorAction SilentlyContinue){
            $propsKey = Join-Path $guidKey.PSPath 'Properties'
            $fxKey    = Join-Path $guidKey.PSPath 'FxProperties'
            $name = '(unknown)'; $instId = '(unknown)'; $disableSysFx = '(unset)'; $hasClsid = $false

            # Read Properties once (null-safe)
            $propsObj = $null
            if(Test-Path $propsKey){ $propsObj = Get-ItemProperty -Path $propsKey -EA SilentlyContinue }
            if($propsObj){
                # Endpoint display name
                $n1 = $propsObj.'{a45c254e-df1c-4efd-8020-67d146a850e0},2'
                $n2 = $propsObj.'{b3f8fa53-0004-438e-9003-51a46e139bfc},14' # fallback
                if([string]::IsNullOrWhiteSpace($n1) -and -not [string]::IsNullOrWhiteSpace($n2)) { $n1 = $n2 }
                if(-not [string]::IsNullOrWhiteSpace($n1)){ $name = $n1.Trim() }

                # Device InstanceId mirror
                $iid = $propsObj.'{78C34FC8-104A-4D11-9F5B-700F2848BCA5},256'
                if(-not [string]::IsNullOrWhiteSpace($iid)){ $instId = $iid.Trim() }

                # Disable_SysFx may sit here on some systems
                if($propsObj.PSObject.Properties.Match('Disable_SysFx')){
                    $disableSysFx = [string]$propsObj.'Disable_SysFx'
                }
            }

            # FxProperties scan (null-safe, detect CLSID)
            $fxObj = $null
            if(Test-Path $fxKey){ $fxObj = Get-ItemProperty -Path $fxKey -EA SilentlyContinue }
            if($fxObj){
                # Primary source for Disable_SysFx
                if($fxObj.PSObject.Properties.Match('Disable_SysFx')){ $disableSysFx = [string]$fxObj.'Disable_SysFx' }

                $propNames = $fxObj.PSObject.Properties |
                    Where-Object { $_.MemberType -eq 'NoteProperty' } |
                    ForEach-Object { $_.Name }
                foreach($pn in $propNames){
                    try{
                        $v = $fxObj.$pn
                        if(($v -is [string]) -and ($v.Trim().ToLower() -eq $ApoClsid.ToLower())){ $hasClsid = $true; break }
                        if(($v -is [string[]]) -and ($v -contains $ApoClsid)){ $hasClsid = $true; break }
                    }catch{}
                }
            }

            $list += [pscustomobject]@{
                Kind = $kind
                EndpointGuid = $guidKey.PSChildName
                Name = $name
                InstanceId = $instId
                HasClsid = $hasClsid
                Disable_SysFx = $disableSysFx
                FxKey = $fxKey
                PropsKey = $propsKey
            }
        }
    }
    return $list
}

function Get-DeviceFx0ForInstance {
    param([string]$InstanceId)
    $path = "HKLM:\SYSTEM\CurrentControlSet\Enum\" + $InstanceId + "\Device Parameters\FX\0"
    if(Test-Path $path){ return Get-Item $path } else { return $null }
}

# --------------------- IPolicyConfig (Vista/Win10) ---------------------
if (-not ("PC.IPolicyConfigVista" -as [type])) {
Add-Type -Language CSharp -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace PC {
    [ComImport, Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")]
    public class _CPolicyConfigVistaClient {}

    [ComImport, Guid("F8679F50-850A-41CF-9C72-430F290290C8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPolicyConfigVista {
        int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string id, ERole role);
    }

    [ComImport, Guid("294935CE-F637-4E7C-A41B-AB255460B862"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPolicyConfig10 {
        int Unused1(); int Unused2(); int Unused3(); int Unused4(); int Unused5(); int Unused6(); int Unused7();
        int GetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string id, int role, IntPtr ppwfx);
        int SetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string id, IntPtr pwfxIn, IntPtr pwfxOut);
        int GetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string id, int role, IntPtr pDefault, IntPtr pMinimum);
        int SetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string id, IntPtr pPeriod);
        int GetShareMode([MarshalAs(UnmanagedType.LPWStr)] string id, IntPtr pMode);
        int SetShareMode([MarshalAs(UnmanagedType.LPWStr)] string id, IntPtr pMode);
        int GetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string id, ref PropertyKey key, IntPtr pv);
        int SetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string id, ref PropertyKey key, IntPtr pv);
        int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string id, ERole role);
        int Unused8(); int Unused9();
    }

    public enum ERole { eConsole=0, eMultimedia=1, eCommunications=2 }

    [StructLayout(LayoutKind.Sequential)]
    public struct PropertyKey { public Guid fmtid; public int pid; }
}
"@
} else {
    WriteInfo 'PolicyConfig types already loaded (skipped Add-Type).'
}

function Set-DefaultEndpointById {
    param([string]$EndpointGuid)
    $epId = $EndpointGuid
    # Vista interface (works broadly)
    try{
        $vista = [PC.IPolicyConfigVista][Activator]::CreateInstance([type]::GetTypeFromCLSID([Guid]'870af99c-171d-4f9e-af0d-e63df40c2bc9'))
        foreach($role in 0,1,2){ $null = $vista.SetDefaultEndpoint($epId, [PC.ERole]$role) }
        WriteInfo "Set default endpoint (Vista policy) => $epId"
        return
    }catch{}
    # Win10 interface
    try{
        $pc10 = [PC.IPolicyConfig10][Activator]::CreateInstance([type]::GetTypeFromCLSID([Guid]'870af99c-171d-4f9e-af0d-e63df40c2bc9'))
        foreach($role in 0,1,2){ $null = $pc10.SetDefaultEndpoint($epId, [PC.ERole]$role) }
        WriteInfo "Set default endpoint (Win10 policy) => $epId"
    }catch{
        WriteWarn ("SetDefaultEndpoint failed: " + $_.Exception.Message)
    }
}

# --------------------- COM/PE helpers ---------------------
function Get-RegisteredDllPathForClsid {
    param([string]$clsid)
    $k = "Registry::HKEY_CLASSES_ROOT\\CLSID\\" + $clsid + "\\InprocServer32"
    if(Test-Path $k){ return (Get-ItemProperty $k -Name '(default)').'(default)' } else { return $null }
}

function Parse-PEHeader {
    param([string]$path)
    $fs = [System.IO.File]::Open($path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::Read)
    try{
        $br = New-Object System.IO.BinaryReader($fs)
        # DOS
        if([System.Text.Encoding]::ASCII.GetString($br.ReadBytes(2)) -ne 'MZ'){ return $null }
        $fs.Seek(0x3C,0) | Out-Null
        $e_lfanew = $br.ReadInt32()
        # NT headers
        $fs.Seek($e_lfanew,0) | Out-Null
        $sig = $br.ReadBytes(4)
        if($sig[0] -ne 0x50 -or $sig[1] -ne 0x45 -or $sig[2] -ne 0x00 -or $sig[3] -ne 0x00){ return $null }
        $machine = $br.ReadUInt16()       # 0x8664 for AMD64
        $br.ReadUInt16() | Out-Null       # NumberOfSections
        $time = $br.ReadUInt32()          # TimeDateStamp (epoch seconds)
        $fs.Seek(0x10,1) | Out-Null       # Skip to OptionalHeader.Magic
        $magic = $br.ReadUInt16()         # 0x20B => PE32+
        $isPE32Plus = ($magic -eq 0x20B)
        # Subsystem offset from OptionalHeader start
        $subsysOff = if($isPE32Plus){ 0x5C } else { 0x44 }
        # We already consumed 2 bytes of OptionalHeader; seek to Subsystem accordingly
        $fs.Seek($subsysOff - 2,1) | Out-Null
        $subsystem = $br.ReadUInt16()
        [pscustomobject]@{
            Machine = ('0x{0:X4}' -f $machine)
            PE32Plus = $isPE32Plus
            Subsystem = ('0x{0:X4}' -f $subsystem)
            TimeStamp = ([DateTime]::UnixEpoch).AddSeconds([double]$time).ToString('yyyy-MM-dd HH:mm:ss')
        }
    } finally { $fs.Dispose() }
}

# --------------------- Optional: Mirror device FX to endpoint FxProperties ---------------------
function Mirror-DeviceFx0-To-EndpointFx {
    param(
        [Parameter(Mandatory=$true)][string]$DeviceInstanceId,
        [Parameter(Mandatory=$true)][pscustomobject]$EndpointObject
    )
    $fx0 = Get-DeviceFx0ForInstance -InstanceId $DeviceInstanceId
    if(-not $fx0){ WriteWarn "  No FX\\0 under device. Skip mirror."; return }
    $fx0Obj = Get-ItemProperty $fx0.PSPath -EA SilentlyContinue
    if(-not $fx0Obj){ WriteWarn "  FX\\0 values not readable."; return }

    $dst = $EndpointObject.FxKey
    if(-not (Test-Path $dst)){ New-Item -Path $dst -Force | Out-Null }

    $props = $fx0Obj.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' }
    foreach($p in $props){
        $name = $p.Name; $val = $p.Value
        try{
            if($val -is [string[]]){
                New-ItemProperty -Path $dst -Name $name -Value ([string[]]$val) -PropertyType MultiString -Force | Out-Null
            } else {
                New-ItemProperty -Path $dst -Name $name -Value ([string]$val) -PropertyType String -Force | Out-Null
            }
            WriteInfo ("  -> FxProperties." + $name + " = " + ((@($val) -join ', ')))
        }catch{
            WriteWarn ("  Failed write " + $name + ": " + $_.Exception.Message)
        }
    }
}

# --------------------- Actions start ---------------------
WriteStep '1) Installed Class=Extension driver packages'
try {
    # Locale-agnostic fallback: print raw pnputil lines (Release Names etc.)
    $lines = & pnputil /enum-drivers 2>$null
    if($lines){ $lines | Where-Object { $_.Trim() -ne '' } | ForEach-Object { WriteInfo $_ } }
} catch { WriteWarn 'pnputil not available or failed.' }

WriteStep '2) Device-side FX (Enum\\...\\Device Parameters\\FX\\0)'
$matched = Get-PnpDevice -EA SilentlyContinue | Where-Object { $_.InstanceId -like "$HwPrefix*" }
if($matched){
    foreach($d in $matched){
        WriteInfo ("- Instance: " + $d.InstanceId)
        $fx0 = Get-DeviceFx0ForInstance -InstanceId $d.InstanceId
        if($fx0){
            $fx0Obj = Get-ItemProperty $fx0.PSPath -EA SilentlyContinue
            if($fx0Obj){
                foreach($np in ($fx0Obj.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' })){
                    $v = $np.Value; $vs = if($v -is [Array]){ $v -join ', ' } else { [string]$v }
                    WriteInfo ("  FX\\0 " + $np.Name + " = " + $vs)
                }
            } else { WriteWarn '  FX\\0 present but not readable.' }
        } else {
            WriteWarn '  No FX\\0 key present (device-side)!'
        }
    }
} else { WriteWarn 'No PnP instance matched HwPrefix.' }

WriteStep '3) Endpoint strict map (by InstanceId -> endpoint)'
$eps = Get-EndpointObjects -Kinds $Kinds
$mapHits = @()
if($eps){ $mapHits = $eps | Where-Object { $_.InstanceId -like "$HwPrefix*" } }
if($mapHits -and ($mapHits | Measure-Object).Count -gt 0){
    $mapHits | Select Kind,EndpointGuid,Name,Disable_SysFx,HasClsid,InstanceId | Format-Table -AutoSize
} else {
    WriteWarn 'No endpoint mapped by InstanceId.'
}

if($SearchFx){
    WriteStep '4) SearchFx — scan all endpoints for your CLSID'
    $hits = @(); if($eps){ $hits = $eps | Where-Object { $_.HasClsid } }
    if($hits -and ($hits | Measure-Object).Count -gt 0){
        $hits | Select Kind,EndpointGuid,Name,Disable_SysFx | Sort-Object Kind,Name | Format-Table -AutoSize
    } else { WriteWarn 'No endpoint with your CLSID in FxProperties.' }
}

if($ForceWriteEndpointFx){
    WriteStep '5) ForceWriteEndpointFx — mirror FX\\0 to endpoint FxProperties'
    $targets = if($mapHits -and ($mapHits | Measure-Object).Count -gt 0){ $mapHits } else { $eps | Where-Object { $_.Kind -eq 'Render' } }
    if($targets){
        foreach($t in $targets){
            WriteInfo ("Endpoint: " + $t.Name + " [" + $t.EndpointGuid + "]")
            if($t.InstanceId -and $t.InstanceId -ne '(unknown)'){
                Mirror-DeviceFx0-To-EndpointFx -DeviceInstanceId $t.InstanceId -EndpointObject $t
            } else {
                WriteWarn '  Missing InstanceId; cannot mirror.'
            }
        }
    } else { WriteWarn '  No endpoint to write.' }
}

if($ClearDisableSysFx){
    WriteStep '6) ClearDisableSysFx — set Disable_SysFx=0 on matched endpoints'
    $targets = if($SearchFx){ $eps | Where-Object {$_.HasClsid -and $_.Kind -eq 'Render'} } else { $mapHits }
    foreach($t in $targets){
        $ok = $false
        foreach($keyPath in @($t.PropsKey, $t.FxKey)){
            if(-not $keyPath){ continue }
            try{
                if(-not (Test-Path $keyPath)){ New-Item -Path $keyPath -Force | Out-Null }
                New-ItemProperty -Path $keyPath -Name 'Disable_SysFx' -Value 0 -PropertyType DWord -Force | Out-Null
                $ok = $true
            }catch{}
            if($ok){ break }
        }
        if($ok){ WriteInfo ("- " + $t.EndpointGuid + " Disable_SysFx=0") } else { WriteWarn ("- " + $t.EndpointGuid + " failed to set Disable_SysFx") }
    }
}

if($SetDefaultToTarget){
    WriteStep '7) SetDefaultToTarget — switch default render endpoint'
    $candidate = $null
    if($eps){
        $candidate = ($eps | Where-Object { $_.Kind -eq 'Render' -and $_.HasClsid -and $_.InstanceId -like "$HwPrefix*" } | Select-Object -First 1)
        if(-not $candidate){ $candidate = ($eps | Where-Object { $_.Kind -eq 'Render' -and $_.HasClsid } | Select-Object -First 1) }
        if(-not $candidate){ $candidate = ($eps | Where-Object { $_.Kind -eq 'Render' -and $_.InstanceId -like "$HwPrefix*" } | Select-Object -First 1) }
    }
    if($candidate){
        WriteInfo ("Target: " + $candidate.Name + " [" + $candidate.EndpointGuid + "]")
        Set-DefaultEndpointById -EndpointGuid $candidate.EndpointGuid
    } else { WriteWarn 'No suitable render endpoint to set as default.' }
}

if($CycleDevice){
    WriteStep '8) CycleDevice — disable/enable matched device'
    $devs = Get-PnpDevice -EA SilentlyContinue | Where-Object { $_.InstanceId -like "$HwPrefix*" }
    if($devs){
        foreach($d in $devs){
            try{ Disable-PnpDevice -InstanceId $d.InstanceId -Confirm:$false -EA Stop; Start-Sleep -Seconds 1; Enable-PnpDevice -InstanceId $d.InstanceId -Confirm:$false -EA Stop; WriteInfo ("Cycled " + $d.InstanceId) } catch { WriteWarn $_.Exception.Message }
        }
    } else { WriteWarn 'No PnP device found to cycle.' }
}

if($RestartAudio){
    WriteStep '9) RestartAudio — restart audiosrv and AudioEndpointBuilder'
    foreach($svc in 'Audiosrv','AudioEndpointBuilder'){
        try{ Restart-Service -Name $svc -Force -EA Stop; WriteInfo ("Restarted service: " + $svc) }catch{ WriteWarn ("Failed to restart " + $svc + ": " + $_.Exception.Message) }
    }
}

WriteStep '10) CLSID/COM registration and DLL path'
$dllPath = Get-RegisteredDllPathForClsid -clsid $ApoClsid
if($dllPath){ WriteInfo ("CLSID InprocServer32 = " + $dllPath) } else { WriteWarn 'CLSID not found under HKCR\\CLSID. COM registration may be missing.' }
if($dllPath -and (Test-Path $dllPath)){
    WriteInfo ("DLL Exists? True  => " + $dllPath)
} else { WriteWarn 'DLL file not found on disk (bitness/path?).' }

if($TestCOM -and $dllPath){
    try{
        $obj = [Activator]::CreateInstance([type]::GetTypeFromCLSID([Guid]$ApoClsid))
        [void]$obj.GetType()
        WriteInfo 'COM CoCreateInstance() OK'
    }catch{ WriteWarn ("COM CoCreateInstance() failed: " + $_.Exception.Message) }
}

if($ParsePE -and $dllPath -and (Test-Path $dllPath)){
    WriteStep '11) ParsePE — APO binary'
    $pe = Parse-PEHeader -path $dllPath
    if($pe){ WriteInfo ("Machine=" + $pe.Machine + ", PE32Plus=" + $pe.PE32Plus + ", Subsystem=" + $pe.Subsystem + ", Time=" + $pe.TimeStamp) } else { WriteWarn 'Failed to parse PE header.' }
}

if($PlayTest){
    WriteStep '12) PlayTest — start audiodg by playing WAV on default device'
    $wav = Join-Path $env:WINDIR 'Media\\Windows Background.wav'
    if(-not (Test-Path $wav)){ $wav = Join-Path $env:WINDIR 'Media\\Windows Ding.wav' }
    try{ $p = New-Object System.Media.SoundPlayer($wav); $p.PlaySync(); WriteInfo ("Played: " + (Split-Path $wav -Leaf)) }
    catch{ WriteWarn ("SoundPlayer failed: " + $_.Exception.Message) }
}

WriteStep '13) Is audiodg.exe running and is your DLL loaded?'
try{
    $adg = Get-Process audiodg -EA SilentlyContinue
    if($adg){ WriteInfo ('audiodg.exe PID=' + $adg.Id) } else { WriteWarn 'audiodg.exe not running.' }
    $tl = & tasklist /fi "imagename eq audiodg.exe" /m $ApoDll 2>$null
    if($tl -and ($tl -match '^audiodg\.exe')){
        $tl | ForEach-Object { WriteInfo $_ }
        WriteInfo ("=> OK: audiodg.exe has loaded " + $ApoDll)
    } else {
        WriteWarn ("audiodg.exe has NOT loaded " + $ApoDll + " yet. Start playback or use -PlayTest.")
    }
}catch{ WriteWarn 'tasklist failed.' }

if($ShowEvents){
    WriteStep '14) Recent audio logs (Audio, Audio-Effects-Manager)'
    foreach($log in 'Microsoft-Windows-Audio/Operational','Microsoft-Windows-Audio-Effects-Manager/Operational'){
        try{
            WriteInfo ('-- Log: ' + $log)
            Get-WinEvent -LogName $log -MaxEvents $EventCount | Select-Object TimeCreated, Id, LevelDisplayName, Message | Format-List
        }catch{ WriteWarn ('Cannot read log ' + $log + ': ' + $_.Exception.Message) }
    }
}

if($ScanSetupLog){
    WriteStep '15) Scan setupapi.dev.log for ExtId/HwPrefix'
    $log = Join-Path $env:WINDIR 'INF\\setupapi.dev.log'
    if(Test-Path $log){
        $pat = [regex]::Escape($HwPrefix) + '|' + [regex]::Escape($ExtId)
        Get-Content -Path $log -Encoding ASCII | Select-String -Pattern $pat -AllMatches | ForEach-Object { $_.Line } | ForEach-Object { WriteInfo $_ }
    } else { WriteWarn 'setupapi.dev.log not found.' }
}

WriteStep 'Done.'
