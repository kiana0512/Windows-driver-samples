# Apo_DeepVerify_Fix_Auto.ps1
# Zero-parameter, end-to-end APO deep verify & (optional) auto-fix.
# Dynamic endpoint GUID discovery (multi-method), robust registry access.
# Save as UTF-8 (no BOM). Run in PowerShell 7+ as Administrator.

#requires -RunAsAdministrator
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# =========================
# [0] Constants (stable, hardcoded)
# =========================
# Your USB function interface (MI_00) instance (hardcoded)
$UsbMi00InstanceId   = 'USB\VID_0A67&PID_30A2&MI_00\7&3B1FF4EF&0&0000'
# Expected ContainerId (set $null to skip strict matching)
$ExpectedContainerId = '{16b5dc8e-d125-5b68-bfb3-10182a74f929}'

# APO (hardcoded)
$ApoClsid          = '{8E3E0B71-5B8A-45C9-9B3D-3A2E5B418A10}'
$ProcessingModePM7 = '{C18E2F7E-933D-4965-B7D1-1EEF228D2AF3}'
$ApoDllName        = 'MyCompanyEfxApo.dll'
$ApoDllPath        = 'C:\Windows\System32\MyCompanyEfxApo.dll'
$InfName           = 'MyCompanyUsbApoExt.inf'

# Policy (embedded; no parameters)
$AutoFix               = $true   # auto-fix FxProperties mismatch
$RestartAudioAfterFix  = $true   # restart audio services after fix

# =========================
# [1] Output helpers
# =========================
function _W([string]$tag,[string]$msg,[ConsoleColor]$fg='Gray'){
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  Write-Host "[$tag] $msg" -ForegroundColor $fg
}
function INFO($m){ _W 'INFO' $m 'Gray' }
function STEP($m){ _W 'STEP' $m 'Cyan' }
function PASS($m){ _W 'PASS' $m 'Green' }
function WARN($m){ _W 'WARN' $m 'Yellow' }
function FAIL($m){ _W 'FAIL' $m 'Red' }

function Invoke-Try([ScriptBlock]$b,[string]$desc){
  try{ & $b }catch{ FAIL "$desc => $($_.Exception.Message)"; throw }
}

# =========================
# [2] Utilities
# =========================
function Get-ContainerId-ByPnP(){
  $prop = Get-PnpDeviceProperty -InstanceId $UsbMi00InstanceId -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction Stop
  return $prop.Data.ToString()
}

function Get-PnP-Children(){
  $p = Get-PnpDeviceProperty -InstanceId $UsbMi00InstanceId -KeyName 'DEVPKEY_Device_Children' -ErrorAction Stop
  return @($p.Data)
}

function Extract-Guid-From-MMDev-Instance([string]$mmdevInstance){
  # e.g. SWD\MMDEVAPI\{0.0.0.00000000}.{E1F9...}
  if($mmdevInstance -match '\.\{([0-9A-Fa-f-]{36})\}$'){
    return '{' + $Matches[1].ToUpper() + '}'
  }
  return $null
}

function Get-MMDevices-Path([string]$flow,[string]$guidCurly){
  # flow: Render / Capture
  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$flow\$guidCurly"
}

function Restart-AudioServices(){
  STEP "Restart Windows Audio services"
  sc.exe stop audiosrv | Out-Null
  sc.exe stop audioendpointbuilder | Out-Null
  Start-Sleep -Seconds 1
  sc.exe start audioendpointbuilder | Out-Null
  sc.exe start audiosrv | Out-Null
  PASS "Audio services restarted"
}

function Grant-RegTree([string]$path){
  if(!(Test-Path $path)){ return }
  $sys = New-Object System.Security.Principal.NTAccount('NT AUTHORITY','SYSTEM')
  $adm = New-Object System.Security.Principal.NTAccount('BUILTIN','Administrators')
  $acl = Get-Acl $path
  $acl.SetOwner($sys)
  $r1 = New-Object System.Security.AccessControl.RegistryAccessRule($sys,'FullControl','ContainerInherit,ObjectInherit','None','Allow')
  $r2 = New-Object System.Security.AccessControl.RegistryAccessRule($adm,'FullControl','ContainerInherit,ObjectInherit','None','Allow')
  $acl.SetAccessRule($r1) | Out-Null
  $acl.AddAccessRule($r2) | Out-Null
  Set-Acl -Path $path -AclObject $acl
}

function Ensure-REG-SZ([string]$regPath,[string]$name,[string]$expected){
  $cur = $null
  try{
    $cur = Get-ItemPropertyValue -Path $regPath -Name $name -ErrorAction Stop
  }catch{}
  if($null -eq $cur){
    if(!(Test-Path $regPath)){ New-Item -Path $regPath -Force | Out-Null }
    New-ItemProperty -Path $regPath -Name $name -PropertyType String -Value $expected -Force | Out-Null
    return $true
  }elseif($cur -ne $expected){
    Set-ItemProperty -Path $regPath -Name $name -Value $expected
    return $true
  }
  return $false
}

# Get unnamed/default registry value; try HKCR/HKLM; expand env vars.
function Get-RegistryDefault([string]$keyPath){
  try{
    $key = Get-Item -Path $keyPath -ErrorAction Stop
    $val = $key.GetValue('', $null, 'DoNotExpandEnvironmentNames')
    if($null -ne $val){ return [string]$val }
  }catch{}
  return $null
}
function Expand-EnvPath([string]$p){
  if([string]::IsNullOrWhiteSpace($p)){ return $p }
  return [Environment]::ExpandEnvironmentVariables($p)
}

function Test-Audiodg-HasModule([string]$moduleName){
  $p = Get-Process -Name audiodg -ErrorAction SilentlyContinue
  if(!$p){ return $false }
  foreach($proc in $p){
    try{
      $m = $proc.Modules | Where-Object { $_.ModuleName -ieq $moduleName }
      if($m){ return $true }
    }catch{}
  }
  return $false
}

# Robust helper: find MMDevices GUID by matching any Properties value to a given InstanceId.
function Find-MMDevice-ByInstanceId([string]$flow,[string]$inst){
  $base = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$flow"
  if(!(Test-Path $base)){ return $null }
  foreach($k in Get-ChildItem $base -ErrorAction SilentlyContinue){
    $props = Join-Path $k.PSPath 'Properties'
    if(!(Test-Path $props)){ continue }
    # Preferred property: PKEY_Device_InstanceId
    $iid = $null
    try{
      $iid = Get-ItemPropertyValue -Path $props -Name '{78C34FC8-104A-4D11-9F5B-700F2848BCA5},256' -ErrorAction Stop
    }catch{}
    if($iid -and ($iid -ieq $inst)){ return $k.PSChildName }
    # Fallback: scan all REG_SZ values and compare
    try{
      $all = Get-ItemProperty -Path $props -ErrorAction SilentlyContinue
      if($all){
        foreach($pname in ($all.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name)){
          $val = $null
          try{ $val = Get-ItemPropertyValue -Path $props -Name $pname -ErrorAction Stop }catch{}
          if(($val -is [string]) -and ($val -ieq $inst)){ return $k.PSChildName }
        }
      }
    }catch{}
  }
  return $null
}

# =========================
# [3] Environment & services
# =========================
STEP "Environment"
$psver = $PSVersionTable.PSVersion.ToString()
$arch  = if([Environment]::Is64BitProcess){ 'X64' } else { 'X86' }
$admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if(!$admin){ FAIL "Run as Administrator"; throw }
PASS "PS=$psver; Arch=$arch; Admin=$admin"

Invoke-Try {
  $s1 = Get-Service -Name audiosrv
  $s2 = Get-Service -Name audioendpointbuilder
  if($s1.Status -ne 'Running' -or $s2.Status -ne 'Running'){
    WARN "Audio services not fully running, starting..."
    if($s2.Status -ne 'Running'){ Start-Service audioendpointbuilder }
    if($s1.Status -ne 'Running'){ Start-Service audiosrv }
  }
  PASS "AudioSrv=$($s1.Status); AudioEndpointBuilder=$($s2.Status)"
} "Check/start audio services"

# =========================
# [4] COM / DLL / DriverStore
# =========================
STEP "COM/PE and signature"

# 4.1 COM InprocServer32 default value (HKCR preferred, HKLM fallback)
$clsidKeyHKCR = "HKCR:\CLSID\$ApoClsid\InprocServer32"
$clsidKeyHKLM = "HKLM:\SOFTWARE\Classes\CLSID\$ApoClsid\InprocServer32"

$comPath = Get-RegistryDefault $clsidKeyHKCR
if(-not $comPath){ $comPath = Get-RegistryDefault $clsidKeyHKLM }

if($comPath){
  $comPathExpanded = Expand-EnvPath $comPath
  PASS "COM InprocServer32 => $comPathExpanded"
}else{
  FAIL "COM InprocServer32 missing default value (HKCR/HKLM) for $ApoClsid"
}

# 4.2 DLL presence & signature
if(Test-Path $ApoDllPath){
  PASS "DLL Exists => $ApoDllPath"
  $sig = Get-AuthenticodeSignature -FilePath $ApoDllPath
  $status = $sig.Status
  $signer = if($sig.SignerCertificate){ $sig.SignerCertificate.Subject } else { '' }
  PASS "DLL Signature => Status=$status; Signer=$signer"
}else{
  FAIL "DLL not found => $ApoDllPath"
}

# 4.3 DriverStore
Invoke-Try {
  $infFound = (pnputil /enum-drivers) -match [Regex]::Escape($InfName) | Select-Object -First 1
  if($infFound){ PASS "DriverStore => found '$InfName' = True" } else { WARN "DriverStore => '$InfName' not found (non-fatal)" }
} "Enumerate drivers"

# =========================
# [5] Device tree & dynamic endpoint discovery (multi-method)
# =========================
STEP "Device tree and dynamic endpoint discovery (multi-method)"

# 5.1 MI_00 PnP presence
$mi = Get-PnpDevice -InstanceId $UsbMi00InstanceId -ErrorAction SilentlyContinue
if(!$mi){ FAIL "USB MI_00 not found: $UsbMi00InstanceId"; throw }
PASS "USB MI_00 present: $($mi.Status) | $($mi.InstanceId)"

# 5.2 ContainerId (method A: PnP)
$cidA = Get-ContainerId-ByPnP
INFO "ContainerId (PnP) => $cidA"
if($ExpectedContainerId){
  if($cidA -ieq $ExpectedContainerId){ PASS "ContainerId matches expected" } else { WARN "ContainerId differs from expected: $cidA <> $ExpectedContainerId" }
}

# 5.3 Children (method A: PnP)
$children = Get-PnP-Children
INFO "Children => $($children -join '; ')"

# Extract render/capture endpoints
$mmdevRenderA  = $children | Where-Object { $_ -like 'SWD\MMDEVAPI\{0.0.0.00000000}.*' } | Select-Object -First 1
$mmdevCaptureA = $children | Where-Object { $_ -like 'SWD\MMDEVAPI\{0.0.1.00000000}.*' } | Select-Object -First 1

if($mmdevRenderA){ PASS "Render(A: PnP) => $mmdevRenderA" } else { FAIL "Render(A: PnP) => not found" }
if($mmdevCaptureA){ PASS "Capture(A: PnP) => $mmdevCaptureA" } else { FAIL "Capture(A: PnP) => not found" }

$rguidA = if($mmdevRenderA){ Extract-Guid-From-MMDev-Instance $mmdevRenderA }
$cguidA = if($mmdevCaptureA){ Extract-Guid-From-MMDev-Instance $mmdevCaptureA }
INFO "GUID(A) => Render=$rguidA | Capture=$cguidA"

# 5.4 Method B: MMDevices (reverse by InstanceId), fully guarded
$rguidB = if($mmdevRenderA){ Find-MMDevice-ByInstanceId 'Render' $mmdevRenderA } else { $null }
$cguidB = if($mmdevCaptureA){ Find-MMDevice-ByInstanceId 'Capture' $mmdevCaptureA } else { $null }

if($rguidB){ PASS "Render(B: MMDevices) => $rguidB" } else { WARN "Render(B: MMDevices) => not found by InstanceId (not fatal)" }
if($cguidB){ PASS "Capture(B: MMDevices) => $cguidB" } else { WARN "Capture(B: MMDevices) => not found by InstanceId (not fatal)" }

# 5.5 Final endpoint GUIDs (prefer A, fallback B)
$RenderGuid  = if($rguidA){ $rguidA } elseif($rguidB){ $rguidB } else { $null }
$CaptureGuid = if($cguidA){ $cguidA } elseif($cguidB){ $cguidB } else { $null }

if($RenderGuid){ PASS "Render GUID (resolved) => $RenderGuid" } else { FAIL "Render GUID => unresolved" }
if($CaptureGuid){ PASS "Capture GUID (resolved) => $CaptureGuid" } else { FAIL "Capture GUID => unresolved" }

# 5.6 Presence & ContainerId check for endpoints
if($mmdevRenderA){
  $r_dev = Get-PnpDevice -InstanceId $mmdevRenderA -ErrorAction SilentlyContinue
  if($r_dev){ PASS "Render endpoint present (PnP) => OK" } else { FAIL "Render endpoint missing (PnP)" }
  $r_cid = (Get-PnpDeviceProperty -InstanceId $mmdevRenderA -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction SilentlyContinue).Data
  if($r_cid){
    if(!$ExpectedContainerId -or ($r_cid -ieq $cidA)){ PASS "Render ContainerId OK => $r_cid" } else { WARN "Render ContainerId differs => $r_cid" }
  }
}
if($mmdevCaptureA){
  $c_dev = Get-PnpDevice -InstanceId $mmdevCaptureA -ErrorAction SilentlyContinue
  if($c_dev){ PASS "Capture endpoint present (PnP) => OK" } else { FAIL "Capture endpoint missing (PnP)" }
  $c_cid = (Get-PnpDeviceProperty -InstanceId $mmdevCaptureA -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction SilentlyContinue).Data
  if($c_cid){
    if(!$ExpectedContainerId -or ($c_cid -ieq $cidA)){ PASS "Capture ContainerId OK => $c_cid" } else { WARN "Capture ContainerId differs => $c_cid" }
  }
}

# =========================
# [6] Device-level FxProperties (check & optional fix)
# =========================
STEP "Device-level FxProperties (7/15/PM7) check"
$fxKey = "HKLM:\SYSTEM\CurrentControlSet\Enum\$UsbMi00InstanceId\Device Parameters\FxProperties"
$fxChanged = $false
Invoke-Try {
  if(!(Test-Path $fxKey)){
    WARN "FxProperties key missing: $fxKey"
    if($AutoFix){
      STEP "Create FxProperties and grant permissions"
      Grant-RegTree ("HKLM:\SYSTEM\CurrentControlSet\Enum\" + $UsbMi00InstanceId)
      New-Item -Path $fxKey -Force | Out-Null
      $fxChanged = $true
    }
  }

  $v7  = $null; $v15 = $null; $pm7 = $null
  try{ $v7  = Get-ItemPropertyValue -Path $fxKey -Name '7'   -ErrorAction Stop }catch{}
  try{ $v15 = Get-ItemPropertyValue -Path $fxKey -Name '15'  -ErrorAction Stop }catch{}
  try{ $pm7 = Get-ItemPropertyValue -Path $fxKey -Name 'PM7' -ErrorAction Stop }catch{}

  $ok7  = ($v7  -ieq $ApoClsid)
  $ok15 = ($v15 -ieq $ApoClsid)
  $okpm = ($pm7 -ieq $ProcessingModePM7)

  if($ok7 -and $ok15 -and $okpm){
    PASS ("FX\0 => 7='{0}'  15='{1}'  PM7='{2}'" -f $v7,$v15,$pm7)
  }else{
    WARN ("FxProperties mismatch: 7='{0}'  15='{1}'  PM7='{2}' (expect EFX={3}, PM7={4})" -f $v7,$v15,$pm7,$ApoClsid,$ProcessingModePM7)
    if($AutoFix){
      STEP "Write expected FxProperties"
      Grant-RegTree ("HKLM:\SYSTEM\CurrentControlSet\Enum\" + $UsbMi00InstanceId)
      $c1 = Ensure-REG-SZ $fxKey '7'   $ApoClsid
      $c2 = Ensure-REG-SZ $fxKey '15'  $ApoClsid
      $c3 = Ensure-REG-SZ $fxKey 'PM7' $ProcessingModePM7
      if($c1 -or $c2 -or $c3){ $fxChanged = $true }
      PASS ("Set: 7='{0}'  15='{1}'  PM7='{2}'" -f $ApoClsid,$ApoClsid,$ProcessingModePM7)
    }
  }
} "FxProperties check/fix"

# =========================
# [7] Endpoint-level: DisableEnhancements & key existence
# =========================
STEP "Endpoint-level checks (DisableEnhancements, MMDevices existence)"

function Read-DisableEnhancements([string]$flow,[string]$guidCurly){
  $key = Get-MMDevices-Path $flow $guidCurly
  if(!(Test-Path $key)){ return $null }
  $p = Join-Path $key 'Properties'
  if(!(Test-Path $p)){ return $null }
  try{ return Get-ItemPropertyValue -Path $p -Name 'DisableEnhancements' -ErrorAction Stop }catch{ return $null }
}

if($RenderGuid){
  $rKey = Get-MMDevices-Path 'Render' $RenderGuid
  if(Test-Path $rKey){ PASS "Render MMDevices key => $rKey" } else { WARN "Render MMDevices key missing => $rKey" }
  $de_r = Read-DisableEnhancements 'Render' $RenderGuid
  if($de_r -ne $null){
    $msg = if($de_r -eq 0){ 'Enabled' } elseif($de_r -eq 1){ 'Disabled' } else { "Unknown($de_r)" }
    PASS "EP(Render) DisableEnhancements=$de_r ($msg)"
  }else{
    WARN "EP(Render) DisableEnhancements not found (treat as enabled)"
  }
}

if($CaptureGuid){
  $cKey = Get-MMDevices-Path 'Capture' $CaptureGuid
  if(Test-Path $cKey){ PASS "Capture MMDevices key => $cKey" } else { WARN "Capture MMDevices key missing => $cKey" }
  $de_c = Read-DisableEnhancements 'Capture' $CaptureGuid
  if($de_c -ne $null){
    $msg = if($de_c -eq 0){ 'Enabled' } elseif($de_c -eq 1){ 'Disabled' } else { "Unknown($de_c)" }
    PASS "EP(Capture) DisableEnhancements=$de_c ($msg)"
  }else{
    WARN "EP(Capture) DisableEnhancements not found (treat as enabled)"
  }
}

# =========================
# [8] Policy keys
# =========================
STEP "Policy keys (system audio policies)"
$polKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Audio'
$DisableLegacyAudioEffects = $null; $DisableSystemEffects = $null; $EnableCompositeFx = $null
try{ $DisableLegacyAudioEffects = Get-ItemPropertyValue -Path $polKey -Name 'DisableLegacyAudioEffects' -ErrorAction Stop }catch{}
try{ $DisableSystemEffects      = Get-ItemPropertyValue -Path $polKey -Name 'DisableSystemEffects'      -ErrorAction Stop }catch{}
try{ $EnableCompositeFx         = Get-ItemPropertyValue -Path $polKey -Name 'EnableCompositeFx'         -ErrorAction Stop }catch{}
INFO ("Policy: DisableLegacyAudioEffects={0}; DisableSystemEffects={1}; EnableCompositeFx={2}" -f $DisableLegacyAudioEffects,$DisableSystemEffects,$EnableCompositeFx)

# =========================
# [9] Audio event log (recent 120)
# =========================
STEP "Collect Microsoft-Windows-Audio/Operational recent events"
try{
  $logs = Get-WinEvent -LogName 'Microsoft-Windows-Audio/Operational' -MaxEvents 120 -ErrorAction Stop |
          Select-Object TimeCreated, Id, LevelDisplayName, Message
  PASS "Fetched $($logs.Count) events (suppressing verbose print; use Out-GridView if needed)"
}catch{
  WARN "Get-WinEvent failed: $($_.Exception.Message)"
}

# =========================
# [10] audiodg module presence
# =========================
STEP "audiodg module presence"
$loaded = Test-Audiodg-HasModule $ApoDllName
if($loaded){
  PASS "audiodg has $ApoDllName"
}else{
  WARN "audiodg not showing $ApoDllName (start playback/recording, then re-check)"
}

# Optional restart after FxProperties change
if($AutoFix -and $RestartAudioAfterFix -and $fxChanged){
  Restart-AudioServices
  Start-Sleep -Seconds 1
  if(Test-Audiodg-HasModule $ApoDllName){ PASS "After restart: audiodg has $ApoDllName" } else { WARN "After restart: still not seen (trigger a shared-mode stream)" }
}

# =========================
# [11] Summary
# =========================
STEP "Summary"
INFO "USB MI_00: $UsbMi00InstanceId"
INFO "ContainerId(PnP): $cidA"
INFO "EFX CLSID: $ApoClsid | PM7: $ProcessingModePM7"
INFO "APO DLL: $ApoDllPath"
INFO "INF: $InfName"
INFO "Render GUID:  $RenderGuid"
INFO "Capture GUID: $CaptureGuid"
PASS "Apo DeepVerify finished."
