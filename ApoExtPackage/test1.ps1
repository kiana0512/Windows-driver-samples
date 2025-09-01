<# 
  Apo_DeepVerify_Fixed.ps1
  仅用固定锚点做完整验证：功能 devnode（MEDIA）、Render/Capture 两个端点 GUID、CLSID、DLL、INF
  不枚举、不搜索、不依赖可变 GUID。
#>

# ===== 固定锚点（按你给的信息写死） =====
$UsbInstance = 'USB\VID_0A67&PID_30A2&MI_00\7&3B1FF4EF&0&0000'  # 功能 devnode（MEDIA）
$RenderGuid  = 'e1f9ea7e-1a96-42ef-bead-e850164bb076'            # Render 端点 GUID（来自 Endpoint_Render.txt）
$CaptureGuid = '48fb70d7-bfac-4ab1-bcdd-34f44a30e26c'            # Capture 端点 GUID（来自 Endpoint_Capture.txt）
$ApoClsid    = '{8E3E0B71-5B8A-45C9-9B3D-3A2E5B418A10}'
$DllName     = 'MyCompanyEfxApo.dll'
$InfName     = 'MyCompanyUsbApoExt.inf'

$ErrorActionPreference = 'Stop'
function _ok([string]$m){ Write-Host "[PASS] $m" -ForegroundColor Green }
function _ng([string]$m){ Write-Host "[FAIL] $m" -ForegroundColor Red }
function _wr([string]$m){ Write-Host "[WARN] $m" -ForegroundColor Yellow }
function _info([string]$m){ Write-Host "[INFO] $m" -ForegroundColor Cyan }

# 小工具
function Read-Dword($path,$name){
  try{ $v=(Get-ItemProperty $path -ErrorAction Stop).$name; if($v -is [int]){return $v}else{return $null} }catch{ return $null }
}
function Read-Default($path){
  try{ return (Get-ItemProperty $path -ErrorAction Stop)."(default)" }catch{ return $null }
}

# === 0) 环境/服务/COM/DLL ===
$arch  = if([Environment]::Is64BitProcess){"X64"}else{"X86"}
$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
_ok ("Apo DeepVerify start: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
_ok ("PSVersion={0}; Arch={1}; Admin={2}" -f $PSVersionTable.PSVersion, $arch, $admin)
if(-not $admin){ _wr "建议以管理员身份运行。" }

try{
  $as  = Get-Service -Name 'audiosrv' -ErrorAction SilentlyContinue
  $aeb = Get-Service -Name 'AudioEndpointBuilder' -ErrorAction SilentlyContinue
  if($as -and $as.Status -eq 'Running'){ _ok ("AudioSrv Running => OK | Status={0}" -f $as.Status) } else { _ng ("AudioSrv Running => NG | Status={0}" -f $as.Status) }
  if($aeb -and $aeb.Status -eq 'Running'){ _ok ("AudioEndpointBuilder Running => OK | Status={0}" -f $aeb.Status) } else { _ng ("AudioEndpointBuilder Running => NG | Status={0}" -f $aeb.Status) }
}catch{ _wr ("读取服务状态失败：{0}" -f $_.Exception.Message) }

# COM 注册（x64 视图）
try{
  $cls = "HKLM:\SOFTWARE\Classes\CLSID\$ApoClsid\InprocServer32"
  $inproc = Read-Default $cls
  if($inproc -and (Test-Path $inproc)){ _ok ("COM x64 InprocServer32 => OK | {0}" -f $inproc) }
  else { _ng "COM x64 InprocServer32 => NG | 值缺失或 DLL 不存在" }
}catch{ _ng ("COM x64 InprocServer32 => NG | {0}" -f $_.Exception.Message) }

# DLL 存在/签名
$sysdll = Join-Path $env:WINDIR "System32\$DllName"
if(Test-Path $sysdll){ _ok ("DLL Exists (System32) => OK | {0}" -f $sysdll) } else { _ng ("DLL Exists (System32) => NG | {0}" -f $sysdll) }
try{
  $sig = Get-AuthenticodeSignature -FilePath $sysdll -ErrorAction Stop
  _ok ("DLL Signature => OK | Status={0}; Signer={1}" -f $sig.Status, ($sig.SignerCertificate.Subject -as [string]))
}catch{ _wr ("DLL Signature => 读取失败：{0}" -f $_.Exception.Message) }

# === 1) 设备功能 devnode（MEDIA）下的 FX\0 / EP\0 / Legacy FX ===
$devRoot = "HKLM:\SYSTEM\CurrentControlSet\Enum\$UsbInstance"

# ContainerId（对齐端点）
try{
  $cid = (Get-ItemProperty $devRoot -ErrorAction Stop).ContainerID
  if($cid){ _ok ("Device ContainerId => {0}" -f $cid) } else { _wr "读取 ContainerId 失败。" }
}catch{ _wr ("读取 ContainerId 异常：{0}" -f $_.Exception.Message) }

# FX\0（EFX ,7 + Composite ,15 + PM7）
$fx0 = Join-Path $devRoot 'Device Parameters\FX\0'
if(Test-Path $fx0){
  $p = Get-ItemProperty $fx0 -ErrorAction SilentlyContinue
  $v7  = $p.'{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},7'
  $v15 = ($p.'{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},15' -join ';')
  $pm7 = ($p.'{D3993A3F-99C2-4402-B5EC-A92A0367664B},7' -join ';')
  if( ($v7 -like "*$ApoClsid*") -or ($v15 -like "*$ApoClsid*") ){
    _ok ("Device FX\0 => 7='{0}'  15='{1}'  PM7='{2}'" -f $v7,$v15,$pm7)
  } else {
    _ng ("Device FX\0 => 未见你的 CLSID | 7='{0}' 15='{1}' PM7='{2}'" -f $v7,$v15,$pm7)
  }
} else {
  _ng "Device FX\0 => 不存在（AEB 不会镜像）。"
}

# EP\0（禁用增强=0）
$ep0 = Join-Path $devRoot 'Device Parameters\EP\0'
if(Test-Path $ep0){
  $enh = Read-Dword $ep0 '{1DA5D803-D492-4EDD-8C23-E0C0FFEE7F0E},5'
  if($enh -eq 0){ _ok "EP\0 => DisableEnhancements=0（系统效果启用）" }
  else { _wr ("EP\0 => DisableEnhancements={0}（若为1，系统效果管线会被关）" -f $enh) }
}else{
  _wr "EP\0 => 未设置（非致命，仅用于确保系统效果没被策略关闭）。"
}

# Legacy FX（可选参考：Dll/EFX/Order）
$fxLegacy = Join-Path $devRoot ("Device Parameters\FX\$ApoClsid")
if(Test-Path $fxLegacy){
  $dll   = (Get-ItemProperty $fxLegacy -ErrorAction SilentlyContinue).Dll
  $efx   = (Get-ItemProperty $fxLegacy -ErrorAction SilentlyContinue).EFX
  $order = (Get-ItemProperty $fxLegacy -ErrorAction SilentlyContinue).Order
  _info ("Legacy FX => Dll='{0}' EFX={1} Order={2}" -f $dll,$efx,$order)
}

# === 2) 端点 FxProperties（仅读你给的两个 GUID，不做任何搜索） ===
function Check-EndpointFx([string]$kind,[string]$guid,[string]$expectCid){
  $base = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$kind\$guid"
  $props= Join-Path $base 'Properties'
  $fx   = Join-Path $base 'FxProperties'
  if(-not (Test-Path $base)){ _ng ("Endpoint({0}) => 不存在：{1}" -f $kind,$guid); return }

  $fn = (Get-ItemProperty $base -ErrorAction SilentlyContinue).'FriendlyName'
  $cid= (Get-ItemProperty $props -ErrorAction SilentlyContinue).'{b3f8fa53-0004-438e-9003-51a46e139bfc},2'  # AssociatedContainerId
  $iid= (Get-ItemProperty $props -ErrorAction SilentlyContinue).'{78C34FC8-104A-4D11-9F5B-700F2848BCA5},256' # DeviceInstanceId

  if($expectCid -and $cid -and ($cid -ne $expectCid)){ _wr ("Endpoint({0}) => ContainerId 不同：Dev={1} | EP={2}" -f $kind,$expectCid,$cid) }
  if(-not (Test-Path $fx)){ _ng ("Endpoint({0}) => 无 FxProperties（尚未镜像？） | {1}" -f $kind,$fn); return }

  $v7  = (Get-ItemProperty $fx -ErrorAction SilentlyContinue).'{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},7'
  $v15 = (Get-ItemProperty $fx -ErrorAction SilentlyContinue).'{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},15'
  $v15s= ($v15 -join ';')
  $hit = ($v7 -like "*$ApoClsid*") -or ($v15s -like "*$ApoClsid*")

  if($hit){ _ok ("Endpoint({0}) => 已镜像 CLSID | Name='{1}' | 7='{2}' | 15='{3}'" -f $kind,$fn,$v7,$v15s) }
  else    { _ng ("Endpoint({0}) => 未见 CLSID | Name='{1}' | 7='{2}' | 15='{3}'" -f $kind,$fn,$v7,$v15s) }

  if($iid -and ($iid -notlike 'USB\VID_0A67&PID_30A2*')){
    _wr ("Endpoint({0}) => 绑定实例ID与设备不符：{1}" -f $kind,$iid)
  }
}

Check-EndpointFx -kind 'Render'  -guid $RenderGuid  -expectCid $cid
Check-EndpointFx -kind 'Capture' -guid $CaptureGuid -expectCid $cid

# === 3) DriverStore/策略位 ===
try{
  $ps = (& pnputil /enum-drivers) 2>$null
  $foundInf = ($ps -match [regex]::Escape($InfName)) -ne $null
  if($foundInf){ _ok ("DriverStore => found '{0}' = True" -f $InfName) } else { _wr ("DriverStore => 未在仓库找到 {0}" -f $InfName) }
}catch{ _wr ("pnputil 查询失败：{0}" -f $_.Exception.Message) }

$pol1 = Read-Dword "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Audio" "DisableLegacyAudioEffects"
$pol2 = Read-Dword "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Audio" "DisableSystemEffects"
$pol3 = Read-Dword "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Audio" "EnableCompositeFx"
_info ("Policy: DisableLegacyAudioEffects={0}; DisableSystemEffects={1}; EnableCompositeFx={2}" -f $pol1,$pol2,$pol3)

# === 4) 探针播放 + audiodg 模块双重检测 ===
try{
  $v = New-Object -ComObject SAPI.SpVoice
  $null = $v.Speak('APO probe test audio.')
  Start-Sleep -Milliseconds 600
}catch{ _wr ("SAPI 播放探针失败：{0}" -f $_.Exception.Message) }

$loaded = $false
try{
  $mods = (Get-Process -Name audiodg -ErrorAction SilentlyContinue | ForEach-Object { $_.Modules })
  if($mods){ $loaded = ($mods | Where-Object { $_.ModuleName -ieq $DllName -or $_.FileName -like ("*\" + $DllName) }) -ne $null }
}catch{ }
if(-not $loaded){
  $out = (tasklist /m $DllName) 2>$null
  if($out){ $loaded = ($out -match 'audiodg\.exe') -ne $null }
}
if($loaded){ _ok 'audiodg Loaded APO DLL => True' } else { _ng 'audiodg Loaded APO DLL => False（若端点未镜像，请先重建端点；确认使用非 RAW 端点播放）' }

_ok 'DeepVerify finished.'

