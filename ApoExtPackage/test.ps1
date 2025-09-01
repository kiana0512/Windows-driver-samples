<# 
  PostInstall_ApoDeepVerify_NoParams.ps1
  固定锚点校验（无参数 / 不枚举 / 不依赖端点 GUID）

  做什么：
    1) 检查服务、COM、DLL、签名
    2) 仅在指定功能 devnode 下读取 FX\0：EFX ,7（单）、,15（Composite）、PM7（SupportedModes）
    3) 播放一次探针后，严格检查 audiodg.exe 是否加载 DLL（先 Get-Process -Module，再回退 tasklist /m）
    4) 可选：一键重建设备容器（默认注释）

  你可以按需只改 5 个常量：$UsbVid / $UsbPid / $UsbMi / $UsbInstance / $ApoClsid / $DllName / $InfName
#>

# ===== 固定常量（仅使用稳定锚点；避免使用 $PID 这个保留变量） =====
$UsbVid      = "0A67"
$UsbPid      = "30A2"
$UsbMi       = "MI_00"
$UsbInstance = "USB\VID_0A67&PID_30A2&MI_00\7&3B1FF4EF&0&0000"   # 你的功能 devnode（MEDIA）
$ApoClsid    = "{8E3E0B71-5B8A-45C9-9B3D-3A2E5B418A10}"
$DllName     = "MyCompanyEfxApo.dll"
$InfName     = "MyCompanyUsbApoExt.inf"

$ErrorActionPreference = "Stop"
function _ok([string]$m){ Write-Host "[PASS] $m" -ForegroundColor Green }
function _ng([string]$m){ Write-Host "[FAIL] $m" -ForegroundColor Red }
function _wr([string]$m){ Write-Host "[WARN] $m" -ForegroundColor Yellow }
function _info([string]$m){ Write-Host "[INFO] $m" -ForegroundColor Cyan }

# === 0) 环境 ===
$arch  = if([Environment]::Is64BitProcess){"X64"}else{"X86"}
$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
_ok ("Apo DeepVerify start: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
_ok ("PSVersion={0}; Arch={1}; Admin={2}" -f $PSVersionTable.PSVersion, $arch, $admin)
if(-not $admin){ _wr "建议以管理员身份运行。" }

# === 1) 基础服务 ===
try{
  $as  = Get-Service -Name "audiosrv" -ErrorAction SilentlyContinue
  $aeb = Get-Service -Name "AudioEndpointBuilder" -ErrorAction SilentlyContinue
  if($as -and $as.Status -eq "Running"){ _ok ("AudioSrv Running => OK | Status={0}" -f $as.Status) } else { _ng ("AudioSrv Running => NG | Status={0}" -f $as.Status) }
  if($aeb -and $aeb.Status -eq "Running"){ _ok ("AudioEndpointBuilder Running => OK | Status={0}" -f $aeb.Status) } else { _ng ("AudioEndpointBuilder Running => NG | Status={0}" -f $aeb.Status) }
}catch{ _wr ("读取服务状态失败：{0}" -f $_.Exception.Message) }

# === 2) COM & DLL ===
try{
  $cls = "HKLM:\SOFTWARE\Classes\CLSID\$ApoClsid\InprocServer32"
  $inproc = (Get-ItemProperty $cls -ErrorAction Stop)."(default)"
  if($inproc -and (Test-Path $inproc)){
    _ok ("COM x64 InprocServer32 => OK | {0}" -f $inproc)
  } else {
    _ng "COM x64 InprocServer32 => NG | 值缺失或 DLL 不存在"
  }
}catch{ _ng ("COM x64 InprocServer32 => NG | {0}" -f $_.Exception.Message) }

$sysdll = Join-Path $env:WINDIR "System32\$DllName"
if(Test-Path $sysdll){ _ok ("DLL Exists (System32) => OK | {0}" -f $sysdll) } else { _ng ("DLL Exists (System32) => NG | {0}" -f $sysdll) }
try{
  $sig = Get-AuthenticodeSignature -FilePath $sysdll -ErrorAction Stop
  _ok ("DLL Signature => OK | Status={0}; Signer={1}" -f $sig.Status, ($sig.SignerCertificate.Subject -as [string]))
}catch{ _wr ("DLL Signature => 读取失败：{0}" -f $_.Exception.Message) }

# === 3) 设备侧 FX\0（只读你给的功能 devnode；不枚举、不查 MMDevices） ===
function Read-FX0([string]$inst){
  $k = "HKLM:\SYSTEM\CurrentControlSet\Enum\$inst\Device Parameters\FX\0"
  if(-not (Test-Path $k)){ return $null }
  $p = Get-ItemProperty $k -ErrorAction SilentlyContinue
  [pscustomobject]@{
    K7  = $p.'{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},7'
    K15 = ($p.'{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},15' -join ';')
    PM7 = ($p.'{D3993A3F-99C2-4402-B5EC-A92A0367664B},7' -join ';')
  }
}
$fx = Read-FX0 $UsbInstance
if($fx){
  $has7  = -not [string]::IsNullOrWhiteSpace($fx.K7)
  $has15 = -not [string]::IsNullOrWhiteSpace($fx.K15)
  if($has7 -or $has15){
    _ok ("Device FX\0 => 7='{0}'  15='{1}'  PM7='{2}'" -f $fx.K7, $fx.K15, $fx.PM7)
  } else {
    _ng ("Device FX\0 => 缺少 EFX（7/15 均空），AEB 不会镜像。")
  }
} else {
  _ng ("Device FX\0 => 未找到（请核对 InstanceId 是否正确）")
}

# === 4) SetupAPI/DriverStore（可快速确认 INF 是否在仓库） ===
try{
  $ps = (& pnputil /enum-drivers) 2>$null
  $foundInf = ($ps -match [regex]::Escape($InfName)) -ne $null
  if($foundInf){ _ok ("DriverStore => found '{0}' = True" -f $InfName) } else { _wr ("DriverStore => 未在仓库找到 {0}" -f $InfName) }
}catch{ _wr ("pnputil 查询失败：{0}" -f $_.Exception.Message) }

# === 5) 播放探针并准确检查 audiodg 是否加载 DLL ===
try{
  $v = New-Object -ComObject SAPI.SpVoice
  $null = $v.Speak("APO probe test audio.")
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
if($loaded){ _ok "audiodg Loaded APO DLL => True" } else { _ng "audiodg Loaded APO DLL => False（若 FX\0 有值但仍为 False，请先重建端点或换非 RAW 端点播放）" }

# === 6) 可选：一键重建设备容器（默认注释；不依赖端点 GUID） ===
<#
_info "开始重建（仅按 VID/PID 前缀）：停止音频服务 → rescan → 重启容器"
try{ & net stop audiosrv /y | Out-Null; & net stop audioendpointbuilder | Out-Null }catch{}
try{
  & devcon rescan | Out-Null
  & net start audioendpointbuilder | Out-Null
  & net start audiosrv | Out-Null
  & devcon restart ("USB\VID_{0}&PID_{1}*" -f $UsbVid,$UsbPid) | Out-Null
  _ok "容器重启完成。重新播放后再看 audiodg 模块。"
}catch{ _wr ("重建出现异常：{0}" -f $_.Exception.Message) }
#>

_ok "DeepVerify finished."
