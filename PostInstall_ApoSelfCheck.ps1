<# 
PostInstall_ApoSelfCheck.ps1
用途：启动后/安装后，对 APO 全链路做自检与追因。仅读取，不修改系统。
建议：以管理员运行。若注册为“登录时”自检任务，可附带 -Quiet。
#>

param(
  [string]$Instance           = 'USB\VID_0A67&PID_30A2&MI_00\7&3b1ff4ef&0&0000',   # 你的 USB 设备实例
  [ValidateSet('Render','Capture')] [string]$Kind = 'Render',
  [string]$EndpointGuid       = '{4ffd4a6d-2616-44a3-bc17-5e00afefb449}',          # 已知端点（可留空自动匹配）
  [string]$ExpectedClsid      = '{8E3E0B71-5B8A-45C9-9B3D-3A2E5B418A10}',          # 你的 EFX CLSID
  [string]$DllName            = 'MyCompanyEfxApo.dll',                              # 你的 DLL 名
  [string]$InfName            = 'MyCompanyUsbApoExt.inf',                           # 便于检查 pnputil 输出
  [switch]$Quiet,                                                                     # 控制台少量输出
  [switch]$PlaybackTest      = $true,                                               # 触发 1 段播放以检测 audiodg 模块
  [switch]$RegisterStartup,                                                          # 注册“登录时”自检计划任务
  [string]$LogDir            = "$env:ProgramData\MyCompany\ApoSelfTest"             # 日志目录
)

# ========= 基础工具 =========
$ErrorActionPreference = 'Stop'
$PSStyle.OutputRendering = 'Host'
function New-Dir([string]$p){ if(-not (Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } }
New-Dir $LogDir
$stamp   = (Get-Date).ToString('yyyyMMdd_HHmmss')
$LogPath = Join-Path $LogDir "ApoSelfTest_$stamp.txt"

$S_OK    = 0; $S_WARN = 1; $S_FAIL = 2
$Results = [System.Collections.Generic.List[object]]::new()

function Write-Log([string]$msg, [int]$sev = $S_OK){
  $tag = @('PASS','WARN','FAIL')[$sev]
  $prefix = "[$tag] $msg"
  $prefix | Out-File -FilePath $LogPath -Encoding utf8 -Append
  if(-not $Quiet){
    switch($sev){
      0 { Write-Host $prefix -ForegroundColor Green }
      1 { Write-Host $prefix -ForegroundColor Yellow }
      2 { Write-Host $prefix -ForegroundColor Red }
      default { Write-Host $prefix }
    }
  }
}
function Add-Result([string]$name, [bool]$ok, [string]$detail){
  $sev = $ok ? $S_OK : $S_FAIL
  Write-Log "$name => $(if($ok){'OK'}else{'NG'}) | $detail" $sev
  $Results.Add([pscustomobject]@{ Check=$name; Pass=$ok; Detail=$detail })
}
function Is-Admin {
  $id=[Security.Principal.WindowsIdentity]::GetCurrent()
  (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function RegQueryValue([string]$key,[string]$name){
  $o = cmd /c "reg query `"$key`" /v `"$name`"" 2>&1
  if($LASTEXITCODE -ne 0){ return @{Found=$false;Type=$null;Data=$null} }
  foreach($line in ($o -split "`r?`n")){
    # 匹配：<空格><Name><空格+><Type><空格+><Data>
    if($line -match "^\s*$([Regex]::Escape($name))\s+([A-Z_0-9]+)\s+(.*)$"){
      return @{Found=$true; Type=$matches[2]; Data=$matches[3] }
    }
  }
  return @{Found=$false;Type=$null;Data=$null}
}

function RegListSubkeys([string]$key){
  $out = cmd /c "reg query `"$key`"" 2>&1
  if($LASTEXITCODE -ne 0){ return @() }
  ($out -split "`r?`n") | Where-Object { $_ -match "^[Hh][Kk][Ll][Mm]\\" }
}
function Get-DeviceContainerId([string]$inst){
  try{ (Get-ItemProperty ("HKLM:\SYSTEM\CurrentControlSet\Enum\$inst")).ContainerID }catch{ $null }
}

# ========= 常量（PKEY） =========
$K_EFX7      = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},7'     # EndpointEffectClsid（关键）
$K_EFX15     = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},15'    # Composite Endpoint EFX（1803+）
$K_DEF3      = '{FC1CFC9B-31F9-4C56-9D2C-39A781AB0B2E},3'     # 历史 Default EFX（兼容性）
$K_SYSFX_OFF = '{1DA5D803-D492-4EDD-8C23-E0C0FFEE7F0E},5'     # 禁用增强

# ========= Step 0 环境信息 =========
Write-Log "=== APO 开机自检开始：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
Write-Log "PSVersion=$($PSVersionTable.PSVersion); Arch=$([Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture); Admin=$(Is-Admin)"
Write-Log "Params: Instance='$Instance'; Kind=$Kind; EndpointGuid='$EndpointGuid'; CLSID='$ExpectedClsid'; DllName='$DllName'; InfName='$InfName'"
if(-not (Is-Admin)){ Write-Log "建议以管理员运行，某些查询需要管理员权限。" $S_WARN }

# ========= Step 1 服务状态 =========
try{
  $svc1 = Get-Service -Name audiosrv -ErrorAction Stop
  $svc2 = Get-Service -Name audioendpointbuilder -ErrorAction Stop
  Add-Result 'AudioSrv Running' ($svc1.Status -eq 'Running') "Status=$($svc1.Status)"
  Add-Result 'AudioEndpointBuilder Running' ($svc2.Status -eq 'Running') "Status=$($svc2.Status)"
}catch{
  Add-Result 'Audio Services Present' $false "无法查询音频服务：$($_.Exception.Message)"
}

# ========= Step 2 COM/CLSID 注册（改为 .NET 读取，避免本地化/换行问题）=========
try{
  Add-Type -AssemblyName 'Microsoft.Win32.Registry'
  $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
               [Microsoft.Win32.RegistryHive]::LocalMachine,
               [Microsoft.Win32.RegistryView]::Registry64)
  $rk  = $baseKey.OpenSubKey("SOFTWARE\Classes\CLSID\$ExpectedClsid\InprocServer32")
  $dll = $rk?.GetValue($null)          # (默认)
  $tm  = $rk?.GetValue('ThreadingModel')
  $expDll = "C:\Windows\System32\$DllName"
  $okCom  = ($dll -ieq $expDll -and $tm -match 'Both')
  Add-Result 'COM x64 InprocServer32' $okCom "Dll='$dll'; TM='$tm' (期望='$expDll', Both)"
}catch{
  Add-Result 'COM x64 InprocServer32' $false "读取失败：$($_.Exception.Message)"
}



# ========= Step 3 DLL 签名/存在 =========
try{
  $expDll = "C:\Windows\System32\$DllName"
  $fi = Get-Item -Path $expDll -ErrorAction Stop
  Add-Result 'DLL Exists (System32)' $true $fi.FullName
  $sig = Get-AuthenticodeSignature -FilePath $fi.FullName
  Add-Result 'DLL Signature' ($sig.Status -eq 'Valid' -or $sig.Status -eq 'NotSigned') "Status=$($sig.Status); Signer=$($sig.SignerCertificate.Subject)"
}catch{
  Add-Result 'DLL Exists (System32)' $false "缺失：$expDll"
}

# ========= Step 4 设备 ContainerId & 参与的 MEDIA 功能实例 =========
$cid = Get-DeviceContainerId $Instance
$cidDisp = $cid
Add-Result 'Device ContainerId' ($cid -ne $null) "$cidDisp"
$mediaFuncs = @()
try{
  $mediaFuncs = Get-PnpDevice -Class MEDIA -PresentOnly -ErrorAction Stop | Where-Object {
    (Get-PnpDeviceProperty -InstanceId $_.InstanceId -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction SilentlyContinue).Data -eq $cid
  }
  Add-Result 'MEDIA Functions in Same Container' ($mediaFuncs.Count -gt 0) "Count=$($mediaFuncs.Count)"
}catch{
  Add-Result 'MEDIA Functions in Same Container' $false "查询失败：$($_.Exception.Message)"
}

# ========= Step 5 检查每个 MEDIA 实例的 FX\0（是否有 7/15 指向 CLSID）=========
$fxHit = $false
foreach($m in $mediaFuncs){
  $fxKey = "HKLM\SYSTEM\CurrentControlSet\Enum\$($m.InstanceId)\Device Parameters\FX\0"
  $v7  = RegQueryValue $fxKey $K_EFX7
  $v15 = RegQueryValue $fxKey $K_EFX15
  $v3  = RegQueryValue $fxKey $K_DEF3
  $msg = "FX0('$($m.InstanceId)'): 7='$($v7.Data)' 15='$($v15.Data)' 3='$($v3.Data)'"
  $ok  = ($v15.Found -and $v15.Data -ieq $ExpectedClsid) -or ($v7.Found -and $v7.Data -ieq $ExpectedClsid) -or ($v3.Found -and $v3.Data -ieq $ExpectedClsid)
  if($ok){ $fxHit = $true }
  Add-Result 'Device FX\0 → EFX present' $ok $msg
}
if(-not $mediaFuncs -or -not $fxHit){
  Write-Log "提示：若此处 FAIL，多半是 INF 写到了非参与端点生成的接口（常见：只写 MI_00）。" $S_WARN
}

# ========= Step 6 端点定位（SWD\MMDEVAPI + 可选 EndpointGuid）=========
$epPick     = $null
$mmRoot     = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Kind"
$allEps     = @()
if(Test-Path $mmRoot){ $allEps = Get-ChildItem $mmRoot | Select-Object -ExpandProperty PSChildName }
function Get-PnpEp([string]$g){
  $gid = $g.Trim('{}').ToUpper()
  $eid = "SWD\MMDEVAPI\{0.0.0.00000000}.{$gid}"  # 注意这里有花括号
  try{
    $dev = Get-PnpDevice -InstanceId $eid -PresentOnly -ErrorAction Stop
    $cid2 = (Get-PnpDeviceProperty -InstanceId $eid -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction SilentlyContinue).Data
    [pscustomobject]@{ InstanceId=$dev.InstanceId; ContainerId=$cid2; Status=$dev.Status }
  }catch{ $null }
}

if($EndpointGuid){
  $epPick = $EndpointGuid
}else{
  # 无指定：优先默认端点（Role:0/1/2）且 ContainerId 匹配
  $cands = foreach($g in $allEps){
    $rk = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Kind\$g"
    $p  = Get-ItemProperty $rk -ErrorAction SilentlyContinue
    $pp = Get-ItemProperty (Join-Path $rk 'Properties') -ErrorAction SilentlyContinue
    [pscustomobject]@{
      Guid = $g
      Default = ($p.'Role:0' -eq 1 -or $p.'Role:1' -eq 1 -or $p.'Role:2' -eq 1)
      PKEY_Device_ContainerId = $pp.'{A45C254E-DF1C-4EFD-8020-67D146A850E0},10'
    }
  }
  $epPick = ($cands | Where-Object { $_.Default -and $_.PKEY_Device_ContainerId -ieq $cid } | Select-Object -First 1).Guid
  if(-not $epPick){ $epPick = ($cands | Where-Object { $_.PKEY_Device_ContainerId -ieq $cid } | Select-Object -First 1).Guid }
}
$pnpe = if($epPick){ Get-PnpEp $epPick } else { $null }
Add-Result 'Endpoint Resolved' ($null -ne $pnpe) "Guid='$epPick'; PnP.Status='$($pnpe?.Status)'; CID='$($pnpe?.ContainerId)'"

# ========= Step 7 端点 FxProperties/Properties 检查 =========
if($pnpe -ne $null){
  $ekFx = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Kind\$epPick\FxProperties"
  $ekPr = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Kind\$epPick\Properties"
  $e7   = RegQueryValue $ekFx $K_EFX7
  $e15  = RegQueryValue $ekFx $K_EFX15
  $e3   = RegQueryValue $ekFx $K_DEF3
  $dEnh = RegQueryValue $ekPr $K_SYSFX_OFF
  $okFx = ($e15.Found -and $e15.Data -ieq $ExpectedClsid) -or ($e7.Found -and $e7.Data -ieq $ExpectedClsid) -or ($e3.Found -and $e3.Data -ieq $ExpectedClsid)
  Add-Result 'Endpoint FxProperties → EFX present' $okFx ("7='$($e7.Data)' 15='$($e15.Data)' 3='$($e3.Data)'")
  $okEn = ($dEnh.Found -and [int]("0x"+($dEnh.Data -replace '^0x','')) -eq 0) -or ($dEnh.Data -eq '0')
  Add-Result 'Endpoint DisableEnhancements=0' $okEn "Value='$($dEnh.Data)'"
  # 子键模式检查（RAW/Media/Comm 等）
  $subs = RegListSubkeys $ekFx
  $miss = @()
  foreach($sk in $subs){
    $s7  = RegQueryValue $sk $K_EFX7
    $s15 = RegQueryValue $sk $K_EFX15
    if(-not (($s15.Found -and $s15.Data -ieq $ExpectedClsid) -or ($s7.Found -and $s7.Data -ieq $ExpectedClsid))){
      $miss += $sk
    }
  }
  Add-Result 'Endpoint ProcessingMode Subkeys OK' ($miss.Count -eq 0) ("UncheckedOrMissing=" + ($miss.Count))
}else{
  Write-Log "提示：端点未解析；若你已知 GUID，请用 -EndpointGuid '{...}' 指定。" $S_WARN
}

# ========= Step 8 驱动包存在性（pnputil 枚举）=========
try{
  $pnpo = cmd /c "pnputil /enum-drivers" 2>&1
  $hasInf = (($pnpo -join "`n") -match [regex]::Escape($InfName))
  Add-Result 'Extension INF Installed' ([bool]$hasInf) ("pnputil found '$InfName' = " + [bool]$hasInf)
}catch{
  Add-Result 'Extension INF Installed' $false "pnputil 不可用或失败：$($_.Exception.Message)"
}


# ========= Step 9 播放触发 & audiodg 模块加载 =========
$loaded = $false
if($PlaybackTest){
  try{
    $v = New-Object -ComObject SAPI.SpVoice
    $null = $v.Speak('APO self test playback.')
    Start-Sleep 1
    $loaded = ((tasklist /m $DllName) -match 'audiodg.exe')
    Add-Result 'audiodg Loaded APO DLL' $loaded "tasklist /m $DllName -> $loaded"
  }catch{
    Add-Result 'Playback Trigger' $false "SAPI 播放失败：$($_.Exception.Message)"
  }
}else{
  Write-Log "已跳过 PlaybackTest（未触发模块装载检测）。" $S_WARN
}

# ========= Step 10 最近音频日志（可选）=========
try{
  wevtutil sl Microsoft-Windows-Audio/Operational /e:true | Out-Null
  $events = Get-WinEvent -FilterHashtable @{ LogName='Microsoft-Windows-Audio/Operational'; StartTime=(Get-Date).AddMinutes(-5) } -ErrorAction SilentlyContinue
  $hit = $events | Where-Object { $_.Message -match 'APO|FX|EFX|RAW|Activate|CLSID' } | Select-Object -First 8
  if($hit){
    "---- Recent Audio Operational events (<=8) ----" | Out-File -FilePath $LogPath -Append -Encoding utf8
    $hit | ForEach-Object {
      ("[{0}] {1} {2}`n{3}`n" -f $_.TimeCreated, $_.LevelDisplayName, $_.Id, $_.Message) |
        Out-File -FilePath $LogPath -Append -Encoding utf8
    }
    Add-Result 'Audio Operational Log Available' $true "Found=$($hit.Count)"
  }else{
    Add-Result 'Audio Operational Log Available' $false "最近 5 分钟未捕获相关事件（正常也可能无事件）"
  }
}catch{
  Add-Result 'Audio Operational Log Available' $false "查询失败：$($_.Exception.Message)"
}

# ========= 汇总 =========
$failed = $Results | Where-Object { -not $_.Pass }
$overall = ($failed.Count -eq 0)
Write-Log "=== 总体结果：$(if($overall){'PASS'}else{'FAIL'}) ==="
"摘要：
$( $Results | ForEach-Object { "- [$($_.Check)]: " + ($(if($_.Pass){'PASS'}else{'FAIL'})) + " | " + $_.Detail } | Out-String )
日志文件: $LogPath
" | Out-File -FilePath $LogPath -Append -Encoding utf8

if(-not $Quiet){
  Write-Host "`n==== 总结（关键环节）====" -ForegroundColor Cyan
  $Results | Format-Table -AutoSize
  Write-Host "`n日志文件: $LogPath" -ForegroundColor Gray
}

# ========= 可选：注册“登录时”自检计划任务（最高权限）=========
if($RegisterStartup){
  try{
    $scriptPath = $MyInvocation.MyCommand.Path
    $taskName = 'ApoSelfCheck_AtLogon'
    $tr = "powershell -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -Instance `"$Instance`" -Kind $Kind -EndpointGuid `"$EndpointGuid`" -ExpectedClsid `"$ExpectedClsid`" -DllName `"$DllName`" -InfName `"$InfName`" -Quiet"
    schtasks /create /tn $taskName /ru "$env:USERNAME" /RL HIGHEST /sc ONLOGON /tr "$tr" /F | Out-Null
    Write-Log "已注册计划任务 '$taskName'（登录触发，自检日志输出到 $LogDir）。"
  }catch{
    Write-Log "注册计划任务失败：$($_.Exception.Message)" $S_FAIL
  }
}
