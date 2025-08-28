<#  ApoEndToEnd_v5_2.ps1
    - 自检 →（可选）端点重建 →（可选）ETW 采集 →（可选，仅调试）SYSTEM 镜像
#>

param(
  [Parameter(Mandatory=$true)][string]$Instance,
  [ValidateSet('Render','Capture')][string]$Kind='Render',
  [Parameter(Mandatory=$true)][string]$ExpectedClsid,
  [Parameter(Mandatory=$true)][string]$DllName,
  [string]$InfName='MyCompanyUsbApoExt.inf',
  [switch]$RebuildEndpoints,
  [switch]$CaptureEtw,
  [switch]$ForceMirror,
  [switch]$NoPlayback,
  [switch]$Quiet,
  [string]$LogDir="$env:ProgramData\MyCompany\ApoSelfTest"
)

$ErrorActionPreference='Stop'
function New-Dir([string]$p){ if(-not (Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } }
New-Dir $LogDir
$stamp=(Get-Date).ToString('yyyyMMdd_HHmmss')
$Log    = Join-Path $LogDir "ApoE2E_$stamp.txt"
$EtwEtl = Join-Path $LogDir "ApoE2E_$stamp.etl"
$TmpDir = "$env:SystemRoot\Temp"  # 短路径，避免 /TR 超长
New-Dir $TmpDir

$S_OK=0;$S_WARN=1;$S_FAIL=2
function WL([string]$msg,[int]$sev=$S_OK){
  $tag=@('PASS','WARN','FAIL')[$sev]
  $line="[$tag] $msg"
  $line | Out-File -FilePath $Log -Append -Encoding utf8
  if(-not $Quiet){
    $c = if($sev -eq $S_OK){'Green'}elseif($sev -eq $S_WARN){'Yellow'}else{'Red'}
    Write-Host $line -ForegroundColor $c
  }
}
function Is-Admin { $id=[Security.Principal.WindowsIdentity]::GetCurrent(); (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) }

# PKEY 名
$K_EFX7      = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},7'
$K_EFX15     = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},15'
$K_DEF3      = '{FC1CFC9B-31F9-4C56-9D2C-39A781AB0B2E},3'
$K_SYSFX_OFF = '{1DA5D803-D492-4EDD-8C23-E0C0FFEE7F0E},5'

$MMRoot = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Kind"
$ExpDll = "C:\Windows\System32\$DllName"

# Registry（.NET）
Add-Type -AssemblyName 'Microsoft.Win32.Registry' | Out-Null
function Open-BaseKey([string]$h,[string]$v){
  [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::$h,[Microsoft.Win32.RegistryView]::$v)
}
function Get-RegValue([string]$h,[string]$v,[string]$sub,[string]$name){
  try{
    $rk=(Open-BaseKey $h $v).OpenSubKey($sub)
    if(-not $rk){ return @{Found=$false;Data=$null;Kind=$null} }
    $val=$rk.GetValue($name,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    if($null -eq $val){ return @{Found=$false;Data=$null;Kind=$null} }
    $kind=$rk.GetValueKind($name)
    $data = switch($kind){ 'MultiString'{($val -join ';')} 'Binary'{($val|%{ $_.ToString('X2') }) -join ''} default{[string]$val} }
    @{Found=$true;Data=$data;Kind=$kind}
  }catch{ @{Found=$false;Data=$null;Kind=$null} }
}

# ETW
$EtwName="ApoE2E_Trace"
function Start-AudioEtw {
  if(-not $CaptureEtw){ return }
  try{
    $providers=(logman query providers) 2>$null
    $want=@('Microsoft-Windows-Audio','Microsoft-Windows-Audio-Client','Microsoft-Windows-Audio-Device Graph','Microsoft-Windows-Audio-UAP')
    $have=$want | Where-Object { $providers -match [regex]::Escape($_) }
    if(-not $have){ WL "未发现音频 ETW Provider，跳过采集。" $S_WARN; return }
    $args=@('start',"$EtwName")
    foreach($p in $have){ $args+=@('-p',"`"$p`"",'0xFFFFFFFF','5') }
    $args+='-o',"$EtwEtl",'-ets'
    logman @args | Out-Null
    WL "ETW 开始：$($have -join ', ') → $EtwEtl"
  }catch{ WL "启动 ETW 失败：$($_.Exception.Message)" $S_WARN }
}
function Stop-AudioEtw {
  if(-not $CaptureEtw){ return }
  try{ logman stop "$EtwName" -ets | Out-Null; WL "ETW 停止：$EtwEtl" }catch{}
}

# 播放触发
function Trigger-Playback {
  if($NoPlayback){ WL '跳过播放触发（-NoPlayback）。' $S_WARN; return }
  try{ $v=New-Object -ComObject SAPI.SpVoice; $null=$v.Speak('APO end to end test audio.'); Start-Sleep 1 }catch{ WL "SAPI 播放失败：$($_.Exception.Message)" $S_WARN }
}

# 服务控制
function Stop-AudioServices { try{ cmd /c "net stop audiosrv" | Out-Null }catch{}; try{ cmd /c "net stop audioendpointbuilder" | Out-Null }catch{} }
function Start-AudioServices { try{ cmd /c "net start audioendpointbuilder" | Out-Null }catch{}; try{ cmd /c "net start audiosrv" | Out-Null }catch{} }

# 自检
function Run-SelfTest {
  WL "=== APO 一体化自检开始：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
  WL "PS=$($PSVersionTable.PSVersion); Arch=$([Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture); Admin=$(Is-Admin)"
  WL "Params: Instance='$Instance'; Kind=$Kind; CLSID='$ExpectedClsid'; DllName='$DllName'; INF='$InfName'"

  try{ $svc1=Get-Service audiosrv; WL "AudioSrv: $($svc1.Status)" }catch{ WL "AudioSrv 查询失败：$($_.Exception.Message)" $S_WARN }
  try{ $svc2=Get-Service audioendpointbuilder; WL "AudioEndpointBuilder: $($svc2.Status)" }catch{ WL "AEB 查询失败：$($_.Exception.Message)" $S_WARN }

  # COM
  $dll='';$tm=''
  try{
    $rk=(Open-BaseKey 'LocalMachine' 'Registry64').OpenSubKey("SOFTWARE\Classes\CLSID\$ExpectedClsid\InprocServer32")
    if($rk){ $dll=$rk.GetValue($null); $tm=$rk.GetValue('ThreadingModel') }
  }catch{}
  if(($dll -ieq $ExpDll) -and ($tm -match 'Both')){ WL "COM x64 InprocServer32 => OK | Dll='$dll'; TM='$tm'" }
  else{ WL "COM x64 InprocServer32 => NG | Dll='$dll'; TM='$tm'（期望'$ExpDll', Both）" $S_FAIL }

  # DLL
  if(Test-Path $ExpDll){ WL "DLL Exists (System32) => OK | $ExpDll"
    try{ $sig=Get-AuthenticodeSignature $ExpDll; WL "DLL Signature => $($sig.Status)" }catch{ WL "DLL Signature 检查失败：$($_.Exception.Message)" $S_WARN }
  }else{ WL "DLL Exists (System32) => NG | 缺失：$ExpDll" $S_FAIL }

  # ContainerId
  $cid=$null
  try{ $cid=(Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Enum\$Instance").ContainerID; WL "Device ContainerId => $cid" }
  catch{ WL "读取 ContainerId 失败：$($_.Exception.Message)" $S_FAIL; return }

  # MI_00 FX\0
  $fx0Sub="SYSTEM\CurrentControlSet\Enum\$Instance\Device Parameters\FX\0"
  $v7= Get-RegValue 'LocalMachine' 'Registry64' $fx0Sub $K_EFX7
  $v15=Get-RegValue 'LocalMachine' 'Registry64' $fx0Sub $K_EFX15
  $v3= Get-RegValue 'LocalMachine' 'Registry64' $fx0Sub $K_DEF3
  $okFx0 = (($v15.Found -and $v15.Data -ieq $ExpectedClsid) -or ($v7.Found -and $v7.Data -ieq $ExpectedClsid) -or ($v3.Found -and $v3.Data -ieq $ExpectedClsid))
  if($okFx0){ WL "MI_00 FX\\0 => OK | 7='$($v7.Data)' 15='$($v15.Data)' 3='$($v3.Data)'" }
  else{ WL "MI_00 FX\\0 => NG | 7='$($v7.Data)' 15='$($v15.Data)' 3='$($v3.Data)'" $S_FAIL }

  # MEDIA
  try{
    $media=Get-PnpDevice -Class MEDIA -PresentOnly | Where-Object {
      (Get-PnpDeviceProperty -InstanceId $_.InstanceId -KeyName 'DEVPKEY_Device_ContainerId' -EA SilentlyContinue).Data -eq $cid
    }
    WL "MEDIA Functions in Container => $($media.Count)"
    foreach($m in $media){ WL "  MEDIA: $($m.InstanceId)" }
  }catch{ WL "查询 MEDIA 失败：$($_.Exception.Message)" $S_WARN }

  # 端点（PnP → MMDevices）
  $endpoints=@()
  try{
    $allEp = Get-PnpDevice -Class AudioEndpoint -PresentOnly
    foreach($e in $allEp){
      $cid2=(Get-PnpDeviceProperty -InstanceId $e.InstanceId -KeyName 'DEVPKEY_Device_ContainerId' -EA SilentlyContinue).Data
      if(-not $cid2){ continue }
      if( ($cid2.ToString().ToLower() -eq $cid.ToString().ToLower()) ){
        $isRender = ($e.InstanceId -match '\{0\.0\.0\.00000000\}')
        if( ($Kind -eq 'Render' -and $isRender) -or ($Kind -eq 'Capture' -and -not $isRender) ){
          $endpoints += $e
        }
      }
    }
    WL "Endpoints in Container ($Kind) => $($endpoints.Count)"
  }catch{ WL "枚举端点失败：$($_.Exception.Message)" $S_FAIL }

  foreach($e in $endpoints){
    $guid = ($e.InstanceId -replace '.*\.\{','{').ToLower()
    $fxSub="SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Kind\$guid\FxProperties"
    $prSub="SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Kind\$guid\Properties"
    $e7 = Get-RegValue 'LocalMachine' 'Registry64' $fxSub $K_EFX7
    $e15= Get-RegValue 'LocalMachine' 'Registry64' $fxSub $K_EFX15
    $e3 = Get-RegValue 'LocalMachine' 'Registry64' $fxSub $K_DEF3
    $den= Get-RegValue 'LocalMachine' 'Registry64' $prSub $K_SYSFX_OFF
    $okFx = (($e15.Found -and $e15.Data -ieq $ExpectedClsid) -or ($e7.Found -and $e7.Data -ieq $ExpectedClsid) -or ($e3.Found -and $e3.Data -ieq $ExpectedClsid))
    if($okFx){ WL "Endpoint[$guid] FxProperties => OK | 7='$($e7.Data)' 15='$($e15.Data)' 3='$($e3.Data)'" }
    else{ WL "Endpoint[$guid] FxProperties => NG | 7='$($e7.Data)' 15='$($e15.Data)' 3='$($e3.Data)'" $S_FAIL }
    $okDen = ($den.Found -and ($den.Data -eq '0' -or $den.Data -eq '0x0'))
    if($okDen){ WL "Endpoint[$guid] DisableEnhancements=0 => OK | Value='$($den.Data)'" }
    else{ WL "Endpoint[$guid] DisableEnhancements=0 => NG | Value='$($den.Data)'" $S_WARN }
  }

  # 驱动包存在性
  try{
    $pnpo=cmd /c "pnputil /enum-drivers" 2>&1
    $hasInf=(($pnpo -join "`n") -match [regex]::Escape($InfName))
    $msg = if($hasInf){'Yes'}else{'No'}
    WL "Extension INF Installed => $msg"
  }catch{ WL "pnputil 查询失败：$($_.Exception.Message)" $S_WARN }

  # 播放 & audiodg 模块
  Trigger-Playback
  try{
    $loaded=((tasklist /m $DllName) -match 'audiodg.exe')
    $msg = if($loaded){'True'}else{'False'}
    WL "audiodg Loaded APO DLL => $msg"
  }catch{ WL "tasklist 检查失败：$($_.Exception.Message)" $S_WARN }
}

# 端点重建
function Rebuild-Endpoints {
  WL "开始端点重建（$Kind）……"
  $cid=(Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Enum\$Instance").ContainerID
  $targets = Get-PnpDevice -Class AudioEndpoint -PresentOnly | Where-Object {
    (Get-PnpDeviceProperty -InstanceId $_.InstanceId -KeyName 'DEVPKEY_Device_ContainerId' -EA SilentlyContinue).Data -eq $cid
  } | Where-Object {
    $isRender = ($_.InstanceId -match '\{0\.0\.0\.00000000\}')
    ($Kind -eq 'Render' -and $isRender) -or ($Kind -eq 'Capture' -and -not $isRender)
  }
  if(-not $targets){ WL "容器中找不到 $Kind 端点，取消重建。" $S_WARN; return }

  Stop-AudioServices
  $devcon = (Get-Command devcon -ErrorAction SilentlyContinue)

  if($devcon){
    foreach($t in $targets){
      WL "devcon remove @`"$($t.InstanceId)`""
      try{ devcon remove "@$($t.InstanceId)" | Out-Null }catch{ WL "devcon remove 失败：$($_.Exception.Message)" $S_WARN }
    }
    try{ devcon rescan | Out-Null }catch{}
  } else {
    # 用 pnputil /remove-device；若不可用，再走 SYSTEM 删除注册表键
    $pnputilOk=$false
    try{
      foreach($t in $targets){
        cmd /c "pnputil /remove-device `"$($t.InstanceId)`"" | Out-Null
      }
      $pnputilOk=$true
    }catch{}
    if(-not $pnputilOk){
      WL "pnputil 不可用，改用 SYSTEM 删除 MMDevices 端点键。"
      $keys=@()
      foreach($t in $targets){
        $guid = ($t.InstanceId -replace '.*\.\{','{').ToLower()
        $keys += "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Kind\$guid"
      }
      $payload = @'
param()
$Keys = @(__KEYS__)
$ErrorActionPreference='Stop'
foreach($k in $Keys){ if(Test-Path $k){ Remove-Item -Path $k -Recurse -Force } }
'@
      $keysLit = ($keys | ForEach-Object { '"{0}"' -f $_ }) -join ','
      $body = $payload.Replace('__KEYS__',$keysLit)
      $tmp = Join-Path $TmpDir "ApoSysDel_$stamp.ps1"
      $body | Out-File -FilePath $tmp -Encoding utf8
      schtasks /Create /TN ApoE2E_SysDel /SC ONCE /ST 23:59 /TR "powershell -NoProfile -ExecutionPolicy Bypass -File `"$tmp`"" /RU SYSTEM /RL HIGHEST | Out-Null
      schtasks /Run /TN ApoE2E_SysDel | Out-Null
      Start-Sleep 2
    }
  }

  Start-AudioServices
  WL "端点重建流程完成。"
}

# SYSTEM 镜像（调试）
function Force-Mirror-System {
  WL "开始 SYSTEM 镜像（在服务运行状态下直接写入端点；含 DEFAULT 处理模式）……" $S_WARN

  $K7  = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},7'
  $K15 = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},15'
  $K3  = '{FC1CFC9B-31F9-4C56-9D2C-39A781AB0B2E},3'
  $KDE = '{1DA5D803-D492-4EDD-8C23-E0C0FFEE7F0E},5'
  $MODE_DEFAULT = '{C18E2F7E-933D-4965-B7D1-1EEF228D2AF3}'

  $cid=(Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Enum\$Instance").ContainerID
  $eps = Get-PnpDevice -Class AudioEndpoint -PresentOnly | Where-Object {
    (Get-PnpDeviceProperty -InstanceId $_.InstanceId -KeyName 'DEVPKEY_Device_ContainerId' -EA SilentlyContinue).Data -eq $cid
  } | Where-Object {
    $isRender = ($_.InstanceId -match '\{0\.0\.0\.00000000\}')
    ($Kind -eq 'Render' -and $isRender) -or ($Kind -eq 'Capture' -and -not $isRender)
  }
  if(-not $eps){ WL "容器内无 $Kind 端点，取消镜像。" $S_FAIL; return }

  $ids = $eps | Select-Object -ExpandProperty InstanceId
  $idsLit = ($ids | ForEach-Object { '"{0}"' -f $_ }) -join ','

  $TmpDir = "$env:SystemRoot\Temp"
  if(-not (Test-Path $TmpDir)){ New-Item -ItemType Directory -Path $TmpDir -Force | Out-Null }
  $tn  = "ApoE2E_SysMirror_" + (Get-Date -Format 'HHmmss')   # 唯一任务名
  $ok  = Join-Path "$env:ProgramData\MyCompany\ApoSelfTest" ("SysMirror_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".ok")
  $etl = Join-Path "$env:ProgramData\MyCompany\ApoSelfTest" ("SysMirror_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")

  $payload = @'
param()
$ErrorActionPreference='Stop'
$Kind = "__KIND__"
$Clsid = "__CLSID__"
$EpIds = @(__EPIDS__)
$OkFile="__OKFILE__"
$LogFile="__LOGFILE__"

$K7  = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},7'
$K15 = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},15'
$K3  = '{FC1CFC9B-31F9-4C56-9D2C-39A781AB0B2E},3'
$KDE = '{1DA5D803-D492-4EDD-8C23-E0C0FFEE7F0E},5'
$MODE_DEFAULT = '{C18E2F7E-933D-4965-B7D1-1EEF228D2AF3}'

$done=@()
foreach($eid in $EpIds){
  $g = ($eid -replace '.*\.\{','{').ToLower()
  $root="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Kind\$g"
  if(-not (Test-Path $root)){ continue }
  $fx=Join-Path $root 'FxProperties'
  $pr=Join-Path $root 'Properties'
  if(-not (Test-Path $fx)){ New-Item -Path $fx -Force | Out-Null }
  if(-not (Test-Path $pr)){ New-Item -Path $pr -Force | Out-Null }
  New-ItemProperty -Path $fx -Name $K7  -Value $Clsid    -PropertyType String      -Force | Out-Null
  New-ItemProperty -Path $fx -Name $K15 -Value @($Clsid) -PropertyType MultiString -Force | Out-Null
  New-ItemProperty -Path $fx -Name $K3  -Value $Clsid    -PropertyType String      -Force | Out-Null
  New-ItemProperty -Path $pr -Name $KDE -Value 0          -PropertyType DWord       -Force | Out-Null

  $modeFx = Join-Path $root ("ProcessingModes\" + $MODE_DEFAULT + "\FxProperties")
  if(-not (Test-Path $modeFx)){ New-Item -Path $modeFx -Force | Out-Null }
  New-ItemProperty -Path $modeFx -Name $K7  -Value $Clsid    -PropertyType String      -Force | Out-Null
  New-ItemProperty -Path $modeFx -Name $K15 -Value @($Clsid) -PropertyType MultiString -Force | Out-Null
  New-ItemProperty -Path $modeFx -Name $K3  -Value $Clsid    -PropertyType String      -Force | Out-Null

  $done += $g
}
"[OK] mirrored: " + ($done -join ',') | Out-File -FilePath $OkFile -Encoding utf8 -Force
'@

  $sys = Join-Path $TmpDir ("ApoSysMirror_" + (Get-Date -Format 'HHmmss') + ".ps1")
  $body = $payload.Replace('__KIND__',$Kind).Replace('__CLSID__',$ExpectedClsid).Replace('__EPIDS__',$idsLit).Replace('__OKFILE__',$ok).Replace('__LOGFILE__',$etl)
  $body | Out-File -FilePath $sys -Encoding utf8

  # 关键：/F 强制覆盖，避免 (Y/N) 提示；用唯一任务名
  schtasks /Create /TN $tn /SC ONCE /ST 23:59 /TR "powershell -NoProfile -ExecutionPolicy Bypass -File `"$sys`"" /RU SYSTEM /RL HIGHEST /F | Out-Null
  schtasks /Run /TN $tn | Out-Null

  # 简单等待与确认
  Start-Sleep 2
  $limit = (Get-Date).AddSeconds(8)
  while((Get-Date) -lt $limit){
    if(Test-Path $ok){ break }
    Start-Sleep 1
  }
  if(Test-Path $ok){
    $okMsg = Get-Content $ok -Raw
    WL "SYSTEM 镜像完成 => $okMsg"
  }else{
    WL "SYSTEM 镜像未确认（可能任务未执行）。建议再次运行或检查计划任务历史：$tn" $S_WARN
  }

  WL "SYSTEM 镜像脚本已运行。随后立即复检。"
}



# 主控
function Main {
  Start-AudioEtw
  try{
    WL "=== APO 一体化自检开始：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
    Run-SelfTest

    if($RebuildEndpoints){
      Rebuild-Endpoints
      Run-SelfTest
    }

    if($ForceMirror){
      Force-Mirror-System
      Run-SelfTest
    }

  } finally {
    Stop-AudioEtw
    WL "日志文件: $Log"
    if($CaptureEtw -and (Test-Path $EtwEtl)){ WL "ETW: $EtwEtl" }
  }
}

Main
