param(
  # 设备实例作为锚点（你已确认是 MI_00，这里就写 MI_00）
  [Parameter(Mandatory=$true)]
  [string]$Instance,

  [ValidateSet('Render','Capture')]
  [string]$Kind = 'Render',

  [Parameter(Mandatory=$true)]
  [string]$ExpectedClsid,          # 例如 '{8E3E0B71-5B8A-45C9-9B3D-3A2E5B418A10}'

  [Parameter(Mandatory=$true)]
  [string]$DllName,                # 例如 'MyCompanyEfxApo.dll'

  [string]$InfName = 'MyCompanyUsbApoExt.inf',

  [switch]$NoPlayback,             # 加上该开关则跳过播放触发
  [switch]$Quiet,

  [string]$LogDir = "$env:ProgramData\MyCompany\ApoSelfTest"
)

# ================= 基础 & 日志 =================
$ErrorActionPreference = 'Stop'
function New-Dir([string]$p){ if(-not (Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } }
New-Dir $LogDir
$stamp   = (Get-Date).ToString('yyyyMMdd_HHmmss')
$LogPath = Join-Path $LogDir "ApoSelfTest_$stamp.txt"

# 统一 PASS/WARN/FAIL
$S_OK=0; $S_WARN=1; $S_FAIL=2
$Results = [System.Collections.Generic.List[object]]::new()

function WL([string]$msg, [int]$sev=$S_OK){
  $tag = @('PASS','WARN','FAIL')[$sev]
  $line = "[$tag] $msg"
  $line | Out-File -FilePath $LogPath -Encoding utf8 -Append
  if(-not $Quiet){
    $color = if($sev -eq $S_OK){'Green'} elseif($sev -eq $S_WARN){'Yellow'} else {'Red'}
    Write-Host $line -ForegroundColor $color
  }
}
function Add-Result([string]$name, [bool]$ok, [string]$detail){
  $sev = if($ok){ $S_OK } else { $S_FAIL }
  WL "$name => $(if($ok){'OK'}else{'NG'}) | $detail" $sev
  $Results.Add([pscustomobject]@{ Check=$name; Pass=$ok; Detail=$detail })
}
function Is-Admin {
  $id=[Security.Principal.WindowsIdentity]::GetCurrent()
  (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ================= 常量 PKEY（字符串名就含花括号与逗号） =================
$K_EFX7      = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},7'      # EndpointEffectClsid
$K_EFX15     = '{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},15'     # Composite EndpointEffectClsid (Win10 1803+)
$K_DEF3      = '{FC1CFC9B-31F9-4C56-9D2C-39A781AB0B2E},3'      # 旧 Default EFX（兼容）
$K_SYSFX_OFF = '{1DA5D803-D492-4EDD-8C23-E0C0FFEE7F0E},5'      # 禁用增强

$MMRoot = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Kind"
$expDll = "C:\Windows\System32\$DllName"

# ================= 注册表帮助（.NET，避免本地化/对齐问题） =================
Add-Type -AssemblyName 'Microsoft.Win32.Registry' | Out-Null
function Open-BaseKey([string]$hive,[string]$view){
  $hv=[Microsoft.Win32.RegistryHive]::$hive
  $vw=[Microsoft.Win32.RegistryView]::$view
  [Microsoft.Win32.RegistryKey]::OpenBaseKey($hv,$vw)
}
function Get-RegValue([string]$hive,[string]$view,[string]$subkey,[string]$name){
  try{
    $rk=(Open-BaseKey $hive $view).OpenSubKey($subkey)
    if(-not $rk){ return @{Found=$false; Data=$null; Kind=$null} }
    $val = $rk.GetValue($name,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    if($null -eq $val){ return @{Found=$false; Data=$null; Kind=$null} }
    $kind = $rk.GetValueKind($name)
    $data = switch($kind){
      'MultiString' { ($val -join ';') }
      'Binary'      { ($val | ForEach-Object { $_.ToString('X2') }) -join '' }
      default       { [string]$val }
    }
    return @{Found=$true; Data=$data; Kind=$kind}
  }catch{
    return @{Found=$false; Data=$null; Kind=$null}
  }
}
function List-SubKeys([string]$hive,[string]$view,[string]$subkey){
  try{ ((Open-BaseKey $hive $view).OpenSubKey($subkey))?.GetSubKeyNames() }catch{ @() }
}

# ================= 打印头 =================
WL "=== APO 开机自检（v4）开始：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
WL "PSVersion=$($PSVersionTable.PSVersion); Arch=$([Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture); Admin=$(Is-Admin)"
WL "Params: Instance='$Instance'; Kind=$Kind; CLSID='$ExpectedClsid'; DllName='$DllName'; InfName='$InfName'"

# ================= 服务状态 =================
try{
  $svc1 = Get-Service -Name audiosrv -ErrorAction Stop
  $svc2 = Get-Service -Name audioendpointbuilder -ErrorAction Stop
  Add-Result 'AudioSrv Running' ($svc1.Status -eq 'Running') "Status=$($svc1.Status)"
  Add-Result 'AudioEndpointBuilder Running' ($svc2.Status -eq 'Running') "Status=$($svc2.Status)"
}catch{
  Add-Result 'Audio Services Present' $false "无法查询音频服务：$($_.Exception.Message)"
}

# ================= COM 注册（主分支 + 矩阵只写入日志） =================
# 主分支：HKLM\SOFTWARE\Classes（x64）
$dll=$null; $tm=$null
try{
  $rk = (Open-BaseKey 'LocalMachine' 'Registry64').OpenSubKey("SOFTWARE\Classes\CLSID\$ExpectedClsid\InprocServer32")
  if($rk){
    $dll = $rk.GetValue($null)                # (默认)
    $tm  = $rk.GetValue('ThreadingModel')
  }
}catch{}
$primaryOk = ($dll -ieq $expDll -and $tm -match 'Both')
Add-Result 'COM x64 InprocServer32 (Primary)' $primaryOk "Dll='$dll'; TM='$tm' (期望='$expDll', Both)"

# 其他视图矩阵（写入日志，便于排查假阴性）
$comMatrix = @(
  @{Hive='LocalMachine'; View='Registry64'; Path="SOFTWARE\Classes\CLSID\$ExpectedClsid\InprocServer32"},
  @{Hive='ClassesRoot';  View='Default';    Path="CLSID\$ExpectedClsid\InprocServer32"},
  @{Hive='CurrentUser';  View='Default';    Path="Software\Classes\CLSID\$ExpectedClsid\InprocServer32"},
  @{Hive='LocalMachine'; View='Registry32'; Path="SOFTWARE\Classes\CLSID\$ExpectedClsid\InprocServer32"}
) | ForEach-Object {
  $dv = Get-RegValue $_.Hive $_.View $_.Path $null
  $tm2= Get-RegValue $_.Hive $_.View $_.Path 'ThreadingModel'
  [pscustomobject]@{ Hive=$_.Hive; View=$_.View; Path=$_.Path; Dll=$dv.Data; ThreadingModel=$tm2.Data; Found=$dv.Found }
}
"`n--- COM Registration Matrix ---`n$($comMatrix | Format-Table -AutoSize | Out-String)" |
  Out-File -FilePath $LogPath -Append -Encoding utf8

# ================= DLL 存在/签名 =================
try{
  $fi = Get-Item -Path $expDll -ErrorAction Stop
  Add-Result 'DLL Exists (System32)' $true $fi.FullName
  $sig = Get-AuthenticodeSignature -FilePath $fi.FullName
  Add-Result 'DLL Signature' ($sig.Status -in 'Valid','NotSigned') "Status=$($sig.Status)"
}catch{
  Add-Result 'DLL Exists (System32)' $false "缺失：$expDll"
}

# ================= ContainerId（锚点） =================
$cid=$null
try{
  $cid = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Enum\$Instance").ContainerID
  Add-Result 'Device ContainerId' ($cid -ne $null) "$cid"
}catch{
  Add-Result 'Device ContainerId' $false "读取失败：$($_.Exception.Message)"
}

# ================= 单独检查：你指定的 MI_00 的 FX\0（直击要害） =================
if($cid){
  $fx0Sub = "SYSTEM\CurrentControlSet\Enum\$Instance\Device Parameters\FX\0"
  $v7  = Get-RegValue 'LocalMachine' 'Registry64' $fx0Sub $K_EFX7
  $v15 = Get-RegValue 'LocalMachine' 'Registry64' $fx0Sub $K_EFX15
  $v3  = Get-RegValue 'LocalMachine' 'Registry64' $fx0Sub $K_DEF3
  $okFx0 = (($v15.Found -and $v15.Data -ieq $ExpectedClsid) -or ($v7.Found -and $v7.Data -ieq $ExpectedClsid) -or ($v3.Found -and $v3.Data -ieq $ExpectedClsid))
  Add-Result "MI_00 FX\0 → EFX present" $okFx0 ("7='$($v7.Data)' 15='$($v15.Data)' 3='$($v3.Data)'")
  if(-not $okFx0){
    WL "提示：若此处为 NG，说明 INF 并未把 ,7/15 写入该实例（即便 [Models] 有 MI_00，也可能未命中该功能设备，建议用 devcon update 精确更新到此实例）。" $S_WARN
  }
}

# ================= 同容器的 MEDIA 功能：逐条 FX\0 检查 =================
$media=@()
try{
  $media = Get-PnpDevice -Class MEDIA -PresentOnly -ErrorAction Stop | Where-Object {
    (Get-PnpDeviceProperty -InstanceId $_.InstanceId -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction SilentlyContinue).Data -eq $cid
  }
  Add-Result 'MEDIA Functions in Same Container' ($media.Count -gt 0) "Count=$($media.Count)"
}catch{
  Add-Result 'MEDIA Functions in Same Container' $false "查询失败：$($_.Exception.Message)"
}

$foundAnyFx = $false
foreach($m in $media){
  $fxSub = "SYSTEM\CurrentControlSet\Enum\$($m.InstanceId)\Device Parameters\FX\0"
  $m7  = Get-RegValue 'LocalMachine' 'Registry64' $fxSub $K_EFX7
  $m15 = Get-RegValue 'LocalMachine' 'Registry64' $fxSub $K_EFX15
  $m3  = Get-RegValue 'LocalMachine' 'Registry64' $fxSub $K_DEF3
  $okM = (($m15.Found -and $m15.Data -ieq $ExpectedClsid) -or ($m7.Found -and $m7.Data -ieq $ExpectedClsid) -or ($m3.Found -and $m3.Data -ieq $ExpectedClsid))
  if($okM){ $foundAnyFx = $true }
  Add-Result "Device FX\0 → EFX present" $okM "[$($m.InstanceId)] 7='$($m7.Data)' 15='$($m15.Data)' 3='$($m3.Data)'"
}
if((-not $foundAnyFx) -and $media.Count -gt 0){
  WL "提示：容器内 MEDIA 功能均未发现 ,7/15=你的 CLSID。若确实是 MI_00 参与 Render，请使用：`devcon update .\MyCompanyUsbApoExt.inf \"$Instance\"` 强制把 INF 命中到这条实例，然后重启该实例。" $S_WARN
}

# ================= 端点发现（以 PnP 为准，再映射回 MMDevices） =================
$endpoints=@()
try{
  # 先枚举 MMDevices 下本 Kind 的所有端点 GUID
  $kids = Get-ChildItem $MMRoot -ErrorAction Stop
  foreach($k in $kids){
    $g = $k.PSChildName            # 形如 {4ffd4a6d-...}
    $eid = "SWD\MMDEVAPI\{0.0.0.00000000}." + $g.ToUpper()  # 对应的 PnP 端点实例

    try{
      $dev  = Get-PnpDevice -InstanceId $eid -PresentOnly -ErrorAction Stop
      $cid2 = (Get-PnpDeviceProperty -InstanceId $eid -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction SilentlyContinue).Data
      if($cid2 -and ($cid2.ToString().ToLower() -eq $cid.ToString().ToLower())){
        # 同容器，收录
        $endpoints += [pscustomobject]@{
          Guid       = $g
          PnpStatus  = $dev.Status
          PnpProblem = $dev.Problem
        }
      }
    }catch{
      # 端点可能不属于本容器或已隐藏，忽略
    }
  }

  Add-Result 'Endpoints in Same Container' ($endpoints.Count -gt 0) "Count=$($endpoints.Count)"
}catch{
  Add-Result 'Endpoints in Same Container' $false "枚举失败：$($_.Exception.Message)"
}

# 对每个端点做 FxProperties/Properties 检查（其余逻辑不变）
foreach($ep in $endpoints){
  $g = $ep.Guid
  $fxSub = "SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Kind\$g\FxProperties"
  $prSub = "SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Kind\$g\Properties"

  $e7  = Get-RegValue 'LocalMachine' 'Registry64' $fxSub $K_EFX7
  $e15 = Get-RegValue 'LocalMachine' 'Registry64' $fxSub $K_EFX15
  $e3  = Get-RegValue 'LocalMachine' 'Registry64' $fxSub $K_DEF3
  $den = Get-RegValue 'LocalMachine' 'Registry64' $prSub $K_SYSFX_OFF

  $okFx = (($e15.Found -and $e15.Data -ieq $ExpectedClsid) -or
           ($e7.Found  -and $e7.Data  -ieq $ExpectedClsid) -or
           ($e3.Found  -and $e3.Data  -ieq $ExpectedClsid))
  Add-Result "Endpoint[$g] FxProperties → EFX present" $okFx ("7='$($e7.Data)' 15='$($e15.Data)' 3='$($e3.Data)'")

  $okEn = ($den.Found -and ($den.Data -eq '0' -or $den.Data -eq '0x0'))
  Add-Result "Endpoint[$g] DisableEnhancements=0" $okEn "Value='$($den.Data)'"

  # 顺便把 PnP 状态也记到日志里，方便看端点是否被隐藏/禁用
  WL "Endpoint[$g] PnP => Status=$($ep.PnpStatus); Problem=$($ep.PnpProblem)" $S_OK
}


  # 子键（RAW/Media/Communications 等）逐一检查是否继承了 7/15
  $subs = List-SubKeys 'LocalMachine' 'Registry64' $fxSub
  $miss = @()
  foreach($s in $subs){
    $sk = "$fxSub\$s"
    $s7  = Get-RegValue 'LocalMachine' 'Registry64' $sk $K_EFX7
    $s15 = Get-RegValue 'LocalMachine' 'Registry64' $sk $K_EFX15
    if(-not (($s15.Found -and $s15.Data -ieq $ExpectedClsid) -or ($s7.Found -and $s7.Data -ieq $ExpectedClsid))){
      $miss += $s
    }
  }
  Add-Result "Endpoint[$g] ProcessingMode Subkeys OK" ($miss.Count -eq 0) ("UncheckedOrMissing=" + $miss.Count)


# ================= 驱动包存在（pnputil） =================
try{
  $pnpo = cmd /c "pnputil /enum-drivers" 2>&1
  $hasInf = (($pnpo -join "`n") -match [regex]::Escape($InfName))
  Add-Result 'Extension INF Installed' ([bool]$hasInf) ("pnputil found '$InfName' = " + [bool]$hasInf)
}catch{
  Add-Result 'Extension INF Installed' $false "pnputil 失败：$($_.Exception.Message)"
}

# ================= 播放触发 & audiodg 模块加载 =================
if(-not $NoPlayback){
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
  WL '已跳过播放触发（-NoPlayback）。' $S_WARN
}

# ================= 可用事件日志 =================
try{
  wevtutil sl Microsoft-Windows-Audio/Operational /e:true | Out-Null
  $hit = Get-WinEvent -FilterHashtable @{ LogName='Microsoft-Windows-Audio/Operational'; StartTime=(Get-Date).AddMinutes(-5) } -ErrorAction SilentlyContinue |
         Where-Object { $_.Message -match 'APO|FX|EFX|RAW|Activate|CLSID' } | Select-Object -First 8
  $okEvt = ($hit -ne $null -and $hit.Count -gt 0)
  Add-Result 'Audio Operational Log Available' $okEvt ("Found=$($hit?.Count)")
  if($okEvt){
    "---- Recent Audio Operational (<=8) ----`n$($hit|Out-String)" | Out-File -FilePath $LogPath -Append -Encoding utf8
  }
}catch{
  Add-Result 'Audio Operational Log Available' $false "查询失败：$($_.Exception.Message)"
}

# ================= 汇总 =================
$failed = $Results | Where-Object { -not $_.Pass }
$overall = ($failed.Count -eq 0)
WL "=== 总体结果：$(if($overall){'PASS'}else{'FAIL'}) ==="
if(-not $Quiet){
  Write-Host "`n==== 总结（关键环节）====" -ForegroundColor Cyan
  $Results | Format-Table -AutoSize
  Write-Host "`n日志文件: $LogPath" -ForegroundColor Gray
}
