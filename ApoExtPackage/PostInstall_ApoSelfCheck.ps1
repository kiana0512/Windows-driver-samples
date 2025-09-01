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
$UsbMi00InstanceId = 'USB\VID_0A67&PID_30A2&MI_00\7&3B1FF4EF&0&0000'
# Expected ContainerId (set $null to skip strict matching)
$ExpectedContainerId = '{16b5dc8e-d125-5b68-bfb3-10182a74f929}'

# APO (hardcoded)
$ApoClsid = '{8E3E0B71-5B8A-45C9-9B3D-3A2E5B418A10}'
$ProcessingModePM7 = '{C18E2F7E-933D-4965-B7D1-1EEF228D2AF3}'
$ApoDllName = 'MyCompanyEfxApo.dll'
$ApoDllPath = 'C:\Windows\System32\MyCompanyEfxApo.dll'
$InfName = 'MyCompanyUsbApoExt.inf'

# Policy (embedded; no parameters)
$AutoFix = $true   # auto-fix FxProperties mismatch (via INF reinstall)
$RestartAudioAfterFix = $true   # restart audio services after fix

# =========================
# [1] Output helpers
# =========================
function _W([string]$tag, [string]$msg, [ConsoleColor]$fg = 'Gray') {
  Write-Host "[$tag] $msg" -ForegroundColor $fg
}
function INFO($m) { _W 'INFO' $m 'Gray' }
function STEP($m) { _W 'STEP' $m 'Cyan' }
function PASS($m) { _W 'PASS' $m 'Green' }
function WARN($m) { _W 'WARN' $m 'Yellow' }
function FAIL($m) { _W 'FAIL' $m 'Red' }

function Invoke-Try([ScriptBlock]$b, [string]$desc) {
  try { & $b }catch { FAIL "$desc => $($_.Exception.Message)"; throw }
}

# =========================
# [2] Utilities
# =========================
function Get-ContainerId-ByPnP() {
  $prop = Get-PnpDeviceProperty -InstanceId $UsbMi00InstanceId -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction Stop
  return $prop.Data.ToString()
}

function Get-PnP-Children() {
  # 优先：直接读 DEVPKEY_Device_Children
  try {
    $props = Get-PnpDeviceProperty -InstanceId $UsbMi00InstanceId -KeyName 'DEVPKEY_Device_Children' -ErrorAction Stop
    $vals = @()
    foreach ($p in @($props)) {
      # 统一当作数组处理
      try {
        $d = $p | Select-Object -ExpandProperty Data -ErrorAction Stop
        if ($d) {
          if ($d -is [array]) { $vals += $d } else { $vals += @($d) }
        }
      }
      catch {}
    }
    if ($vals.Count -gt 0) { return $vals }
  }
  catch {}

  # 兜底：按 ContainerId 反查所有 AudioEndpoint 设备
  try {
    $cid = Get-ContainerId-ByPnP
    $eps = Get-PnpDevice -Class AudioEndpoint -PresentOnly -ErrorAction Stop
    $list = @()
    foreach ($ep in @($eps)) {
      try {
        $cid2 = (Get-PnpDeviceProperty -InstanceId $ep.InstanceId -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction Stop).Data
        if ($cid2 -and ($cid2.ToString().ToUpper() -eq $cid.ToUpper())) {
          $list += $ep.InstanceId
        }
      }
      catch {}
    }
    return $list
  }
  catch {
    return @()
  }
}

# 判断 FxProperties 三元组是否满足：7/15 含 CLSID 且 PM7 等于期望
function Test-FxTriplet {
  param([hashtable]$Fx, [string]$ExpectedEfx, [string]$ExpectedPM7)
  function _hasGuid([object]$v, [string]$g) {
    if ($null -eq $v) { return $false }
    $G = $g.ToUpper()
    if ($v -is [string[]]) { return (($v | ForEach-Object { $_.Trim().ToUpper() }) -contains $G) }
    if ($v -is [string]) { return (($v -split '[;\s,]+' | Where-Object { $_ -ne '' } | ForEach-Object { $_.Trim().ToUpper() }) -contains $G) }
    return $false
  }
  $ok7 = _hasGuid $Fx['7']  $ExpectedEfx
  $ok15 = _hasGuid $Fx['15'] $ExpectedEfx
  $okPM7 = ($Fx['PM7'] -is [string]) -and ($Fx['PM7'] -ieq $ExpectedPM7)
  return ($ok7 -and $ok15 -and $okPM7)
}

# 读取硬件键根部 FX\0（不少机型用它而不是 Device Parameters\FxProperties）
function Get-FX0 {
  param([string]$InstanceId)
  $fx0Key = "HKLM:\SYSTEM\CurrentControlSet\Enum\$InstanceId\FX\0"
  $props = Get-ItemProperty -Path $fx0Key -ErrorAction SilentlyContinue
  if (-not $props) { return @{} }
  # 这些名字与你 INF 里写入的属性 GUID 一一对应
  $EFX7 = $props.PSObject.Properties['{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},7']  ; if ($EFX7) { $EFX7 = $EFX7.Value } else { $EFX7 = $null }
  $EFX15 = $props.PSObject.Properties['{D04E05A6-594B-4FB6-A80D-01AF5EED7D1D},15'] ; if ($EFX15) { $EFX15 = $EFX15.Value } else { $EFX15 = $null }
  $MODE7 = $props.PSObject.Properties['{D3993A3F-99C2-4402-B5EC-A92A0367664B},7']  ; if ($MODE7) { $MODE7 = $MODE7.Value } else { $MODE7 = $null }
  return @{ 'EFX7' = $EFX7; 'EFX15' = $EFX15; 'MODE7' = $MODE7 }
}

# 判断 FX\0 是否满足：EFX7/EFX15 含 CLSID 且 MODE7 含期望 PM7
function Test-FX0Triplet {
  param([hashtable]$Fx0, [string]$ExpectedEfx, [string]$ExpectedPM7)
  function _has([object]$v, [string]$g) {
    if ($null -eq $v) { return $false }
    $G = $g.ToUpper()
    if ($v -is [string[]]) { return (($v | ForEach-Object { $_.Trim().ToUpper() }) -contains $G) }
    if ($v -is [string]) { return (($v -split '[;\s,]+' | Where-Object { $_ -ne '' } | ForEach-Object { $_.Trim().ToUpper() }) -contains $G) }
    return $false
  }
  $okEfx = (_has $Fx0['EFX7']  $ExpectedEfx) -or (_has $Fx0['EFX15'] $ExpectedEfx)
  $okPM7 = _has $Fx0['MODE7'] $ExpectedPM7
  return ($okEfx -and $okPM7)
}

#（可复用你已有的）在同一 ContainerId 内遍历兄弟 devnode 的辅助：
function Find-FxProps-InContainer {
  param([string]$AnchorInstanceId, [string]$ExpectedEfx, [string]$ExpectedPM7)
  try { $cid = (Get-PnpDeviceProperty -InstanceId $AnchorInstanceId -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction Stop).Data.ToString() } catch { return $null }
  $cands = Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue
  foreach ($d in @($cands)) {
    try {
      $cid2 = (Get-PnpDeviceProperty -InstanceId $d.InstanceId -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction Stop).Data.ToString()
      if ($cid2 -ne $cid) { continue }
      $fx = Get-FxProps -UsbMi00InstanceId $d.InstanceId
      if (Test-FxTriplet -Fx $fx -ExpectedEfx $ExpectedEfx -ExpectedPM7 $ExpectedPM7) {
        return @{ InstanceId = $d.InstanceId; Fx = $fx }
      }
    }
    catch {}
  }
  return $null
}

function Find-FX0-InContainer {
  param([string]$AnchorInstanceId, [string]$ExpectedEfx, [string]$ExpectedPM7)
  try { $cid = (Get-PnpDeviceProperty -InstanceId $AnchorInstanceId -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction Stop).Data.ToString() } catch { return $null }
  $cands = Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue
  foreach ($d in @($cands)) {
    try {
      $cid2 = (Get-PnpDeviceProperty -InstanceId $d.InstanceId -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction Stop).Data.ToString()
      if ($cid2 -ne $cid) { continue }
      $fx0 = Get-FX0 -InstanceId $d.InstanceId
      if (Test-FX0Triplet -Fx0 $fx0 -ExpectedEfx $ExpectedEfx -ExpectedPM7 $ExpectedPM7) {
        return @{ InstanceId = $d.InstanceId; Fx0 = $fx0 }
      }
    }
    catch {}
  }
  return $null
}

function Extract-Guid-From-MMDev-Instance([string]$mmdevInstance) {
  # e.g. SWD\MMDEVAPI\{0.0.0.00000000}.{E1F9...}
  if ($mmdevInstance -match '\.\{([0-9A-Fa-f-]{36})\}$') {
    return '{' + $Matches[1].ToUpper() + '}'
  }
  return $null
}

function Get-MMDevices-Path([string]$flow, [string]$guidCurly) {
  # flow: Render / Capture
  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$flow\$guidCurly"
}

function Restart-AudioServices() {
  STEP "Restart Windows Audio services"
  sc.exe stop audiosrv | Out-Null
  sc.exe stop audioendpointbuilder | Out-Null
  Start-Sleep -Seconds 1
  sc.exe start audioendpointbuilder | Out-Null
  sc.exe start audiosrv | Out-Null
  PASS "Audio services restarted"
}

# Legacy (kept but unused for Enum ACL)
function Grant-RegTree([string]$path) {
  if (!(Test-Path $path)) { return }
  $adm = New-Object System.Security.Principal.NTAccount('BUILTIN', 'Administrators')
  $sys = New-Object System.Security.Principal.NTAccount('NT AUTHORITY', 'SYSTEM')
  $acl = Get-Acl $path
  try { $acl.SetOwner($adm) } catch {}
  $inherit = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit
  $prop = [System.Security.AccessControl.PropagationFlags]::None
  $rAdm = New-Object System.Security.AccessControl.RegistryAccessRule($adm, 'FullControl', $inherit, $prop, 'Allow')
  $rSys = New-Object System.Security.AccessControl.RegistryAccessRule($sys, 'FullControl', $inherit, $prop, 'Allow')
  $acl.SetAccessRule($rAdm) | Out-Null
  $acl.AddAccessRule($rSys) | Out-Null
  try { Set-Acl -Path $path -AclObject $acl } catch {}
}

function Ensure-REG-SZ([string]$regPath, [string]$name, [string]$expected) {
  $cur = $null
  try {
    $cur = Get-ItemPropertyValue -Path $regPath -Name $name -ErrorAction Stop
  }
  catch {}
  if ($null -eq $cur) {
    if (!(Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    New-ItemProperty -Path $regPath -Name $name -PropertyType String -Value $expected -Force | Out-Null
    return $true
  }
  elseif ($cur -ne $expected) {
    Set-ItemProperty -Path $regPath -Name $name -Value $expected
    return $true
  }
  return $false
}

# Get unnamed/default registry value; try HKCR/HKLM; expand env vars.
function Get-RegistryDefault([string]$keyPath) {
  try {
    $key = Get-Item -Path $keyPath -ErrorAction Stop
    $val = $key.GetValue('', $null, 'DoNotExpandEnvironmentNames')
    if ($null -ne $val) { return [string]$val }
  }
  catch {}
  return $null
}
function Expand-EnvPath([string]$p) {
  if ([string]::IsNullOrWhiteSpace($p)) { return $p }
  return [Environment]::ExpandEnvironmentVariables($p)
}

function Test-Audiodg-HasModule([string]$moduleName) {
  $p = Get-Process -Name audiodg -ErrorAction SilentlyContinue
  if (!$p) { return $false }
  foreach ($proc in $p) {
    try {
      $m = $proc.Modules | Where-Object { $_.ModuleName -ieq $moduleName }
      if ($m) { return $true }
    }
    catch {}
  }
  return $false
}

# Robust helper: find MMDevices GUID by matching any Properties value to a given InstanceId.
function Find-MMDevice-ByInstanceId([string]$flow, [string]$inst) {
  $base = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$flow"
  if (!(Test-Path $base)) { return $null }
  foreach ($k in Get-ChildItem $base -ErrorAction SilentlyContinue) {
    $props = Join-Path $k.PSPath 'Properties'
    if (!(Test-Path $props)) { continue }
    # Preferred property: PKEY_Device_InstanceId
    $iid = $null
    try {
      $iid = Get-ItemPropertyValue -Path $props -Name '{78C34FC8-104A-4D11-9F5B-700F2848BCA5},256' -ErrorAction Stop
    }
    catch {}
    if ($iid -and ($iid -ieq $inst)) { return $k.PSChildName }
    # Fallback: scan all REG_SZ values and compare
    try {
      $all = Get-ItemProperty -Path $props -ErrorAction SilentlyContinue
      if ($all) {
        foreach ($pname in ($all.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name)) {
          $val = $null
          try { $val = Get-ItemPropertyValue -Path $props -Name $pname -ErrorAction Stop }catch {}
          if (($val -is [string]) -and ($val -ieq $inst)) { return $k.PSChildName }
        }
      }
    }
    catch {}
  }
  return $null
}

# === New: read FxProperties (no '?.' operator) ===
# 读取设备级 FxProperties，显式判空（不使用 ?.)
function Get-FxProps {
  param([string]$UsbMi00InstanceId)
  $fxKey = "HKLM:\SYSTEM\CurrentControlSet\Enum\$UsbMi00InstanceId\Device Parameters\FxProperties"
  $props = Get-ItemProperty -Path $fxKey -ErrorAction SilentlyContinue
  if (-not $props) { return @{} }
  $p7 = $props.PSObject.Properties['7']
  $p15 = $props.PSObject.Properties['15']
  $pPM = $props.PSObject.Properties['PM7']
  return @{
    '7'   = if ($p7) { $p7.Value }  else { $null }
    '15'  = if ($p15) { $p15.Value } else { $null }
    'PM7' = if ($pPM) { $pPM.Value } else { $null }
  }
}
# 判断 7/15/PM7 是否符合预期（兼容 REG_SZ / REG_MULTI_SZ / 分隔串）
function Test-FxTriplet {
  param(
    [hashtable]$Fx,
    [string]$ExpectedEfx,
    [string]$ExpectedPM7
  )
  function _hasGuid([object]$v, [string]$g) {
    if ($null -eq $v) { return $false }
    $G = $g.ToUpper()
    if ($v -is [string[]]) { return (($v | ForEach-Object { $_.Trim().ToUpper() }) -contains $G) }
    if ($v -is [string]) { return (($v -split '[;\s,]+' | Where-Object { $_ -ne '' } | ForEach-Object { $_.Trim().ToUpper() }) -contains $G) }
    return $false
  }
  $ok7 = _hasGuid $Fx['7']  $ExpectedEfx
  $ok15 = _hasGuid $Fx['15'] $ExpectedEfx
  $okPM7 = ($Fx['PM7'] -is [string]) -and ($Fx['PM7'] -ieq $ExpectedPM7)
  return ($ok7 -and $ok15 -and $okPM7)
}

# 在同一 ContainerId 下的兄弟 devnode 里寻找已写好的 FxProperties
function Find-FxProps-InContainer {
  param(
    [string]$AnchorInstanceId,   # 例如 MI_00
    [string]$ExpectedEfx,
    [string]$ExpectedPM7
  )
  try {
    $cid = (Get-PnpDeviceProperty -InstanceId $AnchorInstanceId -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction Stop).Data.ToString()
  }
  catch { return $null }

  $cands = Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue
  foreach ($d in @($cands)) {
    $iid = $d.InstanceId
    if ([string]::IsNullOrWhiteSpace($iid)) { continue }
    try {
      $cid2 = (Get-PnpDeviceProperty -InstanceId $iid -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction Stop).Data.ToString()
    }
    catch { continue }
    if ($cid2 -ne $cid) { continue }

    # 只在同一个 Container 内检查
    $fx = Get-FxProps -UsbMi00InstanceId $iid
    if (Test-FxTriplet -Fx $fx -ExpectedEfx $ExpectedEfx -ExpectedPM7 $ExpectedPM7) {
      return @{ InstanceId = $iid; Fx = $fx }
    }
  }
  return $null
}

# 兼容 REG_SZ / REG_MULTI_SZ / 分号/逗号分隔串：判断是否“包含该 GUID”
function Test-ValueHasGuid {
  param([object]$Value, [string]$Guid)
  if ($null -eq $Value) { return $false }
  $g = $Guid.ToUpper()
  if ($Value -is [string[]]) {
    return (($Value | ForEach-Object { $_.Trim().ToUpper() }) -contains $g)
  }
  elseif ($Value -is [string]) {
    $parts = $Value -split '[;\s,]+' | Where-Object { $_ -ne '' }
    return (($parts | ForEach-Object { $_.Trim().ToUpper() }) -contains $g)
  }
  else { return $false }
}

# 检查 INF 是否在 [*.NT.HW] 段里向 HKR,"Device Parameters\FxProperties" 写入（硬件键）
function Inspect-InfForFxRegHW {
  param([string]$InfPath)

  if (!(Test-Path -LiteralPath $InfPath)) {
    return @{ HasHWAddReg = $false; Reason = "INF not found: $InfPath" }
  }

  # 读入并拆行
  $raw = Get-Content -LiteralPath $InfPath -Raw
  $lines = $raw -split "`r?`n"

  $hwSections = @()
  $currentSection = $null
  $inHW = $false

  # 被 .HW 节 AddReg= 引用到的目标小节集合
  $addRegTargets = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)

  # -------- Pass 1：在 *.NT*.HW 里找直接 HKR，并收集 AddReg 目标 --------
  foreach ($line in $lines) {
    $m = [regex]::Match($line, '^\s*\[\s*([^\]]+)\s*\]')
    if ($m.Success) {
      $currentSection = $m.Groups[1].Value
      $inHW = ($currentSection -match '\.NT([^\]]*)?\.HW\b')
      if ($inHW) { $hwSections += $currentSection }
      continue
    }
    if (-not $inHW) { continue }

    $stripped = ($line -replace ';.*$', '').Trim()
    if ([string]::IsNullOrWhiteSpace($stripped)) { continue }

    # 直接在 .NT*.HW 节里写 HKR, ..., Device Parameters\FxProperties
    if ($stripped -match '(?i)^\s*HKR\s*,\s*.*Device\s*Parameters\\FxProperties') {
      return @{ HasHWAddReg = $true; Path = "direct in [$currentSection]" }
    }

    # 收集 AddReg 目标（允许多目标，用逗号分隔）
    $m2 = [regex]::Match($stripped, '^\s*AddReg\s*=\s*(.+)$', 'IgnoreCase')
    if ($m2.Success) {
      $targets = ($m2.Groups[1].Value -split '\s*,\s*') | Where-Object { $_ -ne '' }
      foreach ($t in $targets) { $addRegTargets.Add($t) | Out-Null }
    }
  }

  # -------- Pass 2：到被 AddReg 引用的小节中查找 HKR, ..., Device Parameters\FxProperties --------
  if ($addRegTargets.Count -gt 0) {
    $active = $false
    $secName = $null
    foreach ($line in $lines) {
      $m = [regex]::Match($line, '^\s*\[\s*([^\]]+)\s*\]')
      if ($m.Success) {
        $secName = $m.Groups[1].Value
        $active = $addRegTargets.Contains($secName)
        continue
      }
      if (-not $active) { continue }

      $stripped = ($line -replace ';.*$', '').Trim()
      if ([string]::IsNullOrWhiteSpace($stripped)) { continue }

      if ($stripped -match '(?i)^\s*HKR\s*,\s*.*Device\s*Parameters\\FxProperties') {
        return @{ HasHWAddReg = $true; Path = "via [$secName]" }
      }
    }
  }

  # 没找到
  $why = if ($hwSections.Count -eq 0) { "*.NT*.HW section not present" } else { "no HKR Device Parameters\\FxProperties under any *.NT*.HW (nor its AddReg targets)" }
  return @{ HasHWAddReg = $false; Reason = $why }
}

# === New: locate published oem*.inf for a given INF name ===
function Get-InfPublishedNames {
  param([string]$InfName)

  $txt = (pnputil /enum-drivers | Out-String)
  $lines = $txt -split "`r?`n"

  $pubs = @()
  $currPub = $null
  $currInf = $null

  foreach ($ln in $lines) {
    if ([string]::IsNullOrWhiteSpace($ln)) {
      if ($currPub -and $currInf) {
        if ($currInf -ieq $InfName) { $pubs += $currPub }
      }
      $currPub = $null
      $currInf = $null
      continue
    }

    # 例："Published Name : oem10.inf" 或 "发布名称: oem10.inf"
    if ($ln -match '(?im)\b(oem\d+\.inf)\b') {
      $currPub = $Matches[1].ToLower()
      continue
    }

    # 例："Original Name : MyCompanyUsbApoExt.inf" 或 "原始名称: MyCompanyUsbApoExt.inf"
    # 捕获“非 oem*.inf”的 .inf 名称
    if ($ln -match '(?im)\b(?!oem)\w[\w\-\._]*\.inf\b') {
      $val = $Matches[0]
      if ($val -notmatch '^(?i)oem\d+\.inf$') {
        $currInf = $val
        continue
      }
    }
  }

  # flush 最后一个 block（文件末尾没有空行时）
  if ($currPub -and $currInf) {
    if ($currInf -ieq $InfName) { $pubs += $currPub }
  }

  return $pubs | Select-Object -Unique
}


# === New: Reapply Fx properties by reinstalling Extension INF (handles exit 259) ===
function Reapply-FxProps-via-Inf {
  param(
    [string]$InfPath,
    [string]$UsbMi00InstanceId,
    [string]$InfNameParam
  )
  if (!(Test-Path $InfPath)) { throw "INF not found: $InfPath" }

  Write-Host "[INFO] Reinstalling Extension INF to stamp FxProperties..."
  $args = "/add-driver `"$InfPath`" /install"
  $p = Start-Process -FilePath "pnputil.exe" -ArgumentList $args -NoNewWindow -PassThru -Wait

  if ($p.ExitCode -eq 0) {
    Start-Sleep -Seconds 1
  }
  elseif ($p.ExitCode -eq 259) {
    Write-Warning "pnputil reported 'already the best driver' (259); forcing reinstall..."
    $pubs = @((Get-InfPublishedNames -InfName $InfNameParam))
    if ($pubs.Count -eq 0) {
      throw "Cannot locate published name (oem*.inf) for $InfNameParam"
    }
    foreach ($pub in $pubs) {
      Write-Host "[INFO] Deleting driver package $pub ..."
      $pDel = Start-Process -FilePath "pnputil.exe" -ArgumentList "/delete-driver $pub /uninstall /force" -NoNewWindow -PassThru -Wait
      if ($pDel.ExitCode -ne 0) {
        Write-Warning "Delete-driver $pub returned $($pDel.ExitCode) (continuing)"
      }
    }

    Start-Sleep -Seconds 1
    $p2 = Start-Process -FilePath "pnputil.exe" -ArgumentList $args -NoNewWindow -PassThru -Wait
    if ($p2.ExitCode -ne 0 -and $p2.ExitCode -ne 259) {
      throw "pnputil reinstall failed (exit=$($p2.ExitCode))"
    }
  }
  else {
    throw "pnputil /install failed (exit=$($p.ExitCode))"
  }

  # refresh device
  try {
    Import-Module PnpDevice -ErrorAction SilentlyContinue | Out-Null
    Disable-PnpDevice -InstanceId $UsbMi00InstanceId -Confirm:$false -ErrorAction Stop
    Start-Sleep 1
    Enable-PnpDevice  -InstanceId $UsbMi00InstanceId -Confirm:$false -ErrorAction Stop
  }
  catch {
    pnputil /scan-devices | Out-Null
  }
}

function Try-PlaySystemSound() {
  try {
    Add-Type -AssemblyName System.Media | Out-Null
    $wav = Join-Path $env:WINDIR 'Media\Windows Background.wav'
    if (Test-Path $wav) {
      (New-Object System.Media.SoundPlayer($wav)).PlaySync()
    }
  }
  catch {}
}

# =========================
# [3] Environment & services
# =========================
STEP "Environment"
$psver = $PSVersionTable.PSVersion.ToString()
$arch = if ([Environment]::Is64BitProcess) { 'X64' } else { 'X86' }
$admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (!$admin) { FAIL "Run as Administrator"; throw }
PASS "PS=$psver; Arch=$arch; Admin=$admin"

Invoke-Try {
  $s1 = Get-Service -Name audiosrv
  $s2 = Get-Service -Name audioendpointbuilder
  if ($s1.Status -ne 'Running' -or $s2.Status -ne 'Running') {
    WARN "Audio services not fully running, starting..."
    if ($s2.Status -ne 'Running') { Start-Service audioendpointbuilder }
    if ($s1.Status -ne 'Running') { Start-Service audiosrv }
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
if (-not $comPath) { $comPath = Get-RegistryDefault $clsidKeyHKLM }

if ($comPath) {
  $comPathExpanded = Expand-EnvPath $comPath
  PASS "COM InprocServer32 => $comPathExpanded"
}
else {
  FAIL "COM InprocServer32 missing default value (HKCR/HKLM) for $ApoClsid"
}

# 4.2 DLL presence & signature
if (Test-Path $ApoDllPath) {
  PASS "DLL Exists => $ApoDllPath"
  $sig = Get-AuthenticodeSignature -FilePath $ApoDllPath
  $status = $sig.Status
  $signer = if ($sig.SignerCertificate) { $sig.SignerCertificate.Subject } else { '' }
  PASS "DLL Signature => Status=$status; Signer=$signer"
}
else {
  FAIL "DLL not found => $ApoDllPath"
}

# 4.3 DriverStore
Invoke-Try {
  $infFound = (pnputil /enum-drivers) -match [Regex]::Escape($InfName) | Select-Object -First 1
  if ($infFound) { PASS "DriverStore => found '$InfName' = True" } else { WARN "DriverStore => '$InfName' not found (non-fatal)" }
} "Enumerate drivers"

# =========================
# [5] Device tree & dynamic endpoint discovery (multi-method)
# =========================
STEP "Device tree and dynamic endpoint discovery (multi-method)"

# 5.1 MI_00 PnP presence
$mi = Get-PnpDevice -InstanceId $UsbMi00InstanceId -ErrorAction SilentlyContinue
if (!$mi) { FAIL "USB MI_00 not found: $UsbMi00InstanceId"; throw }
PASS "USB MI_00 present: $($mi.Status) | $($mi.InstanceId)"

# 5.2 ContainerId (method A: PnP)
$cidA = Get-ContainerId-ByPnP
INFO "ContainerId (PnP) => $cidA"
if ($ExpectedContainerId) {
  if ($cidA -ieq $ExpectedContainerId) { PASS "ContainerId matches expected" } else { WARN "ContainerId differs from expected: $cidA <> $ExpectedContainerId" }
}

# 5.3 Children (method A: PnP)
$children = Get-PnP-Children
INFO "Children => $($children -join '; ')"

# Extract render/capture endpoints
$mmdevRenderA = $children | Where-Object { $_ -like 'SWD\MMDEVAPI\{0.0.0.00000000}.*' } | Select-Object -First 1
$mmdevCaptureA = $children | Where-Object { $_ -like 'SWD\MMDEVAPI\{0.0.1.00000000}.*' } | Select-Object -First 1

if ($mmdevRenderA) { PASS "Render(A: PnP) => $mmdevRenderA" } else { FAIL "Render(A: PnP) => not found" }
if ($mmdevCaptureA) { PASS "Capture(A: PnP) => $mmdevCaptureA" } else { FAIL "Capture(A: PnP) => not found" }

$rguidA = if ($mmdevRenderA) { Extract-Guid-From-MMDev-Instance $mmdevRenderA }
$cguidA = if ($mmdevCaptureA) { Extract-Guid-From-MMDev-Instance $mmdevCaptureA }
INFO "GUID(A) => Render=$rguidA | Capture=$cguidA"

# 5.4 Method B: MMDevices (reverse by InstanceId), fully guarded
$rguidB = if ($mmdevRenderA) { Find-MMDevice-ByInstanceId 'Render' $mmdevRenderA } else { $null }
$cguidB = if ($mmdevCaptureA) { Find-MMDevice-ByInstanceId 'Capture' $mmdevCaptureA } else { $null }

if ($rguidB) { PASS "Render(B: MMDevices) => $rguidB" } else { WARN "Render(B: MMDevices) => not found by InstanceId (not fatal)" }
if ($cguidB) { PASS "Capture(B: MMDevices) => $cguidB" } else { WARN "Capture(B: MMDevices) => not found by InstanceId (not fatal)" }

# 5.5 Final endpoint GUIDs (prefer A, fallback B)
$RenderGuid = if ($rguidA) { $rguidA } elseif ($rguidB) { $rguidB } else { $null }
$CaptureGuid = if ($cguidA) { $cguidA } elseif ($cguidB) { $cguidB } else { $null }

if ($RenderGuid) { PASS "Render GUID (resolved) => $RenderGuid" } else { FAIL "Render GUID => unresolved" }
if ($CaptureGuid) { PASS "Capture GUID (resolved) => $CaptureGuid" } else { FAIL "Capture GUID => unresolved" }

# 5.6 Presence & ContainerId check for endpoints
if ($mmdevRenderA) {
  $r_dev = Get-PnpDevice -InstanceId $mmdevRenderA -ErrorAction SilentlyContinue
  if ($r_dev) { PASS "Render endpoint present (PnP) => OK" } else { FAIL "Render endpoint missing (PnP)" }
  $r_cid = (Get-PnpDeviceProperty -InstanceId $mmdevRenderA -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction SilentlyContinue).Data
  if ($r_cid) {
    if (!$ExpectedContainerId -or ($r_cid -ieq $cidA)) { PASS "Render ContainerId OK => $r_cid" } else { WARN "Render ContainerId differs => $r_cid" }
  }
}
if ($mmdevCaptureA) {
  $c_dev = Get-PnpDevice -InstanceId $mmdevCaptureA -ErrorAction SilentlyContinue
  if ($c_dev) { PASS "Capture endpoint present (PnP) => OK" } else { FAIL "Capture endpoint missing (PnP)" }
  $c_cid = (Get-PnpDeviceProperty -InstanceId $mmdevCaptureA -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction SilentlyContinue).Data
  if ($c_cid) {
    if (!$ExpectedContainerId -or ($c_cid -ieq $cidA)) { PASS "Capture ContainerId OK => $c_cid" } else { WARN "Capture ContainerId differs => $c_cid" }
  }
}

# =========================
# [6] Device-level FxProperties (7/15/PM7) check —— 兼容容器兄弟、支持 FX\0 与 Legacy FX\{CLSID}
# =========================
STEP "Device-level FxProperties (7/15/PM7) check"

$fxChanged   = $false
$expectedEfx = $ApoClsid
$expectedPM7 = $ProcessingModePM7

function Format-Value { param($v)
  if ($null -eq $v) { return '' }
  elseif ($v -is [string[]]) { return ($v -join ';') }
  else { return [string]$v }
}

# 旧式 FX\{CLSID} 读取（Dll/EFX/Order）
function Get-LegacyFxLocal { param([string]$InstanceId, [string]$Clsid)
  $k = "HKLM:\SYSTEM\CurrentControlSet\Enum\$InstanceId\FX\$Clsid"
  $p = Get-ItemProperty -Path $k -ErrorAction SilentlyContinue
  if (-not $p) { return @{} }
  $dll   = $p.PSObject.Properties['Dll'];   if ($dll)   { $dll   = $dll.Value }   else { $dll = $null }
  $efx   = $p.PSObject.Properties['EFX'];   if ($efx)   { $efx   = $efx.Value }   else { $efx = $null }
  $order = $p.PSObject.Properties['Order']; if ($order) { $order = $order.Value } else { $order = $null }
  return @{ Dll = $dll; EFX = $efx; Order = $order }
}

# 在同容器下找任一 devnode 命中 Legacy FX\{CLSID}
function Find-LegacyFxInContainerLocal { param([string]$AnchorInstanceId, [string]$Clsid, [string]$DllName)
  try { $cid = (Get-PnpDeviceProperty -InstanceId $AnchorInstanceId -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction Stop).Data.ToString() } catch { return $null }
  $devs = Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue
  foreach ($d in @($devs)) {
    $iid = $d.InstanceId
    if ([string]::IsNullOrWhiteSpace($iid)) { continue }
    try {
      $cid2 = (Get-PnpDeviceProperty -InstanceId $iid -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction Stop).Data.ToString()
    } catch { continue }
    if ($cid2 -ne $cid) { continue }
    $fxL = Get-LegacyFxLocal -InstanceId $iid -Clsid $Clsid
    if ($fxL.Count -gt 0) {
      $dllVal = Format-Value $fxL['Dll']
      $efxVal = $fxL['EFX']
      $okDll  = ($dllVal -ne '') -and ($dllVal -ieq $DllName)
      $okEfx  = ($efxVal -ne $null) -and ([int]$efxVal -eq 1)
      if ($okDll -or $okEfx) {
        return @{ InstanceId = $iid; Legacy = $fxL }
      }
    }
  }
  return $null
}

# ---------- 1) 先查锚点 MI_00 的 Device Parameters\FxProperties ----------
$fxAnchor = Get-FxProps -UsbMi00InstanceId $UsbMi00InstanceId
$anchor7   = Format-Value $fxAnchor['7']
$anchor15  = Format-Value $fxAnchor['15']
$anchorPM7 = Format-Value $fxAnchor['PM7']

if (Test-FxTriplet -Fx $fxAnchor -ExpectedEfx $expectedEfx -ExpectedPM7 $expectedPM7) {
  PASS "FXProps@anchor => 7='$anchor7'  15='$anchor15'  PM7='$anchorPM7'"
}
else {
  WARN "FxProperties mismatch on anchor: 7='$anchor7'  15='$anchor15'  PM7='$anchorPM7' (expect EFX contains $expectedEfx, PM7=$expectedPM7)"

  # ---------- 2) 同容器兄弟 devnode 的 Device Parameters\FxProperties ----------
  $hitFx = Find-FxProps-InContainer -AnchorInstanceId $UsbMi00InstanceId -ExpectedEfx $expectedEfx -ExpectedPM7 $expectedPM7
  if ($hitFx -ne $null) {
    $s7  = Format-Value $hitFx.Fx['7'];  $s15 = Format-Value $hitFx.Fx['15'];  $sPM = Format-Value $hitFx.Fx['PM7']
    PASS "FXProps@sibling: InstanceId=$($hitFx.InstanceId) | 7='$s7'  15='$s15'  PM7='$sPM'"
  }
  else {
    # ---------- 3) 强制重装 INF ----------
    if ($AutoFix) {
      STEP "Reinstall Extension INF to stamp FxProperties"
      $infPath = Join-Path $PSScriptRoot $InfName
      try { Reapply-FxProps-via-Inf -InfPath $infPath -UsbMi00InstanceId $UsbMi00InstanceId -InfNameParam $InfName }
      catch { FAIL "FxProperties reinstall => $($_.Exception.Message)"; throw }
      Start-Sleep -Seconds 1
    }

    # ---------- 4) 重装后再查锚点/兄弟（Device Parameters\FxProperties） ----------
    $fxAfter = Get-FxProps -UsbMi00InstanceId $UsbMi00InstanceId
    $a7  = Format-Value $fxAfter['7'];  $a15 = Format-Value $fxAfter['15'];  $aPM = Format-Value $fxAfter['PM7']
    if (Test-FxTriplet -Fx $fxAfter -ExpectedEfx $expectedEfx -ExpectedPM7 $expectedPM7) {
      PASS "Stamped FXProps@anchor: 7='$a7'  15='$a15'  PM7='$aPM'"
      $fxChanged = $true
    }
    else {
      $hitFx2 = Find-FxProps-InContainer -AnchorInstanceId $UsbMi00InstanceId -ExpectedEfx $expectedEfx -ExpectedPM7 $expectedPM7
      if ($hitFx2 -ne $null) {
        $s2_7 = Format-Value $hitFx2.Fx['7']; $s2_15 = Format-Value $hitFx2.Fx['15']; $s2_PM = Format-Value $hitFx2.Fx['PM7']
        PASS "Stamped FXProps@sibling: InstanceId=$($hitFx2.InstanceId) | 7='$s2_7'  15='$s2_15'  PM7='$s2_PM'"
        $fxChanged = $true
      }
      else {
        # ---------- 5) 回退：硬件键根部 FX\0 ----------
        INFO "FxProperties not found; fallback to FX\\0 check (device root)."
        $fx0Anchor = Get-FX0 -InstanceId $UsbMi00InstanceId
        if (Test-FX0Triplet -Fx0 $fx0Anchor -ExpectedEfx $expectedEfx -ExpectedPM7 $expectedPM7) {
          $f7 = Format-Value $fx0Anchor['EFX7']; $f15 = Format-Value $fx0Anchor['EFX15']; $fPM = Format-Value $fx0Anchor['MODE7']
          PASS "FX0@anchor => EFX7='$f7'  EFX15='$f15'  MODE7='$fPM'"
          $fxChanged = $true
        }
        else {
          $hitFx0 = Find-FX0-InContainer -AnchorInstanceId $UsbMi00InstanceId -ExpectedEfx $expectedEfx -ExpectedPM7 $expectedPM7
          if ($hitFx0 -ne $null) {
            $g7 = Format-Value $hitFx0.Fx0['EFX7']; $g15 = Format-Value $hitFx0.Fx0['EFX15']; $gPM = Format-Value $hitFx0.Fx0['MODE7']
            PASS "FX0@sibling: InstanceId=$($hitFx0.InstanceId) | EFX7='$g7'  EFX15='$g15'  MODE7='$gPM'"
            $fxChanged = $true
          }
          else {
            # ---------- 6) 最后回退：Legacy FX\{CLSID} ----------
            INFO "FX\\0 not found; fallback to Legacy FX\\$ApoClsid check."
            $legacyA = Get-LegacyFxLocal -InstanceId $UsbMi00InstanceId -Clsid $ApoClsid
            $legacyHit = $false
            if ($legacyA.Count -gt 0) {
              $ld = Format-Value $legacyA['Dll']; $le = $legacyA['EFX']; $lo = $legacyA['Order']
              if ( ($ld -ne '' -and $ld -ieq $ApoDllName) -or ($le -ne $null -and ([int]$le -eq 1)) ) {
                PASS "LegacyFX@anchor => Dll='$ld'  EFX='$le'  Order='$lo'"
                $legacyHit = $true
              }
            }
            if (-not $legacyHit) {
              $hitLegacy = Find-LegacyFxInContainerLocal -AnchorInstanceId $UsbMi00InstanceId -Clsid $ApoClsid -DllName $ApoDllName
              if ($hitLegacy -ne $null) {
                $ld2 = Format-Value $hitLegacy.Legacy['Dll']; $le2 = $hitLegacy.Legacy['EFX']; $lo2 = $hitLegacy.Legacy['Order']
                PASS "LegacyFX@sibling: InstanceId=$($hitLegacy.InstanceId) | Dll='$ld2'  EFX='$le2'  Order='$lo2'"
                $legacyHit = $true
              }
            }
            if (-not $legacyHit) {
              FAIL "[FAIL] No Device Parameters\\FxProperties, no FX\\0, and no Legacy FX\\$ApoClsid found."
              throw "Fx configuration not applied on any devnode in container"
            } else {
              $fxChanged = $true
            }
          }
        }
      }
    }
  }
}




# =========================
# [7] Endpoint-level: DisableEnhancements & key existence
# =========================
STEP "Endpoint-level checks (DisableEnhancements, MMDevices existence)"

function Read-DisableEnhancements([string]$flow, [string]$guidCurly) {
  $key = Get-MMDevices-Path $flow $guidCurly
  if (!(Test-Path $key)) { return $null }
  $p = Join-Path $key 'Properties'
  if (!(Test-Path $p)) { return $null }
  try { return Get-ItemPropertyValue -Path $p -Name 'DisableEnhancements' -ErrorAction Stop }catch { return $null }
}

if ($RenderGuid) {
  $rKey = Get-MMDevices-Path 'Render' $RenderGuid
  if (Test-Path $rKey) { PASS "Render MMDevices key => $rKey" } else { WARN "Render MMDevices key missing => $rKey" }
  $de_r = Read-DisableEnhancements 'Render' $RenderGuid
  if ($de_r -ne $null) {
    $msg = if ($de_r -eq 0) { 'Enabled' } elseif ($de_r -eq 1) { 'Disabled' } else { "Unknown($de_r)" }
    PASS "EP(Render) DisableEnhancements=$de_r ($msg)"
  }
  else {
    WARN "EP(Render) DisableEnhancements not found (treat as enabled)"
  }
}

if ($CaptureGuid) {
  $cKey = Get-MMDevices-Path 'Capture' $CaptureGuid
  if (Test-Path $cKey) { PASS "Capture MMDevices key => $cKey" } else { WARN "Capture MMDevices key missing => $cKey" }
  $de_c = Read-DisableEnhancements 'Capture' $CaptureGuid
  if ($de_c -ne $null) {
    $msg = if ($de_c -eq 0) { 'Enabled' } elseif ($de_c -eq 1) { 'Disabled' } else { "Unknown($de_c)" }
    PASS "EP(Capture) DisableEnhancements=$de_c ($msg)"
  }
  else {
    WARN "EP(Capture) DisableEnhancements not found (treat as enabled)"
  }
}

# =========================
# [8] Policy keys
# =========================
STEP "Policy keys (system audio policies)"
$polKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Audio'
$DisableLegacyAudioEffects = $null; $DisableSystemEffects = $null; $EnableCompositeFx = $null
try { $DisableLegacyAudioEffects = Get-ItemPropertyValue -Path $polKey -Name 'DisableLegacyAudioEffects' -ErrorAction Stop }catch {}
try { $DisableSystemEffects = Get-ItemPropertyValue -Path $polKey -Name 'DisableSystemEffects'      -ErrorAction Stop }catch {}
try { $EnableCompositeFx = Get-ItemPropertyValue -Path $polKey -Name 'EnableCompositeFx'         -ErrorAction Stop }catch {}
INFO ("Policy: DisableLegacyAudioEffects={0}; DisableSystemEffects={1}; EnableCompositeFx={2}" -f $DisableLegacyAudioEffects, $DisableSystemEffects, $EnableCompositeFx)

# =========================
# [9] Audio event log (recent 120)
# =========================
STEP "Collect Microsoft-Windows-Audio/Operational recent events"
try {
  $logs = Get-WinEvent -LogName 'Microsoft-Windows-Audio/Operational' -MaxEvents 120 -ErrorAction Stop |
  Select-Object TimeCreated, Id, LevelDisplayName, Message
  PASS "Fetched $($logs.Count) events (suppressing verbose print; use Out-GridView if needed)"
}
catch {
  WARN "Get-WinEvent failed: $($_.Exception.Message)"
}

# =========================
# [10] audiodg module presence
# =========================
STEP "audiodg module presence"
$loaded = Test-Audiodg-HasModule $ApoDllName
if ($loaded) {
  PASS "audiodg has $ApoDllName"
}
else {
  WARN "audiodg not showing $ApoDllName (start playback/recording, then re-check)"
}

# Optional restart after FxProperties change
if ($AutoFix -and $RestartAudioAfterFix -and $fxChanged) {
  Restart-AudioServices
  Start-Sleep -Seconds 1
  if (-not (Test-Audiodg-HasModule $ApoDllName)) {
    Try-PlaySystemSound
  }
  if (Test-Audiodg-HasModule $ApoDllName) {
    PASS "After restart: audiodg has $ApoDllName"
  }
  else {
    WARN "After restart: still not seen (trigger a shared-mode stream)"
  }
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
