<#
.SYNOPSIS
    把 dsh-onenote 装进指定的 DSH profile。

.DESCRIPTION
    1. 把本目录同步到 %USERPROFILE%\.dsh\plugins-src\dsh-onenote（插件暂存区）
    2. 复制进 <profile>\node_modules\dsh-onenote
    3. 在 profile 的 package.json 里登记 file: 依赖和 dsh.profile.bundles 条目
       （改之前会备份成 package.json.bak-dsh-onenote）
    4. 如果能找到 dsh CLI，再跑一次 pnpm 安装，把 pnpm-lock.yaml 补齐
       （私有 registry 用 -Registry 指定，否则 pnpm 会去 registry.npmjs.org 打 404）
    5. 打印重启提示

    插件仓库里不存任何机器绝对路径：cordis.patch.yml 的 extraSections 默认是空数组。
    需要挂本机独立分区时，用 -ExtraSectionPath 传入，安装脚本会写进「装好的那份」配置里。

    装完必须重启 DSH 桌面端，bundle 列表在启动时组合。
#>
[CmdletBinding()]
param(
    [string]$ProfileName = 'desktop',
    [string]$Registry = '',
    [string]$DshCli = '',
    [string]$ExtraSectionName = '',
    [string]$ExtraSectionPath = ''
)

$ErrorActionPreference = 'Stop'

$source = $PSScriptRoot
# DSH_HOME wins: when it is set, the desktop app boots the profile under it and
# a profile under %USERPROFILE%\.dsh is never loaded (that mismatch silently
# installs into the wrong home). Fall back to the legacy default.
$dshHome = $env:DSH_HOME
if ([string]::IsNullOrWhiteSpace($dshHome)) { $dshHome = Join-Path $env:USERPROFILE '.dsh' }
$dshHome = $dshHome.TrimEnd('\', '/')
$profileDir = Join-Path $dshHome "profiles\$ProfileName"
$staging = Join-Path $dshHome 'plugins-src\dsh-onenote'
$target = Join-Path $profileDir 'node_modules\dsh-onenote'
$pkgPath = Join-Path $profileDir 'package.json'

if (-not (Test-Path -LiteralPath $pkgPath)) {
    throw "profile not found: $profileDir (expected $pkgPath)"
}
Write-Host "DSH home: $dshHome"

function Invoke-Robocopy([string]$from, [string]$to) {
    robocopy $from $to /MIR /XD node_modules .git /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy failed with exit code $LASTEXITCODE ($from -> $to)" }
    $global:LASTEXITCODE = 0
}

Write-Host "[1/5] staging  $source -> $staging"
New-Item -ItemType Directory -Force -Path $staging | Out-Null
Invoke-Robocopy $source $staging

Write-Host "[2/5] install  $staging -> $target"
if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force }
New-Item -ItemType Directory -Force -Path $target | Out-Null
Invoke-Robocopy $staging $target

# Machine-specific section paths stay OUT of the repo: the shipped
# cordis.patch.yml has `extraSections: []`, and this step fills in the local
# copy inside node_modules when the caller supplies one.
if ($ExtraSectionPath -ne '') {
    $patchPath = Join-Path $target 'cordis.patch.yml'
    if (Test-Path -LiteralPath $patchPath) {
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $name = $ExtraSectionName
        if ([string]::IsNullOrWhiteSpace($name)) { $name = [IO.Path]::GetFileNameWithoutExtension($ExtraSectionPath) }
        $block = "        extraSections:`n          - name: '" + $name + "'`n            path: '" + $ExtraSectionPath + "'"
        $text = [IO.File]::ReadAllText($patchPath, $utf8NoBom)
        if ($text -match '(?m)^\s*extraSections:\s*\[\]\s*$') {
            $text = [regex]::Replace($text, '(?m)^\s*extraSections:\s*\[\]\s*$', $block)
            [IO.File]::WriteAllText($patchPath, $text, $utf8NoBom)
            Write-Host "  extraSections: $name -> $ExtraSectionPath"
        }
    }
}

Write-Host "[3/5] register in $pkgPath"
Copy-Item -LiteralPath $pkgPath -Destination "$pkgPath.bak-dsh-onenote" -Force
$utf8 = New-Object System.Text.UTF8Encoding($false)
$pkg = [IO.File]::ReadAllText($pkgPath, $utf8) | ConvertFrom-Json

$depValue = 'file:' + ($staging -replace '\\', '/')
if ($null -eq $pkg.PSObject.Properties['dependencies']) {
    $pkg | Add-Member -NotePropertyName dependencies -NotePropertyValue ([pscustomobject]@{})
}
if ($pkg.dependencies.PSObject.Properties.Name -contains 'dsh-onenote') {
    $pkg.dependencies.'dsh-onenote' = $depValue
} else {
    $pkg.dependencies | Add-Member -NotePropertyName 'dsh-onenote' -NotePropertyValue $depValue
}

if ($null -eq $pkg.PSObject.Properties['dsh'] -or $null -eq $pkg.dsh.PSObject.Properties['profile']) {
    throw "unexpected profile package.json: missing dsh.profile"
}
$bundles = @($pkg.dsh.profile.bundles)
if ($bundles -notcontains 'dsh-onenote') {
    $pkg.dsh.profile.bundles = @($bundles + 'dsh-onenote')
}

$json = $pkg | ConvertTo-Json -Depth 12
[IO.File]::WriteAllText($pkgPath, $json, $utf8)

Write-Host "[4/5] pnpm install (keeps pnpm-lock.yaml in sync)"
$cli = $DshCli
if ($cli -eq '') {
    $command = Get-Command dsh -ErrorAction SilentlyContinue
    if ($null -ne $command) { $cli = $command.Source }
}
if ($cli -ne '' -and (Test-Path -LiteralPath $cli)) {
    $spec = $depValue
    $arguments = @('plugin', '--profile', $ProfileName, 'add')
    if ($Registry -ne '') { $arguments += @('--registry', $Registry) }
    $arguments += $spec
    & $cli @arguments
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "pnpm add failed (exit $LASTEXITCODE); the manual copy above still works, but re-run the pnpm add later to repair pnpm-lock.yaml."
    }
} else {
    Write-Warning "dsh CLI not found; skipping pnpm add. Run this once to sync pnpm-lock.yaml:"
    Write-Warning "  dsh plugin --profile $ProfileName add --registry $Registry `"$depValue`""
}

Write-Host "[5/5] done"
Write-Host ""
Write-Host "  dependency : $depValue"
Write-Host "  bundles    : $((@($pkg.dsh.profile.bundles)) -join ', ')"
Write-Host ""
Write-Host "重启 DSH 桌面端后，onenote_* 工具才会出现（profile 的 bundle 列表在启动时组合）。"
Write-Host "卸载：删掉 node_modules\dsh-onenote、package.json 里的 dependency 与 bundles 条目，再重启。"
