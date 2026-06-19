<#
.SYNOPSIS
    VMware vCenter 一键巡检脚本 (PowerShell + REST API + 可选 PowerCLI 回退)

.DESCRIPTION
    通过 vCenter REST API (vSphere Automation API 8.0+) 拉取数据,
    生成工程师风 HTML 巡检报告。零依赖 (PowerCLI 是可选回退)。
    兼容 Windows PowerShell 5.1 与 PowerShell 7+。

    v1.1: 检测到 PowerCLI 时自动补采 REST 拿不到的数据：
      - 单 ESXi 主机实时 CPU/Mem/Uptime/Build (REST 8.0 已 deprecated)
      - VM 快照清单 (含大小 + 年龄 + 链深度, REST 不暴露)
      - 当前 Triggered Alarms (REST 不暴露)

.PARAMETER VCenter
    vCenter IP 或 FQDN, 例: vcenter.lab.local

.PARAMETER Username
    用户名, 例: administrator@vsphere.local

.PARAMETER Password
    密码

.PARAMETER Output
    输出 HTML 路径; 留空则写入脚本目录
    report_<vcenter>_<yyyy-MM-dd>.html

.PARAMETER ToolsSampleSize
    VMware Tools 抽样数量 (开机 VM 抽样调用 /tools 接口), 默认 16

.PARAMETER SkipToolsSample
    跳过 VMware Tools 抽样, 节省 ~10s

.PARAMETER Quiet
    静默运行, 不输出进度

.PARAMETER UsePowerCLI
    强制启用 PowerCLI 回退; 未安装时打印安装提示。

.PARAMETER SkipPowerCLI
    强制跳过 PowerCLI 回退 (即使已安装), 只用纯 REST。

.EXAMPLE
    .\vcenter_inspect.ps1 -VCenter vcenter.lab.local -Username administrator@vsphere.local -Password '***'

.EXAMPLE
    # 显式启用 PowerCLI 回退补全快照/Alarm
    .\vcenter_inspect.ps1 -VCenter vcenter.lab.local -Username root@vsphere.local -Password '***' -UsePowerCLI

.NOTES
    Author : Claude Opus 4.7
    Date   : 2026-06-19
    Style  : 仿 linux_inspect.sh v2.4 工程师风
    Version: v1.1.0 (PowerCLI 回退层)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $VCenter,
    [Parameter(Mandatory=$true)] [string] $Username,
    [Parameter(Mandatory=$true)] [string] $Password,
    [string] $Output,
    [int]    $ToolsSampleSize = 16,
    [switch] $SkipToolsSample,
    [switch] $Quiet,
    [switch] $DebugDump,
    # v1.1: PowerCLI 回退 — 自动检测；显式开关
    [switch] $UsePowerCLI,    # 强制启用 (即使未装也尝试 Install-Module 提示)
    [switch] $SkipPowerCLI,   # 强制跳过 (即使已装)
    # v1.2: 主题
    [ValidateSet('light','dark','minimal','amber')]
    [string] $Theme = 'light',
    [string] $AccentColor     # 可选 hex (#1565c0) 单独覆盖主色,不切整套主题
)

# ============================================================================
#  全局初始化
# ============================================================================
$ErrorActionPreference  = 'Stop'
$ProgressPreference     = 'SilentlyContinue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$OutputEncoding         = [System.Text.Encoding]::UTF8

# 跳过 SSL 证书校验 (vCenter 默认自签)
try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
} catch {}
if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type) {
    Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) { return true; }
}
"@
}
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

# ============================================================================
#  日志输出
# ============================================================================
$Script:StepIdx   = 0
$Script:StepTotal = 14
$Script:T0        = Get-Date
$Script:PCLIEnabled = $false   # v1.1: PowerCLI 回退是否启用

function Log-Banner {
    if ($Quiet) { return }
    $line = '═' * 70
    Write-Host ''
    Write-Host $line -ForegroundColor DarkCyan
    Write-Host (' vCenter Inspect  v1.1.0 |  target: {0}   |  {1}' -f $VCenter, (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -ForegroundColor Cyan
    Write-Host (' steps: {0}                  |  user  : {1}' -f $Script:StepTotal, $Username) -ForegroundColor DarkGray
    Write-Host $line -ForegroundColor DarkCyan
    Write-Host ''
}
function Log-Step([string]$msg) {
    $Script:StepIdx++
    if ($Quiet) { return }
    $pct = [int](100 * $Script:StepIdx / $Script:StepTotal)
    Write-Host ("  [{0,2}/{1}] ({2,3}%) {3}" -f $Script:StepIdx, $Script:StepTotal, $pct, $msg) -ForegroundColor Cyan
}
function Log-Info([string]$msg) { if (-not $Quiet) { Write-Host "         $msg" -ForegroundColor DarkGray } }
function Log-Warn([string]$msg) { Write-Host "  [WARN] $msg" -ForegroundColor Yellow }
function Log-Err ([string]$msg) { Write-Host "  [ERR ] $msg" -ForegroundColor Red }

# ============================================================================
#  低层 HTTP 调用 (手控 UTF-8, 兼容 5.1 / 7+)
# ============================================================================
$Script:VCBase   = "https://$VCenter"
$Script:Session  = $null
# API style: 'v8' = vCenter 7.0+ /api/...  |  'v6' = 6.5/6.7 /rest/...
# Login-VC 会自动探测并设置
$Script:ApiStyle = 'v8'
# Debug dump (-DebugDump 开启时, 每个 endpoint 调用追加 [code, path, body 前 300B] 到日志)
$Script:DebugLog = $null
if ($DebugDump) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
    $Script:DebugLog = Join-Path $scriptDir ("vcenter_debug_{0}_{1}.log" -f ($VCenter -replace '[^\w\.\-]','_'), (Get-Date -Format 'yyyyMMdd_HHmmss'))
    "vcenter_inspect debug log — $VCenter — $(Get-Date)" | Out-File -FilePath $Script:DebugLog -Encoding UTF8
    "" | Out-File -FilePath $Script:DebugLog -Append -Encoding UTF8
}
function Write-Dump {
    param([string]$Path, [int]$Code, [string]$Body)
    if (-not $Script:DebugLog) { return }
    $snippet = ''
    if ($Body) { $snippet = $Body.Substring(0, [math]::Min(300, $Body.Length)) -replace "`r?`n",' ' }
    "[{0,3}]  {1,-55}  {2}" -f $Code, $Path, $snippet | Out-File -FilePath $Script:DebugLog -Append -Encoding UTF8
}

# v6 模式 (vCenter 6.5/6.7) 不支持的 endpoint (子串匹配; 命中则直接返回 null, 不发请求)
# 标准: 这些 endpoint 在 6.5 不存在或为 RPC-style 与 v8 GET 不兼容
$Script:V6Unsupported = @(
    '/access/ssh','/access/shell','/access/dcui','/access/consolecli',
    '/certificate-management',
    '/recovery/backup',
    '/vcenter/deployment',
    '/appliance/update',
    '/appliance/networking',          # 6.7+ 才有
    '/appliance/services',            # 6.7+ 才有
    '/appliance/ntp',                 # 6.5 是 RPC POST, 与 GET 不兼容
    '/appliance/timesync',            # 同上
    '/health/system/lastcheck'        # 6.5 端点路径不同
)

function Invoke-VCRaw {
    param([string]$Path, [string]$Method='GET', [string]$AuthBasic, [string]$Body)
    $url = "$Script:VCBase$Path"
    $req = [System.Net.HttpWebRequest]::Create($url)
    $req.Method            = $Method
    $req.Timeout           = 30000
    $req.ReadWriteTimeout  = 30000
    $req.UserAgent         = 'vcenter_inspect/1.0'
    $req.Accept            = 'application/json'
    $req.KeepAlive         = $true
    if ($AuthBasic)      { $req.Headers.Add('Authorization', "Basic $AuthBasic") }
    if ($Script:Session) { $req.Headers.Add('vmware-api-session-id', $Script:Session) }
    if ($Body) {
        $req.ContentType   = 'application/json'
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
        $req.ContentLength = $bytes.Length
        $s = $req.GetRequestStream(); $s.Write($bytes,0,$bytes.Length); $s.Close()
    } elseif ($Method -eq 'POST') {
        # 明确 Content-Length: 0, 避免某些 vCenter 后端对空 POST 返回 5xx
        $req.ContentLength = 0
    }
    try {
        $resp = $req.GetResponse()
        $sr = New-Object System.IO.StreamReader($resp.GetResponseStream(), [System.Text.Encoding]::UTF8)
        $text = $sr.ReadToEnd(); $sr.Close(); $resp.Close()
        Write-Dump -Path $Path -Code ([int]$resp.StatusCode) -Body $text
        return [pscustomobject]@{ Code = [int]$resp.StatusCode; Text = $text; Error = $null }
    } catch [System.Net.WebException] {
        $we = $_.Exception
        $code = 0; $body = $null
        if ($we.Response) {
            try { $code = [int]$we.Response.StatusCode } catch {}
            try {
                $sr2 = New-Object System.IO.StreamReader($we.Response.GetResponseStream(), [System.Text.Encoding]::UTF8)
                $body = $sr2.ReadToEnd(); $sr2.Close()
            } catch {}
            try { $we.Response.Close() } catch {}
        }
        Write-Dump -Path $Path -Code $code -Body ("ERR: " + $we.Message + " | " + $body)
        return [pscustomobject]@{ Code = $code; Text = $body; Error = $we.Message }
    } catch {
        Write-Dump -Path $Path -Code 0 -Body ("EXCEPTION: " + $_.Exception.Message)
        return [pscustomobject]@{ Code = 0; Text = $null; Error = $_.Exception.Message }
    }
}

# 带自动重试的封装: 对 0 / 5xx 瞬时错误重试, 401/403/404 不重试
function Invoke-VCRetry {
    param([string]$Path, [string]$Method='GET', [string]$AuthBasic, [string]$Body, [int]$Retries=3)
    $attempt = 0; $r = $null
    while ($attempt -lt $Retries) {
        $attempt++
        $r = Invoke-VCRaw -Path $Path -Method $Method -AuthBasic $AuthBasic -Body $Body
        if ($r.Code -ge 200 -and $r.Code -lt 300) { return $r }
        # 不重试: 4xx (认证 / 不存在 / 客户端错)
        if ($r.Code -ge 400 -and $r.Code -lt 500) { return $r }
        # 重试: 0 (网络错), 5xx (服务端瞬时错)
        if ($attempt -lt $Retries) {
            $wait = [int]([math]::Pow(2, $attempt - 1))   # 1s, 2s, 4s
            Log-Warn ("HTTP $($r.Code) on $Path  attempt $attempt/$Retries — retry in ${wait}s")
            Start-Sleep -Seconds $wait
        }
    }
    return $r
}

# 将 v8 风格 /api/... 路径转换为当前 ApiStyle 下的真实路径
function Resolve-VCPath {
    param([string]$Path)
    if ($Script:ApiStyle -ne 'v6') { return $Path }
    if ($Path -eq '/api/session') { return '/rest/com/vmware/cis/session' }
    if ($Path -like '/api/vcenter/vm/*/tools')                { return $null }   # 6.5 无
    if ($Path -like '/api/vcenter/vm/*/snapshot')             { return $null }   # 6.5 无
    if ($Path -like '/api/vcenter/vm/*/guest/*')              { return $null }   # 6.5 无
    if ($Path -like '/api/vcenter/vm/*/hardware')             { return $null }   # 6.5 无
    foreach ($p in $Script:V6Unsupported) {
        if ($Path -like "*$p*") { return $null }
    }
    return ($Path -replace '^/api/', '/rest/')
}

function Invoke-VC {
    param([string]$Path)
    $real = Resolve-VCPath -Path $Path
    if (-not $real) { return $null }   # v6 模式不支持的端点直接 skip
    $r = Invoke-VCRetry -Path $real
    if ($r.Code -ge 200 -and $r.Code -lt 300 -and $r.Text) {
        try {
            $obj = $r.Text | ConvertFrom-Json
            # v6 模式自动 unwrap {"value": ...} 包装
            if ($Script:ApiStyle -eq 'v6' -and $obj -and $obj.PSObject -and ($obj.PSObject.Properties.Name -contains 'value')) {
                return $obj.value
            }
            return $obj
        } catch { return $null }
    }
    return $null
}

function Login-VC {
    $pair = "${Username}:${Password}"
    $b64  = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($pair))

    # 先试 v8 (/api/session, vCenter 7.0+); 2 次重试即可, 避免 6.5 卡 4 次
    $Script:ApiStyle = 'v8'
    $r = Invoke-VCRetry -Path '/api/session' -Method 'POST' -AuthBasic $b64 -Retries 2
    if ($r.Code -eq 201 -and $r.Text) {
        $sid = $r.Text.Trim().Trim('"')
        $Script:Session = $sid
        Log-Info ("API style: v8  (vCenter 7.0+, endpoint /api/...)")
        return $sid
    }

    # v8 失败 → fallback v6 (/rest/com/vmware/cis/session, vCenter 6.5/6.7)
    Log-Info ("v8 /api/session 失败 (HTTP $($r.Code)), fallback 尝试 v6 /rest/com/vmware/cis/session ...")
    $Script:ApiStyle = 'v6'
    $r2 = Invoke-VCRetry -Path '/rest/com/vmware/cis/session' -Method 'POST' -AuthBasic $b64 -Retries 3
    if ($r2.Code -ge 200 -and $r2.Code -lt 300 -and $r2.Text) {
        $sid = $null
        try {
            $obj = $r2.Text | ConvertFrom-Json
            if ($obj -and $obj.PSObject.Properties.Name -contains 'value') {
                $sid = "$($obj.value)"
            }
        } catch {}
        if (-not $sid) { $sid = $r2.Text.Trim().Trim('"') }
        $Script:Session = $sid
        Log-Info ("API style: v6  (vCenter 6.5 / 6.7, endpoint /rest/...)")
        Log-Warn ("v6 模式: 部分章节 (cert / backup / access / VMware Tools / Deployment / Update) 在 6.5 REST 未暴露, 报告将标注 N/A")
        return $sid
    }

    # 两种都失败 → 给精确诊断
    $best = if ($r2.Code -gt 0) { $r2 } else { $r }
    $hint = ''
    if ($best.Code -eq 401)     { $hint = "  >> 401: 用户名 / 密码错误, 或账号已锁定 (vCenter SSO 默认 5 次失败锁 5 分钟)" }
    elseif ($best.Code -eq 403) { $hint = "  >> 403: 账号无 API 权限" }
    elseif ($best.Code -ge 500) { $hint = "  >> $($best.Code): vCenter 服务端错误 (sts-idmd / vapi-endpoint 可能异常), 已重试仍失败, 建议: ssh root@$VCenter; service-control --status; service-control --restart vmware-vapi-endpoint" }
    elseif ($best.Code -eq 0)   { $hint = "  >> 网络不可达: 检查 $VCenter 路由 / 防火墙 / TLS" }
    elseif ($best.Code -eq 404) { $hint = "  >> 404: 该 vCenter REST API 端点都不可用, 可能 < 6.5 版本, 此版本只能用 PowerCLI" }
    $detail = if ($best.Text) { " body=$($best.Text.Substring(0, [math]::Min(200, $best.Text.Length)))" } else { '' }
    throw "vCenter 登录失败 (v8=HTTP$($r.Code) / v6=HTTP$($r2.Code)): $($best.Error)$detail`n$hint"
}
function Logout-VC {
    if (-not $Script:Session) { return }
    $p = if ($Script:ApiStyle -eq 'v6') { '/rest/com/vmware/cis/session' } else { '/api/session' }
    Invoke-VCRaw -Path $p -Method 'DELETE' | Out-Null
    $Script:Session = $null
}

# ============================================================================
#  辅助工具
# ============================================================================
function Format-Bytes([double]$bytes) {
    if ($bytes -eq $null -or $bytes -le 0) { return '0 B' }
    $u = @('B','KB','MB','GB','TB','PB'); $i = 0
    while ($bytes -ge 1024 -and $i -lt $u.Length-1) { $bytes /= 1024; $i++ }
    return ('{0:N2} {1}' -f $bytes, $u[$i])
}
function Format-DaysFromSec([double]$sec) {
    if (-not $sec) { return 'N/A' }
    $d = [math]::Floor($sec / 86400)
    $h = [math]::Floor(($sec % 86400) / 3600)
    return "{0} 天 {1} 小时" -f $d, $h
}
function Html-Encode([object]$s) {
    if ($s -eq $null) { return '' }
    $str = "$s"
    return ([System.Net.WebUtility]::HtmlEncode($str))
}
function Badge {
    param([string]$Text, [string]$Kind='gray')
    return "<span class='badge badge-$Kind'>$([System.Net.WebUtility]::HtmlEncode($Text))</span>"
}
function Health-Badge([string]$state) {
    $s = "$state".ToLower()
    $kind = switch ($s) {
        'green'  { 'green' }
        'yellow' { 'amber' }
        'orange' { 'amber' }
        'red'    { 'red'   }
        'gray'   { 'gray'  }
        default  { 'gray'  }
    }
    return Badge -Text $s.ToUpper() -Kind $kind
}

# ============================================================================
#  PowerCLI 回退层 (v1.1)
#  补 REST 在 8.0 拿不到的: 快照大小 / 当前 Alarm / 单 host 实时 CPU/Mem
# ============================================================================
function Try-LoadPowerCLI {
    if ($SkipPowerCLI) { return $false }
    $mod = Get-Module -ListAvailable VMware.VimAutomation.Core -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $mod) {
        if ($UsePowerCLI) {
            Log-Warn 'PowerCLI 未安装。先运行: Install-Module VMware.PowerCLI -Scope CurrentUser -Force'
        } else {
            Log-Info 'PowerCLI 未安装,跳过快照/Alarm/Host 实时采集 (如需启用: Install-Module VMware.PowerCLI)'
        }
        return $false
    }
    try {
        Import-Module VMware.VimAutomation.Core -ErrorAction Stop -WarningAction SilentlyContinue
        # 关掉证书校验 + CEIP 询问
        Set-PowerCLIConfiguration -InvalidCertificateAction Ignore   -Confirm:$false -Scope Session | Out-Null
        Set-PowerCLIConfiguration -ParticipateInCeip       $false    -Confirm:$false -Scope Session | Out-Null
        Set-PowerCLIConfiguration -DefaultVIServerMode     Single    -Confirm:$false -Scope Session | Out-Null
        Log-Info ("PowerCLI 已加载 (VimAutomation.Core $($mod.Version))")
        return $true
    } catch {
        Log-Warn ("PowerCLI 加载失败: $($_.Exception.Message)")
        return $false
    }
}

function Connect-PowerCLI {
    try {
        $sec = ConvertTo-SecureString $Password -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential($Username, $sec)
        $srv = Connect-VIServer -Server $VCenter -Credential $cred -ErrorAction Stop -WarningAction SilentlyContinue
        if ($srv) {
            Log-Info ("PowerCLI 已连接 vCenter (build $($srv.Build))")
            return $true
        }
    } catch {
        Log-Warn ("PowerCLI Connect-VIServer 失败: $($_.Exception.Message)")
    }
    return $false
}

function Disconnect-PowerCLI {
    try { Disconnect-VIServer -Server $VCenter -Confirm:$false -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
}

function Collect-PowerCLI {
    if (-not (Try-LoadPowerCLI))   { return }
    if (-not (Connect-PowerCLI))   { return }
    $Script:PCLIEnabled = $true

    # ---- 1. 单 Host 实时 (NumCpu/Cpu/Mem/Uptime/Build/Vendor/Model) ----
    Log-Step '[PCLI] 单 Host 实时数据'
    try {
        $hostList = New-Object System.Collections.Generic.List[psobject]
        foreach ($h in (Get-VMHost -ErrorAction Stop)) {
            $cpuTotalMhz = [int]$h.CpuTotalMhz
            $cpuUsedMhz  = [int]$h.CpuUsageMhz
            $memTotalGB  = [math]::Round($h.MemoryTotalGB, 1)
            $memUsedGB   = [math]::Round($h.MemoryUsageGB, 1)
            $cpuPct      = if ($cpuTotalMhz -gt 0) { [math]::Round(100 * $cpuUsedMhz / $cpuTotalMhz, 1) } else { 0 }
            $memPct      = if ($memTotalGB  -gt 0) { [math]::Round(100 * $memUsedGB  / $memTotalGB,  1) } else { 0 }
            $uptimeDays  = 0
            try {
                $bt = $h.ExtensionData.Runtime.BootTime
                if ($bt) { $uptimeDays = [math]::Round(((Get-Date).ToUniversalTime() - $bt.ToUniversalTime()).TotalDays, 1) }
            } catch {}
            $hostList.Add([pscustomobject]@{
                Name        = $h.Name
                ConnState   = "$($h.ConnectionState)"
                PowerState  = "$($h.PowerState)"
                NumCpu      = [int]$h.NumCpu
                CpuUsedMhz  = $cpuUsedMhz
                CpuTotalMhz = $cpuTotalMhz
                CpuPct      = $cpuPct
                MemUsedGB   = $memUsedGB
                MemTotalGB  = $memTotalGB
                MemPct      = $memPct
                UptimeDays  = $uptimeDays
                Version     = "$($h.Version)"
                Build       = "$($h.Build)"
                Vendor      = "$($h.Manufacturer)"
                Model       = "$($h.Model)"
            })
        }
        $Data.PCLI_Hosts = @($hostList)
        Log-Info ("拉到 {0} 台 ESXi 实时数据" -f $hostList.Count)
    } catch {
        Log-Warn ("Get-VMHost 失败: $($_.Exception.Message)")
        $Data.PCLI_Hosts = @()
    }

    # ---- 2. VM 快照 (Get-Snapshot per VM 太慢, 一次拉全部 VM 的) ----
    Log-Step '[PCLI] VM 快照'
    try {
        $allVMs = Get-VM -ErrorAction Stop
        $snaps  = New-Object System.Collections.Generic.List[psobject]
        $bySrc  = @{}   # vmName -> snap count
        foreach ($vm in $allVMs) {
            $ss = @($vm | Get-Snapshot -ErrorAction SilentlyContinue)
            if ($ss.Count -eq 0) { continue }
            $bySrc[$vm.Name] = $ss.Count
            foreach ($s in $ss) {
                $created = $null
                try { $created = [DateTime]$s.Created } catch {}
                $ageDays = if ($created) { [math]::Round(((Get-Date) - $created).TotalDays, 1) } else { $null }
                $snaps.Add([pscustomobject]@{
                    VM        = $vm.Name
                    Snapshot  = "$($s.Name)"
                    Created   = if ($created) { $created.ToString('yyyy-MM-dd HH:mm') } else { '—' }
                    AgeDays   = $ageDays
                    SizeGB    = [math]::Round([double]$s.SizeGB, 2)
                    PowerState= "$($vm.PowerState)"
                    Desc      = "$($s.Description)"
                })
            }
        }
        $totalGB = ($snaps | Measure-Object -Property SizeGB -Sum).Sum
        if (-not $totalGB) { $totalGB = 0 }
        $Data.PCLI_Snapshots = [pscustomobject]@{
            All        = @($snaps)
            VmCount    = $bySrc.Count
            Count      = $snaps.Count
            TotalGB    = [math]::Round($totalGB, 2)
            OldestDays = if ($snaps.Count -gt 0) { ($snaps | Measure-Object -Property AgeDays -Maximum).Maximum } else { 0 }
            MaxChain   = if ($bySrc.Count -gt 0) { ($bySrc.Values | Measure-Object -Maximum).Maximum } else { 0 }
            Top10      = @($snaps | Sort-Object -Property SizeGB -Descending | Select-Object -First 10)
        }
        Log-Info ("拉到 {0} 个快照, 涉及 {1} 台 VM, 共 {2} GB" -f $snaps.Count, $bySrc.Count, $Data.PCLI_Snapshots.TotalGB)
    } catch {
        Log-Warn ("Get-Snapshot 失败: $($_.Exception.Message)")
        $Data.PCLI_Snapshots = $null
    }

    # ---- 3. 当前 Triggered Alarms ----
    Log-Step '[PCLI] 当前告警 (Triggered Alarms)'
    try {
        $triggered = New-Object System.Collections.Generic.List[psobject]
        # 遍历所有有 TriggeredAlarmState 的对象
        $entities = @()
        $entities += @(Get-Datacenter -ErrorAction SilentlyContinue)
        $entities += @(Get-Cluster    -ErrorAction SilentlyContinue)
        $entities += @(Get-VMHost     -ErrorAction SilentlyContinue)
        $entities += @(Get-Datastore  -ErrorAction SilentlyContinue)
        $entities += @(Get-VM         -ErrorAction SilentlyContinue)
        foreach ($e in $entities) {
            $tas = $null
            try { $tas = $e.ExtensionData.TriggeredAlarmState } catch {}
            if (-not $tas) { continue }
            foreach ($t in @($tas)) {
                $alarmName = '—'
                try {
                    $alm = Get-View -Id $t.Alarm -ErrorAction SilentlyContinue
                    if ($alm) { $alarmName = "$($alm.Info.Name)" }
                } catch {}
                $time = $null
                try { $time = [DateTime]$t.Time } catch {}
                $triggered.Add([pscustomobject]@{
                    Entity    = "$($e.Name)"
                    EntityType= ($e.GetType().Name -replace 'Impl$','')
                    Alarm     = $alarmName
                    Status    = "$($t.OverallStatus)"
                    Time      = if ($time) { $time.ToString('yyyy-MM-dd HH:mm') } else { '—' }
                    AgeHours  = if ($time) { [math]::Round(((Get-Date) - $time).TotalHours, 1) } else { $null }
                    Acked     = [bool]$t.Acknowledged
                })
            }
        }
        $Data.PCLI_Alarms = @($triggered | Sort-Object @{e='Status';desc=$true}, Time -Descending)
        $redN  = @($triggered | Where-Object { "$($_.Status)" -eq 'red'    }).Count
        $yelN  = @($triggered | Where-Object { "$($_.Status)" -eq 'yellow' }).Count
        Log-Info ("当前告警 {0} 条 (red={1} / yellow={2})" -f $triggered.Count, $redN, $yelN)
    } catch {
        Log-Warn ("Triggered Alarms 拉取失败: $($_.Exception.Message)")
        $Data.PCLI_Alarms = @()
    }

    Disconnect-PowerCLI
}

# ============================================================================
#  数据采集
# ============================================================================
$Data = [ordered]@{}

function Collect-All {

    Log-Step 'Login vCenter'
    [void](Login-VC)

    Log-Step '1. 系统版本与运行时间'
    $Data.SysVersion   = Invoke-VC '/api/appliance/system/version'
    $Data.SysUptime    = Invoke-VC '/api/appliance/system/uptime'
    $Data.SysTime      = Invoke-VC '/api/appliance/system/time'
    $Data.Deployment   = Invoke-VC '/api/vcenter/deployment'
    $Data.Update       = Invoke-VC '/api/appliance/update'

    Log-Step '2. Appliance 健康'
    $Data.Health = [ordered]@{
        System          = Invoke-VC '/api/appliance/health/system'
        Storage         = Invoke-VC '/api/appliance/health/storage'
        Memory          = Invoke-VC '/api/appliance/health/mem'
        Swap            = Invoke-VC '/api/appliance/health/swap'
        Load            = Invoke-VC '/api/appliance/health/load'
        DatabaseStorage = Invoke-VC '/api/appliance/health/database-storage'
        Applmgmt        = Invoke-VC '/api/appliance/health/applmgmt'
        SoftwarePackages= Invoke-VC '/api/appliance/health/software-packages'
        LastCheck       = Invoke-VC '/api/appliance/health/system/lastcheck'
    }

    Log-Step '3. 网络 / DNS / NTP / 时间同步'
    $Data.Networking    = Invoke-VC '/api/appliance/networking'
    $Data.NetInterfaces = Invoke-VC '/api/appliance/networking/interfaces'
    $Data.NTP           = Invoke-VC '/api/appliance/ntp'
    $Data.Timesync      = Invoke-VC '/api/appliance/timesync'

    Log-Step '4. 访问入口 (SSH/Shell/DCUI/Console)'
    $Data.Access = [ordered]@{
        SSH        = Invoke-VC '/api/appliance/access/ssh'
        Shell      = Invoke-VC '/api/appliance/access/shell'
        DCUI       = Invoke-VC '/api/appliance/access/dcui'
        ConsoleCli = Invoke-VC '/api/appliance/access/consolecli'
    }

    Log-Step '5. vCenter 证书'
    $Data.CertTLS     = Invoke-VC '/api/vcenter/certificate-management/vcenter/tls'

    Log-Step '6. 备份策略与历史'
    $Data.BackupJobs      = Invoke-VC '/api/appliance/recovery/backup/job'
    $Data.BackupSchedules = Invoke-VC '/api/appliance/recovery/backup/schedules'

    Log-Step '7. Datacenter / Cluster / Folder / 资源池'
    $Data.Datacenters = Invoke-VC '/api/vcenter/datacenter'
    $Data.Clusters    = Invoke-VC '/api/vcenter/cluster'
    $Data.Folders     = Invoke-VC '/api/vcenter/folder'
    $Data.ResPools    = Invoke-VC '/api/vcenter/resource-pool'

    Log-Step '8. ESXi 主机'
    $Data.Hosts = Invoke-VC '/api/vcenter/host'

    Log-Step '9. Datastore'
    $Data.Datastores = Invoke-VC '/api/vcenter/datastore'

    Log-Step '10. 网络 (Portgroup / dvSwitch)'
    $Data.Networks = Invoke-VC '/api/vcenter/network'

    Log-Step '11. 虚拟机清单'
    $Data.VMs = Invoke-VC '/api/vcenter/vm'
    Log-Info ("拉到 {0} 台 VM" -f ($Data.VMs | Measure-Object).Count)

    Log-Step '12. Appliance 服务列表'
    $Data.Services = Invoke-VC '/api/appliance/services'

    Log-Step '13. VMware Tools 抽样'
    $Data.ToolsSample = @()
    if (-not $SkipToolsSample -and $Data.VMs) {
        $onVMs = @($Data.VMs | Where-Object { $_.power_state -eq 'POWERED_ON' })
        $take  = [math]::Min($ToolsSampleSize, $onVMs.Count)
        if ($take -gt 0) {
            $pick = $onVMs | Select-Object -First $take
            $i = 0; $n = $pick.Count
            foreach ($v in $pick) {
                $i++
                if (-not $Quiet) {
                    Write-Host ("         tools sample [{0,2}/{1}] {2}" -f $i, $n, $v.name) -ForegroundColor DarkGray
                }
                $t = Invoke-VC "/api/vcenter/vm/$($v.vm)/tools"
                $Data.ToolsSample += [pscustomobject]@{
                    VM        = $v.name
                    VMId      = $v.vm
                    RunState  = if ($t) { $t.run_state }       else { 'N/A' }
                    Version   = if ($t) { $t.version }         else { 'N/A' }
                    Status    = if ($t) { $t.version_status }  else { 'N/A' }
                    Install   = if ($t) { $t.install_type }    else { 'N/A' }
                    Policy    = if ($t) { $t.upgrade_policy }  else { 'N/A' }
                }
            }
        }
    } else { Log-Info '已跳过 (SkipToolsSample 或无 VM)' }

    Log-Step '14. 关闭 vCenter session'
    Logout-VC

    # ---- 数据采集汇总诊断 ----
    if (-not $Quiet) {
        Write-Host ''
        Write-Host '  ─── 数据采集汇总 ──────────────────────────' -ForegroundColor DarkCyan
        $verLine = 'null'
        if ($Data.SysVersion -and $Data.SysVersion.version) { $verLine = "$($Data.SysVersion.version)  build $($Data.SysVersion.build)" }
        Write-Host ("    API style  : {0}" -f $Script:ApiStyle) -ForegroundColor Gray
        Write-Host ("    Version    : {0}" -f $verLine) -ForegroundColor Gray
        foreach ($k in @('Datacenters','Clusters','Folders','ResPools','Hosts','Datastores','Networks','VMs','Services','BackupJobs','NetInterfaces','ToolsSample')) {
            $v = $Data.$k
            $c = 'null'
            if ($null -ne $v) { $c = ($v | Measure-Object).Count }
            Write-Host ("    {0,-13}: {1}" -f $k, $c) -ForegroundColor Gray
        }
        Write-Host '  ──────────────────────────────────────────' -ForegroundColor DarkCyan
        if ($Script:DebugLog) { Write-Host ("  Debug log  : {0}" -f $Script:DebugLog) -ForegroundColor DarkYellow }
        Write-Host ''
    }
}

# ============================================================================
#  评估与告警生成
# ============================================================================
function Eval-Findings {
    $issues = New-Object System.Collections.Generic.List[psobject]
    function Add-Issue([string]$Severity, [string]$Area, [string]$Title, [string]$Detail) {
        $issues.Add([pscustomobject]@{
            Severity = $Severity   # critical / warn / info
            Area     = $Area
            Title    = $Title
            Detail   = $Detail
        })
    }

    # health
    foreach ($k in @('System','Storage','Memory','Swap','Load','DatabaseStorage','Applmgmt','SoftwarePackages')) {
        $v = $Data.Health.$k
        if (-not $v) { continue }
        $s = "$v".ToLower()
        if ($s -eq 'yellow' -or $s -eq 'orange') { Add-Issue 'warn'     "健康" ("Appliance Health: $k = $s") "建议进入 VAMI (5480) 查看具体告警明细" }
        elseif ($s -eq 'red')                    { Add-Issue 'critical' "健康" ("Appliance Health: $k = RED") "立即处理, 可能影响 vCenter 可用性" }
    }

    # ntp / timesync
    $ntpCount = @($Data.NTP).Count
    $tsState  = "$($Data.Timesync)".ToUpper()
    if ($ntpCount -eq 0)            { Add-Issue 'warn'     '时间同步' 'NTP 服务器列表为空' '建议配置 ≥ 2 个 NTP 源, 避免 vCenter 与 ESXi 时钟漂移导致 SSO / 证书异常' }
    if ($tsState -ne 'NTP')         { Add-Issue 'warn'     '时间同步' "Timesync 模式 = $tsState" '推荐 NTP 模式; HOST/DISABLED 在跨集群迁移或主机时间漂移时可能引发故障' }

    # DNS hostname / DNS server
    if ($Data.Networking) {
        $hn = "$($Data.Networking.dns.hostname)"
        if ($hn -eq 'localhost' -or [string]::IsNullOrWhiteSpace($hn)) {
            Add-Issue 'warn' 'DNS' "DNS hostname 仍是 '$hn'" '建议为 VCSA 配置正式 FQDN, 否则证书 / SSO / 备份恢复均可能异常'
        }
        $srvs = @($Data.Networking.dns.servers)
        if ($srvs.Count -lt 2) {
            Add-Issue 'info' 'DNS' ("DNS 服务器数 = " + $srvs.Count) '生产环境建议 ≥ 2 个 DNS 服务器避免单点'
        }
    }

    # access
    if ("$($Data.Access.SSH)"  -eq 'True') { Add-Issue 'warn' '访问' 'SSH 已启用'  '生产环境建议默认关闭, 排错时临时打开后及时关闭' }
    if ("$($Data.Access.DCUI)" -eq 'True') { Add-Issue 'info' '访问' 'DCUI 已启用' 'DCUI 用于本地控制台紧急恢复, 远程运维场景可关闭' }

    # cert
    if ($Data.CertTLS) {
        try {
            $vt = [DateTime]::Parse($Data.CertTLS.valid_to)
            $days = ($vt - (Get-Date).ToUniversalTime()).TotalDays
            if ($days -lt 30)       { Add-Issue 'critical' '证书' ('vCenter TLS 证书将于 {0:N0} 天后过期' -f $days) '需立即续签 / 替换' }
            elseif ($days -lt 90)   { Add-Issue 'warn'     '证书' ('vCenter TLS 证书将于 {0:N0} 天后过期' -f $days) '建议尽快续签' }
            if ($Data.CertTLS.issuer_dn -match 'localhost') {
                Add-Issue 'info' '证书' '当前 vCenter TLS 证书为自签 (VMSCA / localhost 默认 CA)' '生产可考虑替换为企业 CA 或公共 CA 签发的证书'
            }
        } catch {}
    }

    # backup
    $bjCount = @($Data.BackupJobs).Count
    $bsCount = if ($Data.BackupSchedules) { ([pscustomobject]$Data.BackupSchedules).PSObject.Properties.Name.Count } else { 0 }
    if ($bjCount -eq 0) { Add-Issue 'warn' '备份' 'VAMI 备份历史为空' '建议在 VAMI (5480) → Backup 配置 SFTP/NFS/SMB 周期备份, 灾备无备份风险大' }
    if ($bsCount -eq 0) { Add-Issue 'warn' '备份' '未配置任何备份计划'   '同上' }

    # cluster HA/DRS
    if ($Data.Clusters) {
        foreach ($c in @($Data.Clusters)) {
            $hostCount = @($Data.Hosts).Count
            if (-not $c.ha_enabled  -and $hostCount -ge 2) { Add-Issue 'warn' '集群' "Cluster '$($c.name)' 未启用 HA"  '≥ 2 节点集群推荐启用 HA 保障 VM 故障迁移' }
            if (-not $c.drs_enabled -and $hostCount -ge 2) { Add-Issue 'info' '集群' "Cluster '$($c.name)' 未启用 DRS" '推荐启用 DRS 自动均衡负载 (FullyAutomated 或 PartiallyAutomated)' }
        }
    }

    # datastore 使用率
    if ($Data.Datastores) {
        foreach ($ds in @($Data.Datastores)) {
            if ($ds.capacity -gt 0) {
                $used = ($ds.capacity - $ds.free_space) / $ds.capacity
                if ($used -ge 0.90)      { Add-Issue 'critical' 'Datastore' ("$($ds.name) 使用率 {0:P1}" -f $used) '已超 90%, 强烈建议立即清理或扩容; thin VM 可能因写满导致 VM 停机' }
                elseif ($used -ge 0.80)  { Add-Issue 'warn'     'Datastore' ("$($ds.name) 使用率 {0:P1}" -f $used) '已超 80%, 建议清理快照 / 旧 VM 或扩容' }
                elseif ($used -ge 0.70)  { Add-Issue 'info'     'Datastore' ("$($ds.name) 使用率 {0:P1}" -f $used) '使用率较高, 关注后续增长' }
            }
        }
    }

    # tools sample
    if ($Data.ToolsSample.Count -gt 0) {
        $oldTools = @($Data.ToolsSample | Where-Object { $_.Status -in @('SUPPORTED_OLD','UNMANAGED_OLD') })
        if ($oldTools.Count -gt 0) {
            Add-Issue 'info' 'VMware Tools' ("$($oldTools.Count) / $($Data.ToolsSample.Count) 抽样 VM Tools 版本偏旧") '建议升级 VMware Tools 至最新 LTS; open-vm-tools 可走 yum/apt 更新'
        }
    }

    # update
    if ($Data.Update -and "$($Data.Update.state)" -notin @('UP_TO_DATE','')) {
        Add-Issue 'info' '补丁' ("vCenter 更新状态: $($Data.Update.state)") '可在 VAMI (5480) → Update 查看可用补丁'
    }

    # ---------- v1.1: PowerCLI 维度 ----------
    if ($Script:PCLIEnabled) {
        # 单 host 实时
        if ($Data.PCLI_Hosts) {
            foreach ($h in @($Data.PCLI_Hosts)) {
                if ($h.MemPct -ge 90) { Add-Issue 'critical' 'ESXi 主机' ("$($h.Name) 内存使用率 {0}%" -f $h.MemPct) '已超 90%, VM swap/balloon 风险大, 建议立即均衡负载或扩容内存' }
                elseif ($h.MemPct -ge 80) { Add-Issue 'warn' 'ESXi 主机' ("$($h.Name) 内存使用率 {0}%" -f $h.MemPct) '已超 80%, 建议关注内存增长趋势' }
                if ($h.CpuPct -ge 85) { Add-Issue 'warn' 'ESXi 主机' ("$($h.Name) CPU 使用率 {0}%" -f $h.CpuPct) '采样瞬时值, 持续高位需检查 vCPU 超分配或负载倾斜' }
                if ("$($h.ConnState)" -in @('Disconnected','NotResponding')) {
                    Add-Issue 'critical' 'ESXi 主机' ("$($h.Name) 连接状态 = $($h.ConnState)") '主机异常脱管, 立即排查网络 / vpxa / hostd'
                }
                if ($h.UptimeDays -gt 365) {
                    Add-Issue 'info' 'ESXi 主机' ("$($h.Name) uptime {0} 天" -f $h.UptimeDays) '运行 > 1 年, 建议安排窗口期重启或打补丁'
                }
            }
        }
        # 快照健康
        if ($Data.PCLI_Snapshots) {
            $sn = $Data.PCLI_Snapshots
            if ($sn.OldestDays -gt 90)  { Add-Issue 'critical' '快照' ("最老快照已存在 {0} 天" -f $sn.OldestDays) '快照 > 90 天会拖累性能并放大磁盘占用, 立即合并或删除' }
            elseif ($sn.OldestDays -gt 30) { Add-Issue 'warn' '快照' ("最老快照已存在 {0} 天" -f $sn.OldestDays) '快照 > 30 天建议清理; 长期快照非备份机制' }
            if ($sn.MaxChain -gt 3)     { Add-Issue 'warn' '快照' ("单 VM 最长快照链 = $($sn.MaxChain)") '快照链 > 3 层导致 I/O 严重劣化, 建议合并' }
            if ($sn.TotalGB -gt 1024)   { Add-Issue 'info' '快照' ("快照总占用 {0} GB" -f $sn.TotalGB) '考虑定期清理或转 RBP 备份' }
            if ($sn.Count -eq 0)        { Add-Issue 'info' '快照' '当前无任何快照' '快照为空属正常运维状态' }
        }
        # 告警
        if ($Data.PCLI_Alarms) {
            $red = @($Data.PCLI_Alarms | Where-Object { "$($_.Status)" -eq 'red' })
            $yel = @($Data.PCLI_Alarms | Where-Object { "$($_.Status)" -eq 'yellow' })
            if ($red.Count -gt 0) { Add-Issue 'critical' 'Alarm' ("当前 RED 告警 {0} 条" -f $red.Count) '立即在 vSphere Client → Monitor → Triggered Alarms 中处理' }
            if ($yel.Count -gt 5) { Add-Issue 'warn'     'Alarm' ("当前 YELLOW 告警 {0} 条" -f $yel.Count) '黄色告警偏多, 建议梳理并设置告警抑制 / 修复' }
        }
    }

    return ,$issues
}

# ============================================================================
#  HTML 渲染
# ============================================================================
function Render-Report {

    # ---- 预处理 ----
    $now    = Get-Date
    $sv     = $Data.SysVersion
    $up     = [double]([string]$Data.SysUptime -replace '[^\d.eE+\-]','')
    if (-not $up) { $up = 0 }
    $time   = $Data.SysTime
    $hosts  = @($Data.Hosts)
    $clu    = @($Data.Clusters)
    $ds     = @($Data.Datastores)
    $vms    = @($Data.VMs)
    $nets   = @($Data.Networks)
    $svcs   = $Data.Services
    $dcs    = @($Data.Datacenters)

    $vmOn   = @($vms | Where-Object { $_.power_state -eq 'POWERED_ON'  })
    $vmOff  = @($vms | Where-Object { $_.power_state -eq 'POWERED_OFF' })
    $vmSusp = @($vms | Where-Object { $_.power_state -eq 'SUSPENDED'   })
    $cpuSum = ($vms | Measure-Object -Property cpu_count       -Sum).Sum
    $memSum = ($vms | Measure-Object -Property memory_size_MiB -Sum).Sum
    $cpuOn  = ($vmOn | Measure-Object -Property cpu_count       -Sum).Sum
    $memOn  = ($vmOn | Measure-Object -Property memory_size_MiB -Sum).Sum
    $dsCap  = ($ds  | Measure-Object -Property capacity   -Sum).Sum
    $dsFree = ($ds  | Measure-Object -Property free_space -Sum).Sum

    # service 统计
    $svcStarted = 0; $svcStopped = 0; $svcOther = 0; $svcRows = @()
    if ($svcs) {
        $names = $svcs.PSObject.Properties.Name
        foreach ($n in $names) {
            $s = $svcs.$n
            switch ("$($s.state)") {
                'STARTED' { $svcStarted++ }
                'STOPPED' { $svcStopped++ }
                default   { $svcOther++   }
            }
        }
        # 核心 vCenter 服务过滤
        $coreList = @(
            'vmware-vpxd','vmware-vpostgres','vmware-vapi-endpoint','vmware-sts-idmd',
            'vmware-cis-license','vmware-content-library','vmware-eam','vmware-perfcharts',
            'vmware-rbd-watchdog','vmware-sps','vmware-vsm','vmware-updatemgr',
            'vmware-vsan-health','vmware-vmcam','vmware-analytics','vmware-postgres-archiver',
            'vmware-stsd','vmware-trustmanagement','vsphere-ui','vsphere-client'
        )
        foreach ($n in $coreList) {
            if ($svcs.PSObject.Properties.Name -contains $n) {
                $svcRows += [pscustomobject]@{ Name=$n; State=$svcs.$n.state; Desc=$svcs.$n.description }
            }
        }
    }
    $svcTotal = $svcStarted + $svcStopped + $svcOther

    # 证书剩余天数
    $certDays = 'N/A'; $certKind = 'gray'
    if ($Data.CertTLS) {
        try {
            $vt = [DateTime]::Parse($Data.CertTLS.valid_to)
            $cd = [int]($vt - (Get-Date).ToUniversalTime()).TotalDays
            $certDays = $cd
            if ($cd -lt 30)       { $certKind = 'red' }
            elseif ($cd -lt 90)   { $certKind = 'amber' }
            else                  { $certKind = 'green' }
        } catch {}
    }

    # ---- 评估 ----
    $findings = Eval-Findings
    $critCnt = @($findings | Where-Object Severity -eq 'critical').Count
    $warnCnt = @($findings | Where-Object Severity -eq 'warn'    ).Count
    $infoCnt = @($findings | Where-Object Severity -eq 'info'    ).Count

    $overall = if     ($critCnt -gt 0) { '严重' }
               elseif ($warnCnt -gt 0) { '警告' }
               else                    { '正常' }
    $overallKind = if     ($critCnt -gt 0) { 'red' }
                   elseif ($warnCnt -gt 0) { 'amber' }
                   else                    { 'green' }

    # ============================================================
    #  HTML 模板
    # ============================================================
    # ---- 主题色板 (v1.2): 4 套全部内嵌, 通过 [data-theme] 切换 ----
    $palLight = @'
  --bg:#f5f7fa;--bg-side:#ffffff;--bg-card:#ffffff;--bg-card-2:#f9fafb;
  --fg:#0f172a;--fg-mute:#475569;--fg-dim:#94a3b8;
  --border:#e2e8f0;--border-strong:#cbd5e1;
  --accent:#1565c0;--accent-2:#1976d2;
  --green:#16a34a;--amber:#d97706;--red:#dc2626;--gray:#64748b;--blue:#1565c0;
  --green-bg:rgba(22,163,74,.10);--amber-bg:rgba(217,119,6,.10);
  --red-bg:rgba(220,38,38,.10);--gray-bg:rgba(100,116,139,.10);--blue-bg:rgba(21,101,192,.10);
'@
    $palDark = @'
  --bg:#0f1419;--bg-side:#0a0d12;--bg-card:#161b22;--bg-card-2:#1c2330;
  --fg:#d4d9e0;--fg-mute:#7a8290;--fg-dim:#5a6270;
  --border:#2a313c;--border-strong:#3a414c;
  --accent:#58a6ff;--accent-2:#79b8ff;
  --green:#3fb950;--amber:#d29922;--red:#f85149;--gray:#6e7681;--blue:#58a6ff;
  --green-bg:rgba(63,185,80,.12);--amber-bg:rgba(210,153,34,.12);
  --red-bg:rgba(248,81,73,.12);--gray-bg:rgba(110,118,129,.15);--blue-bg:rgba(88,166,255,.12);
'@
    $palMinimal = @'
  --bg:#fafafa;--bg-side:#f4f4f5;--bg-card:#ffffff;--bg-card-2:#f7f7f8;
  --fg:#18181b;--fg-mute:#52525b;--fg-dim:#a1a1aa;
  --border:#e4e4e7;--border-strong:#d4d4d8;
  --accent:#27272a;--accent-2:#3f3f46;
  --green:#16a34a;--amber:#a16207;--red:#b91c1c;--gray:#71717a;--blue:#1d4ed8;
  --green-bg:rgba(22,163,74,.08);--amber-bg:rgba(161,98,7,.08);
  --red-bg:rgba(185,28,28,.08);--gray-bg:rgba(113,113,122,.08);--blue-bg:rgba(29,78,216,.08);
'@
    $palAmber = @'
  --bg:#fdfaf3;--bg-side:#f7f0e2;--bg-card:#ffffff;--bg-card-2:#fbf4e6;
  --fg:#3f2a14;--fg-mute:#78623d;--fg-dim:#a89378;
  --border:#e7d9bf;--border-strong:#d4c19c;
  --accent:#b45309;--accent-2:#d97706;
  --green:#15803d;--amber:#b45309;--red:#b91c1c;--gray:#8a7656;--blue:#1565c0;
  --green-bg:rgba(21,128,61,.10);--amber-bg:rgba(180,83,9,.14);
  --red-bg:rgba(185,28,28,.10);--gray-bg:rgba(138,118,86,.12);--blue-bg:rgba(21,101,192,.10);
'@
    # 可选 -AccentColor 覆盖所有主题的 accent
    $accentOverride = ''
    if ($AccentColor) {
        $accentOverride = "`n:root,[data-theme]{--accent:$AccentColor !important;--accent-2:$AccentColor !important}"
    }

    $css = @"
:root,[data-theme="light"]{
$palLight  --mono:'JetBrains Mono','Fira Code','SF Mono','Consolas','Cascadia Mono',monospace;
  --sans:-apple-system,BlinkMacSystemFont,'Segoe UI','Microsoft YaHei','PingFang SC',Roboto,sans-serif;
}
[data-theme="dark"]{
$palDark}
[data-theme="minimal"]{
$palMinimal}
[data-theme="amber"]{
$palAmber}$accentOverride
"@ + @'
*{box-sizing:border-box}
html,body{margin:0;padding:0;background:var(--bg);color:var(--fg);font-family:var(--sans);font-size:14px;line-height:1.55;-webkit-font-smoothing:antialiased}
a{color:var(--accent);text-decoration:none}
a:hover{text-decoration:underline}
code,pre,.mono{font-family:var(--mono)}
hr{border:none;border-top:1px solid var(--border);margin:32px 0}

/* sidebar */
.sidebar{position:fixed;top:0;left:0;width:260px;height:100vh;background:var(--bg-side);border-right:1px solid var(--border);overflow-y:auto;padding:20px 0;z-index:10}
.sidebar .brand{padding:0 20px 20px;border-bottom:1px solid var(--border);margin-bottom:12px}
.sidebar .brand .logo{display:flex;align-items:center;gap:10px}
.sidebar .brand .logo svg{width:28px;height:28px;flex-shrink:0}
.sidebar .brand .name{font-weight:600;font-size:15px;color:var(--fg)}
.sidebar .brand .sub{font-size:11px;color:var(--fg-mute);margin-top:2px;font-family:var(--mono)}
.sidebar nav{padding:0 8px}
.sidebar nav a{display:flex;align-items:center;gap:8px;padding:7px 12px;color:var(--fg-mute);font-size:13px;border-radius:5px;border-left:2px solid transparent;margin:1px 0;transition:none}
.sidebar nav a .num{font-family:var(--mono);font-size:11px;color:var(--fg-dim);min-width:18px}
.sidebar nav a:hover{background:var(--bg-card);color:var(--fg);text-decoration:none}
.sidebar nav a.active{background:var(--bg-card);color:var(--accent);border-left-color:var(--accent)}
.sidebar nav a.active .num{color:var(--accent)}

/* main */
.main{margin-left:260px;max-width:1280px;padding:32px 40px 80px}

/* banner */
.banner{background:var(--bg-card);border:1px solid var(--border);border-left:3px solid var(--accent);border-radius:8px;padding:24px 28px;margin-bottom:32px}
.banner .tag{display:inline-block;background:var(--accent-2);color:#001020;font-size:11px;font-weight:700;padding:3px 10px;border-radius:12px;letter-spacing:.4px}
.banner h1{margin:8px 0 6px;font-size:22px;font-weight:600;letter-spacing:-.2px}
.banner .sub{color:var(--fg-mute);font-size:13px;font-family:var(--mono);margin-bottom:18px}
.banner .meta{display:grid;grid-template-columns:repeat(3,1fr);gap:12px 32px;padding-top:16px;border-top:1px solid var(--border)}
.banner .meta dt{font-size:11px;color:var(--fg-mute);text-transform:uppercase;letter-spacing:.5px;margin-bottom:3px}
.banner .meta dd{margin:0;font-size:14px;font-family:var(--mono);color:var(--fg);word-break:break-all}

/* section */
section{margin-bottom:40px}
section h2{display:flex;align-items:center;gap:12px;font-size:18px;font-weight:600;margin:0 0 16px;padding:0 0 10px;border-bottom:1px solid var(--border)}
section h2 .num{display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;background:var(--bg-card);color:var(--accent);font-size:13px;font-family:var(--mono);border:1px solid var(--border-strong);border-radius:6px}
section h2 .lbl{font-size:11px;font-weight:500;color:var(--fg-mute);font-family:var(--mono);margin-left:auto;letter-spacing:.4px}

/* summary cards */
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;margin-bottom:18px}
.card{background:var(--bg-card);border:1px solid var(--border);border-radius:6px;padding:18px;display:flex;align-items:center;gap:16px}
.card .icon{flex-shrink:0;width:44px;height:44px;border-radius:6px;display:flex;align-items:center;justify-content:center}
.card .icon svg{width:22px;height:22px}
.card .icon-green{background:var(--green-bg)}.card .icon-green svg{fill:var(--green)}
.card .icon-amber{background:var(--amber-bg)}.card .icon-amber svg{fill:var(--amber)}
.card .icon-red  {background:var(--red-bg)  }.card .icon-red   svg{fill:var(--red)}
.card .icon-blue {background:var(--blue-bg) }.card .icon-blue  svg{fill:var(--blue)}
.card .icon-gray {background:var(--gray-bg) }.card .icon-gray  svg{fill:var(--gray)}
.card .body{flex:1;min-width:0}
.card .body .num{font-size:26px;font-weight:600;font-family:var(--mono);line-height:1;color:var(--fg)}
.card .body .lbl{font-size:12px;color:var(--fg-mute);margin-top:4px}
.card .body .sub{font-size:11px;color:var(--fg-dim);margin-top:6px;font-family:var(--mono)}

/* info grid */
.info-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:0;background:var(--bg-card);border:1px solid var(--border);border-radius:6px;overflow:hidden}
.info-grid > div{padding:14px 18px;border-right:1px solid var(--border);border-bottom:1px solid var(--border)}
.info-grid > div:last-child{border-right:none}
.info-grid dt{font-size:11px;color:var(--fg-mute);text-transform:uppercase;letter-spacing:.4px;margin:0 0 4px}
.info-grid dd{margin:0;font-family:var(--mono);font-size:13px;color:var(--fg);word-break:break-all}

/* table */
table{width:100%;border-collapse:collapse;background:var(--bg-card);border:1px solid var(--border);border-radius:6px;overflow:hidden;font-size:13px}
thead th{background:var(--bg-card-2);color:var(--fg);font-weight:600;text-align:left;padding:10px 14px;border-bottom:1px solid var(--border);font-size:12px;text-transform:uppercase;letter-spacing:.3px}
tbody td{padding:9px 14px;border-bottom:1px solid var(--border);font-family:var(--mono);font-size:12.5px;vertical-align:middle}
tbody tr:nth-child(2n){background:rgba(255,255,255,.015)}
tbody tr:hover{background:rgba(88,166,255,.05)}
tbody tr:last-child td{border-bottom:none}
td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}
td.w{font-family:var(--sans);font-size:13px}

/* badge */
.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;font-family:var(--mono);letter-spacing:.3px;line-height:1.5;border:1px solid transparent}
.badge-green{background:var(--green-bg);color:var(--green);border-color:rgba(63,185,80,.3)}
.badge-amber{background:var(--amber-bg);color:var(--amber);border-color:rgba(210,153,34,.3)}
.badge-red  {background:var(--red-bg);  color:var(--red);  border-color:rgba(248,81,73,.3)}
.badge-gray {background:var(--gray-bg); color:var(--fg-mute);border-color:rgba(110,118,129,.3)}
.badge-blue {background:var(--blue-bg); color:var(--blue); border-color:rgba(88,166,255,.3)}

/* progress bar */
.bar{position:relative;height:6px;background:var(--gray-bg);border-radius:3px;overflow:hidden;width:100%;min-width:80px}
.bar > i{display:block;height:100%;background:var(--green);transition:none}
.bar.amber > i{background:var(--amber)}
.bar.red > i{background:var(--red)}

/* findings */
.findings{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}
.find-col{background:var(--bg-card);border:1px solid var(--border);border-radius:6px;padding:18px}
.find-col h3{margin:0 0 12px;font-size:13px;font-weight:600;color:var(--fg);display:flex;align-items:center;gap:8px;padding-bottom:8px;border-bottom:1px solid var(--border)}
.find-col h3 .dot{width:8px;height:8px;border-radius:50%}
.find-col.short h3 .dot{background:var(--red)}
.find-col.mid   h3 .dot{background:var(--amber)}
.find-col.long  h3 .dot{background:var(--blue)}
.find-col ul{margin:0;padding:0;list-style:none}
.find-col li{padding:8px 0;border-bottom:1px dashed var(--border);font-size:12.5px;color:var(--fg);line-height:1.5}
.find-col li:last-child{border-bottom:none}
.find-col li .area{font-family:var(--mono);font-size:11px;color:var(--fg-mute);text-transform:uppercase;letter-spacing:.3px;display:block;margin-bottom:2px}
.find-col .empty{color:var(--fg-dim);font-style:italic;font-size:12.5px}

/* disclaimer */
.disclaimer{background:var(--bg-card-2);border:1px solid var(--border);border-left:3px solid var(--gray);border-radius:6px;padding:20px 24px;color:var(--fg-mute);font-size:12.5px;line-height:1.7}
.disclaimer h3{margin:0 0 8px;color:var(--fg);font-size:13px;font-weight:600}

/* footer */
footer{text-align:center;color:var(--fg-dim);font-size:11px;font-family:var(--mono);margin-top:40px;padding-top:20px;border-top:1px solid var(--border)}

/* theme switcher (v1.2): 右上角悬浮 4 色圆点 */
.theme-switch{position:fixed;top:14px;right:16px;display:flex;gap:6px;background:var(--bg-card);border:1px solid var(--border);border-radius:24px;padding:6px;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.08)}
.theme-switch button{width:22px;height:22px;border:2px solid transparent;border-radius:50%;cursor:pointer;padding:0;transition:transform .15s,border-color .15s;outline:none}
.theme-switch button:hover{transform:scale(1.15)}
.theme-switch button.active{border-color:var(--fg);transform:scale(1.1)}

/* misc */
.k{color:var(--fg-mute)}
.muted{color:var(--fg-mute)}
.right{text-align:right}
.nowrap{white-space:nowrap}
.kbd{font-family:var(--mono);font-size:11px;padding:1px 6px;background:var(--bg-card-2);border:1px solid var(--border);border-radius:3px}

/* print */
@media print{
  .sidebar,.theme-switch{display:none}
  .main{margin-left:0;max-width:none;padding:20px}
  body{background:#fff;color:#000}
  .card,.banner,.info-grid,table,.find-col,.disclaimer{background:#fff;border-color:#ccc;color:#000}
  thead th{background:#f3f3f3;color:#000}
  section h2{color:#000}
}
'@

    $logoSvg = @'
<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path fill="#58a6ff" d="M3 3h6v6H3zM15 3h6v6h-6zM3 15h6v6H3zM15 15h6v6h-6z"/><path fill="#3fb950" d="M11 11h2v2h-2z"/></svg>
'@

    # 章节定义 (编号, ID, 标题)
    $toc = @(
        @('1','sec-overview',  'vCenter 概览'),
        @('2','sec-health',    'Appliance 健康'),
        @('3','sec-network',   '网络与 DNS'),
        @('4','sec-ntp',       'NTP 与时间'),
        @('5','sec-access',    '访问入口'),
        @('6','sec-cert',      '证书有效期'),
        @('7','sec-backup',    '备份策略'),
        @('8','sec-topo',      'Datacenter 与拓扑'),
        @('9','sec-hosts',     'ESXi 主机'),
        @('10','sec-ds',       'Datastore'),
        @('11','sec-net',      '网络'),
        @('12','sec-vm-sum',   'VM 总览'),
        @('13','sec-vm-list',  'VM 列表'),
        @('14','sec-tools',    'VMware Tools 抽样'),
        @('15','sec-services', 'Appliance 服务'),
        @('16','sec-snap',     'VM 快照健康'),
        @('17','sec-alarm',    'Alarm 当前告警'),
        @('18','sec-find',     '总体建议'),
        @('19','sec-discl',    '免责声明')
    )

    $sb = New-Object System.Text.StringBuilder

    # ---- HTML head ----
    [void]$sb.AppendLine(@"
<!DOCTYPE html>
<html lang="zh-CN" data-theme="$Theme">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>vCenter 巡检报告 — $(Html-Encode $VCenter) — $(Get-Date -Format 'yyyy-MM-dd')</title>
<style>$css</style>
</head>
<body>
<!-- 主题切换器 (右上角悬浮, localStorage 持久化) -->
<div class="theme-switch" id="themeSwitch" title="切换主题">
  <button data-th="light"   style="background:#1565c0" title="Light"></button>
  <button data-th="dark"    style="background:#58a6ff" title="Dark NOC"></button>
  <button data-th="minimal" style="background:#27272a" title="Minimal"></button>
  <button data-th="amber"   style="background:#b45309" title="Amber"></button>
</div>
"@)

    # ---- sidebar ----
    [void]$sb.AppendLine('<aside class="sidebar">')
    [void]$sb.AppendLine('  <div class="brand">')
    [void]$sb.AppendLine("    <div class='logo'>$logoSvg<div><div class='name'>vCenter Inspect</div><div class='sub'>v1.0</div></div></div>")
    [void]$sb.AppendLine('  </div>')
    [void]$sb.AppendLine('  <nav>')
    foreach ($t in $toc) {
        [void]$sb.AppendLine("    <a href='#$($t[1])'><span class='num'>$($t[0]).</span><span>$(Html-Encode $t[2])</span></a>")
    }
    [void]$sb.AppendLine('  </nav>')
    [void]$sb.AppendLine('</aside>')
    [void]$sb.AppendLine('<main class="main">')

    # ---- banner ----
    $verStr  = if ($sv) { "$($sv.version)  build $($sv.build)" } else { 'N/A' }
    $prod    = if ($sv) { "$($sv.product)"                       } else { 'VMware vCenter Server' }
    $relDate = if ($sv) { "$($sv.releasedate)"                   } else { 'N/A' }
    $instAt  = if ($sv) { "$($sv.install_time)"                  } else { 'N/A' }
    $vcTime  = if ($time) { "$($time.date) $($time.time) ($($time.timezone))" } else { 'N/A' }

    [void]$sb.AppendLine(@"
<div class="banner">
  <span class="tag">VCENTER 巡检</span>
  <h1>$(Html-Encode $prod) — $(Html-Encode $VCenter)</h1>
  <div class="sub">$(Html-Encode $verStr) &nbsp;·&nbsp; release $(Html-Encode $relDate) &nbsp;·&nbsp; total findings: <span class="badge badge-$overallKind">$overall</span></div>
  <dl class="meta">
    <div><dt>目标</dt><dd>$(Html-Encode $VCenter)</dd></div>
    <div><dt>采集账号</dt><dd>$(Html-Encode $Username)</dd></div>
    <div><dt>报告生成</dt><dd>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')</dd></div>
    <div><dt>vCenter 版本</dt><dd>$(Html-Encode $verStr)</dd></div>
    <div><dt>vCenter 时间</dt><dd>$(Html-Encode $vcTime)</dd></div>
    <div><dt>安装时间</dt><dd>$(Html-Encode $instAt)</dd></div>
  </dl>
</div>
"@)

    # =========================================================
    #  Summary 卡片
    # =========================================================
    $findKind = 'green'
    if ($critCnt -gt 0) { $findKind = 'red' }
    elseif ($warnCnt -gt 0) { $findKind = 'amber' }
    $sumCards = @(
        @{ icon='cluster';   kind='blue';  num=$hosts.Count;       lbl='ESXi 主机';    sub="$($clu.Count) 集群" },
        @{ icon='server';    kind='blue';  num=$vms.Count;         lbl='VM 总数';      sub="开机 $($vmOn.Count) / 关机 $($vmOff.Count) / 挂起 $($vmSusp.Count)" },
        @{ icon='database';  kind='blue';  num=$ds.Count;          lbl='Datastore';    sub="$(Format-Bytes $dsCap) 总容量" },
        @{ icon='network';   kind='blue';  num=$nets.Count;        lbl='Port Group';   sub="" },
        @{ icon='warn';      kind=$findKind; num=($critCnt + $warnCnt + $infoCnt); lbl='Findings'; sub="严重 $critCnt / 警告 $warnCnt / 提示 $infoCnt" },
        @{ icon='clock';     kind='gray';  num=([math]::Floor($up/86400)); lbl='Uptime (天)'; sub="$(Format-DaysFromSec $up)" }
    )
    $iconSvg = @{
        cluster  = '<svg viewBox="0 0 24 24"><path d="M6 2a4 4 0 100 8 4 4 0 000-8zm12 0a4 4 0 100 8 4 4 0 000-8zM6 14a4 4 0 100 8 4 4 0 000-8zm12 0a4 4 0 100 8 4 4 0 000-8z"/></svg>'
        server   = '<svg viewBox="0 0 24 24"><path d="M3 3h18v6H3zm0 8h18v6H3zm0 8h18v2H3zM6 5v2h2V5zm0 8v2h2v-2z"/></svg>'
        database = '<svg viewBox="0 0 24 24"><path d="M12 2C6.48 2 3 4.24 3 6v12c0 1.76 3.48 4 9 4s9-2.24 9-4V6c0-1.76-3.48-4-9-4zm0 18c-4.42 0-7-1.74-7-2v-1.41A12.69 12.69 0 0012 18a12.69 12.69 0 007-1.41V18c0 .26-2.58 2-7 2zm0-4c-4.42 0-7-1.74-7-2v-1.41A12.69 12.69 0 0012 14a12.69 12.69 0 007-1.41V14c0 .26-2.58 2-7 2zm0-4c-4.42 0-7-1.74-7-2V8.59A12.69 12.69 0 0012 10a12.69 12.69 0 007-1.41V10c0 .26-2.58 2-7 2zm0-4c-4.42 0-7-1.74-7-2s2.58-2 7-2 7 1.74 7 2-2.58 2-7 2z"/></svg>'
        network  = '<svg viewBox="0 0 24 24"><path d="M12 4l8 4-8 4-8-4 8-4zm0 6l8 4-8 4-8-4 8-4zm0 6l8 4-8 4-8-4 8-4z"/></svg>'
        warn     = '<svg viewBox="0 0 24 24"><path d="M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z"/></svg>'
        clock    = '<svg viewBox="0 0 24 24"><path d="M12 2a10 10 0 100 20 10 10 0 000-20zm0 18a8 8 0 110-16 8 8 0 010 16zm.5-13H11v6l5.2 3.2.8-1.3-4.5-2.7V7z"/></svg>'
    }
    [void]$sb.AppendLine('<div class="cards">')
    foreach ($c in $sumCards) {
        [void]$sb.AppendLine(@"
  <div class="card">
    <div class="icon icon-$($c.kind)">$($iconSvg[$c.icon])</div>
    <div class="body"><div class="num">$($c.num)</div><div class="lbl">$($c.lbl)</div><div class="sub">$(Html-Encode $c.sub)</div></div>
  </div>
"@)
    }
    [void]$sb.AppendLine('</div>')

    # =========================================================
    #  Section 1 — 概览
    # =========================================================
    $updateBadgeHtml = '—'
    if ($Data.Update) {
        $updKind = 'amber'
        if ("$($Data.Update.state)" -eq 'UP_TO_DATE') { $updKind = 'green' }
        $updateBadgeHtml = Badge -Text $Data.Update.state -Kind $updKind
    }
    [void]$sb.AppendLine(@"
<section id="sec-overview">
  <h2><span class="num">1</span>vCenter 概览<span class="lbl">/api/appliance/system/version</span></h2>
  <div class="info-grid">
    <div><dt>产品</dt><dd>$(Html-Encode $prod)</dd></div>
    <div><dt>版本</dt><dd>$(Html-Encode $sv.version)</dd></div>
    <div><dt>Build</dt><dd>$(Html-Encode $sv.build)</dd></div>
    <div><dt>发布日期</dt><dd>$(Html-Encode $sv.releasedate)</dd></div>
    <div><dt>部署形态</dt><dd>$(Html-Encode $sv.type)</dd></div>
    <div><dt>Summary</dt><dd>$(Html-Encode $sv.summary)</dd></div>
    <div><dt>安装时间</dt><dd>$(Html-Encode $sv.install_time)</dd></div>
    <div><dt>Uptime</dt><dd>$(Format-DaysFromSec $up) <span class="muted">($([math]::Round($up,0)) s)</span></dd></div>
    <div><dt>vCenter 当前时间</dt><dd>$(Html-Encode $vcTime)</dd></div>
    <div><dt>报告生成时间</dt><dd>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</dd></div>
    <div><dt>补丁状态</dt><dd>$updateBadgeHtml <span class="muted">last check $(Html-Encode $Data.Update.latest_query_time)</span></dd></div>
    <div><dt>整体评估</dt><dd>$(Badge -Text $overall -Kind $overallKind) &nbsp; $((Badge -Text "严重 $critCnt" -Kind 'red')) $((Badge -Text "警告 $warnCnt" -Kind 'amber')) $((Badge -Text "提示 $infoCnt" -Kind 'blue'))</dd></div>
  </div>
</section>
"@)

    # =========================================================
    #  Section 2 — Health
    # =========================================================
    [void]$sb.AppendLine(@"
<section id="sec-health">
  <h2><span class="num">2</span>Appliance 健康<span class="lbl">/api/appliance/health/*</span></h2>
  <div class="info-grid">
    <div><dt>System</dt>           <dd>$(Health-Badge $Data.Health.System)</dd></div>
    <div><dt>Storage</dt>          <dd>$(Health-Badge $Data.Health.Storage)</dd></div>
    <div><dt>Memory</dt>           <dd>$(Health-Badge $Data.Health.Memory)</dd></div>
    <div><dt>Swap</dt>             <dd>$(Health-Badge $Data.Health.Swap)</dd></div>
    <div><dt>Load</dt>             <dd>$(Health-Badge $Data.Health.Load)</dd></div>
    <div><dt>Database Storage</dt> <dd>$(Health-Badge $Data.Health.DatabaseStorage)</dd></div>
    <div><dt>Applmgmt</dt>         <dd>$(Health-Badge $Data.Health.Applmgmt)</dd></div>
    <div><dt>Software Packages</dt><dd>$(Health-Badge $Data.Health.SoftwarePackages)</dd></div>
    <div><dt>Last Check</dt>       <dd>$(Html-Encode $Data.Health.LastCheck)</dd></div>
  </div>
  <p class="muted" style="margin-top:10px;font-size:12px">注: GREEN = 正常, YELLOW/ORANGE = 警告, RED = 严重; 明细可在 VAMI (https://$VCenter:5480) 查看。</p>
</section>
"@)

    # =========================================================
    #  Section 3 — 网络与 DNS
    # =========================================================
    $net  = $Data.Networking
    $nics = @($Data.NetInterfaces)
    $dnsM = if ($net) { "$($net.dns.mode)" } else { 'N/A' }
    $dnsS = if ($net) { ($net.dns.servers -join ', ') } else { 'N/A' }
    $hostName = if ($net) { "$($net.dns.hostname)" } else { 'N/A' }
    $hnKind = if ($hostName -eq 'localhost' -or [string]::IsNullOrWhiteSpace($hostName)) { 'amber' } else { 'green' }

    $nicRows = ''
    foreach ($n in $nics) {
        $ipMode = "$($n.ipv4.mode)"
        $ipAddr = "$($n.ipv4.address)/$($n.ipv4.prefix)"
        $gw     = "$($n.ipv4.default_gateway)"
        $stKind = if ("$($n.status)" -eq 'up') { 'green' } else { 'gray' }
        $nicRows += @"
    <tr>
      <td>$(Html-Encode $n.name)</td>
      <td>$(Html-Encode $n.mac)</td>
      <td>$(Badge -Text $n.status -Kind $stKind)</td>
      <td>$(Html-Encode $ipMode)</td>
      <td>$(Html-Encode $ipAddr)</td>
      <td>$(Html-Encode $gw)</td>
    </tr>
"@
    }

    [void]$sb.AppendLine(@"
<section id="sec-network">
  <h2><span class="num">3</span>网络与 DNS<span class="lbl">/api/appliance/networking</span></h2>
  <div class="info-grid" style="margin-bottom:14px">
    <div><dt>Hostname</dt><dd>$(Badge -Text $hostName -Kind $hnKind)</dd></div>
    <div><dt>DNS 模式</dt><dd>$(Html-Encode $dnsM)</dd></div>
    <div><dt>DNS 服务器</dt><dd>$(Html-Encode $dnsS)</dd></div>
    <div><dt>vCenter Base URL</dt><dd>$(Html-Encode $net.vcenter_base_url)</dd></div>
  </div>
  <table>
    <thead><tr><th>网卡</th><th>MAC</th><th>状态</th><th>模式</th><th>IPv4</th><th>网关</th></tr></thead>
    <tbody>$nicRows</tbody>
  </table>
</section>
"@)

    # =========================================================
    #  Section 4 — NTP / Timesync
    # =========================================================
    $ntpList = @($Data.NTP)
    $tsMode  = "$($Data.Timesync)"
    $ntpKind = if ($ntpList.Count -eq 0) { 'amber' } else { 'green' }
    $tsKind  = if ($tsMode -eq 'NTP')    { 'green' } else { 'amber' }
    $ntpDisplay = if ($ntpList.Count -eq 0) { '(空 — 未配置 NTP 服务器)' } else { $ntpList -join ', ' }

    [void]$sb.AppendLine(@"
<section id="sec-ntp">
  <h2><span class="num">4</span>NTP 与时间同步<span class="lbl">/api/appliance/ntp · /timesync</span></h2>
  <div class="info-grid">
    <div><dt>NTP 服务器</dt><dd>$(Badge -Text ("count=$($ntpList.Count)") -Kind $ntpKind) &nbsp; $(Html-Encode $ntpDisplay)</dd></div>
    <div><dt>Timesync 模式</dt><dd>$(Badge -Text $tsMode -Kind $tsKind)</dd></div>
    <div><dt>vCenter 时间</dt><dd>$(Html-Encode $vcTime)</dd></div>
    <div><dt>本机时间</dt><dd>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')</dd></div>
  </div>
  <p class="muted" style="margin-top:10px;font-size:12px">说明: vCenter 与 ESXi 必须保持时间同步, 否则会导致 SSO 异常、证书校验失败、备份恢复时戳错乱。推荐 timesync = NTP, 至少 2 个上游源。</p>
</section>
"@)

    # =========================================================
    #  Section 5 — 访问入口
    # =========================================================
    $boolBadge = { param($b, $danger='amber')
        $v = "$b"
        if ($v -eq 'True' -or $v -eq 'true') { return Badge -Text 'enabled' -Kind $danger }
        return Badge -Text 'disabled' -Kind 'green'
    }
    $shellEn = if ($Data.Access.Shell) { $Data.Access.Shell.enabled } else { $null }

    [void]$sb.AppendLine(@"
<section id="sec-access">
  <h2><span class="num">5</span>访问入口<span class="lbl">/api/appliance/access/*</span></h2>
  <div class="info-grid">
    <div><dt>SSH</dt><dd>$(& $boolBadge $Data.Access.SSH 'amber')</dd></div>
    <div><dt>Bash Shell</dt><dd>$(& $boolBadge $shellEn 'amber')</dd></div>
    <div><dt>DCUI</dt><dd>$(& $boolBadge $Data.Access.DCUI 'amber')</dd></div>
    <div><dt>Console CLI</dt><dd>$(& $boolBadge $Data.Access.ConsoleCli 'amber')</dd></div>
  </div>
  <p class="muted" style="margin-top:10px;font-size:12px">建议: 生产环境默认关闭 SSH 与 Bash Shell, 仅在排错时临时启用; DCUI 用于本机控制台救援, 远程纯运维环境可关闭。</p>
</section>
"@)

    # =========================================================
    #  Section 6 — 证书
    # =========================================================
    $crt = $Data.CertTLS
    [void]$sb.AppendLine(@"
<section id="sec-cert">
  <h2><span class="num">6</span>vCenter TLS 证书<span class="lbl">/api/vcenter/certificate-management/vcenter/tls</span></h2>
  <div class="info-grid">
    <div><dt>剩余有效期</dt><dd>$(Badge -Text "$certDays 天" -Kind $certKind)</dd></div>
    <div><dt>有效期始</dt><dd>$(Html-Encode $crt.valid_from)</dd></div>
    <div><dt>有效期止</dt><dd>$(Html-Encode $crt.valid_to)</dd></div>
    <div><dt>颁发者</dt><dd>$(Html-Encode $crt.issuer_dn)</dd></div>
    <div><dt>主题</dt><dd>$(Html-Encode $crt.subject_dn)</dd></div>
    <div><dt>SAN</dt><dd>$(Html-Encode ($crt.subject_alternative_name -join ', '))</dd></div>
    <div><dt>签名算法</dt><dd>$(Html-Encode $crt.signature_algorithm)</dd></div>
    <div><dt>指纹 (SHA1)</dt><dd>$(Html-Encode $crt.thumbprint)</dd></div>
  </div>
</section>
"@)

    # =========================================================
    #  Section 7 — 备份
    # =========================================================
    $bjList = @($Data.BackupJobs)
    $bsCount = 0
    if ($Data.BackupSchedules) {
        try { $bsCount = ([pscustomobject]$Data.BackupSchedules).PSObject.Properties.Name.Count } catch {}
    }
    $bjKind  = 'green'; if ($bjList.Count -eq 0) { $bjKind = 'red' }
    $bsKind  = 'green'; if ($bsCount      -eq 0) { $bsKind = 'red' }
    $bkOverall = Badge -Text '已配置' -Kind 'green'
    if ($bjList.Count -eq 0 -and $bsCount -eq 0) { $bkOverall = Badge -Text '未配置任何备份' -Kind 'red' }

    [void]$sb.AppendLine(@"
<section id="sec-backup">
  <h2><span class="num">7</span>VAMI 备份策略<span class="lbl">/api/appliance/recovery/backup/*</span></h2>
  <div class="info-grid">
    <div><dt>备份计划数</dt><dd>$(Badge -Text "$bsCount" -Kind $bsKind)</dd></div>
    <div><dt>历史备份任务数</dt><dd>$(Badge -Text "$($bjList.Count)" -Kind $bjKind)</dd></div>
    <div><dt>状态</dt><dd>$bkOverall</dd></div>
  </div>
  <p class="muted" style="margin-top:10px;font-size:12px">建议: 在 VAMI (https://$VCenter:5480) → Backup 配置 SFTP/NFS/SMB 周期备份, 至少每周一次, 保留 ≥ 3 份。</p>
</section>
"@)

    # =========================================================
    #  Section 8 — 拓扑
    # =========================================================
    $topoRows = ''
    foreach ($dc in $dcs) { $topoRows += "<tr><td><b>Datacenter</b></td><td>$(Html-Encode $dc.name)</td><td>$(Html-Encode $dc.datacenter)</td></tr>" }
    foreach ($cl in $clu) {
        $haB  = if ($cl.ha_enabled) { Badge -Text 'enabled' -Kind 'green' } else { Badge -Text 'disabled' -Kind 'amber' }
        $drsB = if ($cl.drs_enabled){ Badge -Text 'enabled' -Kind 'green' } else { Badge -Text 'disabled' -Kind 'amber' }
        $topoRows += "<tr><td><b>Cluster</b></td><td>$(Html-Encode $cl.name)</td><td>$(Html-Encode $cl.cluster) &nbsp; HA: $haB &nbsp; DRS: $drsB</td></tr>"
    }
    foreach ($f in @($Data.Folders)) { $topoRows += "<tr><td>Folder ($($f.type))</td><td>$(Html-Encode $f.name)</td><td>$(Html-Encode $f.folder)</td></tr>" }
    foreach ($rp in @($Data.ResPools)) { $topoRows += "<tr><td>ResourcePool</td><td>$(Html-Encode $rp.name)</td><td>$(Html-Encode $rp.resource_pool)</td></tr>" }

    [void]$sb.AppendLine(@"
<section id="sec-topo">
  <h2><span class="num">8</span>Datacenter / Cluster / Folder<span class="lbl">/api/vcenter/{datacenter,cluster,folder,resource-pool}</span></h2>
  <table>
    <thead><tr><th style="width:140px">类型</th><th>名称</th><th>对象 ID / 属性</th></tr></thead>
    <tbody>$topoRows</tbody>
  </table>
</section>
"@)

    # =========================================================
    #  Section 9 — ESXi 主机
    # =========================================================
    $hostRows = ''
    foreach ($h in $hosts) {
        $csKind = if ("$($h.connection_state)" -eq 'CONNECTED') { 'green' } elseif ("$($h.connection_state)" -eq 'DISCONNECTED') { 'red' } else { 'amber' }
        $psKind = if ("$($h.power_state)"      -eq 'POWERED_ON') { 'green' } else { 'gray' }
        $hostRows += @"
    <tr>
      <td>$(Html-Encode $h.host)</td>
      <td>$(Html-Encode $h.name)</td>
      <td>$(Badge -Text $h.connection_state -Kind $csKind)</td>
      <td>$(Badge -Text $h.power_state -Kind $psKind)</td>
    </tr>
"@
    }
    # PowerCLI 实时数据 (v1.1)
    $pcliHostBlock = ''
    if ($Script:PCLIEnabled -and $Data.PCLI_Hosts) {
        $rowsPcli = ''
        foreach ($h in @($Data.PCLI_Hosts)) {
            $csKind = if ("$($h.ConnState)" -eq 'Connected') { 'green' } elseif ("$($h.ConnState)" -in @('Disconnected','NotResponding')) { 'red' } else { 'amber' }
            $cpuKind = if ($h.CpuPct -ge 85) { 'red' } elseif ($h.CpuPct -ge 70) { 'amber' } else { 'green' }
            $memKind = if ($h.MemPct -ge 90) { 'red' } elseif ($h.MemPct -ge 80) { 'amber' } else { 'green' }
            $rowsPcli += @"
    <tr>
      <td class="w">$(Html-Encode $h.Name)</td>
      <td>$(Html-Encode $h.Version) <span class="muted">build $(Html-Encode $h.Build)</span></td>
      <td>$(Html-Encode $h.Vendor) $(Html-Encode $h.Model)</td>
      <td>$($h.NumCpu) vCPU · $(Badge -Text ("{0}%" -f $h.CpuPct) -Kind $cpuKind) <span class="muted">$($h.CpuUsedMhz) / $($h.CpuTotalMhz) MHz</span></td>
      <td>$(Badge -Text ("{0}%" -f $h.MemPct) -Kind $memKind) <span class="muted">$($h.MemUsedGB) / $($h.MemTotalGB) GB</span></td>
      <td>$($h.UptimeDays) 天</td>
      <td>$(Badge -Text $h.ConnState -Kind $csKind)</td>
    </tr>
"@
        }
        $pcliHostBlock = @"
  <h3 style="margin:24px 0 10px;font-size:14px;color:var(--text-2)">实时运行数据 <span class="muted" style="font-weight:400;font-size:12px">via PowerCLI</span></h3>
  <table>
    <thead><tr><th>主机</th><th>ESXi 版本</th><th>硬件</th><th>CPU</th><th>内存</th><th>Uptime</th><th>状态</th></tr></thead>
    <tbody>$rowsPcli</tbody>
  </table>
"@
    } elseif (-not $Script:PCLIEnabled) {
        $pcliHostBlock = '<p class="muted" style="margin-top:10px;font-size:12px"><b>提示</b>: 启用 PowerCLI 回退 (-UsePowerCLI 或安装 VMware.PowerCLI) 可补全单主机 CPU/Mem/Uptime 等 REST 不暴露的实时数据。</p>'
    }

    [void]$sb.AppendLine(@"
<section id="sec-hosts">
  <h2><span class="num">9</span>ESXi 主机<span class="lbl">/api/vcenter/host</span></h2>
  <table>
    <thead><tr><th>Host ID</th><th>名称 / IP</th><th>连接状态</th><th>电源状态</th></tr></thead>
    <tbody>$hostRows</tbody>
  </table>
$pcliHostBlock
</section>
"@)

    # =========================================================
    #  Section 10 — Datastore
    # =========================================================
    $dsRows = ''
    foreach ($d in $ds) {
        $cap  = [double]$d.capacity
        $free = [double]$d.free_space
        $used = if ($cap -gt 0) { ($cap - $free) / $cap } else { 0 }
        $pct  = [math]::Round($used * 100, 1)
        $bar  = if     ($used -ge 0.9) { 'red'   }
                elseif ($used -ge 0.8) { 'amber' }
                elseif ($used -ge 0.7) { 'amber' }
                else                   { 'green' }
        $dsRows += @"
    <tr>
      <td>$(Html-Encode $d.name)</td>
      <td>$(Html-Encode $d.type)</td>
      <td class="num">$(Format-Bytes $cap)</td>
      <td class="num">$(Format-Bytes $free)</td>
      <td class="num">$(Format-Bytes ($cap - $free))</td>
      <td class="num">$pct %</td>
      <td style="width:140px"><div class="bar $bar"><i style="width:$([math]::Min($pct,100))%"></i></div></td>
    </tr>
"@
    }
    $dsUsed = $dsCap - $dsFree
    $dsPct  = if ($dsCap -gt 0) { [math]::Round($dsUsed / $dsCap * 100, 1) } else { 0 }
    [void]$sb.AppendLine(@"
<section id="sec-ds">
  <h2><span class="num">10</span>Datastore<span class="lbl">/api/vcenter/datastore</span></h2>
  <div class="info-grid" style="margin-bottom:14px">
    <div><dt>Datastore 数</dt><dd>$($ds.Count)</dd></div>
    <div><dt>总容量</dt><dd>$(Format-Bytes $dsCap)</dd></div>
    <div><dt>已用</dt><dd>$(Format-Bytes $dsUsed)</dd></div>
    <div><dt>空闲</dt><dd>$(Format-Bytes $dsFree)</dd></div>
    <div><dt>整体使用率</dt><dd>$dsPct %</dd></div>
  </div>
  <table>
    <thead><tr><th>名称</th><th>类型</th><th class="num">容量</th><th class="num">空闲</th><th class="num">已用</th><th class="num">使用率</th><th>占比</th></tr></thead>
    <tbody>$dsRows</tbody>
  </table>
</section>
"@)

    # =========================================================
    #  Section 11 — 网络 (Portgroup)
    # =========================================================
    $netRows = ''
    $stdPg = @($nets | Where-Object { $_.type -eq 'STANDARD_PORTGROUP' })
    $dvPg  = @($nets | Where-Object { $_.type -eq 'DISTRIBUTED_PORTGROUP' })
    foreach ($n in $nets) {
        $netRows += "<tr><td>$(Html-Encode $n.name)</td><td>$(Html-Encode $n.type)</td><td>$(Html-Encode $n.network)</td></tr>"
    }
    [void]$sb.AppendLine(@"
<section id="sec-net">
  <h2><span class="num">11</span>网络 / Portgroup<span class="lbl">/api/vcenter/network</span></h2>
  <div class="info-grid" style="margin-bottom:14px">
    <div><dt>Portgroup 总数</dt><dd>$($nets.Count)</dd></div>
    <div><dt>Standard Portgroup</dt><dd>$($stdPg.Count)</dd></div>
    <div><dt>Distributed Portgroup</dt><dd>$($dvPg.Count)</dd></div>
  </div>
  <table>
    <thead><tr><th>名称</th><th>类型</th><th>网络 ID</th></tr></thead>
    <tbody>$netRows</tbody>
  </table>
</section>
"@)

    # =========================================================
    #  Section 12 — VM 总览 (cards + 资源汇总)
    # =========================================================
    [void]$sb.AppendLine(@"
<section id="sec-vm-sum">
  <h2><span class="num">12</span>虚拟机总览<span class="lbl">/api/vcenter/vm</span></h2>
  <div class="info-grid">
    <div><dt>VM 总数</dt><dd>$($vms.Count)</dd></div>
    <div><dt>开机</dt><dd>$(Badge -Text $vmOn.Count -Kind 'green')</dd></div>
    <div><dt>关机</dt><dd>$(Badge -Text $vmOff.Count -Kind 'gray')</dd></div>
    <div><dt>挂起</dt><dd>$(Badge -Text $vmSusp.Count -Kind 'amber')</dd></div>
    <div><dt>vCPU 总配额 (全部)</dt><dd>$cpuSum</dd></div>
    <div><dt>vCPU (开机)</dt><dd>$cpuOn</dd></div>
    <div><dt>内存总配额 (全部)</dt><dd>$(Format-Bytes ($memSum * 1MB))</dd></div>
    <div><dt>内存 (开机)</dt><dd>$(Format-Bytes ($memOn * 1MB))</dd></div>
  </div>
</section>
"@)

    # =========================================================
    #  Section 13 — VM 列表
    # =========================================================
    $vmRows = ''
    $sortedVms = $vms | Sort-Object @{e='power_state';desc=$true}, @{e='cpu_count';desc=$true}, name
    foreach ($v in $sortedVms) {
        $psKind = if ("$($v.power_state)" -eq 'POWERED_ON') { 'green' } elseif ("$($v.power_state)" -eq 'POWERED_OFF') { 'gray' } else { 'amber' }
        $vmRows += @"
    <tr>
      <td class="w">$(Html-Encode $v.name)</td>
      <td>$(Html-Encode $v.vm)</td>
      <td>$(Badge -Text $v.power_state -Kind $psKind)</td>
      <td class="num">$($v.cpu_count)</td>
      <td class="num">$(Format-Bytes ([double]$v.memory_size_MiB * 1MB))</td>
    </tr>
"@
    }
    [void]$sb.AppendLine(@"
<section id="sec-vm-list">
  <h2><span class="num">13</span>VM 列表<span class="lbl">共 $($vms.Count) 台 · 按电源 + vCPU 排序</span></h2>
  <table>
    <thead><tr><th>名称</th><th>VM ID</th><th>电源</th><th class="num">vCPU</th><th class="num">内存</th></tr></thead>
    <tbody>$vmRows</tbody>
  </table>
</section>
"@)

    # =========================================================
    #  Section 14 — VMware Tools 抽样
    # =========================================================
    $toolRows = ''
    if ($Data.ToolsSample.Count -eq 0) {
        $toolRows = '<tr><td colspan="6" class="muted">未抽样 (SkipToolsSample) 或无开机 VM</td></tr>'
    } else {
        foreach ($t in $Data.ToolsSample) {
            $rsKind = if ("$($t.RunState)" -eq 'RUNNING') { 'green' } else { 'gray' }
            $stKind = switch ("$($t.Status)") {
                'CURRENT'        { 'green' }
                'SUPPORTED_NEW'  { 'green' }
                'SUPPORTED_OLD'  { 'amber' }
                'UNMANAGED_OLD'  { 'amber' }
                'UNMANAGED'      { 'blue'  }
                'NOT_INSTALLED'  { 'red'   }
                default          { 'gray'  }
            }
            $toolRows += @"
    <tr>
      <td class="w">$(Html-Encode $t.VM)</td>
      <td>$(Badge -Text $t.RunState -Kind $rsKind)</td>
      <td>$(Html-Encode $t.Version)</td>
      <td>$(Badge -Text $t.Status -Kind $stKind)</td>
      <td>$(Html-Encode $t.Install)</td>
      <td>$(Html-Encode $t.Policy)</td>
    </tr>
"@
        }
    }
    [void]$sb.AppendLine(@"
<section id="sec-tools">
  <h2><span class="num">14</span>VMware Tools 抽样<span class="lbl">/api/vcenter/vm/{vm}/tools · 样本大小 $($Data.ToolsSample.Count)</span></h2>
  <table>
    <thead><tr><th>VM</th><th>运行</th><th>版本</th><th>状态</th><th>类型</th><th>策略</th></tr></thead>
    <tbody>$toolRows</tbody>
  </table>
</section>
"@)

    # =========================================================
    #  Section 15 — Services
    # =========================================================
    $svcCoreRows = ''
    foreach ($s in $svcRows) {
        $k = if ("$($s.State)" -eq 'STARTED') { 'green' } elseif ("$($s.State)" -eq 'STOPPED') { 'gray' } else { 'amber' }
        $svcCoreRows += "<tr><td>$(Html-Encode $s.Name)</td><td>$(Badge -Text $s.State -Kind $k)</td><td class='w'>$(Html-Encode $s.Desc)</td></tr>"
    }
    [void]$sb.AppendLine(@"
<section id="sec-services">
  <h2><span class="num">15</span>Appliance 服务<span class="lbl">/api/appliance/services</span></h2>
  <div class="info-grid" style="margin-bottom:14px">
    <div><dt>服务总数</dt><dd>$svcTotal</dd></div>
    <div><dt>STARTED</dt><dd>$(Badge -Text "$svcStarted" -Kind 'green')</dd></div>
    <div><dt>STOPPED</dt><dd>$(Badge -Text "$svcStopped" -Kind 'gray')</dd></div>
    <div><dt>OTHER</dt><dd>$(Badge -Text "$svcOther" -Kind 'amber')</dd></div>
  </div>
  <table>
    <thead><tr><th style="width:280px">核心服务</th><th>状态</th><th>描述</th></tr></thead>
    <tbody>$svcCoreRows</tbody>
  </table>
  <p class="muted" style="margin-top:10px;font-size:12px">说明: 仅列出核心 vCenter 服务; 完整列表可在 VAMI 或 SSH `service-control --list` 查看。</p>
</section>
"@)

    # =========================================================
    #  Section 16 — VM 快照健康 (v1.1, PowerCLI)
    # =========================================================
    if ($Script:PCLIEnabled -and $Data.PCLI_Snapshots) {
        $sn = $Data.PCLI_Snapshots
        $oldKind = if ($sn.OldestDays -gt 90) { 'red' } elseif ($sn.OldestDays -gt 30) { 'amber' } else { 'green' }
        $chnKind = if ($sn.MaxChain -gt 3)    { 'amber' } elseif ($sn.MaxChain -gt 1) { 'blue' } else { 'green' }
        $totKind = if ($sn.TotalGB -gt 1024)  { 'amber' } else { 'green' }
        $topRows = ''
        if ($sn.Top10.Count -eq 0) {
            $topRows = '<tr><td colspan="6" class="empty">当前无任何快照,状态健康</td></tr>'
        } else {
            foreach ($s in $sn.Top10) {
                $ageKind = if ($s.AgeDays -gt 90) { 'red' } elseif ($s.AgeDays -gt 30) { 'amber' } else { 'green' }
                $topRows += @"
    <tr>
      <td class="w">$(Html-Encode $s.VM)</td>
      <td class="w">$(Html-Encode $s.Snapshot)</td>
      <td>$(Html-Encode $s.Created)</td>
      <td>$(Badge -Text ("{0} 天" -f $s.AgeDays) -Kind $ageKind)</td>
      <td>$($s.SizeGB) GB</td>
      <td>$(Html-Encode $s.PowerState)</td>
    </tr>
"@
            }
        }
        [void]$sb.AppendLine(@"
<section id="sec-snap">
  <h2><span class="num">16</span>VM 快照健康<span class="lbl">via PowerCLI · Get-Snapshot</span></h2>
  <div class="info-grid" style="margin-bottom:14px">
    <div><dt>有快照的 VM</dt><dd>$($sn.VmCount)</dd></div>
    <div><dt>快照总数</dt><dd>$($sn.Count)</dd></div>
    <div><dt>总占用</dt><dd>$(Badge -Text ("{0} GB" -f $sn.TotalGB) -Kind $totKind)</dd></div>
    <div><dt>最老快照</dt><dd>$(Badge -Text ("{0} 天" -f $sn.OldestDays) -Kind $oldKind)</dd></div>
    <div><dt>最长快照链</dt><dd>$(Badge -Text "$($sn.MaxChain) 层" -Kind $chnKind)</dd></div>
  </div>
  <h3 style="margin:18px 0 8px;font-size:13px;color:var(--text-2)">Top 10 占用最大的快照</h3>
  <table>
    <thead><tr><th>VM</th><th>快照名</th><th>创建时间</th><th>年龄</th><th>大小</th><th>电源</th></tr></thead>
    <tbody>$topRows</tbody>
  </table>
  <p class="muted" style="margin-top:10px;font-size:12px">建议: 快照非备份机制, &gt; 30 天建议合并, &gt; 90 天强烈建议合并或删除; 快照链 &gt; 3 层会显著拖累 I/O。</p>
</section>
"@)
    } else {
        [void]$sb.AppendLine(@"
<section id="sec-snap">
  <h2><span class="num">16</span>VM 快照健康<span class="lbl">via PowerCLI</span></h2>
  <p class="empty">PowerCLI 回退未启用, 跳过快照采集。<br><span class="muted">启用方式: <code>Install-Module VMware.PowerCLI -Scope CurrentUser</code> 或加参数 <code>-UsePowerCLI</code></span></p>
</section>
"@)
    }

    # =========================================================
    #  Section 17 — Alarm 当前告警 (v1.1, PowerCLI)
    # =========================================================
    if ($Script:PCLIEnabled -and $null -ne $Data.PCLI_Alarms) {
        $alarms = @($Data.PCLI_Alarms)
        $redN  = @($alarms | Where-Object { "$($_.Status)" -eq 'red' }).Count
        $yelN  = @($alarms | Where-Object { "$($_.Status)" -eq 'yellow' }).Count
        $alarmRows = ''
        if ($alarms.Count -eq 0) {
            $alarmRows = '<tr><td colspan="6" class="empty">当前无任何 triggered alarm</td></tr>'
        } else {
            foreach ($a in $alarms) {
                $stKind = if ("$($a.Status)" -eq 'red') { 'red' } elseif ("$($a.Status)" -eq 'yellow') { 'amber' } else { 'gray' }
                $ackBg  = if ($a.Acked) { Badge -Text 'ACK' -Kind 'gray' } else { Badge -Text 'NEW' -Kind 'blue' }
                $ageStr = if ($null -ne $a.AgeHours) {
                    if ($a.AgeHours -lt 24) { ("{0:N1} h" -f $a.AgeHours) }
                    else                    { ("{0:N1} d" -f ($a.AgeHours / 24)) }
                } else { '—' }
                $alarmRows += @"
    <tr>
      <td class="w">$(Html-Encode $a.Entity)</td>
      <td>$(Html-Encode $a.EntityType)</td>
      <td class="w">$(Html-Encode $a.Alarm)</td>
      <td>$(Badge -Text $a.Status -Kind $stKind) $ackBg</td>
      <td>$(Html-Encode $a.Time)</td>
      <td>$ageStr</td>
    </tr>
"@
            }
        }
        [void]$sb.AppendLine(@"
<section id="sec-alarm">
  <h2><span class="num">17</span>Alarm 当前告警<span class="lbl">via PowerCLI · TriggeredAlarmState</span></h2>
  <div class="info-grid" style="margin-bottom:14px">
    <div><dt>总告警</dt><dd>$($alarms.Count)</dd></div>
    <div><dt>RED (严重)</dt><dd>$(Badge -Text "$redN" -Kind ($(if($redN -gt 0){'red'}else{'green'})))</dd></div>
    <div><dt>YELLOW (警告)</dt><dd>$(Badge -Text "$yelN" -Kind ($(if($yelN -gt 0){'amber'}else{'green'})))</dd></div>
  </div>
  <table>
    <thead><tr><th>实体</th><th>类型</th><th>Alarm</th><th>状态</th><th>触发时间</th><th>持续</th></tr></thead>
    <tbody>$alarmRows</tbody>
  </table>
  <p class="muted" style="margin-top:10px;font-size:12px">说明: 仅显示当前 triggered 状态; 已 acknowledged 的告警仍会展示, 历史 Event 记录请在 vSphere Client → Monitor → Events 查看。</p>
</section>
"@)
    } else {
        [void]$sb.AppendLine(@"
<section id="sec-alarm">
  <h2><span class="num">17</span>Alarm 当前告警<span class="lbl">via PowerCLI</span></h2>
  <p class="empty">PowerCLI 回退未启用, 跳过 Alarm 采集。<br><span class="muted">启用方式见上一章节。</span></p>
</section>
"@)
    }

    # =========================================================
    #  Section 18 — 总体建议 (was 16)
    # =========================================================
    $short = $findings | Where-Object Severity -eq 'critical'
    $mid   = $findings | Where-Object Severity -eq 'warn'
    $long  = $findings | Where-Object Severity -eq 'info'
    function Render-Findings($arr, $emptyMsg) {
        if (-not $arr -or $arr.Count -eq 0) { return "<p class='empty'>$emptyMsg</p>" }
        $html = '<ul>'
        foreach ($it in $arr) {
            $html += "<li><span class='area'>[$($it.Area)]</span><b>$(Html-Encode $it.Title)</b><br><span class='muted'>$(Html-Encode $it.Detail)</span></li>"
        }
        $html += '</ul>'
        return $html
    }
    [void]$sb.AppendLine(@"
<section id="sec-find">
  <h2><span class="num">18</span>总体建议<span class="lbl">基于本次巡检数据动态生成</span></h2>
  <div class="findings">
    <div class="find-col short">
      <h3><span class="dot"></span>短期 (立即处理 · 严重 $($short.Count))</h3>
      $(Render-Findings $short '本次未发现严重问题。')
    </div>
    <div class="find-col mid">
      <h3><span class="dot"></span>中期 (1-2 周内 · 警告 $($mid.Count))</h3>
      $(Render-Findings $mid '本次未发现警告项。')
    </div>
    <div class="find-col long">
      <h3><span class="dot"></span>长期 (持续改进 · 提示 $($long.Count))</h3>
      $(Render-Findings $long '暂无优化提示。')
    </div>
  </div>
</section>
"@)

    # =========================================================
    #  Section 19 — 免责 (was 17)
    # =========================================================
    $pcliNote = if ($Script:PCLIEnabled) {
        '<p><b>PowerCLI 回退已启用</b>: 已通过 PowerCLI 补采单主机实时数据 (CPU/Mem/Uptime/Build)、VM 快照清单 (含大小 + 年龄) 与当前 Triggered Alarms。</p>'
    } else {
        '<p><b>PowerCLI 回退未启用</b>: 单主机实时数据 / VM 快照 / Alarm 三类数据未采集 (本次仅走 REST API)。如需补全, 请运行 <code>Install-Module VMware.PowerCLI -Scope CurrentUser</code> 后重跑, 或指定参数 <code>-UsePowerCLI</code>。</p>'
    }
    [void]$sb.AppendLine(@"
<section id="sec-discl">
  <h2><span class="num">19</span>免责声明<span class="lbl">disclaimer</span></h2>
  <div class="disclaimer">
    <h3>关于本报告</h3>
    <p>本报告由 <b>vcenter_inspect.ps1 v1.1.0</b> 通过 vCenter REST API (vSphere Automation API) 自动采集生成, 不修改任何 vCenter 配置, 不写入任何文件到 vCenter Appliance。v1.1 引入 PowerCLI 回退层, 补 REST 8.0 不暴露的快照 / Alarm / 单 host 实时数据。</p>
    $pcliNote
    <p><b>REST API 8.0 仍存在的限制</b>:</p>
    <ul>
      <li>单 ESXi 主机详细信息 (CPU/Memory/Build/Uptime/Maintenance Mode) 已 deprecated, v1.1 通过 PowerCLI 补全。</li>
      <li>VM 快照列表 / 大小、Alarm 触发状态在 REST 8.0 未暴露, v1.1 通过 PowerCLI 补全。</li>
      <li>历史 Event / Task / 性能曲线 (CPU/Memory/IOPS) 仍未采集, 建议在 vSphere Client → Monitor → Events / Performance 中查看。</li>
      <li>License 状态在 `/api/vcenter/licensing/licenses` 8.0 返回 404, 仍未采集。</li>
    </ul>
    <p>建议结合 VAMI (https://$VCenter:5480) 与 vSphere Client 手工核对告警明细。本报告中的"总体建议"基于通用最佳实践与本次数据生成, 实际优先级请根据业务场景判断。</p>
  </div>
</section>
"@)

    # ---- footer + scroll-spy ----
    $cost = [int]((Get-Date) - $Script:T0).TotalSeconds
    $pcliFooter = if ($Script:PCLIEnabled) { ' &nbsp;·&nbsp; <span style="color:var(--text-2)">PowerCLI 回退</span>' } else { '' }
    [void]$sb.AppendLine(@"
<footer>
  vcenter_inspect.ps1 v1.1.0 &nbsp;·&nbsp; 采集耗时 ${cost}s &nbsp;·&nbsp; 生成于 $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')$pcliFooter
</footer>
</main>
<script>
(function(){
  var sections = document.querySelectorAll('section[id]');
  var links = {};
  document.querySelectorAll('.sidebar nav a').forEach(function(a){
    links[a.getAttribute('href').slice(1)] = a;
  });
  var io = new IntersectionObserver(function(entries){
    entries.forEach(function(e){
      var id = e.target.id;
      if (e.isIntersecting && links[id]) {
        Object.values(links).forEach(function(x){x.classList.remove('active');});
        links[id].classList.add('active');
      }
    });
  }, { rootMargin: '-40% 0px -55% 0px' });
  sections.forEach(function(s){ io.observe(s); });
})();

// 主题切换 (v1.2): localStorage 持久化, 覆盖 <html data-theme>
(function(){
  var KEY = 'vci_theme';
  var root = document.documentElement;
  var saved = null;
  try { saved = localStorage.getItem(KEY); } catch(e) {}
  if (saved && /^(light|dark|minimal|amber)`$/.test(saved)) {
    root.setAttribute('data-theme', saved);
  }
  var current = root.getAttribute('data-theme') || 'light';
  var sw = document.getElementById('themeSwitch');
  if (!sw) return;
  function paint(){
    sw.querySelectorAll('button').forEach(function(b){
      b.classList.toggle('active', b.dataset.th === current);
    });
  }
  paint();
  sw.addEventListener('click', function(e){
    var b = e.target.closest('button[data-th]');
    if (!b) return;
    current = b.dataset.th;
    root.setAttribute('data-theme', current);
    try { localStorage.setItem(KEY, current); } catch(e) {}
    paint();
  });
})();
</script>
</body></html>
"@)

    return $sb.ToString()
}

# ============================================================================
#  主流程
# ============================================================================
Log-Banner
try {
    Collect-All
} catch {
    Log-Err $_.Exception.Message
    try { Logout-VC } catch {}
    exit 2
}

# v1.1: PowerCLI 回退 (默认: 已装就用, 没装就跳过; -SkipPowerCLI 强制关; -UsePowerCLI 强制开)
try {
    Collect-PowerCLI
} catch {
    Log-Warn ("PowerCLI 回退异常 (跳过, 不影响 REST 主报告): $($_.Exception.Message)")
    try { Disconnect-PowerCLI } catch {}
}

$html = Render-Report

if (-not $Output) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
    $safeVC = ($VCenter -replace '[^\w\.\-]','_')
    $Output = Join-Path $scriptDir ("report_{0}_{1}.html" -f $safeVC, (Get-Date -Format 'yyyy-MM-dd'))
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($Output, $html, $utf8NoBom)

if (-not $Quiet) {
    $cost = [int]((Get-Date) - $Script:T0).TotalSeconds
    Write-Host ''
    Write-Host ('  ✓ 巡检完成 — 耗时 {0}s — 报告已写入:' -f $cost) -ForegroundColor Green
    Write-Host "    $Output" -ForegroundColor White
    Write-Host ''
}
exit 0
