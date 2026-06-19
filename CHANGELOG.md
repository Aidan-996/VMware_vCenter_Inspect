# Changelog

All notable changes to **VMware vCenter Inspect** will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

### Added (v1.2 → 等多 vC 批量做完一起发)

- **报告主题切换** — 新增 `-Theme` 参数，4 套内置主题：
  - `light`（新默认，白底 + 工程师蓝，适合邮件 / 打印）
  - `dark`（深灰 + 亮蓝 NOC，原 v1.0/v1.1 默认）
  - `minimal`（灰白 + 近黑 accent）
  - `amber`（米色 + 琥珀棕，暖色商务）
- **报告内实时切换主题** — 右上角悬浮 4 色圆点切换器，点击即换肤，无需重跑脚本。`localStorage` 持久化用户偏好，再次打开同一份 HTML 自动恢复上次选择
- **主色单独覆盖** — 新增 `-AccentColor` 参数，接 hex 如 `#10b981`，单独覆盖该主题的 accent 而不切整套配色（对所有 4 主题生效）
- HTML `:root` 重构为 ~15 个 CSS variables，4 套主题全部内嵌（`[data-theme="..."]` 切换），新增章节只需写一遍样式即可全主题统一
- 体积成本：报告从 ~42 KB → ~45 KB（+3 KB），打印时 `@media print` 自动隐藏切换器

### Planned

- 多 vCenter 批量模式（`-VCenter @('vc1','vc2','vc3')`）
- 基线对比（跟上次 findings 做 diff，只输出新增 / 已解决告警）
- Telegram / 飞书 / 钉钉告警推送（`-Notify` 参数）
- 配置文件（阈值从 `vcenter_inspect.config.json` 读取）

---

## [1.1.0] — 2026-06-19

### Added

#### PowerCLI 回退层

REST API 8.0 至今未暴露的三类数据，本版本通过可选 PowerCLI 回退补全：

- **单 ESXi 主机实时数据** — `Get-VMHost`：CPU MHz 使用率 / 内存 GB 使用率 / Uptime / Version+Build / 厂商型号 / 连接状态
- **VM 快照清单** — `Get-Snapshot` 全 VM 遍历：每个快照的名称、创建时间、年龄（天）、大小（GB）、所属 VM 电源状态；聚合：有快照 VM 数 / 总快照数 / 总占用 GB / 最老快照年龄 / 最长快照链 / Top10 大快照表
- **当前 Triggered Alarms** — 遍历 Datacenter/Cluster/Host/Datastore/VM 的 `TriggeredAlarmState`：实体名 / 类型 / Alarm 定义名 / 状态（red/yellow/gray）/ 触发时间 / 持续小时 / 已 ACK 标识

#### 自动检测 + 优雅降级

- **默认行为**：检测到 `VMware.VimAutomation.Core` 模块即启用；未装则跳过 PowerCLI 章节，不影响 REST 主报告
- **`-UsePowerCLI`**：强制启用（未装时打印 `Install-Module VMware.PowerCLI` 安装提示）
- **`-SkipPowerCLI`**：强制跳过（即使已装），只走纯 REST

#### HTML 报告新增 2 章节 + 1 个内联模块

- **Section 9 ESXi 主机** 注入 "实时运行数据" 副表（PowerCLI 启用时）
- **Section 16 VM 快照健康**（新）—— KPI 卡片（VM 数 / 快照数 / 总占用 / 最老天数 / 最长链）+ Top10 表
- **Section 17 Alarm 当前告警**（新）—— 红/黄计数 + 告警明细表
- 原 16/17（总体建议 / 免责声明）顺延到 18/19

#### 新增 Eval-Findings 规则（10 条）

| 维度 | 规则 | 级别 |
|---|---|---|
| ESXi Memory | 使用率 ≥ 90% | critical |
| ESXi Memory | 使用率 ≥ 80% | warn |
| ESXi CPU    | 使用率 ≥ 85% | warn |
| ESXi 连接   | Disconnected / NotResponding | critical |
| ESXi Uptime | > 365 天 | info |
| Snapshot    | 最老快照 > 90 天 | critical |
| Snapshot    | 最老快照 > 30 天 | warn |
| Snapshot    | 单 VM 链深 > 3 | warn |
| Snapshot    | 总占用 > 1 TB | info |
| Alarm       | RED 告警 > 0 | critical |
| Alarm       | YELLOW 告警 > 5 | warn |

### Changed

- 脚本 banner 版本 `v1.0` → `v1.1.0`
- 免责声明章节按 PowerCLI 状态动态显示提示文字
- TOC 从 17 项扩展到 19 项

### Compatibility

- PowerCLI 13.x 已验证（VMware.PowerCLI 13.0+）
- 已装 PowerCLI 时整体耗时 +10~20s（取决于 VM 数量与快照数）
- 未装 PowerCLI 行为完全等同 v1.0（无破坏性变更）

---

## [1.0.0] — 2026-05-25

### Added

#### 主脚本 `vcenter_inspect.ps1`（1250 行）

- 17 章节完整巡检报告：vCenter 概览 / Appliance 健康 / 网络 DNS / NTP 时间同步 / 访问入口 / TLS 证书 / VAMI 备份 / 拓扑 / ESXi 主机 / Datastore / Portgroup / VM 总览 / VM 列表 / VMware Tools 抽样 / Appliance 服务 / 总体建议 / 免责声明
- **Dual-mode REST API**：自动 fallback `/api/...`（v8）→ `/rest/com/vmware/cis/...`（v6），覆盖 vCenter 6.5 ~ 8.0
- **18 条 Findings 评估规则**，按 critical / warn / info 分级
- **动态总体建议**：按 findings 等级自动归类成短期 / 中期 / 长期 3 列
- **工程师风 HTML 模板**：扁平卡片 + 单色边框 + 蓝色 metadata banner + 章节编号 + 左侧 TOC scroll-spy（IntersectionObserver，无 jQuery）+ 打印样式（Ctrl+P 出 PDF）
- **8 个命令行参数**：`-VCenter` / `-Username` / `-Password` / `-Output` / `-ToolsSampleSize` / `-SkipToolsSample` / `-DebugDump` / `-Quiet`
- **自动重试机制**：对 0 / 5xx 瞬时错误指数退避（1s / 2s / 4s），扛 sts-idmd 抖动
- **错误诊断 hint**：401 / 403 / 5xx / 网络不可达 / 证书问题分别给出具体修复建议
- **session 自动清理**：脚本退出时主动 DELETE session，不留挂
- **UTF-8 中文支持**：手控 `[System.Net.HttpWebRequest]` + `[Text.Encoding]::UTF8`，绕过 PS 5.1 `Invoke-RestMethod` GBK 误解码

#### 配套转换器

- `html_to_md.ps1` — HTML → Markdown 转换器，针对本项目 HTML 结构正则解析，输出 ~15 KB
- `html_to_docx.ps1` — HTML → Word 转换器，用 Office Word COM 注入样式，输出 ~40 KB

#### 文档

- README.md — 完整使用文档（核心特性 / 快速开始 / 参数 / 章节说明 / Findings 规则 / 兼容性 / 自动化场景 / 开发踩坑 / 路线图）
- 脱敏 demo 报告样张 `report_demo_2026-05-25.html`
- README 截图：`docs/report-preview.png`（首屏 hero）+ `docs/report-preview-full.jpg`（整页超长效果）

### Tested

- vCenter 8.0.3 单 Cluster：2 ESXi / 33 VM / 6 Datastore，**4 秒**采集，42 KB 报告，捕获 11 条 findings
- vCenter 6.5 LTS 多 Cluster：25 ESXi / 39 Datastore / 38 TB，**~30 秒**采集，49 KB 报告，捕获 26 条 findings（6 critical / 15 warn / 5 info）
- Windows PowerShell 5.1（Win11 跳板机）+ PowerShell 7.4

### Known Limitations

REST API 不暴露的部分本工具不采集，需要 PowerCLI / vmodl SOAP 补充：

- 单 ESXi 主机 CPU / Memory / Build / Uptime / Maintenance Mode 实时数据（8.0 端点 deprecated）
- VM 快照大小（REST 可拿快照树，size 字段不暴露）
- Alarm / Event / 告警历史
- License 详情（`/api/vcenter/licensing/licenses` 在 8.0 返回 404）
- 性能历史曲线（stats API 仍是 preview）

7.0 / 7.0U3 路径已通过 dual-mode 自动匹配，缺真实环境实测验证。

---

[Unreleased]: https://github.com/Aidan-996/VMware_vCenter_Inspect/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/Aidan-996/VMware_vCenter_Inspect/releases/tag/v1.0.0
