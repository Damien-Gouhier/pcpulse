# PCPulse

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE)](https://learn.microsoft.com/powershell/)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)]()
[![Version](https://img.shields.io/badge/version-1.0-brightgreen)](CHANGELOG.md)
[![Status](https://img.shields.io/badge/status-pilot-orange)]()

> **Zero-dependency Windows fleet monitoring.**
> One PowerShell script to collect, one PowerShell script to render a self-contained HTML dashboard. That's it.

🇫🇷 [Version française](README.md)

![PCPulse Dashboard - Fleet overview](screenshots/dashboard-overview.png)

> **Note:** the dashboard UI is in French. An English localization is on the roadmap. The rest of the documentation is available in English below.

---

## 🤔 What is PCPulse?

A **Windows fleet monitoring tool** for IT teams who want a fleet health snapshot **without deploying Zabbix, SCCM, or buying a $30k/year solution**.

- **Two PowerShell scripts**, nothing else.
- **No database**, no service, no installed agent.
- **A shared SMB folder** serves as storage.
- **A self-contained HTML report** generated on demand, openable on any PC.

## ✨ What's collected on each PC

| Category | Metrics |
|---|---|
| 🔒 **Security** | EDR status (SentinelOne), offline PCs |
| ⚠️ **Stability** | Application crashes, freezes, BSODs, WHEA fatal/corrected, GPU TDR, thermal throttling |
| ⚡ **Performance** | Boot duration, detailed Boot Performance (MainPath, PostBoot, UserProfile, Explorer init) |
| 🔧 **Hardware wear** | Battery health (wear % + cycles), Disk SMART (wear, temp, errors), aging secondary monitors |
| 📊 **Inventory** | CPU (model, year, age category), RAM, disks, chassis (Laptop/Desktop/AIO), external monitors (EDID) |

## 📸 Preview

### Fleet overview

One row per PC. Row color = alert level. Sort, filters (period, site, CPU), search.

![Fleet view with 5 different sites](screenshots/dashboard-overview.png)

### Per-PC drill-down

Clicking a PC opens 5 tabs for deep-dive: Overview, Stability, Boot, Hardware, Security.

![Drill-down — Stability tab](screenshots/drill-stabilite.png)

![Drill-down — Boot tab with Boot Performance](screenshots/drill-demarrage.png)

![Drill-down — Hardware tab (disk, SMART, battery, monitors)](screenshots/drill-materiel.png)

### Fleet-wide aggregates

Boot type distribution, cross-fleet top crashers, secondary monitor inventory.

![Fleet-wide aggregates](screenshots/agregats-parc.png)

## 🚀 Quick Start — try it in 3 minutes

Before deploying to your fleet, you can see the Dashboard **right now** with the 5 demo JSON files provided.

**Prerequisites**: Windows 10/11 + PowerShell 7 (`winget install Microsoft.PowerShell`).

```powershell
# 1. Clone the repo
git clone https://github.com/Damien-Gouhier/pcpulse.git
cd pcpulse

# 2. Run the Dashboard on demo JSONs
pwsh .\02_Dashboard.ps1 -SharePath ".\examples\demo"
```

> 💡 **If Windows blocks execution** with `cannot be loaded... not digitally signed`, it's normal (default Windows protection). Two options:
> - **One-shot**: add `-ExecutionPolicy Bypass` → `pwsh -ExecutionPolicy Bypass -File .\02_Dashboard.ps1 -SharePath ".\examples\demo"`
> - **Permanent (recommended)**: run once as admin `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned`

The HTML opens automatically in your browser. Explore the 5 example scenarios:

- `LAPTOP-001` → Healthy case (all green)
- `LAPTOP-002` → Multiple alerts (dead battery + BSOD + crashes + PCIe errors)
- `DESKTOP-003` → Aging desktop, disk almost full
- `AIO-004` → All-In-One with an 8-year-old secondary monitor
- `OFFLINE-005` → Laptop not seen for 12 days

## 🏗️ Architecture

```
┌──────────────────────────────────────────────────────────┐
│                     DEPLOYMENT                           │
│   (Intune, SmartDeploy, GPO, or manual scheduled task)   │
└─────────────────────┬────────────────────────────────────┘
                      │
                      ▼
     ┌────────────────────────────────────┐
     │   01_Collector.ps1 on each PC      │
     │   • Scheduled task (SYSTEM)        │
     │   • Runs every 1-4h                │
     │   • Random anti-collision delay    │
     └────────────────┬───────────────────┘
                      │ writes
                      ▼
            ┌──────────────────────┐
            │   \\SERVER\share\    │
            │   ├─ PC1.json        │
            │   ├─ PC2.json        │
            │   ├─ PC3.json        │
            │   └─ ...             │
            └──────────┬───────────┘
                       │ reads
                       ▼
     ┌────────────────────────────────────┐
     │   02_Dashboard.ps1 (admin host)    │
     │   • PowerShell 7                   │
     │   • On demand                      │
     │   • Generates a self-contained HTML│
     └────────────────┬───────────────────┘
                      │
                      ▼
           🌐 PCPulse-Dashboard-*.html
```

### Key characteristics

- **Zero external dependencies**: only native PowerShell and inline HTML/CSS/JS. The generated HTML is self-contained (no CDN, works offline).
- **PS 5.1 compatible** on the Collector side (= native Windows 10/11 fleet, no prerequisite install).
- **Atomic writes**: if the SMB share is unavailable, the Collector buffers locally and catches up on the next run.
- **Backward compatible**: the Dashboard supports older JSON schemas as the project evolves.

## ⚙️ Configuration

Two optional files, placed in `$SharePath` (default `C:\PCPulse`):

- **`config.psd1`** — thresholds, score weights, dashboard title, etc.
  See [`config.psd1.example`](config.psd1.example) as a documented template.
- **`ip-ranges.csv`** — IP / hostname → Site mapping (optional, enables the Site column).
  See [`ip-ranges.example.csv`](ip-ranges.example.csv) and [`ip-ranges.README.md`](ip-ranges.README.md).

Both files are excluded from the repo via `.gitignore` to prevent accidental leaks of real data.

## 🎯 Who is this for?

- **SMB/mid-market sysadmins** (50 to 2000 endpoints) without budget for a commercial monitoring solution
- **Public sector IT teams** (public / para-public organizations) with heterogeneous fleets
- **MSPs / managed service providers** who want a lightweight tool to deploy at multiple clients
- **Homelab / curious sysadmins** who just want to see their machines' health

**Not suitable for**:
- Real-time monitoring (this is a periodic snapshot, not a live feed)
- Push alerting (no Slack / email notifications — it's a dashboard)
- Linux / macOS fleets (Windows only)

## 📦 Deploying to a real fleet

The [Quick Start](#-quick-start--try-it-in-3-minutes) isn't enough for production. For a real rollout:

1. Set up an **SMB share** writable by the fleet's machine accounts (Kerberos auth)
2. Deploy `01_Collector.ps1` to each endpoint + create a **SYSTEM scheduled task** (via Intune, SmartDeploy, GPO…)
3. Configure `config.psd1` and `ip-ranges.csv` to match your environment
4. Run `02_Dashboard.ps1` on demand from an admin host with PowerShell 7

👉 Detailed docs coming in `docs/` (INSTALL, DEPLOYMENT-INTUNE, DEPLOYMENT-SMARTDEPLOY, TROUBLESHOOTING, SECURITY).

## 🛠️ Tech stack

- **PowerShell 5.1** (Collector) / **PowerShell 7** (Dashboard)
- **WMI / CIM** for hardware telemetry
- **Get-WinEvent** for event logs
- **Vanilla HTML / CSS / JS** for the Dashboard (no framework, no bundler)
- **JSON** as exchange format (Collector → Dashboard)

## 🤝 Contributing

Contributions are welcome! To discuss an idea, report a bug, or propose an improvement, open a [GitHub Issue](https://github.com/Damien-Gouhier/pcpulse/issues).

For a Pull Request:
1. Fork the repo
2. Create a branch (`git checkout -b feature/my-feature`)
3. Commit with a clear message
4. Push and open the PR

The project is in **pilot phase**: the roadmap adapts based on field feedback.

## 📄 License

[MIT](LICENSE) — Copyright (c) 2026 Damien Gouhier.

You can use, modify and redistribute this project freely, including in commercial contexts, as long as you keep the copyright notice.

---

*PCPulse — because a healthy fleet is a fleet where users stop calling support.* 💙
