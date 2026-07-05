# PCPulse

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE)](https://learn.microsoft.com/powershell/)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)](.)
[![Version](https://img.shields.io/badge/version-2.2.0-brightgreen)](CHANGELOG.md)
[![Status](https://img.shields.io/badge/status-pilot-orange)](.)

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

## 💡 Why I'm building this

Three reasons, in this order:

1. **Curiosity and learning.** PCPulse is my first real coding project. It's an opportunity to confront real questions: how to think about a data schema that survives evolutions? How to manage a deployment on production endpoints? How to write docs that are actually useful? Etc.

2. **Wanting to give back to open source.** I use free software every day at work. At my modest scale, I'd like to give back a small part of what I receive.

3. **A real-world need.** I manage a fleet of several hundred Windows endpoints and I needed this tool. Rather than depending on a contractor or buying an off-the-shelf solution full of features I'd never use, I preferred to code exactly what I needed. PCPulse is therefore used in production, which forces it to be solid and pragmatic.

If any of these resonate with you, feel free to open an [Issue](https://github.com/Damien-Gouhier/pcpulse/issues) to discuss, contribute, or just share your feedback.

## ✨ What's collected on each PC

| Category | Metrics |
|---|---|
| 🔒 **Security** | EDR status (defined in config.psd1), offline PCs |
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
>
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
     │   • Scheduled task (SYSTEM/gMSA)   │
     │   • Runs every 1-4h                │
     │   • Random anti-collision delay    │
     └────────────────┬───────────────────┘
                      │ writes
                      ▼
            ┌──────────────────────┐
            │   \\SERVER\share\    │
            │   ├─ release\        │ ◄── read-only for clients
            │   │   ├─ 01_Collector.ps1
            │   │   ├─ version.txt
            │   │   └─ KILLSWITCH.txt (optional)
            │   ├─ killed\         │
            │   ├─ logs\           │
            │   ├─ PC1.json        │
            │   ├─ PC2.json        │
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
- **Backward compatible**: the Dashboard accepts JSON schemas 2.1 and 2.2 (during the gradual per-machine rollout).
- **Auto-update**: `PCPulse-Updater.ps1` automatically pulls the latest Collector from `\release\` with SHA256 verification.
- **Killswitch**: remote uninstall via a sentinel file (see below).

## ⚙️ Configuration

Two optional files, placed in `$SharePath` (default `C:\PCPulse`):

- **`config.psd1`** — thresholds, score weights, dashboard title, custom killswitch, etc.
  See [`config.psd1.example`](config.psd1.example) as a documented template.
- **`ip-ranges.csv`** — IP → Site mapping (optional, enables the Site column).
  See [`ip-ranges.example.csv`](ip-ranges.example.csv) and [`ip-ranges.README.md`](ip-ranges.README.md).

Both files are excluded from the repo via `.gitignore` to prevent accidental leaks of real data.

## ☠️ Killswitch — remote uninstall

PCPulse ships with a **killswitch** mechanism that allows you to uninstall Collectors from your entire fleet remotely, without physically touching any machine.

**Use cases**:
- Decommission PCPulse cleanly (replace with another tool)
- Emergency stop if a critical bug is detected
- Major migration (kill v1 → install v2 cleanly)

**How it works**:
1. You drop a sentinel file `KILLSWITCH.txt` in `\\SERVER\PCPulse$\release\` containing the phrase `CONFIRM-UNINSTALL-PCPULSE`
2. On the next hourly cycle, each PC:
   - Detects the file
   - Writes a report to `\killed\<HOSTNAME>.txt`
   - Removes its scheduled task
   - Deletes its `C:\ProgramData\PCPulse\` folder
3. You remove the sentinel file once all PCs have appeared in `\killed\`

**In production**, change the phrase and filename via `config.psd1` to prevent any accidental triggering. See [`config.psd1.example`](config.psd1.example) and [`SECURITY.md`](SECURITY.md) for details.

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

1. Set up an **SMB share** on a Windows server (prefer encrypted SMB3)
2. Run **`Setup-Server.ps1`** on the server to create the structure and hardened ACLs (admin account required, see [`SECURITY.md`](SECURITY.md) for details)
3. Deploy `01_Collector.ps1` + `PCPulse-Updater.ps1` to each endpoint via Intune, SmartDeploy, GPO, etc.
4. Create a **scheduled task** that calls `PCPulse-Updater.ps1 -SharePath \\SERVER\PCPulse$` every 1-4h (as SYSTEM or via gMSA)
5. Configure `config.psd1` and `ip-ranges.csv` to match your environment
6. Run `02_Dashboard.ps1` on demand from an admin host with PowerShell 7

👉 Detailed install guide: [`docs/INSTALL.md`](docs/INSTALL.md). More docs coming (DEPLOYMENT-INTUNE, DEPLOYMENT-SMARTDEPLOY, TROUBLESHOOTING).

## 🔐 Security

PCPulse is designed to be **deployed on production endpoints**. As such, the project takes security seriously.

→ **Before any deployment**, read [`SECURITY.md`](SECURITY.md) which details:
- The **trust model** (who can do what on the share)
- **Recommended ACLs** (and the attack scenario they close)
- The **threat model** of the killswitch
- The **hardening roadmap** (code signing, sanity-checks, etc.)
- How to **report a vulnerability**

`Setup-Server.ps1` automates the application of recommended hardened ACLs on your existing share.

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
