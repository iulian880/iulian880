<div align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="./assets/hero-dark.svg">
    <source media="(prefers-color-scheme: light)" srcset="./assets/hero-light.svg">
    <img alt="Iulian Petcu — Software Engineer · Automation, Fintech & AI Tooling" src="./assets/hero-dark.svg" width="100%">
  </picture>
</div>

<p align="center">
  <a href="mailto:petcuiulian@icloud.com"><strong>petcuiulian@icloud.com</strong></a>
  &nbsp;·&nbsp;
  <a href="#enterprise-automation-1000-hours-saved-annually">Enterprise Impact</a>
  &nbsp;·&nbsp;
  <a href="#featured-systems">Featured Systems</a>
  &nbsp;·&nbsp;
  <a href="#how-i-engineer-software">Engineering Principles</a>
  &nbsp;·&nbsp;
  <a href="#technical-stack">Tech Stack</a>
</p>

---

**Applied Computer Science graduate** with **5+ years in Fortune Global 500 finance operations**, building software where correctness, auditability, and operational controls are non-negotiable.

I started by engineering internal automation suites that **reclaimed 1,000+ hours annually** across finance teams under strict corporate IT lock-downs — and now design and ship standalone **fintech execution engines, local-first desktop systems, and deterministic AI agent tooling** in **Python**, **C# / .NET 10**, and **Node.js 22**.

My edge is owning both sides of the problem: I know financial reconciliation, operational controls, and enterprise failure modes from the inside, and I ship tested, fail-closed software to eliminate them.

---

## Enterprise Automation: 1,000+ Hours Saved Annually

<div align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="./assets/impact-dark.svg">
    <source media="(prefers-color-scheme: light)" srcset="./assets/impact-light.svg">
    <img alt="Enterprise Automation Architecture — 1,000+ hours saved annually" src="./assets/impact-dark.svg" width="100%">
  </picture>
</div>

Built and maintained in production for enterprise finance teams operating on locked-down corporate Windows workstations:

- **End-to-End ERP & ServiceNow Automation:** Engineered a self-contained, declarative Order-to-Cash and ticket-processing engine spanning multiple SAP ERP environments — designed to run with **no admin privileges, no third-party RPA licenses, disabled SAP GUI Scripting, and zero stored credentials**.
- **Hybrid Win32, Edge CDP & Document Pipelines:** Combined focus-safe Win32 automation with Edge WebDriver/CDP (reusing the operator's authenticated corporate SSO session, with Shadow-DOM fallback for ACL-restricted ticket resolution), Outlook COM capture, and a self-deploying embedded Python PDF binder (`pypdf`, `extract-msg`).
- **Zero-Excel Cross-Platform Build & Headless CI:** Kept 100% of automation logic as plain-text Git-tracked modules and built a pure-Python **MS-OVBA 2.4.1 LZ77 + OLE/CFB** compiler that builds `.xlsm` workbooks deterministically on macOS/Linux without Excel — verified by a custom static VBA linter and **20+ headless test suites** (LibreOffice Basic + Node.js `jsdom`).
- **Operational Finance Tooling:** Delivered companion automated invoice/ticket journals, bulk reconciliation tools, and reporting pipelines with Python build automation and `pytest` suites — collectively saving **1,000+ hours per year** across the team.

> *Source code for internal corporate tools remains private to respect confidentiality; architectural walkthroughs and anonymized demonstrations are available during interviews.*

---

## Featured Systems

Four public systems built end-to-end with **1,141 automated tests**, clean architecture, CI verification, and real-world failure handling.

### [`tgmt5`](https://github.com/iulian880/tgmt5) — Risk-Controlled Telegram → MetaTrader 5 Copy-Trading Engine
[![Tests](https://github.com/iulian880/tgmt5/actions/workflows/tests.yml/badge.svg)](https://github.com/iulian880/tgmt5/actions/workflows/tests.yml)
[![Build Windows exe](https://github.com/iulian880/tgmt5/actions/workflows/build.yml/badge.svg)](https://github.com/iulian880/tgmt5/actions/workflows/build.yml)
![Python 3.12+](https://img.shields.io/badge/Python-3.12%2B-3776AB?logo=python&logoColor=white)
![Tests: 299](https://img.shields.io/badge/tests-299%20passing-238636)

Windows desktop execution engine (`PySide6 + qasync`, `Telethon`, `SQLite WAL`, official `MetaTrader5` API) engineered so **every failure mode defaults to doing nothing**:
- **Deterministic signal parser** (zero LLM hallucination risk) with strict geometry, negation, and conditional guards; ambiguous messages route to Manual Review.
- **16-rule persisted Risk Engine** (daily loss & drawdown lockouts, streak circuit breaker, correlated exposure, margin, spread, UTC sessions) + broker-driven `order_calc_profit()` sizing across 6 modes.
- **Idempotent commit-before-send execution:** every order leg commits to SQLite with a unique `TG<sig>L<leg>` tag before dispatch and reconciles positions, pending orders, and 7-day deal history before any bounded retry.
- **Double `AccountGuard`**, monitor-only startup, 3-step typed confirmation for real accounts, OS Keyring secrets with automatic log redaction, and an offline `--sandbox` simulator.

---

### [`ultimate-claude`](https://github.com/iulian880/ultimate-claude) — Deterministic Governance & Evaluation Marketplace for Claude Code
[![Verify marketplace](https://github.com/iulian880/ultimate-claude/actions/workflows/verify.yml/badge.svg)](https://github.com/iulian880/ultimate-claude/actions/workflows/verify.yml)
![Node.js 22](https://img.shields.io/badge/Node.js-22%20ESM-339933?logo=node.js&logoColor=white)
![Zero Dependencies](https://img.shields.io/badge/npm%20dependencies-0-0D9488)
![Tests: 176](https://img.shields.io/badge/tests-176%20passing-238636)

Cross-platform AI agent governance framework (**zero npm dependencies**, verified across Linux, macOS, and Windows CI) that turns best-effort prompting into auditable engineering workflows:
- **Fail-closed `PreToolUse` safety hooks** parsing chained Bash & PowerShell commands, pipes, quotes, aliases, and destructive `git` actions — plus `.xlsm` workbook backup enforcement (`uc-excel`).
- **Cryptographic v2 release gates:** computes content-based Git tree fingerprints via scratch `GIT_OBJECT_DIRECTORY` (never mutating `.git/index` or leaving objects behind) binding spec criteria, before/after code state, and OS platform.
- **Hard-verdict conformance review** (`BLOCK` on unresolved material conditions) and an **offline `stream-json` trace grader** (18 routing scenarios, SHA-256 bound reviews, paired A/B capability tests).

---

### [`finance-os`](https://github.com/iulian880/finance-os) — Local-First Finance Operations Workspace
![C# / .NET 10](https://img.shields.io/badge/C%23-.NET%2010%20WPF-512BD4?logo=dotnet&logoColor=white)
![SQLite + Dapper + FTS5](https://img.shields.io/badge/Storage-SQLite%20WAL%20%2B%20FTS5-003B57?logo=sqlite&logoColor=white)
![Tests: 530](https://img.shields.io/badge/tests-530%20passing-238636)

Single-file portable Windows desktop workspace (`WPF`, `CommunityToolkit.Mvvm`, `Dapper`, `SQLite`) tailored for finance/accounting routines on locked-down machines (no install, no cloud):
- **6-project Clean Architecture** with dependency rules enforced automatically by `FinanceOS.Architecture.Tests` and a domain core decoupled from UI, storage, and Windows APIs.
- **Transactional 50-action Undo** across multi-entity batches, statement-by-statement checksummed SQL migrations, integrity checks on startup, and automated daily/pre-migration backup recovery.
- **Deterministic reminder scheduling** via `TimeProvider` / `FakeTimeProvider` (reconciling Windows scheduled toasts fired while closed), diacritic-insensitive FTS5 command palette, computed Romanian Orthodox Easter holidays, and live zero-restart **EN/RO** localization.

---

### [`noti-relay`](https://github.com/iulian880/noti-relay) — Zero-Dependency Windows Desktop Notification Relay
[![Build Windows exe](https://github.com/iulian880/noti-relay/actions/workflows/build.yml/badge.svg)](https://github.com/iulian880/noti-relay/actions/workflows/build.yml)
![C# / .NET 10](https://img.shields.io/badge/C%23-.NET%2010%20WPF-512BD4?logo=dotnet&logoColor=white)
![Zero Runtime Deps](https://img.shields.io/badge/runtime%20deps-0-0D9488)
![Tests: 136](https://img.shields.io/badge/tests-136%20passing-238636)

Portable Windows tray utility (**zero runtime NuGet dependencies**, signed sparse-MSIX / single `.exe`, no admin rights) forwarding Microsoft Teams messages, Outlook desktop toasts (`UserNotificationListener` WinRT), classic Outlook calendar updates & reminders (leak-free late-bound COM), and laptop low-battery warnings over HTTPS via `ntfy`:
- **Windows DPAPI (`CurrentUser`) encryption** with `ZeroMemory` buffer wiping for per-category ntfy topics and the atomic `.tmp`-swapped durable delivery queue.
- **Resilient delivery & operator controls:** exponential outage backoff, strict FIFO HTTP 429 pacing, launch backlog confirmation gate, instant kill-switch, privacy-first redacting logs, and a read-only 4-step network probe.

---

## Shipped Desktop Interfaces

Real screenshots captured directly from built-in sandbox / seed modes (automatically adapting to GitHub light and dark themes):

<table>
  <tr>
    <td width="50%" valign="top">
      <a href="https://github.com/iulian880/tgmt5">
        <picture>
          <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/iulian880/tgmt5/main/docs/screenshots/01-dashboard.png">
          <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/iulian880/tgmt5/main/docs/screenshots/light/01-dashboard.png">
          <img alt="TGMT5 — Risk-Controlled Trading Dashboard" src="https://raw.githubusercontent.com/iulian880/tgmt5/main/docs/screenshots/01-dashboard.png" width="100%">
        </picture>
      </a>
      <br />
      <sub><strong>TGMT5</strong> — PySide6 trading risk console, live P&amp;L telemetry &amp; emergency controls (<code>--sandbox</code>)</sub>
    </td>
    <td width="50%" valign="top">
      <a href="https://github.com/iulian880/finance-os">
        <picture>
          <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/iulian880/finance-os/main/docs/screenshots/tasks-details-dark.png">
          <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/iulian880/finance-os/main/docs/screenshots/overview-light.png">
          <img alt="Finance OS — Desktop Workspace for Finance Operations" src="https://raw.githubusercontent.com/iulian880/finance-os/main/docs/screenshots/tasks-details-dark.png" width="100%">
        </picture>
      </a>
      <br />
      <sub><strong>Finance OS</strong> — .NET 10 WPF custom design system, SQLite FTS5 &amp; bilingual EN/RO UX (<code>--seed</code>)</sub>
    </td>
  </tr>
</table>

<details>
  <summary>View NotiRelay Windows Desktop UI (<code>noti-relay</code>)</summary>
  <br />
  <p align="center">
    <a href="https://github.com/iulian880/noti-relay">
      <img alt="NotiRelay — Windows Notification Relay UI" src="https://raw.githubusercontent.com/iulian880/noti-relay/main/docs/screenshots/main-active.png" width="520">
    </a>
  </p>
</details>

---

## How I Engineer Software

- **Fail-closed by default:** Ambiguous trading signals escalate to manual review instead of guessing (`tgmt5`); destructive shell chains and unbacked `.xlsm` writes hard-block at the tool boundary (`ultimate-claude`); missed notifications wait for explicit confirmation (`noti-relay`).
- **Idempotency & crash recovery first:** Every order commits to SQLite WAL before broker dispatch (`tgmt5`); persistent queues swap atomically via `.tmp` files (`noti-relay`); SQL migrations enforce checksums inside per-migration transactions (`finance-os`).
- **Deterministic core, thin imperative shell:** Business logic, risk rules, parsers, and scheduling math stay decoupled from UI, OS APIs, and network I/O so they run in fast headless suites (**1,141 public tests + 20+ enterprise suites**).
- **Engineered for real-world constraints:** Single portable executables, minimal or zero runtime dependencies, per-user storage without administrator rights, OS Keyring / DPAPI secret hygiene, and automatic secret redaction in logs.

---

## Technical Stack

| Domain | Technologies & Craft |
|---|---|
| **Languages** | Python (3.12+), C# (.NET 10), JavaScript / Node.js 22 (ESM), TypeScript, SQL, VBA |
| **Desktop, UI & Systems** | PySide6 / Qt (`qasync`), WPF (MVVM, custom design systems), Win32 API, WinRT, Late-Bound COM, DPAPI, Edge WebDriver / CDP |
| **Data & Persistence** | SQLite (WAL, FTS5, `STRICT`, transactional migrations), Dapper, SQLAlchemy, Alembic |
| **AI Tooling & Automation** | Claude Code plugins & fail-closed hooks, offline trace evals, SAP GUI / ServiceNow / PDF pipelines, pure-Python MS-OVBA/CFB tooling |
| **Testing & Delivery** | `pytest`, `xUnit v3`, `node:test`, Architecture Tests, `jsdom`, LibreOffice headless testing, GitHub Actions CI (Linux / macOS / Windows) |
| **Domain Depth** | Finance & accounting operations, Order-to-Cash (O2C), reconciliation & controls, trading risk management, bilingual EN/RO UX |

---

## Background & Contact

- **Education:** B.Sc. in Applied Computer Science — West University of Timișoara (2017–2020)
- **Languages:** Romanian (Native) · English (C1 — Cambridge CAE) · German (B1)
- **Currently Exploring:** RAG pipelines over financial documents, deterministic LLM agent workflows, FastAPI & containerized backend services
- **Availability:** Open to **Remote (EU)** and **Timișoara (Hybrid)** engineering roles — Software Engineering, Automation, Fintech & AI Tooling

📫 **Reach me directly:** [petcuiulian@icloud.com](mailto:petcuiulian@icloud.com)
