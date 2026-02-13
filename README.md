<p align="center">
  <img src="https://img.shields.io/badge/MagicMirror%C2%B2-module-blueviolet" alt="MagicMirror² Module" />
  <img src="https://img.shields.io/badge/platform-Raspberry%20Pi-red" alt="Raspberry Pi" />
  <img src="https://img.shields.io/badge/exchange-EWS%20%2B%20NTLM-blue" alt="Exchange EWS" />
  <img src="https://img.shields.io/badge/license-MIT-green" alt="MIT License" />
</p>

# MMM-Exchange

<p align="center">
  <strong>Display your on-premise Microsoft Exchange calendar on MagicMirror² via EWS SOAP API with NTLM authentication.</strong>
</p>

<p align="center">
  <a href="#features">Features</a> •
  <a href="#installation">Installation</a> •
  <a href="#configuration">Configuration</a> •
  <a href="#troubleshooting">Troubleshooting</a> •
  <a href="#license">License</a>
</p>

<p align="center">
  <strong><a href="README.hu.md">Magyar nyelvű leírás / Hungarian documentation</a></strong>
</p>

---

## Why this module?

Most calendar modules rely on CalDAV, Google Calendar, or iCal URLs. If your organization runs an **on-premise Microsoft Exchange server** that only supports **NTLM authentication**, none of those work.

**MMM-Exchange** talks directly to Exchange via EWS SOAP API with full NTLM handshake support. No cloud, no OAuth, no iCal export — just point it at your Exchange server and it shows your upcoming events.

---

## Features

- **EWS SOAP API** — direct Exchange Web Services integration
- **NTLM authentication** — full 3-step Negotiate/NTLM handshake
- **Smart username parsing** — auto-detects `user@domain.com` or `DOMAIN\user` format
- **Current event highlighting** — ongoing events are visually emphasized
- **Fade-out effect** — long event titles gracefully fade instead of hard clipping
- **Configurable** refresh interval, event count, date range, location display
- **Minimal dependencies** — only 2 packages (xmldom + httpntlm)

---

## Installation

### 1. Clone the module

```bash
cd ~/MagicMirror/modules
git clone https://github.com/bohemtucsok/MMM-Exchange.git
cd MMM-Exchange
npm install
```

### 2. Configure MagicMirror

Add to your `~/MagicMirror/config/config.js`:

```javascript
{
    module: "MMM-Exchange",
    position: "top_left",
    config: {
        username: "user@domain.com",
        password: "password",
        host: "https://your-exchange-server",
        allowInsecureSSL: true
    }
}
```

### 3. Restart MagicMirror

```bash
pm2 restart magicmirror
```

---

## Configuration

| Option | Default | Description |
|--------|---------|-------------|
| `username` | `""` | Exchange username (`user@domain.com` or `DOMAIN\user`) |
| `password` | `""` | Exchange password |
| `host` | `""` | Exchange server URL (e.g. `https://exchange.company.com`) |
| `domain` | `""` | AD domain name (optional, auto-detected from username) |
| `updateInterval` | `300000` (5min) | Calendar refresh interval (ms) |
| `maxEvents` | `5` | Maximum number of displayed events |
| `daysToFetch` | `14` | Number of days to fetch ahead |
| `allowInsecureSSL` | `false` | Allow self-signed SSL certificates |
| `showLocation` | `true` | Show event location below the title |
| `showEnd` | `true` | Show event end time |
| `header` | `"Exchange Calendar"` | Module header text |
| `animationSpeed` | `1000` (1s) | DOM update animation speed (ms) |

---

## Username Formats

The module automatically detects and parses your username:

| Input | Parsed Username | Parsed Domain |
|-------|----------------|---------------|
| `john@company.com` | `john` | `COMPANY` |
| `COMPANY\john` | `john` | `COMPANY` |

If you set the `domain` config option explicitly, it overrides automatic detection.

---

## How It Works

| Layer | Technology |
|-------|-----------|
| Frontend | MagicMirror² Module API (DOM, table rendering) |
| Backend | Node.js node_helper (socket notifications) |
| Auth | NTLM 3-step handshake via httpntlm |
| API | EWS SOAP XML (FindItem + CalendarView) |
| XML Parsing | @xmldom/xmldom |

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| HTTP 401 Unauthorized | Check username/password. Verify the Exchange server supports NTLM. |
| SOAP Fault / Parse Error | Check the `host` URL — the module appends `/EWS/Exchange.asmx` automatically. |
| SSL Error | For self-signed certificates, set `allowInsecureSSL: true` |
| No events showing | Check the date range (`daysToFetch`) and `pm2 logs magicmirror` |
| Behind a proxy | NTLM handshake may break through reverse proxies. Use direct network access. |

---

## Requirements

- MagicMirror² v2.20.0 or later
- On-premise Microsoft Exchange server with EWS enabled
- Direct network access to the Exchange server (proxy not recommended)

---

## Dependencies

| Package | Purpose |
|---------|---------|
| [@xmldom/xmldom](https://www.npmjs.com/package/@xmldom/xmldom) | XML parsing |
| [httpntlm](https://www.npmjs.com/package/httpntlm) | NTLM authentication |

---

## License

[MIT](LICENSE) — use it, fork it, self-host it. Contributions welcome.
