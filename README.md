# MMM-Exchange

MagicMirror² module for displaying on-premise Microsoft Exchange calendar events via EWS (Exchange Web Services) SOAP API with NTLM authentication.

## Features

- Fetches calendar events from Exchange using EWS SOAP API
- NTLM authentication (Negotiate/NTLM)
- Automatic username/domain parsing (`user@domain.com` or `DOMAIN\user` format)
- Configurable refresh interval, number of displayed events, and time range
- Current event highlighting
- Long text fade-out effect with CSS gradient
- Optional location display

## Installation

```bash
cd ~/MagicMirror/modules
git clone https://github.com/bohemtucsok/MMM-Exchange.git
cd MMM-Exchange
npm install
```

## Configuration

Add the following to the modules array in your MagicMirror `config/config.js`:

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

### Configuration Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `username` | String | `""` | Exchange username. Supported formats: `user@domain.com` or `DOMAIN\user` |
| `password` | String | `""` | Exchange password |
| `host` | String | `""` | Exchange server URL (e.g. `https://exchange.company.com`) |
| `domain` | String | `""` | AD domain name (optional, if not derivable from username) |
| `updateInterval` | Number | `300000` | Refresh interval in milliseconds (default: 5 minutes) |
| `maxEvents` | Number | `5` | Maximum number of displayed events |
| `daysToFetch` | Number | `14` | Number of days to fetch ahead |
| `allowInsecureSSL` | Boolean | `false` | Allow self-signed SSL certificates |
| `showLocation` | Boolean | `true` | Show event location |
| `showEnd` | Boolean | `true` | Show event end time |
| `header` | String | `"Exchange Calendar"` | Module header text |
| `animationSpeed` | Number | `1000` | DOM update animation speed (ms) |

### Username Formats

The module automatically detects and parses the username:

- **Email format**: `john@company.com` → username: `john`, domain: `COMPANY`
- **Domain\user format**: `COMPANY\john` → username: `john`, domain: `COMPANY`
- **Explicit domain**: if you set the `domain` config option, it overrides automatic detection

## Requirements

- MagicMirror² v2.20.0 or later
- On-premise Microsoft Exchange server with EWS support
- Direct network access to the Exchange server (proxy not recommended as it may break the NTLM 3-step handshake)

## Troubleshooting

### HTTP 401 Unauthorized
- Check your username and password
- Make sure the Exchange server supports NTLM authentication
- If connecting through a proxy, it may interfere with the NTLM handshake. Use a direct connection instead.

### SOAP Fault / Parse Error
- Verify the `host` URL (the module automatically appends `/EWS/Exchange.asmx`)
- Check the MagicMirror log: `pm2 logs magicmirror`

### SSL Error
- For self-signed certificates, set: `allowInsecureSSL: true`

## Dependencies

- [@xmldom/xmldom](https://www.npmjs.com/package/@xmldom/xmldom) - XML parsing
- [httpntlm](https://www.npmjs.com/package/httpntlm) - NTLM authentication

## License

MIT
