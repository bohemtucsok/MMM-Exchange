# MMM-Exchange

MagicMirror² modul on-premise Microsoft Exchange naptár és feladatlista megjelenítéséhez EWS (Exchange Web Services) SOAP API-n keresztül, NTLM autentikációval.

## Funkciók

- Exchange naptár események lekérése EWS SOAP API-val
- Exchange feladatlista (Tasks) megjelenítése
- NTLM autentikáció (Negotiate/NTLM)
- Automatikus username/domain parsing (`user@domain.com` vagy `DOMAIN\user` formátum)
- Konfiguálható frissítési intervallum, megjelenített események száma, időtartam
- Aktuális esemény kiemelése
- Feladat jelzések: lejárt határidő piros, magas prioritás villám ikonnal
- Hosszú szövegek fade-out effekttel való levágása
- Helyszín megjelenítése (opcionális)

## Telepítés

```bash
cd ~/MagicMirror/modules
git clone https://gitlab.onevps.hu/egyeb_fejlesztesek/magicmirror_exchange.git MMM-Exchange
cd MMM-Exchange
npm install
```

## Konfiguráció

A MagicMirror `config/config.js` fájljában add hozzá a modulok listájához:

```javascript
{
    module: "MMM-Exchange",
    position: "top_left",
    config: {
        username: "felhasznalo@domain.com",
        password: "jelszo",
        host: "https://exchange-szerver-neve",
        allowInsecureSSL: true,
        showTasks: true
    }
}
```

### Konfigurációs opciók

| Opció | Típus | Alapértelmezett | Leírás |
|-------|-------|-----------------|--------|
| `username` | String | `""` | Exchange felhasználónév. Támogatott formátumok: `user@domain.com` vagy `DOMAIN\user` |
| `password` | String | `""` | Exchange jelszó |
| `host` | String | `""` | Exchange szerver URL (pl. `https://exchange.cegnev.hu`) |
| `domain` | String | `""` | AD domain név (opcionális, ha a username-ből nem derül ki) |
| `updateInterval` | Number | `300000` | Frissítési intervallum milliszekundumban (alapértelmezett: 5 perc) |
| `maxEvents` | Number | `5` | Megjelenített események maximális száma |
| `daysToFetch` | Number | `14` | Előre lekérdezett napok száma |
| `allowInsecureSSL` | Boolean | `false` | Self-signed SSL tanúsítvány engedélyezése |
| `showLocation` | Boolean | `true` | Helyszín megjelenítése az esemény alatt |
| `showEnd` | Boolean | `true` | Befejezési időpont megjelenítése |
| `header` | String | `"Exchange Calendar"` | Modul fejléc szövege |
| `showTasks` | Boolean | `false` | Feladatlista megjelenítése a naptár alatt |
| `maxTasks` | Number | `10` | Megjelenített feladatok maximális száma |
| `animationSpeed` | Number | `1000` | DOM frissítés animáció sebessége (ms) |

### Username formátumok

A modul automatikusan felismeri és parse-olja a felhasználónevet:

- **Email formátum**: `gabor@cegnev.com` → username: `gabor`, domain: `CEGNEV`
- **Domain\user formátum**: `CEGNEV\gabor` → username: `gabor`, domain: `CEGNEV`
- **Explicit domain**: ha a `domain` konfigurációs opciót is megadod, az felülírja az automatikus felismerést

## Követelmények

- MagicMirror² v2.20.0 vagy újabb
- On-premise Microsoft Exchange szerver EWS támogatással
- Közvetlen hálózati elérés az Exchange szerverhez (proxy nem ajánlott, mert az NTLM handshake-et megzavarhatja)

## Hibaelhárítás

### HTTP 401 Unauthorized
- Ellenőrizd a felhasználónevet és jelszót
- Győződj meg róla, hogy az Exchange szerver támogatja az NTLM autentikációt
- Ha proxy-n keresztül csatlakozol, az megzavarhatja az NTLM 3-lépéses handshake-et. Használj közvetlen kapcsolatot.

### SOAP Fault / Parse error
- Ellenőrizd a `host` URL-t (a `/EWS/Exchange.asmx` végződést a modul automatikusan hozzáadja)
- Nézd meg a MagicMirror logot: `pm2 logs magicmirror`

### SSL hiba
- Self-signed tanúsítványnál állítsd be: `allowInsecureSSL: true`

## Függőségek

- [@xmldom/xmldom](https://www.npmjs.com/package/@xmldom/xmldom) - XML parsing
- [httpntlm](https://www.npmjs.com/package/httpntlm) - NTLM autentikáció

## Licenc

MIT
