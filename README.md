# Zebra ZD421 Localhost Label Dashboard

This folder converts the three Excel label scripts into a local web page.

Open this on the computer that can reach the Zebra printer at `192.168.1.126:9100`.

## Install

```bash
npm install
```

## Run

```bash
npm start
```

Then open:

```text
http://localhost:3000
```

## Modes

- `test`: first valid Excel row only, quantity forced to 1.
- `first`: first valid Excel row only, uses Excel quantity.
- `all`: every valid Excel row, uses Excel quantity.
- Preview ZPL never prints. It only builds the ZPL and shows it on screen.

## Files

- `server.js`: localhost API server.
- `public/index.html`: dashboard page.
- `lib/printer.js`: Zebra TCP/IP printer sender.
- `labels/product.js`: product label ZPL builder.
- `labels/location.js`: location label ZPL builder.
- `labels/equipment.js`: equipment label ZPL builder.
- `data/*.xlsx`: Excel sources.
- `original-scripts/`: backups of the original scripts.

## Printer IP

Default printer target is:

```text
192.168.1.126:9100
```

You can override it when launching:

```bash
PRINTER_IP=192.168.1.126 PRINTER_PORT=9100 npm start
```

On Windows PowerShell:

```powershell
$env:PRINTER_IP="192.168.1.126"; $env:PRINTER_PORT="9100"; npm start
```
"# Zebra_Barcoding_Function_Test" 
"# label-printer-localhost-labelary" 
