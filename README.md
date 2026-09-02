# ZENOX Number Bot

## What is included

- Member menu: Get Number / Wallet / Status / Support
- Admin-only dashboard
- Owner-only Manage Admins
- Add/delete/toggle authorized panel workers
- Automatic panel inventory synchronization
- XLSX number import/export
- Join requirements
- OTP-group metadata management
- Announcement broadcast
- Support ID: `@labibhosen74`
- Number detail page with Number / Prefix / OTP Group / Change Number / Back

## Important

This version deliberately does NOT capture, display, extract, or forward SMS/OTP verification codes.
The panel monitor is limited to authorized number-inventory data.

## Install

```bash
pip install -r requirements.txt
playwright install chromium
```

## Configure

1. Copy `.env.example` to `.env`.
2. Put your BotFather token in `BOT_TOKEN`.
3. Put the owner's numeric Telegram ID in `OWNER_ID`.
4. Keep `OWNER_USERNAME=labibhosen74` if that is the owner's username.
5. Start:

```bash
python main.py
```

## Panel setup

From Telegram:

Admin Panel -> Worker Management -> Add New Panel

The bot asks:

1. Name
2. URL
3. Username
4. Password

The panel must expose an HTML table where the first three cells are:

`number | country | status`

If your panel uses different selectors/layout, edit the selector variables in `.env` and adapt `PanelMonitor.read_inventory()`.

## XLSX

Recommended header:

`number | country | prefix | otp_group | panel`

Example row:

`+15550000001 | US | ON | Group A | Panel 1`
