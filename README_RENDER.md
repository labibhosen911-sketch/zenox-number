# ZENOX Number Bot — Render test deployment

This package is prepared for a Render **Web Service** test deployment.

## Render settings

- Runtime: Python 3
- Build Command: `pip install -r requirements.txt`
- Start Command: `python start_render.py`
- Plan: Free

Set these Environment Variables in Render:

- `BOT_TOKEN` = your BotFather token
- `OWNER_ID` = your numeric Telegram user ID
- `OWNER_USERNAME` = your Telegram username without `@`

The bot uses polling. `start_render.py` also exposes a small HTTP health endpoint so Render can detect the web service port.

## Important Render limitation

Render's free web services can spin down after 15 minutes without inbound traffic, and their local filesystem is ephemeral. The bot's local SQLite data can therefore be lost after a restart/spin-down. This package is for testing, not reliable production hosting.

This bot version does not capture, extract, display, or forward SMS/OTP verification codes.
