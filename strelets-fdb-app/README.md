# Strelets Integral FDB Mini App

Local web and CLI tools for browsing and editing the **Strelets Integral** Firebird database (`INTEGRAL.FDB`).

## What this is

`INTEGRAL.FDB` is the Firebird database used by **АРМ Стрелец-Интеграл** (Argus-Spectr). It stores:

- system configuration
- device states
- event archive
- users and topology
- graphic plan metadata

## Credentials

| Purpose | Default |
|---------|---------|
| Strelets app login (Конфигуратор) | `2047` / `1111` |
| Firebird SQL (this app) | `SYSDBA` / `masterkey` |

The Strelets login is shown on the dashboard for reference. SQL access uses Firebird credentials from `.env`.

## Requirements

- Python 3.10+
- Firebird 3.0 server

### Install Firebird (Debian/Ubuntu)

```bash
echo 'firebird3.0-server firebird3.0-server/twoauth boolean false' | sudo debconf-set-selections
echo 'firebird3.0-server firebird3.0-server/sysdba_password password masterkey' | sudo debconf-set-selections
echo 'firebird3.0-server firebird3.0-server/sysdba_password_again password masterkey' | sudo debconf-set-selections
sudo apt-get install -y firebird3.0-server firebird3.0-utils
sudo systemctl enable --now firebird3.0
```

## Setup

```bash
cd strelets-fdb-app
cp .env.example .env
pip install -r requirements.txt
```

Place your database file at:

```
strelets-fdb-app/data/INTEGRAL.FDB
```

Update `FDB_PATH` in `.env` if needed.

## Run web UI

```bash
cd strelets-fdb-app
uvicorn app:app --reload --host 0.0.0.0 --port 8080
```

Open: http://localhost:8080

## CLI

```bash
python cli.py info
python cli.py tables
python cli.py schema INT_EVENTS
python cli.py show INT_EVENTS --limit 20
```

## Firebird server notes (Linux)

If connection fails:

1. Ensure Firebird 3 is running: `sudo service firebird3.0 start`
2. Give the `firebird` user read/write access to the `.FDB` file:
   ```bash
   sudo chown firebird:firebird data/INTEGRAL.FDB
   sudo chmod 660 data/INTEGRAL.FDB
   ```
3. For older Strelets databases, enable legacy auth in `/etc/firebird/3.0/firebird.conf`:
   ```
   AuthServer = Legacy_Auth, Srp
   AuthClient = Legacy_Auth, Srp, Srp256
   UserManager = Legacy_UserManager, Srp
   WireCrypt = Enabled
   ```
   Then restart Firebird.


Direct SQL writes can break Strelets Integral integrity (triggers, linked config). For production databases:

1. Work on a **copy** of `INTEGRAL.FDB`
2. Set `READ_ONLY=true` in `.env` for browse-only mode

## Key tables

| Group | Tables |
|-------|--------|
| Events | `INT_EVENTS`, `INT_EVENT_PARSERS` |
| Config | `INT_CONFIGURATION`, `INT_CONFIGURATIONDATA`, `INT_SYSTEMCONFIG` |
| Devices | `INT_STATES`, `PARTITIONS`, `SEGMENTS` |
| Users | `USERS` |
