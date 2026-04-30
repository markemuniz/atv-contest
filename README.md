# ATV Contest App — The Travel Agency DTBK

## What it does
- `/` — Public leaderboard (budtenders see this, read-only, auto-refreshes every minute)
- `/admin` — Password-protected upload page (only you have access)

## Setup

### 1. Install dependencies
```bash
npm install
```

### 2. Set environment variables
Create a `.env` file or set these on your hosting platform:
```
ADMIN_PASSWORD=your_secure_password_here
JWT_SECRET=some_random_long_string_here
PORT=3000
```

### 3. Run locally
```bash
node server.js
```

### 4. Deploy (recommended: Railway)
1. Go to https://railway.app and create a free account
2. Click "New Project" → "Deploy from GitHub repo"
3. Push this folder to a GitHub repo and connect it
4. In Railway, go to Variables and set:
   - `ADMIN_PASSWORD` = your password
   - `JWT_SECRET` = any random string (e.g. use https://randomkeygen.com)
5. Railway auto-detects Node.js and runs `node server.js`
6. Your app gets a public URL like `https://atv-contest.up.railway.app`

## Daily workflow
1. Go to `yoursite.com/admin`
2. Enter your password
3. Upload the day's sales report (.xlsx) + attributed orders (.csv)
4. Click Upload & Process
5. Budtenders visit `yoursite.com` to see the updated leaderboard

## Data
All processed data is stored in `data/leaderboard.json`.
On Railway, data persists as long as the deployment is active.
For permanent storage, consider adding a database (Railway offers free Postgres).

## Security
- Budtenders only access `/` — no upload or delete buttons
- Admin page at `/admin` requires a password
- Sessions expire after 24 hours
- Change `ADMIN_PASSWORD` and `JWT_SECRET` before deploying
