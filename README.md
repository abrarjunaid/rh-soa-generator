# Radiant Homes — SOA Generator

Web app that generates Owner Statement PDFs from P&L workbooks.

## How It Works

1. Upload your P&L workbook (.xlsx)
2. Select the month
3. Choose which units to generate (active units auto-selected)
4. Optionally upload your logo
5. Click Generate → Download all PDFs as a zip

## Run Locally

```bash
# Install dependencies
pip install -r requirements.txt
playwright install chromium

# Run
python app.py

# Open http://localhost:5000
```

## Deploy to Railway

1. Push this folder to a GitHub repo
2. Go to [railway.app](https://railway.app)
3. New Project → Deploy from GitHub repo
4. It will auto-detect the Dockerfile and deploy
5. Your app will be live at `https://your-app.up.railway.app`

## Deploy to Render

1. Push to GitHub
2. Go to [render.com](https://render.com)
3. New Web Service → Connect your repo
4. Environment: Docker
5. It will build and deploy automatically

## Deploy to Fly.io

```bash
flyctl launch
flyctl deploy
```

## Stack

- **Flask** — web framework
- **openpyxl** — Excel parsing (reads P&L workbook with formulas pre-calculated)
- **Playwright + Chromium** — pixel-perfect PDF rendering with Outfit font
- **Docker** — containerized for easy deployment

## File Structure

```
soa-app/
├── app.py              # Flask app + all SOA logic
├── templates/
│   └── index.html      # Frontend UI
├── requirements.txt    # Python dependencies
├── Dockerfile          # Container config
├── railway.toml        # Railway deployment config
└── README.md           # This file
```
