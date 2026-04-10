# CarJoy Pricing Comparison Tool

Upload two HOT SHEET files and get a complete Excel comparison report instantly.

## Setup & Deployment (takes about 10 minutes)

### Step 1 — Get the files onto GitHub

1. Go to github.com and create a free account if you don't have one
2. Click the **+** button → **New repository**
3. Name it `carjoy-comparison`, set it to **Private**, click **Create repository**
4. Click **uploading an existing file**
5. Upload both files: `app.py` and `requirements.txt`
6. Click **Commit changes**

### Step 2 — Deploy on Streamlit Community Cloud

1. Go to share.streamlit.io and sign in with your GitHub account
2. Click **New app**
3. Select your `carjoy-comparison` repository
4. Main file path: `app.py`
5. Click **Deploy**
6. It builds for ~2 minutes, then gives you a live URL

That's it. Share the URL and password with your team.

## Changing the Password

Open `app.py` and find this line near the top:

```python
PASSWORD = "carjoy2025"
```

Change it to whatever you want, save, and push to GitHub. Streamlit redeploys automatically.

## What the Tool Does

- Reads the HOT SHEET tab from both uploaded files
- Matches vehicles on Year + Make + Model + Trim
- Detects MSRP changes, payment changes, added/removed vehicles
- Flags data errors (bad MSRP, date in Make field, missing Year)
- Produces a formatted Excel with all tabs + Summary Dashboard
- Shows a live preview in the browser before downloading
