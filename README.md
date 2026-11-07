# Laptop Black Friday price tracker

Tracks ~11 WoW-capable gaming laptops (budget €1,200–1,600) across DE / UK / IE shops.
Each day it records the cheapest in-stock price as a landed cost in Ireland, rates the deal, and emails you.

| File | What it is |
|---|---|
| `Laptop_Price_Tracker.xlsx` | **Open this.** Dashboard (ratings + advice), Daily Best (chart), Price History, Watchlist, Settings. Rebuilt every run. |
| `laptops.json` | The watchlist: models, WoW score, and your **Excellent / Good / Fair** price targets (landed €). Edit freely. |
| `config.json` | Email address, Black Friday dates, stores + VAT/shipping factors to Ireland. |
| `secrets.json` | Your Gmail **app password** (create from `secrets.example.json`). |
| `data/history.csv` | All price observations (source of truth). |
| `data/incoming/` | One JSON per day from the daily price check. |
| `reports/latest.html` | The same report that gets emailed. |

## How the verdict works
* **Rating:** landed price ≤ Excellent target → *Excellent*; ≤ Good → *Good*; ≤ Fair → *Fair*; above → *Poor*.
* **Advice:** Excellent → **BUY NOW**. Good → buy if it's the lowest price seen so far, otherwise wait for Black Friday (from BF week on: buy). Fair → wait. Poor → skip.
* **Landed price:** German shops' prices get +3.4% (19% DE VAT → 23% IE VAT) + shipping. UK prices: GBP→EUR (ECB rate), −20% UK VAT, +23% IE VAT, + courier fee.

## Commands
```
python tracker.py build                  # rebuild Excel + report from history
python tracker.py daily data/incoming/prices_YYYY-MM-DD.json   # add + build + email
python tracker.py email                  # resend today's email
```

## Email setup (one-off)
1. Google Account → Security → 2-Step Verification → **App passwords** → create one called "Laptop tracker".
2. Copy `secrets.example.json` to `secrets.json` and paste the 16-character password.
3. Run `python tracker.py email` to test.
