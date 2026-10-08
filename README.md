# Kenya diaspora remittances: CBK data and trend charts

Code to reproduce the charts in the article [A record month, a shrinking US corridor: what Kenya's 2026 remittance data are telling us](https://ellyokinyo.com/writing/kenya-remittances-2026.html) by Elly Okinyo (28 September 2026).

The repository holds code only. No data files are committed: the script downloads the data from the Central Bank of Kenya each time it runs, and `data/` is git-ignored.

`remittances_trend.py` downloads the Central Bank of Kenya (CBK) diaspora remittance
data from <https://www.centralbank.go.ke/diaspora-remittances/>, converts it to
USD millions, writes CSVs and draws three charts. No API keys needed.

## Setup (once)

Requires Python 3.11 or newer.

**Windows (PowerShell or Command Prompt)**

    cd path\to\kenya-remittances
    python -m venv .venv
    .venv\Scripts\activate
    pip install -r requirements.txt

**macOS (Terminal)**

    cd path/to/kenya-remittances
    python3 -m venv .venv
    source .venv/bin/activate
    pip install -r requirements.txt

## Run

With the venv activated:

    python remittances_trend.py

(Next time, just activate the venv again and run the script.) It prints a short summary:
latest month and value, trailing 12-month total, and the USA January-to-latest-month
year-on-year change.

## Outputs (created in `data/` next to the script)

- `data/kenya_remittances_tidy.csv`: one row per date and series: `date, series, group, usd_millions, source_table`
- `data/kenya_remittances_wide.csv`: the same data with one column per series
- `data/charts/1_monthly_total_with_12m_sum.png`: monthly total since 2015, plus the 12-month rolling sum in a lower panel
- `data/charts/2_usa_canada_europe.png`: USA, Canada and Europe, monthly values plus 12-month averages
- `data/charts/3_usa_share_of_total.png`: USA share of total diaspora remittances, in % (from Feb 2019, when the CBK country-level Excel file starts)
- `data/raw/`: the downloaded Excel file and a copy of the CBK web page

The charts are 300 dpi PNGs. If the data runs past 1 Jan 2026, a dashed vertical line at
1 Jan 2026 is labelled "1 Jan 2026: US 1% tax on cash-funded remittances". Every chart's
footer reads "Source: Central Bank of Kenya; author's calculations." (other chart notes are
printed to the console instead).

## If the download fails

1. Open <https://www.centralbank.go.ke/diaspora-remittances/> in your browser.
2. Click **"Remittances by Source ('000 USD Equivalent)"** and save the Excel file into
   `data/raw/` (inside this folder; create it if it doesn't exist).
3. Run the script again. If it can't download anything, it uses the newest `.xlsx`/`.xls` file in `data/raw/`.

If CBK ever goes back to an old-style `.xls` file, run `pip install xlrd` once.

## Data notes

- CBK publishes the figures in USD thousands. The script converts them to USD millions.
- The Excel file breaks remittances down by country from Feb 2019 onward. The long-run
  monthly totals (since 2004) and the North America / Europe / Rest of World split come
  from the table on the CBK web page.
- "USA" means the CBK's U.S.A row only. "North America" (web-page table) also includes
  Canada and other Americas. "Europe" means the CBK's Total Europe, which includes the UK and Switzerland.
- CBK revises recent months now and then. Re-run the script before you publish.

## Data sources

- Central Bank of Kenya, Diaspora Remittances page (monthly totals since 2004 and the North America / Europe / Rest of World split): <https://www.centralbank.go.ke/diaspora-remittances/>
- Central Bank of Kenya, "Remittances by Source ('000 USD Equivalent)" Excel file, linked from that page (country-level figures from Feb 2019). The August 2026 release used for the article: <https://www.centralbank.go.ke/wp-content/uploads/2026/09/August2026.xlsx>
