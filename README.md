# Cosmetic formulation crawler (Codespaces)

## 1. Install
```bash
pip install -r requirements.txt
```

## 2. Before running
- Check the target website's current Terms of Use and `/robots.txt`.
- Confirm permission for bulk collection and internal database / AI-ML use.
- Do not use this crawler to bypass login, CAPTCHA, Cloudflare, paywalls, or access controls.
- Replace `YOUR_EMAIL` in `crawler.py` with a contact address you control.
- The script intentionally stops unless `ALLOW_CRAWL=YES` is set.

## 3. Run a small test
```bash
export ALLOW_CRAWL=YES
python crawler.py --max-pages 5 --delay 3
```

## 4. Run selected categories
```bash
python crawler.py \
  --categories \
  https://www.personalcaremagazine.com/formulations-by-category/33/skin-treatment \
  https://www.personalcaremagazine.com/formulations-by-category/15/facial-care-cleansing \
  --max-pages 20 --delay 3
```

## 5. Or use a manually curated detail URL list
Create `detail_urls.txt`, one exact `/formulation-details/...` URL per line:
```text
https://www.personalcaremagazine.com/formulation-details/2905/moisturising-cream
```
Then run:
```bash
python crawler.py --categories --max-pages 100
```

## Output
- `skincare_formulations_crawled.xlsx`
  - `Formulation_Master`
  - `Ingredient_Detail`
  - `Manufacturing_Process`
  - `Crawl_Log`
- `formulations_master.csv`
- `ingredients_detail.csv`

## Important data quality note
The HTML layouts can differ between formulations. Review `Crawl_Log` and records marked
`성분표 구조 확인 필요`. Do not fill missing INCI, percentages, pH, viscosity, appearance,
or stability by guessing. The parser preserves the original amount text (e.g. `q.s.`,
`up to 100`, `as needed`) separately from the numeric amount.
