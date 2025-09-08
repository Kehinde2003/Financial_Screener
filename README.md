# Financial_Screener
A financial screener dashboard using Python &amp; Tableau to score companies across financial, ESG, and tech metrics.
## Features
- **Financial Scoring:** Calculates scores for ROE, FCF, D/E, and Revenue Growth, then aggregates them into a Financial Score.  
- **Tech Readiness:** Sector-specific metric that combines R&D, IT spend as a percentage of revenue, or CapEx, depending on the industry.  
- **ESG:** ESG values are sourced from Yahoo Finance.  
- **Final Score & Classification:** Aggregates all metrics to classify companies as *Strong Buy*, *Buy*, *Hold*, or *Avoid*.  
- **Interactive Dashboard:** Built in Tableau, with filters for sector, classification, and score ranges.

## Files
- `Data_cleaned.xlsx` – Cleaned dataset prepared for scoring.  
- `companies_scores.xlsx` – Processed dataset with calculated scores for dashboard use.  
- `score_calculation.py` – Python script for data processing and scoring logic.  
- `Financial_Screener.twbx` – Tableau dashboard file with interactive visuals.  
- `README.md` – Project overview and usage instructions.

## Usage
1. Run `score_calculation.py` with `Data_cleaned.xlsx` as input to generate `companies_scores.xlsx`.  
2. Open `Financial_Screener.twbx` in Tableau to explore the interactive dashboard.  
3. Use filters to analyze companies by sector, classification, or specific scoring metrics.  
