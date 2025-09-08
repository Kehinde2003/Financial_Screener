import pandas as pd

# 1. Load Excel and get Data_cleaned sheet
file_path = "Data_project.xlsx"
xls = pd.ExcelFile(file_path)
df = pd.read_excel(file_path, sheet_name="Data_cleaned")

# 2. Define quartiles for financial metrics
quartiles = {
    "ROE": [0.063707565, 0.122929415, 0.296961691],
    "DE": [0.276110675, 0.640944195, 1.157623378],
    "FCF": [326027.75, 1814208, 6886000],
    "RevGrowth": [0.01007, 0.0543, 0.1321]  
}

def quartile_score(value, q1, q2, q3):
    if value < q1: return 1
    elif value < q2: return 2
    elif value < q3: return 3
    else: return 4

# 3. Apply scoring
df["ROE_Score"] = df["ROE"].apply(lambda x: quartile_score(x, *quartiles["ROE"]))
df["DE_Score"] = df["DE"].apply(lambda x: quartile_score(x, *quartiles["DE"]))
df["FCF_Score"] = df["FCF"].apply(lambda x: quartile_score(x, *quartiles["FCF"]))
df["RevGrowth_Score"] = df["RevGrowth"].apply(lambda x: quartile_score(x, *quartiles["RevGrowth"]))

df["Financial_Score"] = df[["ROE_Score", "DE_Score", "FCF_Score", "RevGrowth_Score"]].sum(axis=1)

# 4. ESG Score
def ESG_Score(ESG):
    if ESG < 18.825: return 1
    elif ESG < 23.4: return 2
    elif ESG < 27.875: return 3
    else: return 4

df["ESG_Score"] = df["ESG"].apply(ESG_Score)

# 5. R&D score based on sector quartiles
sector_quartiles = {
    "Technology": [0.14, 0.1827, 0.2524],
    "Healthcare": [0.08545, 0.14665, 0.22385],
    "Consumer Goods": [0.001, 0.0069, 0.00999]
}

# 6. CapEx Score
def RD_Score(RD):
    if pd.isna(RD): return 0
    elif RD < 0.06554: return 1
    elif RD < 0.2365: return 2
    elif RD < 0.395675: return 3
    else: return 4

df["RD_Score"] = df["RD"].apply(RD_Score)

# 6. CapEx Score
def CapEx_Score(CapEx):
    if pd.isna(CapEx):   # if no value, return None
        return 0
    elif CapEx < 0.06554:
        return 1
    elif CapEx < 0.2365:
        return 2
    elif CapEx < 0.395675:
        return 3
    else:
        return 4

df['CapEx_Score'] = df['CapEx'].apply(CapEx_Score)

# 7. IT Spend Score
def IT_Score(IT_spend):
    if pd.isna(IT_spend): return 0
    elif IT_spend < 0.06554: return 1
    elif IT_spend < 0.2365: return 2
    elif IT_spend < 0.395675: return 3
    else: return 4

df["IT_Score"] = df["IT_spend"].apply(IT_Score)

# 8. Tech Readiness Score (RD + IT + CapEx)
df["TechReadiness_Score"] = df["RD_Score"].fillna(0) + df["IT_Score"].fillna(0) + df["CapEx_Score"].fillna(0)

# Replace NaN with 0 for the score columns
empty_columns = ["RD", "IT_Score","CapEx_Score", "IT_spend","CapEx","RD_Score"] 
df[empty_columns] = df[empty_columns].replace(['', None], 0).fillna(0)


# 9. Final Score
df["Final_Score"] = (
    0.5 * df["Financial_Score"].fillna(0) +
    0.2 * df["ESG_Score"].fillna(0) +
    0.3 * df["TechReadiness_Score"].fillna(0)
)

# 10. Classification
def classify(score):
    if score >= 7.5: return "Strong Buy"
    elif score >= 6: return "Buy"
    elif score >= 4: return "Hold"
    else: return "Avoid"

df["Classification"] = df["Final_Score"].apply(classify)

# 11. Select only relevant columns for companies_scores sheet (with raw scores)
df_scores = df[[
    "Company", "Sector",
    "ROE", "ROE_Score",
    "FCF", "FCF_Score",
    "DE", "DE_Score",
    "RevGrowth", "RevGrowth_Score",
    "Financial_Score",
    "ESG", "ESG_Score",
    "RD", "RD_Score",
    "CapEx", "CapEx_Score",
    "IT_spend", "IT_Score",
    "TechReadiness_Score",
    "Final_Score",
    "Classification"
]]

# 12. Save results into a new sheet (while keeping old sheets)
with pd.ExcelWriter(file_path, mode="a", if_sheet_exists="replace") as writer:
    df_scores.to_excel(writer, sheet_name="Companies_scores", index=False)