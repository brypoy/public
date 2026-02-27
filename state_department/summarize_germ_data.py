import pandas as pd

df = pd.read_csv("output/germany_data.csv")
df['year'] = df['Date'].astype(str).str[:4]

# Convert numeric columns
for col in df.columns:
    if col not in ['date', 'year', 'german post']:
        df[col] = pd.to_numeric(df[col], errors='coerce')

# Group by year, average numeric columns
summary = df.groupby('year').mean(numeric_only=True).round(2)
summary['count'] = df.groupby('year').size()

summary.to_csv("output/germany_summary_yearly.csv")
print("Saved to output/germany_summary_yearly.csv")
print(summary)