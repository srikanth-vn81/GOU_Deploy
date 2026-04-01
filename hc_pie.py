import pandas as pd

input_path  = r"C:\GOU\GOU_final.xlsx"
output_path = r"C:\GOU\hc_pie.xlsx"

df = pd.read_excel(input_path, sheet_name="Sheet1")

df_crossskill = df[df["CROSS SKILLING - CHECK"] == 9].copy()
df_crossskill["Req HC"] = pd.to_numeric(df_crossskill["Req HC"], errors="coerce").fillna(0)

df_hc = (
    df_crossskill
    .groupby("Staffing Category", dropna=False)["Req HC"]
    .sum()
    .reset_index()
)

df_hc.to_excel(output_path, index=False)
print(df_hc.to_string(index=False))
print(f"\n✅ Exported → {output_path}")