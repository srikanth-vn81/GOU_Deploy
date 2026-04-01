import pandas as pd

input_path  = r"C:\GOU\GOU_final.xlsx"
output_path = r"C:\GOU\lob_pie.xlsx"

df = pd.read_excel(input_path, sheet_name="Sheet1")

df_crossskill = df[df["CROSS SKILLING - CHECK"] == 9].copy()

df_lob = (
    df_crossskill
    .groupby("Staffing Category")
    .size()
    .reset_index(name="Column1")
)

df_lob.to_excel(output_path, index=False)
print(df_lob.to_string(index=False))
print(f"\n✅ Exported → {output_path}")