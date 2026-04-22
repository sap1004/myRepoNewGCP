import pandas as pd
from pathlib import Path

# ----------------------------
# Input file
# ----------------------------
input_file = "cmc_BIAS_summary_report_REGENERATED_REALISTIC.xlsx"

df = pd.read_excel(input_file, sheet_name="Sheet1")

# ----------------------------
# Step 1: Normalize Test Case Name
# ----------------------------
df['Test Case Name'] = (
    df['Test Case Name']
    .astype(str)
    .str.replace("English_US", "English", regex=False)
    .str.replace("English_UK", "English", regex=False)
)

# ----------------------------
# Step 2: Extract Accent & Gender (FIXED)
# Handles cases like:
# xxx_xxx_Australian_FEMALE_Chunk01
# xxx_xxx_English_MALE_Chunk09
# ----------------------------
def extract_accent_gender(name):
    if not isinstance(name, str) or name.strip() == "":
        return pd.Series([None, None])

    parts = name.upper().split('_')

    gender = None
    accent = None

    # Detect gender safely
    for p in parts:
        if p in ["MALE", "FEMALE"]:
            gender = p

    # Detect accent (ignore gender & chunk)
    for p in reversed(parts):
        if p not in ["MALE", "FEMALE"] and not p.startswith("CHUNK"):
            accent = p
            break

    return pd.Series([accent.title(), gender])


df[['Accent', 'Gender']] = df['Test Case Name'].apply(extract_accent_gender)

# ----------------------------
# Step 3: Accent + Gender + Intent
# ----------------------------
grouped = df.groupby(
    ['Accent', 'Gender', 'Expected Head Intent']
).agg(
    Total_Count=('Intent Matching (Pass/Fail)', 'count'),
    Pass_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Pass').sum()),
    Fail_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Fail').sum())
).reset_index()

grouped['Pass Pct'] = (
    grouped['Pass_Count'] / grouped['Total_Count'] * 100
).round(1).astype(str) + '%'

grouped['Fail Pct'] = (
    grouped['Fail_Count'] / grouped['Total_Count'] * 100
).round(1).astype(str) + '%'

# ----------------------------
# Accent + Gender
# ----------------------------
grouped_Acc_gender = df.groupby(
    ['Accent', 'Gender']
).agg(
    Total_Count=('Intent Matching (Pass/Fail)', 'count'),
    Pass_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Pass').sum()),
    Fail_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Fail').sum())
).reset_index()

grouped_Acc_gender['Pass Pct'] = (
    grouped_Acc_gender['Pass_Count'] / grouped_Acc_gender['Total_Count'] * 100
).round(1).astype(str) + '%'

grouped_Acc_gender['Fail Pct'] = (
    grouped_Acc_gender['Fail_Count'] / grouped_Acc_gender['Total_Count'] * 100
).round(1).astype(str) + '%'

# ----------------------------
# Gender
# ----------------------------
grouped_gender = df.groupby(
    ['Gender']
).agg(
    Total_Count=('Intent Matching (Pass/Fail)', 'count'),
    Pass_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Pass').sum()),
    Fail_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Fail').sum())
).reset_index()

grouped_gender['Pass Pct'] = (
    grouped_gender['Pass_Count'] / grouped_gender['Total_Count'] * 100
).round(1).astype(str) + '%'

grouped_gender['Fail Pct'] = (
    grouped_gender['Fail_Count'] / grouped_gender['Total_Count'] * 100
).round(1).astype(str) + '%'

# ----------------------------
# Accent
# ----------------------------
grouped_accent = df.groupby(
    ['Accent']
).agg(
    Total_Count=('Intent Matching (Pass/Fail)', 'count'),
    Pass_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Pass').sum()),
    Fail_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Fail').sum())
).reset_index()

grouped_accent['Pass Pct'] = (
    grouped_accent['Pass_Count'] / grouped_accent['Total_Count'] * 100
).round(1).astype(str) + '%'

grouped_accent['Fail Pct'] = (
    grouped_accent['Fail_Count'] / grouped_accent['Total_Count'] * 100
).round(1).astype(str) + '%'

# ----------------------------
# Intent
# ----------------------------
grouped_intent = df.groupby(
    ['Expected Head Intent']
).agg(
    Total_Count=('Intent Matching (Pass/Fail)', 'count'),
    Pass_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Pass').sum()),
    Fail_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Fail').sum())
).reset_index()

grouped_intent['Pass Pct'] = (
    grouped_intent['Pass_Count'] / grouped_intent['Total_Count'] * 100
).round(1).astype(str) + '%'

grouped_intent['Fail Pct'] = (
    grouped_intent['Fail_Count'] / grouped_intent['Total_Count'] * 100
).round(1).astype(str) + '%'

# ----------------------------
# Gender + Intent
# ----------------------------
grouped_gender_intent = df.groupby(
    ['Gender', 'Expected Head Intent']
).agg(
    Total_Count=('Intent Matching (Pass/Fail)', 'count'),
    Pass_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Pass').sum()),
    Fail_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Fail').sum())
).reset_index()

grouped_gender_intent['Pass Pct'] = (
    grouped_gender_intent['Pass_Count'] / grouped_gender_intent['Total_Count'] * 100
).round(1).astype(str) + '%'

grouped_gender_intent['Fail Pct'] = (
    grouped_gender_intent['Fail_Count'] / grouped_gender_intent['Total_Count'] * 100
).round(1).astype(str) + '%'

# ----------------------------
# Accent + Intent
# ----------------------------
grouped_accent_intent = df.groupby(
    ['Accent', 'Expected Head Intent']
).agg(
    Total_Count=('Intent Matching (Pass/Fail)', 'count'),
    Pass_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Pass').sum()),
    Fail_Count=('Intent Matching (Pass/Fail)', lambda x: (x == 'Fail').sum())
).reset_index()

grouped_accent_intent['Pass Pct'] = (
    grouped_accent_intent['Pass_Count'] / grouped_accent_intent['Total_Count'] * 100
).round(1).astype(str) + '%'

grouped_accent_intent['Fail Pct'] = (
    grouped_accent_intent['Fail_Count'] / grouped_accent_intent['Total_Count'] * 100
).round(1).astype(str) + '%'

# ----------------------------
# Step 4: Write to Excel
# ----------------------------
with pd.ExcelWriter(
    input_file,
    engine='openpyxl',
    mode='a',
    if_sheet_exists='replace'
) as writer:
    grouped.to_excel(writer, sheet_name='summary', index=False)
    grouped_Acc_gender.to_excel(writer, sheet_name='Acc_gender', index=False)
    grouped_gender.to_excel(writer, sheet_name='gender', index=False)
    grouped_accent.to_excel(writer, sheet_name='accent', index=False)
    grouped_intent.to_excel(writer, sheet_name='intent', index=False)
    grouped_gender_intent.to_excel(writer, sheet_name='gender_intent', index=False)
    grouped_accent_intent.to_excel(writer, sheet_name='accent_intent', index=False)

print("✅ Summary sheets generated successfully.")
