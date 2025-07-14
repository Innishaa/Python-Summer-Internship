import pandas as pd

# Load the data

df = pd.read_excel("Masterlist.xlsx")

# Convert First_Trade_Date to datetime

df['First_Trade_Date'] = pd.to_datetime(df['First_Trade_Date'], format='%d-%m-%Y', errors='coerce')

# Base date logic for TD_Eligibility

base_date = pd.to_datetime('2025-02-28')

min_date = base_date - pd.DateOffset(months=3)

max_date = base_date - pd.DateOffset(months=6)

# Apply individual checks

df['Geography_Check'] = (df['Listing_Country'] == 'United States') & (~df['Domicile_Country'].isin(['Ireland']))

df['MarketCap_Check'] = df['Company_Mcap_USD'] >= 1000

df['ADTV_Check'] = df['ADTV_3M'] >= 3

df['TD_Eligibility_Check'] = (

    (df['TD_Eligibility_6M'] == 'Yes') |

    ((df['TD_Eligibility_6M'] == 'No') &

     (df['TD_Eligibility_3M'] == 'Yes') &

     (df['First_Trade_Date'].between(min_date, max_date)))

)

df['Free_Float_Check'] = df['Sel_Free_Float'] >= 10

df['Price_Check'] = df['Price_USD'] < 10000

df['Security_Type_Check'] = df['Security_Type_2'].isin(['SHARE', 'ADR'])

df['Industry_Check'] = df['Factset_Industry'] == "Airlines"

# Final investability check before duplicate filter

df['Final_Composition_Check'] = (

    df['Geography_Check'] &

    df['MarketCap_Check'] &

    df['ADTV_Check'] &

    df['TD_Eligibility_Check'] &

    df['Free_Float_Check'] &

    df['Price_Check'] &

    df['Security_Type_Check'] &

    df['Industry_Check']

)

# Filter rows passing all checks

filtered_temp = df[df['Final_Composition_Check']].copy()

# Sort to prioritize highest ADTV_6M

filtered_temp = filtered_temp.sort_values('ADTV_6M', ascending=False)

# Apply duplicate check on Entity_ID only

filtered_temp['Duplicate_Check'] = ~filtered_temp.duplicated(subset=['Entity_ID'], keep='first')

# Merge Duplicate_Check back into original df

df = df.merge(

    filtered_temp[['Entity_ID', 'Duplicate_Check']],

    on='Entity_ID',

    how='left'

)

df['Duplicate_Check'] = df['Duplicate_Check'].fillna(False)

# Final filter with all checks including duplicate

filtered_df = df[df['Final_Composition_Check'] & df['Duplicate_Check']].copy()

# Select Top 10 by Company Market Cap

df_top10 = filtered_df.sort_values(by='Company_Mcap_USD', ascending=False).head(10).copy()

df_top10['Sel_Free_Float'] = df_top10['Sel_Free_Float'].fillna(0)

# Step: Calculate free-float adjusted market cap

df_top10['ff_mcap'] = df_top10['Security_Mcap_USD'] * (df_top10['Sel_Free_Float'] / 100)

total_ff_mcap = df_top10['ff_mcap'].sum()

df_top10['weight'] = df_top10['ff_mcap'] / total_ff_mcap

# Step: Apply 15% cap

cap = 0.15

df_top10['capped_weight'] = df_top10['weight'].clip(upper=cap)

excess = df_top10['weight'].sum() - df_top10['capped_weight'].sum()

df_top10['final_weight'] = df_top10['capped_weight']

# Redistribute excess weight

while excess > 1e-8:

    eligible = df_top10[df_top10['final_weight'] < cap].copy()

    eligible_total = eligible['final_weight'].sum()

    if eligible_total == 0:

        break

    redistribute = (eligible['final_weight'] / eligible_total) * excess

    temp = eligible['final_weight'] + redistribute

    temp_capped = temp.clip(upper=cap)

    new_excess = (temp - temp_capped).sum()

    df_top10.loc[eligible.index, 'final_weight'] = temp_capped

    excess = new_excess

# Step: Apply 1% floor and redistribute

floor = 0.01

too_low = df_top10['final_weight'] < floor

if too_low.any():

    shortfall = (floor - df_top10.loc[too_low, 'final_weight']).sum()

    df_top10.loc[too_low, 'final_weight'] = floor

    redistribute_from = df_top10[(~too_low) & (df_top10['final_weight'] < cap)]

    available_weight = redistribute_from['final_weight'].sum()

    if available_weight > 0:

        scaling_factor = (available_weight - shortfall) / available_weight

        df_top10.loc[redistribute_from.index, 'final_weight'] *= scaling_factor

    else:

        print("Warning: No securities available to redistribute the floor shortfall.")

# Assign ranks

df_top10['Rank'] = df_top10['final_weight'].rank(method='first', ascending=False).astype(int)

# Write final output to Excel

output_file = "final_airline_index_detailed3.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

    df.to_excel(writer, index=False, sheet_name='Full Universe')

    df_summary = df_top10[['Fsym_ID', 'final_weight', 'Rank']].sort_values(by='Rank').copy()

    df_summary.rename(columns={'final_weight': 'Index Weight (%)'}, inplace=True)

    df_summary['Index Weight (%)'] = (df_summary['Index Weight (%)'] * 100).round(2).astype(str) + '%'

    df_summary.to_excel(writer, index=False, sheet_name='Top 10 Airlines')
 