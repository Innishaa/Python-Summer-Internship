import pandas as pd
import numpy as np
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ---------- Load Data ----------
df = pd.read_excel("Copy of Global E-Commerce Weighting Data.xlsx")
df['Free_Float_MCap'] = df['Security Level Mcap'] * df['FF']

# ---------- Stepwise Capping Function ----------
def apply_stepwise_capping(df, max_iter=30):
    df = df.sort_values(by='Free_Float_MCap', ascending=False).reset_index(drop=True).copy()
    capping_rules = {
        0: 0.08, 1: 0.08, 2: 0.07, 3: 0.065,
        4: 0.06, 5: 0.055, 6: 0.05
    }
    default_cap = 0.045
    df['Cap'] = df.index.map(capping_rules).fillna(default_cap)
    df['Capped_Weight'] = np.minimum(df['Initial_Weight'], df['Cap'])
    for _ in range(max_iter):
        total_weight = df['Capped_Weight'].sum()
        if np.isclose(total_weight, 1.0, atol=1e-8):
            break
        excess = 1.0 - total_weight
        eligible = df[df['Capped_Weight'] < df['Cap']]
        total_room = (eligible['Cap'] - eligible['Capped_Weight']).sum()
        if eligible.empty or total_room == 0 or np.isclose(excess, 0, atol=1e-8):
            break
        df.loc[eligible.index, 'Capped_Weight'] += (
            (eligible['Cap'] - eligible['Capped_Weight']) / total_room
        ) * excess
        df['Capped_Weight'] = np.minimum(df['Capped_Weight'], df['Cap'])
    else:
        raise Exception("Capping stuck after max iterations.")
    return df

# ---------- Final Portfolio Builder ----------
def build_final_portfolio(universe_df, max_attempts=20):
    full_df = universe_df.copy()
    full_df = full_df.sort_values(by='Mcap', ascending=False).reset_index(drop=True)
    attempt = 0
    while attempt < max_attempts:
        attempt += 1
        logging.info(f"Attempt #{attempt}")
        top50 = full_df.iloc[:50].copy()
        if top50.shape[0] < 50:
            fill_count = 50 - top50.shape[0]
            existing_names = top50['Name'].tolist()
            refill_candidates = universe_df[
                (universe_df['Primary Listing'] != 'United States') &
                (~universe_df['Name'].isin(existing_names))
            ].sort_values(by='Mcap', ascending=False)
            fill_df = refill_candidates.head(fill_count)
            logging.info(f"Force-refilling {fill_df.shape[0]} non-US securities: {fill_df['Name'].tolist()}")
            fill_df['Forced_Refill'] = True
            top50['Forced_Refill'] = False
            top50 = pd.concat([top50, fill_df], ignore_index=True)
            top50['Free_Float_MCap'] = top50['Security Level Mcap'] * top50['FF']
            top50['Initial_Weight'] = top50['Free_Float_MCap'] / top50['Free_Float_MCap'].sum()
            top50 = apply_stepwise_capping(top50)
        else:
            top50['Forced_Refill'] = False
        top50['Free_Float_MCap'] = top50['Security Level Mcap'] * top50['FF']
        top50['Initial_Weight'] = top50['Free_Float_MCap'] / top50['Free_Float_MCap'].sum()
        top50 = apply_stepwise_capping(top50)
        top50['Is_US'] = top50['Primary Listing'] == 'United States'
        top50['Capped_Weight_US'] = top50['Capped_Weight'] * top50['Is_US'].astype(float)
        top50['Cumulative_US_Weight'] = top50['Capped_Weight_US'].cumsum()
        us_total_weight = top50.loc[top50['Is_US'], 'Capped_Weight'].sum()
        logging.info(f"US Exposure: {us_total_weight:.4f}")
        if us_total_weight <= 0.50 + 1e-6:
            logging.info("US exposure below threshold. Final portfolio ready.")
            break
        breach_index = top50[top50['Cumulative_US_Weight'] > 0.50].index.min()
        if pd.isna(breach_index):
            raise Exception("Could not detect US breach point.")
        us_slice = top50.iloc[breach_index:]
        us_to_remove = us_slice[us_slice['Is_US']]['Name'].tolist()
        logging.info(f"Removing {len(us_to_remove)} US securities: {us_to_remove}")
        full_df = full_df[~full_df['Name'].isin(us_to_remove)]
        current_names = full_df['Name'].unique()
        non_us_candidates = universe_df[
            (universe_df['Primary Listing'] != 'United States') &
            (~universe_df['Name'].isin(current_names))
        ].sort_values(by='Mcap', ascending=False)
        current_count = full_df.shape[0]
        if current_count < 50:
            fill_count = 50 - current_count
            logging.info(f"Current securities: {current_count}, Need to fill: {fill_count}")
            logging.info(f"Available non-US candidates: {non_us_candidates.shape[0]}")
            fill_df = non_us_candidates.head(fill_count)
            if not fill_df.empty:
                logging.info(f"Adding {fill_df.shape[0]} non-US securities: {fill_df['Name'].tolist()}")
                full_df = pd.concat([full_df, fill_df], ignore_index=True)
            else:
                logging.info("No non-US candidates available to fill the gap.")
        full_df = full_df.sort_values(by='Mcap', ascending=False).reset_index(drop=True)
    else:
        raise Exception("Failed to meet US exposure ≤ 50% in max attempts.")
    final_top50 = full_df.iloc[:50].copy()
    final_top50['Free_Float_MCap'] = final_top50['Security Level Mcap'] * final_top50['FF']
    final_top50['Initial_Weight'] = final_top50['Free_Float_MCap'] / final_top50['Free_Float_MCap'].sum()
    final_top50 = apply_stepwise_capping(final_top50)
    return final_top50

# ---------- Run Final Portfolio Builder ----------
final_df = build_final_portfolio(df)

# ---------- Final Formatting ----------
final_df['Final_Weight'] = final_df['Capped_Weight']
final_df['Final_Weight_%'] = final_df['Final_Weight'] * 100
final_df['Rank'] = final_df['Final_Weight'].rank(ascending=False, method='first').astype(int)
final_df = final_df.round({'Final_Weight': 6, 'Final_Weight_%': 4})
final_df = final_df.sort_values(by='Final_Weight', ascending=False)

# ---------- Final Checks ----------
assert final_df.shape[0] == 50, "Must have 50 securities."
assert np.isclose(final_df['Final_Weight'].sum(), 1.0, atol=1e-6), "Weights do not sum to 100%."
assert final_df['Final_Weight'].max() <= 0.08 + 1e-6, "Weight exceeds 8% cap."
us_weight = final_df.loc[final_df['Primary Listing'] == 'United States', 'Final_Weight'].sum()
assert us_weight <= 0.50 + 1e-6, "US exposure exceeds 50%."

# ---------- Export to Excel ----------
final_df.to_excel("final_weighted_index_final11.xlsx", index=False)
logging.info("Final index created: final_weighted_index_final11.xlsx")
