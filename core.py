import pandas as pd
import numpy as np
from collections import OrderedDict

# ---------- Load Data ----------

df = pd.read_excel("Copy of Global E-Commerce Weighting Data.xlsx")

# Compute Free Float Market Cap (used to calculate weights)

df['Free_Float_MCap'] = df['Security Level Mcap'] * df['FF']

# ---------- Stepwise Capping Function ----------

def apply_stepwise_capping(df, max_iter=30):

    df = df.sort_values(by='Free_Float_MCap', ascending=False).reset_index(drop=True).copy()

    capping_rules = {0: 0.08, 1: 0.08, 2: 0.07, 3: 0.065, 4: 0.06, 5: 0.055, 6: 0.05}

    default_cap = 0.045

    df['Cap'] = df.index.map(capping_rules).fillna(default_cap)

    df['Capped_Weight'] = np.minimum(df['Initial_Weight'], df['Cap'])

    for _ in range(max_iter):

        total_weight = df['Capped_Weight'].sum()

        if np.isclose(total_weight, 1.0, atol=1e-8):

            break

        excess = 1.0 - total_weight

        eligible = df[df['Capped_Weight'] < df['Cap']]

        if eligible.empty or np.isclose(excess, 0, atol=1e-8):

            break

        total_initial_weight = eligible['Initial_Weight'].sum()

        proportions = eligible['Initial_Weight'] / total_initial_weight

        increment = proportions * excess

        df.loc[eligible.index, 'Capped_Weight'] += increment

        df['Capped_Weight'] = np.minimum(df['Capped_Weight'], df['Cap'])

    else:

        raise Exception("Capping stuck after max iterations.")

    return df

# ---------- Portfolio Builder with Stepwise Export ----------

def build_stepwise_portfolio(universe_df):

    steps = OrderedDict()

    full_df = universe_df.copy()

    full_df = full_df.sort_values(by='Mcap', ascending=False).reset_index(drop=True)

    attempt = 0

    while True:

        attempt += 1

        step_label = f"Step_{attempt}"

        # Select top 50 by Mcap

        top50 = full_df.head(50).copy()

        top50['Free_Float_MCap'] = top50['Security Level Mcap'] * top50['FF']

        top50['Initial_Weight'] = top50['Free_Float_MCap'] / top50['Free_Float_MCap'].sum()

        top50 = apply_stepwise_capping(top50)

        top50['Is_US'] = top50['Primary Listing'] == 'United States'

        top50['Capped_Weight_US'] = top50['Capped_Weight'] * top50['Is_US'].astype(float)

        top50['Cumulative_US_Weight'] = top50['Capped_Weight_US'].cumsum()

        steps[step_label] = top50.copy()

        us_total_weight = top50['Capped_Weight_US'].sum()

        if us_total_weight <= 0.50 + 1e-6:

            break

        breach_index = top50[top50['Cumulative_US_Weight'] > 0.50].index.min()
        # Step 1: Remove breached US securities
        us_to_remove = top50.iloc[breach_index:].copy()
        us_to_remove = us_to_remove[us_to_remove['Is_US']]
        us_names_to_remove = us_to_remove['Name'].tolist()
        # Step 2: Keep all original top50 securities except those US names removed
        survivors = top50[~top50['Name'].isin(us_names_to_remove)].copy()
        # Step 3: Add non-US replacements to bring it back to 50
        current_names = survivors['Name'].tolist()
        non_us_candidates = universe_df[
            (universe_df['Primary Listing'] != 'United States') &
            (~universe_df['Name'].isin(current_names))
        ].sort_values(by='Mcap', ascending=False)
        refill_count = 50 - len(survivors)
        refill = non_us_candidates.head(refill_count)
        # Step 4: Combine survivors + refill
        full_df = pd.concat([survivors, refill], ignore_index=True)
        full_df = full_df.sort_values(by='Mcap', ascending=False).reset_index(drop=True)
        
    return steps

# ---------- Run Stepwise Builder ----------

stepwise_sheets = build_stepwise_portfolio(df)

# ---------- Export to Excel ----------

with pd.ExcelWriter("stepwise_portfolio_build2.xlsx") as writer:

    for sheet_name, sheet_df in stepwise_sheets.items():

        sheet_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
 