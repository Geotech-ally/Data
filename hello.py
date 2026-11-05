import numpy as np
import pandas as pd
import os

os.makedirs('data', exist_ok=True)
rng = np.random.default_rng(42)
n = 300
prices = {'Maize Flour': 150, 'Cooking Oil': 320, 'Sugar': 180,
          'Rice': 210, 'Soap': 90, 'Milk': 60}

df = pd.DataFrame({
    'order_id': np.arange(1001, 1001 + n),
    'order_date': pd.to_datetime(rng.choice(pd.date_range('2025-01-01', '2025-12-31'), n)),
    'branch': rng.choice(['Kisumu', 'Siaya', 'Bondo', 'Kakamega'], n),
    'product': rng.choice(list(prices), n),
    'quantity': rng.integers(1, 20, n).astype(float),
    'payment': rng.choice(['M-Pesa', 'Cash', 'Card'], n, p=[0.6, 0.3, 0.1]),
    'customer_rating': rng.integers(1, 6, n).astype(float),
})
df['unit_price'] = df['product'].map(prices)
df['order_date'] = df['order_date'].dt.strftime('%Y-%m-%d')

# ---- add deliberate mess ----
df.loc[rng.choice(n, 20, replace=False), 'customer_rating'] = np.nan
df.loc[rng.choice(n, 6, replace=False), 'quantity'] = np.nan
idx = rng.choice(n, 15, replace=False)
df.loc[idx, 'branch'] = df.loc[idx, 'branch'].str.lower() + ' '
df.loc[rng.choice(n, 10, replace=False), 'payment'] = 'mpesa'
df.loc[7, 'order_date'] = 'not recorded'
df = pd.concat([df, df.sample(8, random_state=1)], ignore_index=True)

df.to_csv('data/sales.csv', index=False)

# ---- Excel workbook with 4 sheets ----
products = pd.DataFrame({
    'product': list(prices),
    'category': ['Grains', 'Oils', 'Baking', 'Grains', 'Household', 'Dairy'],
    'supplier': ['Unga Ltd', 'Pwani Oils', 'Mumias Sugar', 'Pishori Co', 'Bar Soap Co', 'Brookside'],
})
branches = pd.DataFrame({
    'branch': ['Kisumu', 'Siaya', 'Bondo', 'Kakamega'],
    'county': ['Kisumu', 'Siaya', 'Siaya', 'Kakamega'],
    'manager': ['Achieng', 'Otieno', 'Atieno', 'Wafula'],
})
with pd.ExcelWriter('data/shop_report.xlsx') as w:
    df.to_excel(w, sheet_name='Sales', index=False)
    products.to_excel(w, sheet_name='Products', index=False)
    branches.to_excel(w, sheet_name='Branches', index=False)
    df.head(30).to_excel(w, sheet_name='Q1_Report', index=False, startrow=3)
    w.sheets['Q1_Report']['A1'] = 'Q1 Sales Report - Shop Chain Ltd'
    w.sheets['Q1_Report']['A2'] = 'Prepared by head office'
print('Done: data/sales.csv and data/shop_report.xlsx created')