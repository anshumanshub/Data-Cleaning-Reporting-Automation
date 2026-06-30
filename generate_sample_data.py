"""
generate_sample_data.py
------------------------
Creates a deliberately messy CSV (input/sales_data.csv) so the pipeline
has something realistic to clean: duplicates, missing values, inconsistent
date formats, inconsistent phone formats, and numeric outliers.
"""

import numpy as np
import pandas as pd

np.random.seed(42)

n = 200
names = [f"Customer {i}" for i in range(1, n + 1)]
dates = pd.date_range("2024-01-01", periods=n, freq="D").astype(str).tolist()

# Mix up date formats to simulate inconsistency
mixed_dates = []
for i, d in enumerate(dates):
    dt = pd.to_datetime(d)
    if i % 4 == 0:
        mixed_dates.append(dt.strftime("%m/%d/%Y"))
    elif i % 4 == 1:
        mixed_dates.append(dt.strftime("%d-%b-%Y"))
    elif i % 4 == 2:
        mixed_dates.append(dt.strftime("%Y/%m/%d"))
    else:
        mixed_dates.append(dt.strftime("%B %d, %Y"))

phones = []
for i in range(n):
    digits = "".join(np.random.choice(list("0123456789"), 10))
    if i % 3 == 0:
        phones.append(f"({digits[:3]}) {digits[3:6]}-{digits[6:]}")
    elif i % 3 == 1:
        phones.append(f"{digits[:3]}.{digits[3:6]}.{digits[6:]}")
    else:
        phones.append(digits)

amounts = np.random.normal(250, 60, n).round(2)
# Inject some outliers
amounts[5] = 5400.00
amounts[50] = -120.00
amounts[120] = 8800.00

df = pd.DataFrame({
    "customer_name": names,
    "order_date": mixed_dates,
    "phone_number": phones,
    "region": np.random.choice(["North", "South", "East", "West", None], n, p=[0.24, 0.24, 0.24, 0.24, 0.04]),
    "order_amount": amounts,
    "quantity": np.random.randint(1, 20, n),
})

# Inject missing values
for col in ["order_date", "phone_number", "order_amount"]:
    idx = np.random.choice(df.index, size=8, replace=False)
    df.loc[idx, col] = np.nan

# Inject duplicate rows
dupes = df.sample(10, random_state=1)
df = pd.concat([df, dupes], ignore_index=True)

df.to_csv("input/sales_data.csv", index=False)
print(f"Sample data written to input/sales_data.csv ({len(df)} rows)")
