#!/usr/bin/env python
# coding: utf-8

# In[ ]:

import streamlit as st
import pandas as pd
from datetime import date

# === LOAD INPUT ===
task_df = pd.read_excel('DivisionFiles_All.xlsx')

# === DEFINE CONSTANTS ===
div_order = ['Div1', 'Div2', 'Div3', 'Div4']

# Outsource quotas
st.title("Outsource Quota Simulation")

# === USER INPUT for pf12_quotas ===
st.header("Set PF12 Quotas per Year")

# List of years you want users to adjust
pf12_years = [2026, 2027, 2028, 2029, 2030]

# Create an empty dict to hold user inputs
pf12_quotas = {}

for year in pf12_years:
    pf12_quotas[year] = st.number_input(
        f"PF12 Quota for {year}", 
        min_value=0, 
        value=2000 if year == 2027 else 1500, 
        step=100,
        key=f"pf12_{year}"  # unique key added
    )

st.write("### User-defined PF12 Quotas")
st.write(pf12_quotas)


# === USER INPUT for pf11_quotas ===
st.header("Set PF11 Quotas per Year")

# List of years you want users to adjust
pf11_years = [2026, 2027, 2028, 2029, 2030]

# Create an empty dict to hold user inputs
pf11_quotas = {}

for year in pf11_years:
    pf11_quotas[year] = st.number_input(
        f"PF11 Quota for {year}", 
        min_value=0, 
        value=2000 if year == 2027 else 1500, 
        step=100,
        key=f"pf11_{year}"  # unique key added
    )

st.write("### User-defined PF11 Quotas")
st.write(pf11_quotas)


# Thresholds
pf11_thresholds = {
    2025: date(2023, 1, 1),
    2026: date(2024, 1, 1),
    2027: date(2025, 1, 1),
    2028: date(2026, 7, 1),
    2029: date(2027, 7, 1),
    2030: date(2028, 7, 1),
}

pf12_thresholds = {
    2026: date(2024, 1, 1),
    2027: date(2025, 1, 1),
    2028: date(2026, 1, 1),
    2029: date(2027, 7, 1),
    2030: date(2028, 7, 1),
}


# === DATA PREP ===

# Split by S&E type
task_df['S&E Year'] = pd.to_numeric(task_df['S&E Year'], errors='coerce')
task_df_pf11 = task_df[task_df['S&E'] == 'PF11'].copy()
task_df_pf12 = task_df[task_df['S&E'] == 'PF12'].copy()

# Sort PF11 for consistent ordering
task_df_pf11.sort_values(by='S&E Lodge Date', inplace=True)
task_df_pf12.sort_values(by='S&E Lodge Date', inplace=True)


# Initialize outsource flags
task_df_pf11['Outsource S'] = task_df_pf11.get('Outsource S', '')
task_df_pf11['Outsource Year'] = task_df_pf11.get('Outsource Year', pd.NA)
task_df_pf12['Outsource E'] = task_df_pf12.get('Outsource E', '')
task_df_pf12['Outsource Year'] = task_df_pf12.get('Outsource Year', pd.NA)

# === COMMON FUNCTION ===

def apply_quotas_for_year(year, year_qty, task_df, date_threshold, flag_col):
    task_df_remaining = task_df[task_df[flag_col] != 'Y'].copy()
    #task_df_remaining = task_df[(task_df[flag_col] != 'Y') & (task_df['S&E Lodge Date'].dt.date < date_threshold)].copy()

    division_counts = task_df_remaining['Division Transformed'].value_counts().to_dict()
    total_now = len(task_df_remaining)
    div_shares = {div: division_counts.get(div, 0) / total_now for div in div_order}
    '''
    quotas = {}
    allocated = 0
    for div in div_order[:-1]:
        quotas[div] = int(div_shares[div] * year_qty)
        allocated += quotas[div]
    quotas['Div4'] = year_qty - allocated
    '''
    quotas = calculate_division_quotas(div_shares, year_qty, div_order)

    div_counts = {div: 0 for div in div_order}

    for index, task in task_df.iterrows():
        if task.get(flag_col) == 'Y':
            continue
        if task['S&E Lodge Date'].date() < date_threshold:
            continue

        current_div = task['Division Transformed']
        if div_counts.get(current_div, 0) >= quotas.get(current_div, 0):
            continue

        task_df.at[index, 'Outsource Year'] = year
        task_df.at[index, flag_col] = 'Y'
        div_counts[current_div] += 1

        if sum(div_counts.values()) == year_qty:
            break

    return task_df


def calculate_division_quotas(div_shares, year_qty, div_order):
    # Calculate raw (float) quotas
    raw_quotas = {div: div_shares.get(div, 0) * year_qty for div in div_order}

    # Round down to integers
    quotas = {div: int(raw_quotas[div]) for div in div_order}
    allocated = sum(quotas.values())

    # Assign the remainder to the division with the largest share
    if allocated < year_qty:
        largest_div = max(div_shares, key=div_shares.get)
        quotas[largest_div] += year_qty - allocated

    return quotas

# === APPLY PF11 OUTSOURCING ===
for year, qty in pf11_quotas.items():
    task_df_pf11 = apply_quotas_for_year(year, qty, task_df_pf11, pf11_thresholds[year], 'Outsource S')

# === APPLY PF12 OUTSOURCING ===
for year, qty in pf12_quotas.items():
    task_df_pf12 = apply_quotas_for_year(year, qty, task_df_pf12, pf12_thresholds[year], 'Outsource E')

# === FINAL OUTPUT ===
union_df = pd.concat([task_df_pf11, task_df_pf12])
union_df.to_excel('OutsourceE&S.xlsx', index=False)


# In[ ]:




