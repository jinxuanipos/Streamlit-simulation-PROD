import streamlit as st
import pandas as pd
from datetime import date, datetime
import os
import openpyxl
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import io

# === STREAMLIT APP ===
st.title("Running FOA Simulations")

# --- Initialize session state ---
if "submitted" not in st.session_state:
    st.session_state.submitted = False

# Disable widgets AFTER submission
disable_inputs = st.session_state.submitted  # **This disables inputs only after button is clicked**

# --- User Input ---
hire = st.selectbox("Select Hiring Plan", 
                    ["Accelerated - Hire additional 20 by Jan 26",
                     "Moderate - Hire additional 10 by Jan 26",
                     "Slow - Hire additional 20 by Jul 26"],
                   disabled=disable_inputs)

hire_mapping = {
    "Accelerated - Hire additional 20 by Jan 26": "a",
    "Moderate - Hire additional 10 by Jan 26": "b",
    "Slow - Hire additional 20 by Jul 26": "c"
}


stretch = st.slider("Enter % take up of incentive scheme", min_value=0, max_value=100, value=50, disabled=disable_inputs  # **Disabled after submission**)
stretch_v = stretch / 100

pphgrowth = st.slider("Enter PPH Growth Y-o-Y", min_value=0, max_value=100, value=10, disabled=disable_inputs  # **Disabled after submission**)
pphgrowth_v = 1 + pphgrowth / 100

eot = st.selectbox("Select EOT Waiver Success Rate", ["26%", "30%", "35%"],disabled=disable_inputs  # **Disabled after submission**)
file_mapping = {
    "26%": "DivisionFiles_All_26.xlsx",
    "30%": "DivisionFiles_All_30.xlsx",
    "35%": "DivisionFiles_All_35.xlsx"
}

secdivert = st.slider("Enter % of secondary job diversion for 2025-26", min_value=0, max_value=100, value=50, disabled=disable_inputs  # **Disabled after submission**)
secdivert_v = secdivert / 100

# === Capacity Parameters Input ===
excel_path = "Capacity-FOA for Python.xlsx"

# --- Update Excel and store in session_state ---
if st.button("Start Simulation") and not st.session_state.submitted:
    wb = load_workbook(excel_path)
    sheet = wb["Calculate Capacity"]

    sheet["I2"] = hire_mapping.get(hire)
    sheet["I6"] = pphgrowth_v
    sheet["I21"] = secdivert_v
    sheet["I27"] = stretch_v

    # Save to in-memory BytesIO
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)  # Reset pointer to start

    # Store in session state
    st.session_state["updated_excel"] = excel_buffer

    st.success("Excel file updated and stored in memory.")

    # Preview values
    st.write("Updated Values:")
    st.write("I2 (Hire):", sheet["I2"].value)
    st.write("I6 (PPH Growth):", sheet["I6"].value)
    st.write("I21 (Diversion):", sheet["I21"].value)
    st.write("I27 (Stretch):", sheet["I27"].value)

filename = file_mapping.get(eot)
task_df = None

if not os.path.exists(filename):
    st.error(f"File '{filename}' not found in the current directory.")
else:
    task_df = pd.read_excel(filename)

# Read the file into memory for download
    with open(filename, "rb") as f:
        file_bytes = f.read()

   
# === Define Constants ===
div_order = ['Div1', 'Div2', 'Div3', 'Div4']

# Outsource quotas
pf11_quotas = {
    2025: 3000,
    2026: 4200,
    2027: 4656,
    2028: 5232,
    2029: 2000,
    2030: 1000
}


pf12_quotas = {
    2026: 1500,
    2027: 2000,
    2028: 3000,
    2029: 3000,
    2030: 3500
}

# Thresholds
pf11_thresholds = {
    2025: date(2024, 1, 1),
    2026: date(2025, 5, 1),
    2027: date(2026, 5, 1),
    2028: date(2027, 5, 1),
    2029: date(2028, 5, 1),
    2030: date(2029, 5, 1),
}

pf12_thresholds = {
    2026: date(2025, 10, 1),
    2027: date(2026, 10, 1),
    2028: date(2027, 10, 1),
    2029: date(2028, 10, 1),
    2030: date(2029, 10, 1),
}

def calculate_division_quotas(div_shares, year_qty, div_order):
    raw_quotas = {div: div_shares.get(div, 0) * year_qty for div in div_order}
    quotas = {div: int(raw_quotas[div]) for div in div_order}
    allocated = sum(quotas.values())
    if allocated < year_qty:
        largest_div = max(div_shares, key=div_shares.get)
        quotas[largest_div] += year_qty - allocated
    return quotas

def apply_quotas_for_year(year, year_qty, task_df, date_threshold, flag_col):
    task_df_remaining = task_df[task_df[flag_col] != 'Y'].copy()
    division_counts = task_df_remaining['Division Transformed'].value_counts().to_dict()
    total_now = len(task_df_remaining)
    div_shares = {div: division_counts.get(div, 0) / total_now for div in div_order}
    quotas = calculate_division_quotas(div_shares, year_qty, div_order)

    div_counts = {div: 0 for div in div_order}

    for index, task in task_df.iterrows():
        if task.get(flag_col) == 'Y' or task['S&E Lodge Date'].date() < date_threshold:
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

# Proceed only if file was loaded
if task_df is not None:
    task_df['S&E Year'] = pd.to_numeric(task_df['S&E Year'], errors='coerce')
    
    task_df_pf11 = task_df[task_df['S&E'] == 'PF11'].copy()
    task_df_pf12 = task_df[task_df['S&E'] == 'PF12'].copy()
    
    task_df_pf11.sort_values(by='S&E Lodge Date', inplace=True)
    task_df_pf12.sort_values(by='S&E Lodge Date', inplace=True)
    
    task_df_pf11['Outsource S'] = task_df_pf11.get('Outsource S', '')
    task_df_pf11['Outsource Year'] = task_df_pf11.get('Outsource Year', pd.NA)
    task_df_pf12['Outsource E'] = task_df_pf12.get('Outsource E', '')
    task_df_pf12['Outsource Year'] = task_df_pf12.get('Outsource Year', pd.NA)

    # --- Apply quotas with visual progress and displayed quota counts ---
st.subheader("Applying Quotas for PF11 and PF12")
progress_text = "Allocating PF11 and PF12 quotas..."
progress_bar = st.progress(0, text=progress_text)
quota_status = st.empty()

total = len(pf11_quotas) + len(pf12_quotas)
current = 0

# PF11 Quotas
for year, qty in pf11_quotas.items():
    task_df_pf11 = apply_quotas_for_year(year, qty, task_df_pf11, pf11_thresholds[year], 'Outsource S')
    current += 1
    percent = current / total
    progress_bar.progress(percent, text=f"{progress_text} (PF11 - {year})")
    quota_status.markdown(f"✅ **PF11 - {year} Quota Applied**: {qty}")

# PF12 Quotas
for year, qty in pf12_quotas.items():
    task_df_pf12 = apply_quotas_for_year(year, qty, task_df_pf12, pf12_thresholds[year], 'Outsource E')
    current += 1
    percent = current / total
    progress_bar.progress(percent, text=f"{progress_text} (PF12 - {year})")
    quota_status.markdown(f"✅ **PF12 - {year} Quota Applied**: {qty}")

progress_bar.empty()
quota_status.markdown("🎉 All quotas have been applied.")


union_df = pd.concat([task_df_pf11, task_df_pf12])
task_df_inhouse = union_df[union_df['Outsource Year'].isnull()]
divisions = ['Div1', 'Div2', 'Div3', 'Div4']

# Use BytesIO instead of writing to a file
division_buffer = io.BytesIO()
with pd.ExcelWriter(division_buffer, engine='xlsxwriter') as writer:
    for div in divisions:
        div_df = task_df_inhouse[task_df_inhouse['Division Transformed'] == div]
        div_df.to_excel(writer, sheet_name=div, index=False)
division_buffer.seek(0)


if "updated_excel" not in st.session_state:
    st.error("Capacity file not found in memory.")
    st.stop()
 
try:
    excel_buffer = st.session_state["updated_excel"]
    excel_buffer.seek(0)  # Reset buffer pointer before reading
    with pd.ExcelFile(excel_buffer) as xls:
        max_tasks_df = pd.read_excel(xls, sheet_name="Python-FOA", index_col=0)
        max_cap_df = pd.read_excel(xls, sheet_name="Python-Cap", index_col=0)
except Exception as e:
    st.error(f"Failed to load capacity data from in-memory: {e}")
    st.stop()
  
# Step 2: Read calendar file bytes into memory buffer once
with open('WorkingDays25-30_withFY.xlsx', 'rb') as f:
    calendar_bytes = f.read()
calendar_buffer = io.BytesIO(calendar_bytes)

# Step 3: Load capacity dataframes from capacity_buffer
with pd.ExcelFile(capacity_buffer) as xls:
    max_tasks_df = pd.read_excel(xls, sheet_name="Python-FOA", index_col=0)
    max_cap_df = pd.read_excel(xls, sheet_name="Python-Cap", index_col=0)

# Step 4: Load calendar dataframe from calendar_buffer
calendar_buffer.seek(0)  # Important: reset pointer before reading
calendar_df = pd.read_excel(calendar_buffer, sheet_name="2025-2030", parse_dates=['Date'])
working_days_df = calendar_df[calendar_df['NWD_Indicator'] == 'No']


# Define weights
SAndE_Points = {'PF11': 0.97, 'PF12': 0.47}

# FOA scheduling by division
# FOA scheduling by division with nested progress bars
st.subheader("Scheduling FOA by Division")

main_progress = st.progress(0, text="Starting FOA scheduling...")
status_text = st.empty()
total_divs = len(divisions)
xls_divisions = pd.ExcelFile(division_buffer)

division_results_buffers = {}
for i, current_div in enumerate(divisions):
    # Read division sheet from in-memory ExcelFile instead of disk
    div_task_df = pd.read_excel(xls_divisions, sheet_name=current_div)
    div_task_df.sort_values(by='S&E Lodge Date', inplace=True)

    working_day_index = 0
    maxwkdays = len(working_days_df)
    task_completed = 0
    capacity_used = 0

    status_text.markdown(f"📄 Processing **{current_div}**...")
    div_progress = st.progress(0, text=f"Scheduling tasks for {current_div}...")
    total_tasks = len(div_task_df)

    # Default in case no condition met
    foa = pd.NaT
    fy = pd.NA
    for j, (index, task) in enumerate(div_task_df.iterrows()):
        if working_day_index >= maxwkdays:
            break
          
        quarter_label = working_days_df['Quarter'].iloc[working_day_index]
        max_capacity = max_cap_df.loc[quarter_label, current_div] if quarter_label in max_cap_df.index else 0
        max_tasks_per_day = max_tasks_df.loc[quarter_label, current_div] if quarter_label in max_tasks_df.index else 0

        SAndEType = task['S&E']
        SAndEPoint = SAndE_Points.get(SAndEType, 0)
      
        if max_capacity > capacity_used and max_tasks_per_day > task_completed:
            foa = working_days_df['Date'].iloc[working_day_index]
            fy = working_days_df['FY'].iloc[working_day_index]
            capacity_used += SAndEPoint
            task_completed += 1

        elif max_capacity > capacity_used and max_tasks_per_day == task_completed:
            currentday_quarter = quarter_label
            working_day_index += 1
            if working_day_index >= maxwkdays:
                break
            quarter_label = working_days_df['Quarter'].iloc[working_day_index]
            foa = working_days_df['Date'].iloc[working_day_index]
            fy = working_days_df['FY'].iloc[working_day_index]
            capacity_used = SAndEPoint if quarter_label != currentday_quarter else capacity_used + SAndEPoint
            task_completed = 1

        elif max_capacity <= capacity_used:
            next_quarter = 'Q' + str(int(quarter_label[1:]) + 1)
            while working_day_index < maxwkdays and working_days_df['Quarter'].iloc[working_day_index] != next_quarter:
                working_day_index += 1
            if working_day_index >= maxwkdays:
                break
            foa = working_days_df['Date'].iloc[working_day_index]
            fy = working_days_df['FY'].iloc[working_day_index]
            capacity_used = SAndEPoint
            task_completed = 1

        div_task_df.at[index, 'FOA'] = foa
        div_task_df.at[index, 'FY'] = fy

        div_progress.progress((j + 1) / total_tasks, text=f"{current_div}: {j + 1}/{total_tasks} tasks scheduled")

   # Save results to an in-memory buffer instead of a file
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        div_task_df.to_excel(writer, index=False, sheet_name=current_div)
    output_buffer.seek(0)  # Reset pointer to start

    # Store buffer in dictionary for later use
    division_results_buffers[current_div] = output_buffer

    # Update main progress bar
    main_progress.progress((i + 1) / total_divs, text=f"Completed {current_div}")

div_progress.empty()
status_text.markdown("✅ All divisions scheduled.")
main_progress.empty()

# Combine all division buffers into a single Excel file with multiple sheets
combined_buffer = io.BytesIO()
with pd.ExcelWriter(combined_buffer, engine='xlsxwriter') as writer:
    for div_name, buffer in division_results_buffers.items():
        buffer.seek(0)
        df = pd.read_excel(buffer)
        df.to_excel(writer, index=False, sheet_name=div_name)
combined_buffer.seek(0)
 
# Download button for the combined Excel file
st.download_button(
    label="Download Complete Schedule (.xlsx)",
    data=combined_buffer,
    file_name="schedule_output.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


#results here
# --- Helper function to compute average age ---
def compute_avg_age(grouped_df, start_dates, end_dates, quantities, add_months, additional_time):
    avg_list = []
    for i in range(len(quantities)):
        min_date = grouped_df.iloc[i]
        max_date = grouped_df.iloc[i]
        avg_days = ((start_dates[i] - min_date).days + (end_dates[i] - max_date).days) / 2
        avg_list.append(quantities[i] * avg_days)
    avg_age = sum(avg_list) / sum(quantities) / 30.5 + additional_time
    return avg_age

# --- Load inhouse results and combine ---
div_files = []
for div in divisions:  
    buffer = division_results_buffers[div]  # get the BytesIO for this division
    # Read the Excel content from the buffer
    df = pd.read_excel(buffer, sheet_name=div)
    div_files.append(df)
excel_merged = pd.concat(div_files, ignore_index=True)

# --- Add time calculations ---
excel_merged['FOA'] = pd.to_datetime(excel_merged['FOA'], errors='coerce')
excel_merged['S&E Lodge Date'] = pd.to_datetime(excel_merged['S&E Lodge Date'], errors='coerce')
excel_merged['time_c'] = ((excel_merged['FOA'] - excel_merged['S&E Lodge Date']).dt.days / 30.5).clip(lower=1)
excel_merged['original time'] = (excel_merged['FOA'] - excel_merged['S&E Lodge Date']).dt.days / 30.5
#excel_merged.to_excel('Div1-4Combined.xlsx', index=False)

# --- Load outsource data ---
task_df = union_df[union_df['Outsource Year'].notnull()]
task_df_pf12 = task_df[task_df['Outsource E'] == 'Y']
task_df_pf11 = task_df[task_df['Outsource S'] == 'Y']

# --- Define quantities ---
e_qty = [1500, 2000, 3000, 3000, 3500]
s_qty = [3000, 4200, 4656, 5232, 2000, 1000]
outsource_e_time, outsource_s_time = 9, 5


# --- Compute E files returned each year ---
no_of_mths = 15 - outsource_e_time
fy_e = [
    round(e_qty[0]/12 * no_of_mths),
    e_qty[0] - round(e_qty[0]/12 * no_of_mths) + round(e_qty[1]/12 * no_of_mths),
    round(e_qty[1] - e_qty[1]/12 * no_of_mths) + round(e_qty[2]/12 * no_of_mths),
    round(e_qty[2] - e_qty[2]/12 * no_of_mths) + round(e_qty[3]/12 * no_of_mths),
    round(e_qty[3] - e_qty[3]/12 * no_of_mths) + round(e_qty[4]/12 * no_of_mths)
]

# --- Fiscal year start/end dates ---
sdates = [datetime(y, 1, 1) for y in range(2025, 2031)]
edates = [datetime(y, 12, 31) for y in range(2025, 2031)]


# Compute avg_E_age properly using min and max dates per year
min_dates = task_df_pf12.groupby('Outsource Year')['S&E Lodge Date'].min().tolist()
max_dates = task_df_pf12.groupby('Outsource Year')['S&E Lodge Date'].max().tolist()

avg_list = []
for i in range(len(e_qty)):
    avg_days = ((sdates[i+1] - min_dates[i]).days + (edates[i+1] - max_dates[i]).days) / 2
    avg_list.append(e_qty[i] * avg_days)

avg_E_age = sum(avg_list) / sum(e_qty) / 30.5 + outsource_e_time


# --- Calculate average age of outsourced S ---
avg_S_list = []
grp_pf11 = task_df_pf11.groupby('Outsource Year')['S&E Lodge Date']
min_pf11 = grp_pf11.min().tolist()
max_pf11 = grp_pf11.max().tolist()

for i in range(6):
    avg_days = ((sdates[i] - min_pf11[i]).days + (edates[i] - max_pf11[i]).days) / 2
    avg_S_list.append((avg_days / 30.5) + outsource_s_time)
print(avg_S_list)

# --- Merge all FOA counts and averages ---
fy_sums = excel_merged.groupby('FY')['time_c'].sum().to_dict()
fy_counts = excel_merged.groupby('FY')['time_c'].count().to_dict()


# Get value from a specific cell (e.g., B2)
# --- Define PPH projections ---
def projected_pph(year_multiplier):
    return round(pph_base * (pphgrowth_v ** year_multiplier))
pph_base = 714


# --- Updated total_sum_count using both age and time ---
def total_sum_count(fy, s_qty, e_qty=0, s_age=0, e_age=0, s_time=0, e_time=0, year_mult=0):
    projected_sum = projected_pph(year_mult) * 10
    projected_count = projected_pph(year_mult)

    total_sum = (
        fy_sums.get(fy, 0)
        + (s_qty * (s_age + s_time))
        + (e_qty * (e_age + e_time))
        + projected_sum
    )
    total_count = (
        fy_counts.get(fy, 0)
        + s_qty
        + e_qty
        + projected_count
    )
    return total_sum, total_count

# --- Final FOA values per fiscal year ---
foa_values = []
for i, fy in enumerate(['FY25', 'FY26', 'FY27', 'FY28', 'FY29', 'FY30']):
    e = fy_e[i - 1] if i > 0 else 0
    foa_sum, foa_count = total_sum_count(
        fy,
        s_qty[i],
        e_qty=e,
        s_age=avg_S_list[i] - outsource_s_time,
        e_age=avg_E_age - outsource_e_time,
        s_time=outsource_s_time,
        e_time=outsource_e_time,
        year_mult=i + 1
    )
    foa_values.append(round(foa_sum / foa_count, 1))

# --- Append FY23 and FY24 historical data ---
years = ['FY23', 'FY24', 'FY25', 'FY26', 'FY27', 'FY28', 'FY29', 'FY30']
values = [15.4, 19.8] + foa_values

# --- Plot ---
fig, ax = plt.subplots(figsize=(10, 6))
ax.plot(years, values, marker='o')
ax.set_title('FOA')
ax.set_xlabel('Fiscal Year')
ax.set_ylabel('FOA (months)')
ax.grid(True)

for i, v in enumerate(values):
    ax.text(i, v + 0.2, f'{v:.1f}', ha='center')

ax.set_ylim(min(values) - 1, max(values) + 2)
st.pyplot(fig)


