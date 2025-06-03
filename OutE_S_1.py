import streamlit as st
import pandas as pd
from datetime import date, datetime
import os
import openpyxl
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import io

# === STREAMLIT APP ===
st.set_page_config(layout="wide")
st.title("Running FOA Simulations")

# --- First Row: Quadrants 1 and 2 ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("Demand")
    pphgrowth = st.slider("Enter Growth Rate of PPH Usage Rate Y-o-Y", min_value=0, max_value=20, value=10)
    eot = st.selectbox("Select EOT Waiver Success Rate", ["26%", "30%", "35%"])
    Filingsgrowth = st.selectbox("Select projected S&E Growth Y-o-Y 2027-2030 based on projections of patent filings", [
        "High - Upper bound of patent filing forecast",
        "Moderate - Average",
        "Slow - Lower bound of patent filing forecast"
    ])

with col2:
    st.subheader("Capacity")
    secdivert = st.slider("Yearly secondary job diversion for 2025-2026 (%); where 0 = status quo and 100 = divert all secondary jobs", 0, 100, 50)	
    hire = st.selectbox("Select Hiring Plan", [
        "Accelerated - Hire additional 20 by Jan 26",
        "Moderate - Hire additional 10 by Jan 26",
        "Paced - Hire additional 20 by Jul 26"
    ])
    AIgainschoice = st.selectbox("Select projected S&E Productivity gains from PAS and Report Drafter", [
        "Best - 55% by Jan30",
        "Better - 45% by Jan30",
        "Good - 35% by Jan30"
    ])
    st.markdown("**Incentive scheme options explanation**")
    st.caption("Incentive scheme is only valid for 2025-2027. \n_Have grouped 2026 and 2027 together. \n_As the worst case scenario, assume that not meeting baseline target happens for all years.")	
    incentivescheme = st.selectbox("Select success of incentive scheme", [
        "Do not meet baseline target across all years for 2025-2030",
        "Meet baseline target + incentive scheme for 2025, meet baseline target only for 2026 and 2027",
	"Meet baseline target + incentive scheme for 2026 and 2027, meet baseline target only for 2025",
	"Meet baseline target + incentive scheme for 2025, 2026 and 2027",
	"Meet baseline target only for 2025, 2026 and 2027"
    ])
    stretch_2025 = st.slider("Select capacity boost from incentive scheme 2025 (%)", 0, 20, 10)
    stretch_2026onwards = st.slider("Select yearly capacity boost from incentive scheme 2026-2030 (%)", 0, 10, 5)
   	
   

# --- Second Row: Quadrants 3 and 4 ---
col3, col4 = st.columns(2)

with col3:
    st.subheader("Outsource Volumes")
    st.markdown("*Age of files = ... *")	
    # Outsource Search - vary for 2025-2027. 2028-2030 keep constant. 
    Outsource_S_2025 = st.slider("Outsource Search Volume 2025", 0, 4000, 3000, step =100)
    Outsource_S_2026 = st.slider("Outsource Search Volume 2026", 0, 4000, 3000, step =100)
    Outsource_S_2027 = st.slider("Outsource Search Volume 2027", 0, 4000, 3000, step =100)
    Outsource_S_282930 = st.slider("Yearly Outsource Search Volume 2028-2030; equal volumes across all 3 years", 0, 4000, 3000, step =100)
  
with col4:
    st.subheader("Collaboration")
    st.markdown("*Volume of files kept constant at ...., age of files = .... *")	
    # Collaboration Exams - vary turnaround time instead
    Outsource_e_select = st.selectbox("Select partner's turnaround time for Collarboration files", [
        "Fast - 6 months",
        "Moderate - 9 months",
        "Good - 12 months"
    ])



#mapping and calculation for user selected values

# --- MAP the hire selection to respective file names ---
hire_file_mapping = {
    "Accelerated - Hire additional 20 by Jan 26": "DivisionFiles_HighGrowth.xlsx",
    "Moderate - Hire additional 10 by Jan 26": "DivisionFiles_MidGrowth.xlsx",
    "Paced - Hire additional 20 by Jul 26": "DivisionFiles_LowGrowth.xlsx"
}

# --- MAP the EOT selection to sheet names ---
eot_sheet_mapping = {
    "26%": "26",
    "30%": "30",
    "35%": "35"
}

# --- Resolve the Excel file and sheet name ---
excel_file = hire_file_mapping[hire]
sheet_name = eot_sheet_mapping[eot]


# input capacity dictionary
accelerated_worst = {
    "Div1": {2025: 1942, 2026: 2431, 2027: 2784, 2028: 3112, 2029: 3447, 2030: 3490},
    "Div2": {2025: 1844, 2026: 2216, 2027: 2386, 2028: 2719, 2029: 3039, 2030: 3233},
    "Div3": {2025: 1469, 2026: 1728, 2027: 1964, 2028: 2380, 2029: 2776, 2030: 2848},
    "Div4": {2025: 1903, 2026: 2196, 2027: 2565, 2028: 2650, 2029: 2955, 2030: 3027}
}

accelerated_baseline_all = {
    "Div1": {2025: 2035, 2026: 2547, 2027: 2917, 2028: 3260, 2029: 3611, 2030: 3656},
    "Div2": {2025: 1932, 2026: 2322, 2027: 2500, 2028: 2848, 2029: 3184, 2030: 3387},
    "Div3": {2025: 1539, 2026: 1810, 2027: 2058, 2028: 2493, 2029: 2908, 2030: 2984},
    "Div4": {2025: 1994, 2026: 2301, 2027: 2687, 2028: 2776, 2029: 3096, 2030: 3171}
}

accelerated_baseline_incentive_all = {
    "Div1": {2025: 2137, 2026: 2802, 2027: 3209, 2028: 3260, 2029: 3611, 2030: 3656},
    "Div2": {2025: 2028, 2026: 2554, 2027: 2750, 2028: 2848, 2029: 3184, 2030: 3387},
    "Div3": {2025: 1616, 2026: 1991, 2027: 2264, 2028: 2493, 2029: 2908, 2030: 2984},
    "Div4": {2025: 2094, 2026: 2531, 2027: 2955, 2028: 2776, 2029: 3096, 2030: 3171}
}

accelerated_incentive_25 = {
    "Div1": {2025: 2137, 2026: 2547, 2027: 2917, 2028: 3260, 2029: 3611, 2030: 3656},
    "Div2": {2025: 2028, 2026: 2322, 2027: 2500, 2028: 2848, 2029: 3184, 2030: 3387},
    "Div3": {2025: 1616, 2026: 1810, 2027: 2058, 2028: 2493, 2029: 2908, 2030: 2984},
    "Div4": {2025: 2094, 2026: 2301, 2027: 2687, 2028: 2776, 2029: 3096, 2030: 3171}
}

accelerated_incentive_2627 = {
    "Div1": {2025: 2035, 2026: 2802, 2027: 3209, 2028: 3260, 2029: 3611, 2030: 3656},
    "Div2": {2025: 1932, 2026: 2554, 2027: 2750, 2028: 2848, 2029: 3184, 2030: 3387},
    "Div3": {2025: 1539, 2026: 1991, 2027: 2264, 2028: 2493, 2029: 2908, 2030: 2984},
    "Div4": {2025: 1994, 2026: 2531, 2027: 2955, 2028: 2776, 2029: 3096, 2030: 3171}
}

moderate_worst = {
    "Div1": {2025: 1943, 2026: 2428, 2027: 2703, 2028: 2987, 2029: 3213, 2030: 3245},
    "Div2": {2025: 1844, 2026: 2214, 2027: 2398, 2028: 2593, 2029: 2896, 2030: 2989},
    "Div3": {2025: 1469, 2026: 1815, 2027: 1975, 2028: 2346, 2029: 2542, 2030: 2604},
    "Div4": {2025: 1903, 2026: 2283, 2027: 2484, 2028: 2616, 2029: 2720, 2030: 2783}
}

moderate_baseline_all = {
    "Div1": {2025: 2035, 2026: 2544, 2027: 2832, 2028: 3129, 2029: 3366, 2030: 3400},
    "Div2": {2025: 1932, 2026: 2319, 2027: 2512, 2028: 2716, 2029: 3034, 2030: 3131},
    "Div3": {2025: 1539, 2026: 1901, 2027: 2069, 2028: 2458, 2029: 2663, 2030: 2728},
    "Div4": {2025: 1994, 2026: 2392, 2027: 2602, 2028: 2741, 2029: 2850, 2030: 2915}
}

moderate_incentive_all = {
    "Div1": {2025: 2137, 2026: 2799, 2027: 3116, 2028: 3129, 2029: 3366, 2030: 3400},
    "Div2": {2025: 2028, 2026: 2551, 2027: 2763, 2028: 2716, 2029: 3034, 2030: 3131},
    "Div3": {2025: 1616, 2026: 2091, 2027: 2276, 2028: 2458, 2029: 2663, 2030: 2728},
    "Div4": {2025: 2094, 2026: 2631, 2027: 2862, 2028: 2741, 2029: 2850, 2030: 2915}
}

moderate_incentive_25 = {
    "Div1": {2025: 2137, 2026: 2544, 2027: 2832, 2028: 3129, 2029: 3366, 2030: 3400},
    "Div2": {2025: 2028, 2026: 2319, 2027: 2512, 2028: 2716, 2029: 3034, 2030: 3131},
    "Div3": {2025: 1616, 2026: 1901, 2027: 2069, 2028: 2458, 2029: 2663, 2030: 2728},
    "Div4": {2025: 2094, 2026: 2392, 2027: 2602, 2028: 2741, 2029: 2850, 2030: 2915}
}

moderate_incentive_2627 = {
    "Div1": {2025: 2035, 2026: 2799, 2027: 3116, 2028: 3129, 2029: 3366, 2030: 3400},
    "Div2": {2025: 1932, 2026: 2551, 2027: 2763, 2028: 2716, 2029: 3034, 2030: 3131},
    "Div3": {2025: 1539, 2026: 2091, 2027: 2276, 2028: 2458, 2029: 2663, 2030: 2728},
    "Div4": {2025: 1994, 2026: 2631, 2027: 2862, 2028: 2741, 2029: 2850, 2030: 2915}
}

slow_worst = {
    "Div1": {2025: 1943, 2026: 2431, 2027: 2784, 2028: 3112, 2029: 3447, 2030: 3487},
    "Div2": {2025: 1844, 2026: 2216, 2027: 2386, 2028: 2719, 2029: 3039, 2030: 3230},
    "Div3": {2025: 1469, 2026: 1728, 2027: 1964, 2028: 2380, 2029: 2685, 2030: 2846},
    "Div4": {2025: 1903, 2026: 2196, 2027: 2565, 2028: 2650, 2029: 2864, 2030: 3024}
}

slow_baseline_all = {
    "Div1": {2025: 2035, 2026: 2547, 2027: 2917, 2028: 3260, 2029: 3611, 2030: 3653},
    "Div2": {2025: 1932, 2026: 2322, 2027: 2500, 2028: 2848, 2029: 3184, 2030: 3384},
    "Div3": {2025: 1539, 2026: 1810, 2027: 2058, 2028: 2493, 2029: 2813, 2030: 2981},
    "Div4": {2025: 1994, 2026: 2301, 2027: 2687, 2028: 2776, 2029: 3000, 2030: 3168}
}

slow_incentive_all = {
    "Div1": {2025: 2137, 2026: 2802, 2027: 3209, 2028: 3260, 2029: 3611, 2030: 3653},
    "Div2": {2025: 2028, 2026: 2554, 2027: 2750, 2028: 2848, 2029: 3184, 2030: 3384},
    "Div3": {2025: 1616, 2026: 1991, 2027: 2264, 2028: 2493, 2029: 2813, 2030: 2981},
    "Div4": {2025: 2094, 2026: 2531, 2027: 2955, 2028: 2776, 2029: 3000, 2030: 3168}
}

slow_incentive_25 = {
    "Div1": {2025: 2137, 2026: 2547, 2027: 2917, 2028: 3260, 2029: 3611, 2030: 3653},
    "Div2": {2025: 2028, 2026: 2322, 2027: 2500, 2028: 2848, 2029: 3184, 2030: 3384},
    "Div3": {2025: 1616, 2026: 1810, 2027: 2058, 2028: 2493, 2029: 2813, 2030: 2981},
    "Div4": {2025: 2094, 2026: 2301, 2027: 2687, 2028: 2776, 2029: 3000, 2030: 3168}
}

slow_incentive_2627 = {
    "Div1": {2025: 2035, 2026: 2802, 2027: 3209, 2028: 3260, 2029: 3611, 2030: 3653},
    "Div2": {2025: 1932, 2026: 2554, 2027: 2750, 2028: 2848, 2029: 3184, 2030: 3384},
    "Div3": {2025: 1539, 2026: 1991, 2027: 2264, 2028: 2493, 2029: 2813, 2030: 2981},
    "Div4": {2025: 1994, 2026: 2531, 2027: 2955, 2028: 2776, 2029: 3000, 2030: 3168}
}

# Mapping for capacity dictionaries by hire plan and incentive scheme choice
capacity_map = {
    "Accelerated - Hire additional 20 by Jan 26": {
        "Do not meet baseline target across all years for 2025-2030": accelerated_worst,
        "Meet baseline target + incentive scheme for 2025, meet baseline target only for 2026 and 2027": accelerated_incentive_25,
        "Meet baseline target + incentive scheme for 2026 and 2027, meet baseline target only for 2025": accelerated_incentive_2627,
        "Meet baseline target + incentive scheme for 2025, 2026 and 2027": accelerated_baseline_incentive_all,
        "Meet baseline target only for 2025, 2026 and 2027": accelerated_baseline_all,
    },
    "Moderate - Hire additional 10 by Jan 26": {
        "Do not meet baseline target across all years for 2025-2030": moderate_worst,
        "Meet baseline target + incentive scheme for 2025, meet baseline target only for 2026 and 2027": moderate_incentive_25,
        "Meet baseline target + incentive scheme for 2026 and 2027, meet baseline target only for 2025": moderate_incentive_2627,
        "Meet baseline target + incentive scheme for 2025, 2026 and 2027": moderate_incentive_all,
        "Meet baseline target only for 2025, 2026 and 2027": moderate_baseline_all,
    },
    "Paced - Hire additional 20 by Jul 26": {
        "Do not meet baseline target across all years for 2025-2030": slow_worst,
        "Meet baseline target + incentive scheme for 2025, meet baseline target only for 2026 and 2027": slow_incentive_25,
        "Meet baseline target + incentive scheme for 2026 and 2027, meet baseline target only for 2025": slow_incentive_2627,
        "Meet baseline target + incentive scheme for 2025, 2026 and 2027": slow_incentive_all,
        "Meet baseline target only for 2025, 2026 and 2027": slow_baseline_all,
    }
}


# Mapping logic to extract numeric months
turnaround_mapping = {
    "Fast - 6 months": 6,
    "Moderate - 9 months": 9,
    "Good - 12 months": 12
}

# Get numeric value
outsource_e_time = turnaround_mapping[Outsource_e_select]


years = list(range(2025, 2031))  # 2025 to 2030


# --- start calculations
if st.button("Start Simulation"):
    # Resolve file and sheet name based on user input
    excel_file = hire_file_mapping[hire]
    sheet_name = eot_sheet_mapping[eot]
    
    task_df = None
    if not os.path.exists(excel_file):
        st.error(f"File '{excel_file}' not found in the current directory.")
    else:
        try:
            task_df = pd.read_excel(excel_file, sheet_name=sheet_name)
            st.success(f"Successfully loaded '{excel_file}' | Sheet: '{sheet_name}'")
        except Exception as e:
            st.error(f"Failed to read '{sheet_name}' from '{excel_file}': {e}")

    # total capacity
    capacitybydiv = capacity_map[hire][incentivescheme]
    years = list(range(2025, 2031))

    # Initialize a dict to hold combined capacity per year
    totalcapacity = {}
    for year in years:
        total = 0
        for division in capacitybydiv:
            total += capacitybydiv[division].get(year, 0)  # Safe access
        totalcapacity[year] = total

    st.write("Total capacity based on hiring plan and incentive scheme")
    st.write(totalcapacity)

    # --- Define PPH projections ---
    pph_base_rate = 0.063  # 6.3%

    # Convert 'S&E Year' to numeric (safely)
    task_df['S&E Year'] = pd.to_numeric(task_df['S&E Year'], errors='coerce')
    # Filter only years in 2025–2030
    filtered_df = task_df[task_df['S&E Year'].between(2025, 2030)]
    # Group by year and count rows
    year_counts = filtered_df['S&E Year'].value_counts().sort_index()
    # Convert to dictionary
    searchexam_base = year_counts.to_dict()
    
    st.write(searchexam_base)

    projected_pph = {}
    projected_pph_list = []

    for i, year in enumerate(range(2025, 2031)):
    	growth_factor = (1 + pphgrowth / 100) ** i
    	base = searchexam_base.get(year, 0)  # use 0 if year is missing
    	projected_value = base * pph_base_rate * growth_factor
    	projected_pph[year] = projected_value
    	projected_pph_list.append(projected_value)
	st.write("projected_pph_list")
	st.write(projected_pph_list)

    # Adjusting projections (deductions = adjusted values)
    deductions = {}
    for i, year in enumerate(range(2025, 2031)):
        proj_pph = projected_pph[year]
        adjusted_pph = proj_pph * 0.97  # 3% deduction
        deductions[year] = proj_pph + adjusted_pph 
    st.write("pph deductions")
    st.write(deductions)

    # Subtract deductions from capacity year by year 
    adjusted_capacity = [totalcapacity[year] - deductions[year] for year in range(2025, 2031)]
    st.write("adjusted capacity")
    st.write(adjusted_capacity)
	
 

    #calculate AI gains
    est_AI_dict = {
        "pf11": [],
        "pf12": []
    }

    for i in range(len(adjusted_capacity)):
        pf11_val = int(adjusted_capacity[i] * 0.7 / 0.97)
        pf12_val = int(adjusted_capacity[i] * 0.3 / 0.47)
        est_AI_dict["pf11"].append(pf11_val)
        est_AI_dict["pf12"].append(pf12_val)

    # st.write("PF11 and PF12 volumes over years 2025-2030:")
    # st.write(est_AI_dict)

    ai_scenarios = {
    "Best": {
        "PAS - PF11": {2025: 0.0, 2026: 15.0, 2027: 22.5, 2028: 27.5, 2029: 35.0, 2030: 35.0},
        "Report Drafter - PF11": {2025: 0.0, 2026: 10.0, 2027: 12.5, 2028: 15.0, 2029: 17.5, 2030: 20.0},
        "Report Drafter - PF12": {2025: 0.0, 2026: 10.0, 2027: 12.5, 2028: 15.0, 2029: 17.5, 2030: 20.0}},
    "Better": {
        "PAS - PF11": {2025: 0.0, 2026: 10.0, 2027: 15.0, 2028: 20.0, 2029: 25.0, 2030: 25.0},
        "Report Drafter - PF11": {2025: 0.0, 2026: 0.0, 2027: 5.0, 2028: 10.0, 2029: 15.0, 2030: 20.0},
        "Report Drafter - PF12": {2025: 0.0, 2026: 0.0, 2027: 5.0, 2028: 10.0, 2029: 15.0, 2030: 20.0}},
    "Good": {
        "PAS - PF11": {2025: 0.0, 2026: 10.0, 2027: 15.0, 2028: 20.0, 2029: 25.0, 2030: 25.0},
        "Report Drafter - PF11": {2025: 0.0, 2026: 0.0, 2027: 3.0, 2028: 6.0, 2029: 10.0, 2030: 10.0},
        "Report Drafter - PF12": {2025: 0.0, 2026: 0.0, 2027: 3.0, 2028: 6.0, 2029: 10.0, 2030: 10.0}}}
   
    ai_scenario_mapping = {
    "Best - 55% by Jan30": "Best",
    "Better - 45% by Jan30": "Better",
    "Good - 35% by Jan30": "Good"}

    selected_ai_key = ai_scenario_mapping[AIgainschoice]
    ai_dict = ai_scenarios[selected_ai_key]


    ai_gains = {}
    for i, year in enumerate(years):
        pf11 = est_AI_dict["pf11"][i]
        pf12 = est_AI_dict["pf12"][i]

        # Get AI impact percentages for the year
        pas = ai_dict["PAS - PF11"][year] / 100
        rd_pf11 = ai_dict["Report Drafter - PF11"][year] / 100
        rd_pf12 = ai_dict["Report Drafter - PF12"][year] / 100

        # Calculate AI gain
        gain = (pf11 * pas * 0.5) + (pf11 * rd_pf11 * 0.47 * 0.1) + (pf12 * rd_pf12 * 0.47 * 0.1)
        ai_gains[year] = int(gain)  # Store as integer

        # Display result in Streamlit
        st.write("AI Gains from 2025 to 2030:")
        st.write(ai_gains)


    # Initialize dictionary
    secdivert_deductions = {}

    # Apply diversion logic for 2025 to 2030
    for i, year in enumerate(range(2025, 2031)):
        if year in [2025, 2026]:
            diverted_val = adjusted_capacity[i] * 0.25 * secdivert_v
        else:  # For 2027 to 2030
            diverted_val = adjusted_capacity[i] * 0.25
        secdivert_deductions[year] = int(diverted_val)
      
    # Display in Streamlit
    #st.write("Deductions after applying secondary diversion (2025–2026 only):")
    #st.write(secdivert_deductions)


    #qc for out e
    qc_effort = {
        2025: 0,
        2026: 0,
        2027: 262,
        2028: 290,
        2029: 399,
        2030: 465
    }


    for i, year in enumerate(years):
        # Start with current deduction value
        value = adjusted_capacity[i]
        
        # Subtract secdivert only for 2025 and 2026
        if year in secdivert_deductions:
            value -= secdivert_deductions[year]
    
        # Subtract qc_effort
        value -= qc_effort[year]

        # Add AI gains
        value += ai_gains[year]
    
        # Update the cap list
        adjusted_capacity[i] = int(value)  # Ensure it's stored as integer

    # Display updated deductions
    # st.write("Final updated cap after applying secdivert, QC Effort, and AI Gains:")
    # st.write(adjusted_capacity)


    capacity_split = {
        "Div1": {2025: 27.1, 2026: 28.4, 2027: 28.7, 2028: 28.7, 2029: 28.2, 2030: 27.7},
        "Div2": {2025: 25.8, 2026: 25.9, 2027: 24.6, 2028: 25.0, 2029: 24.9, 2030: 25.7},
        "Div3": {2025: 20.5, 2026: 20.2, 2027: 20.3, 2028: 21.9, 2029: 22.7, 2030: 22.6},
        "Div4": {2025: 26.6, 2026: 25.6, 2027: 26.4, 2028: 24.4, 2029: 24.2, 2030: 24.0}
    }


    # Initialize the result dictionary
    capacity_with_incentives = {div: {} for div in capacity_split.keys()}

    for div in capacity_split:
        for year in years:
            # capacity split % for division and year (as decimal)
            split = capacity_split[div][year] / 100
        
            # base capacity (e.g., cap) for the year
            base_capacity = adjusted_capacity[year - 2025]
        
            # incentive for division and year
            incentive_val = selected_incentives[div][year]
        
            # calculate final value
            final_val = int(round(base_capacity * split + incentive_val))
        
            # store in dictionary
            capacity_with_incentives[div][year] = final_val

    # Display result in Streamlit
    # st.write("Capacity split by division with incentives added (2025-2030):")
    # st.write(capacity_with_incentives)


    quarterly_split = {
        "Div1": {"Q1": 30.3, "Q2": 22.1, "Q3": 25.2, "Q4": 22.5},
        "Div2": {"Q1": 29.4, "Q2": 24.8, "Q3": 25.2, "Q4": 20.6},
        "Div3": {"Q1": 24.1, "Q2": 23.2, "Q3": 23.4, "Q4": 29.3},
        "Div4": {"Q1": 25.6, "Q2": 23.2, "Q3": 24.4, "Q4": 26.8}
    }

    # Initialize dict to hold quarterly capacities
    quarterly_capacity = {div: {year: {} for year in range(2025, 2031)} for div in quarterly_split.keys()}

    for div, quarters in quarterly_split.items():
        for year in range(2025, 2031):
            yearly_capacity = capacity_with_incentives[div][year]
            for qtr, pct in quarters.items():
                quarterly_capacity[div][year][qtr] = int(round(yearly_capacity * (pct / 100)))

    # Example output in Streamlit
    # st.write("Quarterly disbursed capacity by division and year:")
    # st.write(quarterly_capacity)


    foa_base = {
        "Div1": {"Q1": 7, "Q2": 6, "Q3": 5, "Q4": 5},
        "Div2": {"Q1": 6, "Q2": 6, "Q3": 5, "Q4": 4},
        "Div3": {"Q1": 6, "Q2": 7, "Q3": 5, "Q4": 7},
        "Div4": {"Q1": 7, "Q2": 7, "Q3": 6, "Q4": 7},
    }


    # Original foa_per_quarter dictionary (before update)
    foa_per_quarter = {
        "Div1": {"Q1": 7, "Q2": 6, "Q3": 5, "Q4": 5},
        "Div2": {"Q1": 6, "Q2": 6, "Q3": 5, "Q4": 4},
        "Div3": {"Q1": 6, "Q2": 7, "Q3": 5, "Q4": 7},
        "Div4": {"Q1": 7, "Q2": 7, "Q3": 6, "Q4": 7},
    }

    # Calculate % gains from 2025 for each year 2026-2030
    base_value = adjusted_capacity[0]  # 2025 value
    percentage_gains = {}
    # st.write("base value to adjust foa is:")
    # st.write(base_value)


    for year_idx in range(1, 6):  # indices 1 to 5 correspond to 2026-2030 
        current_value = adjusted_capacity[year_idx]
        gain = (current_value - base_value) / base_value
        gain = round(gain, 2)  # Round gain to 2 decimal places
        percentage_gains[2025 + year_idx] = max(gain, 0)  # Convert negative gains to 0
 
    # Now update foa_per_quarter for each year 2026-2030
    # We create a nested dict: {year: {division: {quarter: updated_foa}}}
    updated_foa = {}

    for year in range(2025, 2031):
        updated_foa[year] = {}
        if year == 2025:
            # For base year, no change
            updated_foa[year] = foa_per_quarter
        else:
            gain = percentage_gains[year]
            updated_foa[year] = {}
            for div, quarters in foa_per_quarter.items():
                updated_foa[year][div] = {}
                for qtr, count in quarters.items():
                    new_count = int(round(count * (1 + gain)))
                    updated_foa[year][div][qtr] = new_count

    # Optional: Display result for a sample year (e.g., 2026)
    # import pprint
    # pprint.pprint(updated_foa)
    # Example output in Streamlit
    # st.write("Display FOA result:")
    # st.write(updated_foa)
    # st.write("percentage gains:")
    # st.write(percentage_gains)


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
        2026: date(2026, 1, 1),
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
    st.subheader("Starting simulation")

    # PF11 Quotas
    for year, qty in pf11_quotas.items():
        task_df_pf11 = apply_quotas_for_year(year, qty, task_df_pf11, pf11_thresholds[year], 'Outsource S')

    # PF12 Quotas
    for year, qty in pf12_quotas.items():
        task_df_pf12 = apply_quotas_for_year(year, qty, task_df_pf12, pf12_thresholds[year], 'Outsource E')

    st.markdown("🎉 Preparing to start scheduling of FOAs by division...")

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

    
    # Step 2: Read calendar file bytes into memory buffer once
    with open('WorkingDays25-30_withFY.xlsx', 'rb') as f:
        calendar_bytes = f.read()
    calendar_buffer = io.BytesIO(calendar_bytes)

    # Step 4: Load calendar dataframe from calendar_buffer
    calendar_buffer.seek(0)  # Important: reset pointer before reading
    calendar_df = pd.read_excel(calendar_buffer, sheet_name="2025-2030", parse_dates=['Date'])
    working_days_df = calendar_df[calendar_df['NWD_Indicator'] == 'No']


    # Define weights
    SAndE_Points = {'PF11': 0.97, 'PF12': 0.47}

    # FOA scheduling by division
    # FOA scheduling by division with nested progress bars
    st.subheader("Scheduling FOAs by division")
 
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
        # foa = pd.NaT
        # fy = pd.NA
        for j, (index, task) in enumerate(div_task_df.iterrows()):
            if working_day_index >= maxwkdays:
                break

            quarter_label = working_days_df['Quarter'].iloc[working_day_index]
            

            # Get current date to derive year and quarter
            current_date = working_days_df['Date'].iloc[working_day_index]
            current_year = current_date.year
            current_quarter = working_days_df['Quarter'].iloc[working_day_index]

            # Defensive default
            max_capacity = 0
            max_tasks_per_day = 0

            # Use quarterly_capacity and updated_foa
            if (current_div in quarterly_capacity and current_year in quarterly_capacity[current_div] and current_quarter in foa_per_quarter[current_div]):
                max_capacity = quarterly_capacity[current_div][current_year][current_quarter]
                max_tasks_per_day = updated_foa[current_year][current_div][current_quarter]

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

            #elif max_capacity <= capacity_used:
                #next_quarter = 'Q' + str(int(quarter_label[1:]) + 1)
                #while working_day_index < maxwkdays and working_days_df['Quarter'].iloc[working_day_index] != next_quarter:
                    #working_day_index += 1
                #if working_day_index >= maxwkdays:
                    #break
                #foa = working_days_df['Date'].iloc[working_day_index]
                #fy = working_days_df['FY'].iloc[working_day_index]
                #capacity_used = SAndEPoint
                #task_completed = 1

            elif max_capacity <= capacity_used:
                current_quarter = working_days_df['Quarter'].iloc[working_day_index]
                current_year = working_days_df['Date'].iloc[working_day_index].year

                # Look ahead for next available date with a *different quarter or year*
                while working_day_index < maxwkdays:
                  next_quarter = working_days_df['Quarter'].iloc[working_day_index]
                  next_year = working_days_df['Date'].iloc[working_day_index].year
                  if next_quarter != current_quarter or next_year != current_year:
                      break
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
    status_text.markdown("⏳ Calculating results and plotting final FOA graph...")
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
    # st.download_button(
        #label="Download Complete Schedule (.xlsx)",
        #data=combined_buffer,
        #file_name="schedule_output.xlsx",
        #mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    #)


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
    
    # Convert DataFrame to in-memory Excel file (intermediate step for the following download)
    # excel_buffer = io.BytesIO()
    # with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
      # excel_merged.to_excel(writer, index=False)
    # excel_buffer.seek(0)  # Important: move to the beginning of the buffer

    # Download button for the combined Excel file
    #st.download_button(
      #label="Download Excel Merged File (.xlsx)",
      #data=excel_buffer,
      #file_name="excel_merged.xlsx",
      #mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    #)


    # --- Load outsource data ---
    task_df = union_df[union_df['Outsource Year'].notnull()]
    task_df_pf12 = task_df[task_df['Outsource E'] == 'Y']
    task_df_pf11 = task_df[task_df['Outsource S'] == 'Y']

    # --- Define quantities ---
    e_qty = [0, 0, 1583, 2167, 3000, 3083]
    s_qty = [3000, 4200, 4656, 5232, 2000, 1000]
    outsource_e_time, outsource_s_time = 9, 5


    # --- Fiscal year start/end dates ---
    sdates_e = [datetime(2026, 4, 1)] + [datetime(y, 1, 1) for y in range(2027, 2031)]
    sdates = [datetime(y, 1, 1) for y in range(2025, 2031)]
    edates = [datetime(y, 12, 31) for y in range(2025, 2031)]


    # Compute avg_E_age properly using min and max dates per year
    min_dates = task_df_pf12.groupby('Outsource Year')['S&E Lodge Date'].min().tolist()
    max_dates = task_df_pf12.groupby('Outsource Year')['S&E Lodge Date'].max().tolist()

    avg_list = []
    for i in range(len(e_qty) - 1):
        try:
            avg_days = ((sdates_e[i] - min_dates[i]).days + (edates[i+1] - max_dates[i]).days) / 2
        except IndexError:
            avg_days = 0
        avg_list.append(e_qty[i+1] * avg_days)

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
    fy_list = ['FY25', 'FY26', 'FY27', 'FY28', 'FY29', 'FY30']
    for i, fy in enumerate(fy_list):
        # Use e_qty directly
        e = e_qty[i] if i < len(e_qty) else 0
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
