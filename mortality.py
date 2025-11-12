import pandas as pd

def calculate_mortality_dynamic(
    file_path='Hinsdale_excel.xlsx',
    gender='Female',
    start_age=60,
    end_age=65,
    start_year=2022,
    end_year=2030,
    mortality_column='General_EE_Female',
    base_year_of_mortality=2021
):
    # -------------------- Read Sheets --------------------
    mortality_df = pd.read_excel(file_path, sheet_name='Mortality')
    improvement_df = pd.read_excel(file_path, sheet_name='MP-2021')
    
    # Ensure Age is integer
    mortality_df['Age'] = mortality_df['Age'].astype(int)
    improvement_df['Age'] = improvement_df['Age'].astype(int)
    
    # Convert column names to string
    mortality_df.columns = [str(c).strip() for c in mortality_df.columns]
    improvement_df.columns = [str(c).strip() for c in improvement_df.columns]
    
    # Filter improvement factors for the selected gender
    imp_gender = improvement_df[improvement_df['Gender'].str.strip().str.lower() == gender.lower()]
    imp_gender = imp_gender.set_index('Age')
    
    # -------------------- Prepare tables --------------------
    long_rows = []
    pivot_dict = {}
    
    for age in range(start_age, end_age + 1):
        # Base mortality qx at base year
        qx_row = mortality_df.loc[mortality_df['Age'] == age, mortality_column]
        if qx_row.empty:
            print(f"[Warning] Age {age} not found in Mortality sheet")
            continue
        qx_base = qx_row.values[0]
        
        # Get improvement row for age
        if age not in imp_gender.index:
            print(f"[Warning] Age {age} not found in MP-2021 sheet")
            continue
        imp_row = imp_gender.loc[age]
        
        # -------------------- Forward or Backward projection --------------------
        qx_dict = {}  # Store for pivot table
        
        # Backward projection if start_year < base_year_of_mortality
        if start_year <= base_year_of_mortality:
            qx = qx_base
            for year in range(base_year_of_mortality - 1, start_year - 1, -1):
                mx_y = imp_row.get(str(year + 1)) if str(year + 1) in imp_row else imp_row.get(year + 1)
                if mx_y is None:
                    continue
                qxy = qx / (1 - mx_y)  # inverse for backward
                long_rows.append({
                    'Age': age,
                    'Year': year,
                    'qx': round(qx, 6),
                    'mx_y': round(mx_y, 6),
                    'qxy': round(qxy, 6)
                })
                qx = qxy
                qx_dict[year] = round(qxy, 6)
        
        # Forward projection
        qx = qx_base
        for year in range(max(start_year, base_year_of_mortality + 1), end_year + 1):
            mx_y = imp_row.get(str(year)) if str(year) in imp_row else imp_row.get(year)
            if mx_y is None:
                continue
            qxy = qx * (1 - mx_y)
            long_rows.append({
                'Age': age,
                'Year': year,
                'qx': round(qx, 6),
                'mx_y': round(mx_y, 6),
                'qxy': round(qxy, 6)
            })
            qx = qxy
            qx_dict[year] = round(qxy, 6)
        
        pivot_dict[age] = qx_dict
    
    # -------------------- Convert to DataFrames --------------------
    long_df = pd.DataFrame(long_rows)
    pivot_df = pd.DataFrame(pivot_dict).T
    pivot_df.index.name = 'Age'
    
    # -------------------- Cohort-style compact output --------------------
    cohort_rows = []
    for i, age in enumerate(range(start_age, end_age + 1)):
        year = start_year + i
        row = long_df[(long_df['Age'] == age) & (long_df['Year'] == year)]
        if not row.empty:
            qxy = row['qxy'].values[0]
            cohort_rows.append({'Age-Year': f"{age}-{year}", 'qxy': qxy})
    cohort_df = pd.DataFrame(cohort_rows)
    
    # -------------------- Print tables --------------------
    print("\n=== Long-format table ===")
    print(long_df)
    print("\n=== Pivot table (Age x Year) ===")
    print(pivot_df)
    print("\n=== Cohort-style Age-Year-qxy ===")
    print(cohort_df)
    
    return long_df, pivot_df, cohort_df


# -------------------- HARDCODED INPUTS --------------------
file_path = 'Hinsdale_excel.xlsx'
gender = 'Female'
start_age = 60
end_age = 65
start_year = 2022
end_year = 2030
mortality_column = 'General_EE_Female'
base_year_of_mortality = 2021

# -------------------- Run calculation --------------------
long_df, pivot_df, cohort_df = calculate_mortality_dynamic(
    file_path=file_path,
    gender=gender,
    start_age=start_age,
    end_age=end_age,
    start_year=start_year,
    end_year=end_year,
    mortality_column=mortality_column,
    base_year_of_mortality=base_year_of_mortality
)
