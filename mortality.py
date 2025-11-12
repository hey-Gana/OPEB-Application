import pandas as pd

def calculate_mortality(file_path='Hinsdale_excel.xlsx',
                        gender='Female',
                        start_age=60,
                        end_age=65,
                        start_year=2021,
                        end_year=2030,
                        mortality_column='General_EE_Female'):
    
    # -------------------- Read Sheets --------------------
    mortality_df = pd.read_excel(file_path, sheet_name='Mortality')
    improvement_df = pd.read_excel(file_path, sheet_name='MP-2021')
    
    # -------------------- Clean & Prepare --------------------
    # Ensure Age is integer
    mortality_df['Age'] = mortality_df['Age'].astype(int)
    improvement_df['Age'] = improvement_df['Age'].astype(int)
    
    # Strip column names safely (convert to string first)
    mortality_df.columns = [str(c).strip() for c in mortality_df.columns]
    improvement_df.columns = [str(c).strip() for c in improvement_df.columns]
    
    # Filter improvement factors for the selected gender
    imp_gender = improvement_df[improvement_df['Gender'].str.strip().str.lower() == gender.lower()]
    imp_gender = imp_gender.set_index('Age')
    
    # -------------------- Long-format table --------------------
    long_rows = []
    
    # -------------------- Pivot table dictionary --------------------
    pivot_dict = {}
    
    for age in range(start_age, end_age + 1):
        # Get base mortality qx for age
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
        
        qx = qx_base
        pivot_dict[age] = {}
        
        for year in range(start_year + 1, end_year + 1):
            # Dynamically get mx_y (try int or string column)
            mx_y = imp_row.get(str(year)) if str(year) in imp_row else imp_row.get(year)
            if mx_y is None:
                print(f"[Warning] Year {year} not found for age {age} in MP-2021")
                continue
            
            qxy = qx * (1 - mx_y)
            
            # Save in long-format table
            long_rows.append({
                'Age': age,
                'Year': year,
                'qx': round(qx, 6),
                'mx_y': round(mx_y, 6),
                'qxy': round(qxy, 6)
            })
            
            # Save in pivot table
            pivot_dict[age][year] = round(qxy, 6)
            
            # Show accessed cells for verification
            print(f"Accessing Mortality[Age={age}] = {qx:.6f}, MP-2021[Age={age}, Year={year}] = {mx_y:.6f} => qxy = {qxy:.6f}")
            
            # Update qx for next year
            qx = qxy
    
    # -------------------- Convert long-format table --------------------
    long_df = pd.DataFrame(long_rows)
    
    # -------------------- Convert pivot table --------------------
    pivot_df = pd.DataFrame(pivot_dict).T  # Age as rows, Years as columns
    pivot_df.index.name = 'Age'
    
    # -------------------- Print results --------------------
    print("\n=== Long-format table ===")
    print(long_df)
    
    print("\n=== Pivot table (Age x Year) ===")
    print(pivot_df)

    return long_df, pivot_df


# -------------------- HARDCODED INPUTS --------------------
file_path = 'Hinsdale_excel.xlsx'
gender = 'Female'
start_age = 60
end_age = 65
start_year = 2021
end_year = 2030
mortality_column = 'General_EE_Female'

# Run calculation
long_df, pivot_df = calculate_mortality(
    file_path=file_path,
    gender=gender,
    start_age=start_age,
    end_age=end_age,
    start_year=start_year,
    end_year=end_year,
    mortality_column=mortality_column
)
