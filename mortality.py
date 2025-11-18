import pandas as pd
import numpy as np

def calculate_qxy_projection_compact(
    file_path='Hinsdale_excel.xlsx',
    gender='Female',
    person_age=54,
    person_year=2022,
    base_year=2010,
    max_age=65,
    mortality_column='General_EE_Female'
):

    # -------------------- Read Sheets --------------------
    mortality_df = pd.read_excel(file_path, sheet_name='Mortality')
    improvement_df = pd.read_excel(file_path, sheet_name='MP-2021')

    # Ensure correct types
    mortality_df['Age'] = mortality_df['Age'].astype(int)
    improvement_df['Age'] = improvement_df['Age'].astype(int)

    mortality_df.columns = [str(c).strip() for c in mortality_df.columns]
    improvement_df.columns = [str(c).strip() for c in improvement_df.columns]

    # Filter improvement factors for selected gender
    imp = improvement_df[improvement_df['Gender'].str.lower() == gender.lower()]
    imp = imp.set_index('Age')

    # -------------------- PRINT INPUT DATAFRAMES --------------------
    print("\n============ Mortality Sheet ============")
    print(mortality_df)

    print("\n============ MP-2021 (Raw) ============")
    print(improvement_df)

    print("\n============ MP-2021 (Filtered for Gender, Age Indexed) ============")
    print(imp)

    # -------------------- Determine year range --------------------
    start_col_year = min(person_year, base_year)
    end_col_year = person_year + (max_age - person_age)
    year_range = list(range(start_col_year, end_col_year + 1))

    # -------------------- Prepare Full q_x,y Table --------------------
    result_df = pd.DataFrame(index=range(person_age, max_age + 1),
                             columns=['Year'] + year_range)
    result_df.index.name = 'Age'

    for age in range(person_age, max_age + 1):
        result_df.loc[age, 'Year'] = person_year + (age - person_age)

    # -------------------- Compute q_x,y --------------------
    for age in range(person_age, max_age + 1):

        # Base q_x for base year
        qx_base_row = mortality_df.loc[mortality_df['Age'] == age, mortality_column]
        if qx_base_row.empty:
            continue
        qx_base = qx_base_row.values[0]

        # Improvement factor row
        if age not in imp.index:
            continue
        imp_row = imp.loc[age]

        q_map = {base_year: qx_base}

        # ----- BACKWARD PROJECTION -----
        q_prev = qx_base
        for y in range(base_year - 1, start_col_year - 1, -1):
            next_year = y + 1
            m = imp_row.get(str(next_year), np.nan)
            if pd.isna(m):
                q_map[y] = np.nan
                continue
            qy = q_prev / (1 - m)
            q_map[y] = qy
            q_prev = qy

        # ----- FORWARD PROJECTION -----
        q_prev = qx_base
        for y in range(base_year + 1, end_col_year + 1):
            m = imp_row.get(str(y), np.nan)
            if pd.isna(m):
                q_map[y] = np.nan
                continue
            qy = q_prev * (1 - m)
            q_map[y] = qy
            q_prev = qy

        # Save into result_df
        for y in year_range:
            result_df.loc[age, y] = q_map.get(y, np.nan)

    # -------------------- PRINT FULL q_x,y TABLE --------------------
    print("\n============ FULL q_x,y Table ============")
    print(result_df)

    # -------------------- Create Compact Age-Year-qx Table --------------------
    compact_rows = []
    for age in result_df.index:
        year = int(result_df.loc[age, 'Year'])
        qx = result_df.loc[age, year]
        compact_rows.append({'Age': age, 'Year': year, 'qx_year': qx})

    compact_df = pd.DataFrame(compact_rows)
    compact_df['qx_year'] = compact_df['qx_year'].round(8)

    # -------------------- PRINT COMPACT TABLE --------------------
    print("\n============ COMPACT Age-Year-qx Table ============")
    print(compact_df)

    return result_df, compact_df


# -------------------- Example Run --------------------
full_df, compact_df = calculate_qxy_projection_compact(
    file_path='Hinsdale_excel.xlsx',
    gender='Female',
    person_age=54,
    person_year=2022,
    base_year=2010,
    max_age=65,
    mortality_column='General_EE_Female'
)
