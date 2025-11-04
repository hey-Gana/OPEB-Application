from datetime import date
import pandas as pd

# -----------------------------
# Load Hinsdale Excel
# -----------------------------
file_path = "Hinsdale_excel.xlsx"
sheet_name = "Active Census"

df = pd.read_excel(file_path, sheet_name=sheet_name)

# Strip column names to avoid extra spaces
df.columns = df.columns.str.strip()

# Map the exact columns
name_col = 'Payment Form Value'
dob_col = 'DOB'
doh_col = 'DOH'
valuation_col = 'WAGES_PAID ENDING JUNE 30, 2023'

cols_needed = [name_col, dob_col, doh_col, valuation_col, 'EE_NO', 'SEX']
df_subset = df[cols_needed].copy()

# Convert DOB and DOH to datetime
df_subset[dob_col] = pd.to_datetime(df_subset[dob_col])
df_subset[doh_col] = pd.to_datetime(df_subset[doh_col])

# Calculate age at hire
df_subset['Age at Hire'] = df_subset.apply(
    lambda row: row[doh_col].year - row[dob_col].year - ((row[doh_col].month, row[doh_col].day) < (row[dob_col].month, row[dob_col].day)),
    axis=1
)

# -----------------------------
# OPEB Calculation Function 
# -----------------------------
"""
    Pension/CTC table with:
    - Separate salary growth rate (g) and discount rate (r)
    - Salary Increment Factor = (1 + g) ^ time_from_eval
    - Interest Discount = (1 + r) ^ -time_from_eval
    - St = Valuation Salary * Increment Factor * Interest Discount
    - PV of Salary from eval date = St * Interest Discount
    - DOH_Interest Discount = (1 + r) ^ -years of service
    - PV for DOH = St * DOH_Interest Discount
"""
def generate_salary_table_with_growth(date_of_hire, age_when_hired, end_age, discount_rate,
                                      salary_growth_rate, valuation_salary,
                                      evaluation_date):
    start_year = date_of_hire.year
    eval_year = evaluation_date.year
    end_year = start_year + (end_age - age_when_hired)
    years = list(range(start_year, end_year + 1))

    rows = []
    r = discount_rate
    g = salary_growth_rate

    for year in years:
        yos = year - start_year
        member_age = age_when_hired + yos
        time_from_eval = year - eval_year  # negative for past, positive for future

        # Salary Increment Factor
        salary_increment_factor = (1 + g) ** time_from_eval

        # Interest Discount
        interest_discount = (1 + r) ** (-time_from_eval)

        # Salary Transposed to Year (St)
        St = valuation_salary * salary_increment_factor * interest_discount

        # PV from Eval Date
        Pv_eval = St * interest_discount

        # PV from DOH
        trans_from_DOH_to_eval = (1 + r) ** (-yos)
        Sh = St * trans_from_DOH_to_eval

        rows.append({
            "Year": year,
            "Member Age": member_age,
            "Years of Service": yos,
            "Time from Evaluation": time_from_eval,
            "Salary Increment Factor": round(salary_increment_factor, 8),
            "Interest Discount": round(interest_discount, 8),
            "Salary @ Valuation Date": valuation_salary,
            "Salary Transposed {St}": round(St, 8),
            "Pv from Eval date": round(Pv_eval, 8),
            "Translating $ from DOH to Eval": round(trans_from_DOH_to_eval, 8),
            "PV from DOH {Sh}": round(Sh, 8)
        })

    df_emp = pd.DataFrame(rows)
    return df_emp

# -----------------------------
# Loop through employees and calculate
# -----------------------------
"""Constants"""
evaluation_date = date(2022, 1, 1)
end_age = 65
discount_rate = 0.0365
salary_growth_rate = 0.035

with pd.ExcelWriter("Hinsdale_OPEB_Output.xlsx") as writer:
    for idx, row in df_subset.iterrows():
        emp_name = row[name_col]
        valuation_salary = row[valuation_col]
        date_of_hire = row[doh_col]
        age_when_hired = row['Age at Hire']

        print(f"\nCalculating OPEB table for: {emp_name}")

        emp_table = generate_salary_table_with_growth(
            date_of_hire=date_of_hire,
            age_when_hired=age_when_hired,
            end_age=end_age,
            discount_rate=discount_rate,
            salary_growth_rate=salary_growth_rate,
            valuation_salary=valuation_salary,
            evaluation_date=evaluation_date
        )

        # Save each employee's table to a separate sheet
        sheet_name_safe = str(emp_name)[:31]  # Excel sheet name max length
        emp_table.to_excel(writer, sheet_name=sheet_name_safe, index=False)

print("\nOPEB calculations completed and saved to Hinsdale_OPEB_Output.xlsx")
