from datetime import date
import pandas as pd

def generate_salary_table_with_growth(date_of_hire, age_when_hired, end_age, discount_rate,
                                      salary_growth_rate, valuation_salary,
                                      evaluation_date, excel_path="salary_table_growth.xlsx"):
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
    start_year = date_of_hire.year
    eval_year = evaluation_date.year
    
    # Calculate end year from end_age
    end_year = start_year + (end_age - age_when_hired)

    # Generate list of years
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
            "Salary Increment Factor": round(salary_increment_factor, 5),
            "Interest Discount": round(interest_discount, 5),
            "Salary @ Valuation Date": valuation_salary,
            "Salary Transposed {St}": round(St, 5),
            "Pv from Eval date": round(Pv_eval, 5),
            "Translating $ from DOH to Eval": round(trans_from_DOH_to_eval, 5),
            "PV from DOH {Sh}": round(Sh, 5)
        })

    df = pd.DataFrame(rows)

    # Save to Excel
    try:
        df.to_excel(excel_path, index=False)
        print(f"Salary table saved to {excel_path}")
    except Exception as e:
        print("Error saving Excel:", e)

    return df

if __name__ == "__main__":
    date_of_hire = date(2007, 1, 1)
    age_when_hired = 39
    end_age = 65           # target retirement age
    discount_rate = 0.0365  
    salary_growth_rate = 0.035    
    valuation_salary = 61652
    evaluation_date = date(2022, 1, 1)

    df = generate_salary_table_with_growth(
        date_of_hire=date_of_hire,
        age_when_hired=age_when_hired,
        end_age=end_age,
        discount_rate=discount_rate,
        salary_growth_rate=salary_growth_rate,
        valuation_salary=valuation_salary,
        evaluation_date=evaluation_date,
        excel_path="salary_table_growth.xlsx"
    )

    pd.set_option("display.float_format", lambda x: f"{x:,.8f}")
    print(df.to_string(index=False))
