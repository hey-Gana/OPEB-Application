
# Read Active Census #
library(magrittr)
library(dplyr)
library(readxl)
library(tidyverse)
library(lubridate)
library(zoo) 
library(data.table) #added for setDT function 

valuation_date <- as.Date("2022-07-01")
run_name <- "EYDR"
toggle_proval_soa_adj <- TRUE # SET TRUE to apply ProVal's SOA rules to final PREMIUM calc, FALSE for original logic

# Parameters
# ------------------------------------------------------------------
# DYNAMIC ASSUMPTIONS  (replaces old hard-coded defaults list)
# ------------------------------------------------------------------
plan_spec <- list(
  # folder_path       = "C:/Users/fleis/Documents/Rafi 2025/",
  folder_path       = "/Users/ganapathisubramaniam/Downloads/OPEB",
  file_name         = "Hinsdale_excel.xlsx",
  assumptions_sheet = "Assumptions",
  mortality_sheet   = "Mortality",
  improvement_sheet = "MP-2021",
  actives_sheet     = "Active Census",
  retirees_sheet    = "Retiree Census",
  soa_sheet         = "SOA_Adj"
)

assump_raw <- readxl::read_excel(
  file.path(plan_spec$folder_path, plan_spec$file_name),
  sheet = plan_spec$assumptions_sheet,
  col_names = FALSE, range = "A1:C12" #UPDATE IF YOU EDIT ASSUMPTIONS TAB
)

#Read assumptions into R data to use in Actives
assump <- assump_raw %>%
  mutate(
    label = trimws(`...1`),
    sex   = trimws(`...3`)
  ) %>%
  transmute(
    key = case_when(
      label == "Interest rate"        ~ "i_rate",
      label == "Salary increase"      ~ "salary_increase",
      label == "Election Probability" ~ "elect_prob",
      label == "Fraction married"     ~ paste0("frac_married_", tolower(sex)),
      label == "Spouse age difference"~ paste0("sp_age_diff_", tolower(sex)),
      label == "base_year"             ~ "base_year",
      label == "max_benefit_age"       ~ "max_benefit_age",
      label == "PRE_EE"               ~ "premium_ee" 
    ),
    value = as.numeric(`...2`)
  ) %>% filter(!is.na(key))

# Create the final `defaults` list by merging static file paths with dynamic assumptions from Excel
defaults <- c(
  plan_spec,
  as.list(setNames(assump$value, assump$key))       
)

# --- Create Spouse Age Assumption Variables ---
sp_diff_male   <- defaults$sp_age_diff_male   # Age difference for male EEs (-3 in  sheet)
sp_diff_female <- defaults$sp_age_diff_female # Age difference for female EEs (+3 in  sheet)


# Compact date conversion function
# Will standardize date format and handle excel date serials
convert_to_date <- function(x) {
  x_chr <- as.character(x)
  n      <- length(x_chr)
  out <- rep(as.Date(NA), n)                                        # init all NA Date
  is_serial <- !is.na(x_chr) & grepl("^[0-9]+$", x_chr, perl=TRUE)  # identify Excel serials: purely digits (no NAs)
  if (any(is_serial, na.rm=TRUE)) { # if any serials                # serial → Date
    out[is_serial] <- as.Date(
      as.numeric(x_chr[is_serial]),
      origin = "1899-12-30"
    )
  }
  is_text <- !is_serial & !is.na(x_chr)                             # identify text dates (not serial, not NA)        
  
  if (any(is_text)) {
    parsed <- parse_date_time(                                      # parse both m/d/Y and Y-m-d
      x_chr[is_text],
      orders = c("mdy", "Ymd"),
      exact  = FALSE,
      quiet  = TRUE
    )
    out[is_text] <- as_date(parsed)
  }
  out
}


# Define helper function for date/year indexing
round_half_up <- function(x) {
  floor(x + 0.5)
}

# ------------------------------------------------------------------
#  Helper – pad retirement / termination tables and merge qx’s
# ------------------------------------------------------------------
process_qx_data <- function(qx_src, actives_dt, qx_col) {
  
  ## 1 — Dense Age-Service grid for every (Cat, Gender, DOH window)
  full_grid <- CJ(
    Category  = unique(qx_src$Category),
    Gender    = unique(qx_src$Gender),
    DOH_Start = unique(qx_src$DOH_Start),
    DOH_End   = unique(qx_src$DOH_End),
    Age       = 0:120,
    Service   = 0:50,
    unique    = TRUE
  )
  
  ## 2 — Pad and two-way LOCF to fill gaps
  padded <- qx_src[full_grid,
                   on = .(Category, Gender, DOH_Start, DOH_End, Age, Service)]
  
  padded[, (qx_col) := zoo::na.locf(get(qx_col),  na.rm = FALSE),
         by = .(Category, Gender, DOH_Start, DOH_End)]
  padded[, (qx_col) := zoo::na.locf(get(qx_col), fromLast = TRUE, na.rm = FALSE),
         by = .(Category, Gender, DOH_Start, DOH_End)]
  
  ## 3 — Match rates to Actives for the correct DOH window
  actives_dt[, row_id := .I]                   # stable key
  matched <- padded[
    actives_dt,
    on = .(Category, Gender, Age = EE_Age, Service),
    allow.cartesian = TRUE
  ][DOH >= DOH_Start & DOH < DOH_End
  ][, .SD[which.max(DOH_Start)], by = row_id]  # pick latest window
  
  ## 4 — Safe merge-back, no tricky i.* syntax
  actives_dt <- merge(
    actives_dt,
    matched[, .(row_id, tmp_qx = get(qx_col))],
    by = "row_id",
    all.x = TRUE,
    sort = FALSE
  )[, (qx_col) := tmp_qx][, `:=`(tmp_qx = NULL, row_id = NULL)]
  
  invisible(actives_dt)
}


### READ and PREPARE Cencus Data - EDITED FOR PROVAL DATE/YEAR INDEXING###
Actives <- read_excel(file.path(defaults$folder_path, defaults$file_name),
                      sheet = defaults$actives_sheet,
                      col_types = "text") %>% 
  rename(Name = `Lk\177o|6*Woqrkx*V*`,
         Wages_Paid = all_of(names(.)[grepl("wages", names(.), ignore.case = TRUE)])) %>% 
  mutate(
    # a) dates
    across(any_of(c("DOB","DOT","DOR","SP_DOB","DOH")), convert_to_date),
    
    # b) core age calc
    age_val_exact = as.numeric(difftime(valuation_date, DOB, units = "days"))/365.25,
    age_val_rnd   = round_half_up(age_val_exact),
    
    # c) gender FIRST so later code can use it
    Gender = str_to_title(recode(SEX,
                                 "M"="Male","F"="Female",
                                 "m"="Male","f"="Female")),
    SP_Gender = if_else(Gender == "Male", "Female", "Male"), # Define spouse's gender as the opposite 
    
    
    # e) rest of  existing fields …
    hire_year_val_date = as.Date(sprintf("%d-07-01", year(DOH))),
    EntryAge   = round_half_up(as.numeric(difftime(hire_year_val_date, DOB, units="days"))/365.25),
    FundingAge = if_else(DOH <= hire_year_val_date, EntryAge,
                         round_half_up(as.numeric(difftime(hire_year_val_date + years(1), DOB, units="days"))/365.25)),
    funding_diff = FundingAge - EntryAge,
    
    EE_Age  = age_val_rnd,
    Service = as.integer(age_val_rnd - EntryAge),
    BaseSalary = coalesce(readr::parse_number(Wages_Paid), 0),
    
    SP_Age = EE_Age + ifelse(Gender == "Male",
                             sp_diff_female,      # +3 from sheet
                             sp_diff_male),   # –3 from sheet
    
    
    EE_Age_partial = round(age_val_exact - age_val_rnd, 4),
    SP_Age_partial = NA_real_,   # SP_DOB no longer used
    Year = year(valuation_date),
    
    Category = case_when(
      str_to_lower(Instructional) == "instructional"         ~ "Teachers",
      str_detect(str_to_lower(Instructional), "safety|police") ~ "Safety",
      TRUE                                                   ~ "General"
    )
  )


Retirees <- read_excel(file.path(defaults$folder_path, defaults$file_name), 
                       sheet = defaults$retirees_sheet,
                       col_types = "text") %>%
  rename(Name = `Klly~~6Wsmrkov`) %>%
  mutate(across(any_of(c("DOB", "DOT", "DOR", "SP_DOB", "DOH")), convert_to_date),
         Gender = recode(SEX,
                         "M" = "Male", "F" = "Female",
                         "m" = "Male", "f" = "Female"),
         SP_Gender = if_else(Gender == "Male", "Female", "Male"), #New - needed for qx_SP retiree lookups
         Year = year(valuation_date),
         across(.cols = contains("DO"), .fns = ~ as.Date(.)),
         EE_Age = floor(as.numeric(difftime(valuation_date, ymd(DOB), units = "days")) / 365.25),
         EE_Age_partial = round((as.numeric(difftime(valuation_date, ymd(DOB), units = "days")) / 365.25) - EE_Age, 4),
         SP_Age = floor(as.numeric(difftime(valuation_date, ymd(SP_DOB), units = "days")) / 365.25),
         SP_Age_partial = round((as.numeric(difftime(valuation_date, ymd(SP_DOB), units = "days")) / 365.25) - SP_Age, 4)) %>%
  mutate(Category = case_when(
    str_to_lower(Instructional) == "instructional" ~ "Teachers",
    str_detect(str_to_lower(Instructional), "safety|police") ~ "Safety",
    TRUE ~ "General")) %>%
  select(-contains("DOB"), -contains("SP_DOB"),-SEX)

### Actives and Retirees Data cleaning ###


# For Calculate Actives Avg Age for SOA Adjustment
Actives_EE_Age <- Actives %>%  filter(EE_Age < 65) %>% select(EE_Age)
Retirees_EE_Age <- Retirees %>%  filter(EE_Age < 65) %>% select(EE_Age)
Pre_65_Avg_Age <- bind_rows(Actives_EE_Age, Retirees_EE_Age) %>% summarise(Avg_Age = floor(mean(EE_Age, na.rm = TRUE))) %>%  pull(Avg_Age)
rm(Actives_EE_Age, Retirees_EE_Age)
### End Data Cleaning ###

#import soa adj and scale  adj to avg age 
SOA_Adj <- read_excel(file.path(defaults$folder_path, defaults$file_name), sheet = defaults$soa_sheet) %>%
  rename(Age_Adj = Unisex) %>%
  mutate(Avg_Age_Adj =  Age_Adj / Age_Adj[Age == Pre_65_Avg_Age])





## Apply mortality improvement factors (by Gender)
# 1. Read raw mortality table and pivot to long form
raw_mort <- read_excel(file.path(defaults$folder_path, defaults$file_name), sheet = defaults$mortality_sheet)
mortality <- raw_mort %>%
  pivot_longer(-Age, names_to = "Table", values_to = "Rate") %>%
  separate(Table, into = c("Category", "Status", "Gender"), sep = "_", remove = FALSE) %>%
  mutate(
    Gender = recode(Gender,
                    "M" = "Male", "F" = "Female",
                    "m" = "Male", "f" = "Female"),
    Act_Ret = recode(Status, "EE" = "A", "Ret" = "R")
  ) %>%
  rename(q_base = Rate) %>%
  select(Category, Gender, Act_Ret, Age, q_base)

# 2. Load improvement factors and standardize Gender
improvement <- read_excel(file.path(defaults$folder_path, defaults$file_name), sheet = defaults$improvement_sheet) %>%
  pivot_longer(-c(Age, Gender), names_to = "Year", values_to = "m") %>%
  mutate(
    Year = as.integer(Year),
    Gender  = recode(Gender, "M" = "Male", "F" = "Female")
  )
print(summary(improvement$m))    # should say roughly -0.05 … +0.07


defaults$base_year <- 2010

# 3. Merge and compute cumulative projected improvement rates for each Category & Gender
# ---- FIXED mortality-improvement logic ----
mortality_imprv <- mortality %>%                       # <- exists above
  inner_join(improvement, by = c("Age", "Gender")) %>%
  group_by(Category, Gender, Age) %>%
  arrange(Year) %>%
  mutate(
    one_minus_m   = 1 - m,
    cumprod_start = cumprod(one_minus_m),          # product up to *this* year
    base_prod     = cumprod_start[Year == defaults$base_year][1],
    factor        = if_else(Year >= defaults$base_year,
                            cumprod_start / base_prod,      # 2010 forward
                            base_prod / cumprod_start)      # pre-2010
  ) %>%
  ungroup() %>%
  mutate(qx = q_base * factor)
# ---- END fixed block ----



#Adds Run columns for EYDR vs PYDR 
add_run_column_and_rename <- function(original_name) {
  suffix <- sub(".*_(\\w+)$", "\\1", original_name)
  new_name <- sub("_(\\w+)$", "", original_name)
  assign(new_name, dplyr::mutate(get(original_name), Run = suffix), envir = .GlobalEnv)
  rm(list = original_name, envir = .GlobalEnv)  # Remove the original variable
}




retirement_EYDR <- read_excel(
  file.path(defaults$folder_path, defaults$file_name), #updated to pull file path from defaults
  sheet     = "Retirement_EYDR", 
  col_types = "text"
) %>%
  mutate(across(any_of(c("DOH_Start", "DOH_End")), convert_to_date),
         Age = as.numeric(Age)) %>%
  pivot_longer(cols      = matches("^\\d+$"),  # the service columns 0–30
               names_to  = "Service",
               values_to = "qx_ret") %>%
  mutate(Service = as.numeric(Service),
         qx_ret  = as.numeric(qx_ret))
add_run_column_and_rename("retirement_EYDR")

term_EYDR <- read_excel(
  file.path(defaults$folder_path, defaults$file_name), 
  sheet     = "Term_EYDR", 
  col_types = "text"
) %>%
  mutate(across(any_of(c("DOH_Start", "DOH_End")), convert_to_date),
         Age = as.numeric(Age)) %>%
  pivot_longer(cols      = matches("^\\d+$"),  # the service columns 0–30
               names_to  = "Service",
               values_to = "qx_term") %>%
  mutate(Service = as.numeric(Service),
         qx_term = as.numeric(qx_term))
add_run_column_and_rename("term_EYDR")

medical_costs <- read_excel(
  file.path(defaults$folder_path, defaults$file_name), 
  sheet = "Medical_costs"
)









medical_premiums <- medical_costs %>%
  filter(grepl("_", Run)) %>%
  rename(Premium = Rate) %>%
  separate(Run, into = c("Run", "Act_Ret", "EE_SP", "Plan"), sep = "_", extra = "merge")

medical_trend <- medical_costs %>%
  filter(!grepl("_", Run))
rm(medical_costs)



### Build Actives and Retirees tables ###
#— 1. Convert everything to data.table
setDT(Actives)
setDT(Retirees)
setDT(mortality_imprv)
setDT(retirement)
setDT(term)
setDT(medical_trend)
setDT(SOA_Adj)

# ----  fraction‑married flag now that Actives exists 
Actives[, frac_married :=
          fifelse(Gender == "Male",
                  defaults$frac_married_male,
                  defaults$frac_married_female)]



#— 2. Expand Actives from one row per member to one row per year of service
Actives <- Actives[, {
  # Define constants for this member
  age_start <- EE_Age[1] - Service[1]
  l <- 0:(120 - floor(age_start)) # 'l' is the service year, starting from 0
  
  # Create the expanded grid with all core columns
  expanded <- .SD[rep(1L, length(l))][, `:=`(
    Year    = Year[1] - Service[1] + l,
    EE_Age  = age_start + l,
    # old spouse logic
    # SP_Age  = if (!is.na(SP_Age[1])) SP_Age[1] - Service[1] + l else NA_real_,
    SP_Age  = (age_start + l) + ifelse(Gender[1] == "Male",
                                       sp_diff_female,
                                       sp_diff_male),
    Service = l,
    ValuationSalary = BaseSalary[1] *
      (1 + defaults$salary_increase)^(l - Service[1])
  )]
  
  # --- Add back the essential PV and Discount columns ---
  expanded[, years_since_val := ifelse(Year >= year(valuation_date), Year - year(valuation_date), NA_integer_)]
  expanded[, disc_since_val := fifelse(is.na(years_since_val), NA_real_, (1 + defaults$i_rate)^(-years_since_val))]
      #stub year logic for disc_since_service
  expanded[, disc_since_Service := 
             fifelse(Service < funding_diff[1], 0,
                     (1 + defaults$i_rate)^(-(Service - funding_diff[1])))]
  expanded
}, by = EE_NO]

max_age <- 100  # Set your desired maximum age - dont mess with this or else errors in retirees


#— 3. Expand Retirees to one row per year
Retirees <- Retirees[, {
  a <- floor(EE_Age[1])
  if (a >= max_age) {.SD[0]}
  else {l <- 0:(max_age - a)
  expanded <- .SD[rep(1L, length(l))][, `:=`(
    Year    = Year[1] + l,
    EE_Age  = EE_Age[1] + l,
    SP_Age  = if (!is.na(SP_Age[1])) SP_Age[1] + l else NA_real_)]
  expanded[, time_since_val := ifelse(Year >= as.integer(format(valuation_date, "%Y")), Year - as.integer(format(valuation_date, "%Y")), NA_integer_)]
  expanded[, disc_since_val := fifelse(is.na(time_since_val),NA_real_,(1 + defaults$i_rate) ^ (-time_since_val)
  )]
  expanded
  }
}, by = EE_NO]


### Join All decrements and Health costs to Actives and Retirees tables ###
# CORRECT ORDER: These joins now happen AFTER the tables have been expanded by year.

#— 4. Mortality (Act_Ret: Actives="A", Retirees="R") → qx_EE & qx_SP
# Actives
Actives[ mortality_imprv[Act_Ret == "A",
                         .(Category, Gender, Age, Year, qx)],
         on = .(Category,
                Gender,
                EE_Age = Age,
                Year),
         qx_EE := i.qx ]

## 4b.  Spouse mortality  (uses SP_Gender we just added)
Actives[ mortality_imprv[Act_Ret == "A",
                         .(Category, Gender, Age, Year, qx)],
         on = .(Category,
                SP_Gender = Gender,
                SP_Age     = Age,
                Year),
         qx_SP := i.qx ]

# Retirees
Retirees[ mortality_imprv[Act_Ret == "R",
                          .(Category, Gender, Age, Year, qx)],
          on = .(Category,
                 Gender,
                 EE_Age = Age,
                 Year),
          qx_EE := i.qx ]

Retirees[ mortality_imprv[Act_Ret == "R",
                          .(Category, Gender, Age, Year, qx)],
          on = .(Category,
                 SP_Gender = Gender,
                 SP_Age     = Age,
                 Year),
          qx_SP := i.qx ] #retiree spouses




#— 5. Retirement and Termination decrements
#    (The process_qx_data function is correct and joins on the expanded table)
Actives <- process_qx_data(retirement, Actives, "qx_ret")
Actives <- process_qx_data(term, Actives, "qx_term")

# --- Start of New Code Block: Final ProVal Decrement Timing Adjustment ---

# This creates the funding_diff column if it doesn't exist
if (!"funding_diff" %in% names(Actives)) {
  Actives[, funding_diff := FundingAge - EntryAge]
}

## Termination: Apply a universal 1-year lag to correct the lookup offset
Actives[order(EE_NO, Service),
        qx_term := shift(qx_term, 1L, type = "lag", fill = 0),
        by = EE_NO]

## Retirement: Apply a conditional lag to fix the offset ONLY for funding_diff=1 members
Actives[order(EE_NO, Service),
        qx_ret := {
          off <- max(funding_diff[1] - 1L, 0L) 
          shift(qx_ret, off, type = "lag", fill = 0)
        },
        by = EE_NO]



### CORRECTED Health Trend Processing Block ###

# --- Step 1: Read the raw cost data ---
medical_costs <- read_excel(
  file.path(defaults$folder_path, defaults$file_name), 
  sheet = "Medical_costs"
)

# --- Step 2: Define which trend scenario to use ---
run_name <- "EYDR" # Set the single, correct scenario for this valuation

# --- Step 3: Build the clean, shifted, and padded trend table ---

# First, filter for ONLY the specified run's trend data for actives
med_tr_base <- medical_costs %>%
  filter(
    !grepl("_", Run),     # Keep only trend rows (those without underscores)
    Run == run_name,      # Isolate the single, correct scenario (e.g., "EYDR")
    Act_Ret == "A"        # Keep only rates applicable to Actives
  ) %>%
  select(Year_Start, Year_End, Rate)

setDT(med_tr_base)

# Next, apply the 1-year forward shift to match ProVal's timing convention
med_tr_shifted <- med_tr_base[, .(
  Year_Start = Year_Start + 1L, # A rate for 2022 applies to costs in 2023
  Year_End   = Year_End + 1L,
  Rate
)][order(Year_Start)]

# Then, create the historical padding row. All years before our schedule starts
# will use the rate from the first available period (8.2% in this case).
pad <- data.table(
  Year_Start = 1900L,
  Year_End   = med_tr_shifted$Year_Start[1] - 1L,
  Rate       = med_tr_shifted$Rate[1]
)

# Finally, combine the historical padding with the shifted schedule
med_tr_final <- rbindlist(list(pad, med_tr_shifted), use.names = TRUE)

# --- Step 4: Join the corrected trend table to the main Actives table ---
Actives[
  med_tr_final,
  on = .(Year >= Year_Start, Year <= Year_End),
  health_trend := i.Rate
]

# --- Step 5: Clean up the intermediate tables ---
rm(medical_costs, med_tr_base, med_tr_shifted, pad, med_tr_final)





#— 7. SOA Age Adjustment → Avg_Age_Adj
soa <- SOA_Adj[, .(Age, Avg_Age_Adj)]
Actives[
  soa,
  on = .(EE_Age = Age),
  roll = "nearest", rollends = c(TRUE, TRUE),
  Avg_Age_Adj := i.Avg_Age_Adj
]







#— Final Survival Logic
#    This entire block now runs LAST, after all the correct qx rates are in place.

### Blake: 0-out termination if retirement eligible
Actives[qx_ret > 0 & !is.na(qx_ret), qx_term := 0]

### ProVal-style “stub” handling for survival ###
# a. Identify the first service year in which decrements apply
Actives[ , decrement_start_service := ifelse(funding_diff == 1, 2L, 1L) ]

# b. Calculate the one-year survival vector
Actives[ ,
         survival := fifelse(Service < decrement_start_service,
                             1,                                           # Years before decrements kick in are 100%
                             pmax(0, 1 - qx_EE - qx_ret - qx_term))       # Normal decrement years
]

# c. Calculate the initial cumulative survival
Actives[order(EE_NO, Service),
        cum_survival := cumprod(survival),
        by = EE_NO]

# d. Zero-out the bookkeeping row for cases with a 1-year funding gap
Actives[ funding_diff == 1 & Service == 0, cum_survival := 0 ]

# e. Clean up helper column
Actives[, decrement_start_service := NULL]


## --- Pure-mortality survival (no term / retire) ------------------

# a. Identify when decrements begin (same proval rule used before)
Actives[ , decrement_start_service :=
           fifelse(funding_diff == 1, 2L, 1L) ]

# b. One-year mortality-only survival
Actives[ ,
         survival_pure := fifelse(
           Service < decrement_start_service,     # stub rows
           1,                                      # always 1
           pmax(0, 1 - qx_EE)                     # otherwise 1-qx
         )
]

# c. Cumulative mortality-only survival
Actives[order(EE_NO, Service),
        cum_survival_pure := cumprod(survival_pure),
        by = EE_NO]

## --- Spouse pure-mortality survival (mirroring EE logic) ------------------

# 1. Calculate one-year spouse survival, handling stub years and missing rates
Actives[, survival_pure_SP := fifelse(
  Service < fifelse(funding_diff == 1, 2L, 1L), # Check for stub years (no decrements apply)
  1,                                           # Survival is 100% during stub period
  pmax(0, 1 - coalesce(qx_SP, 0))              # Otherwise, it's 1-qx. `coalesce` handles missing qx and spits a 0
)]

# 2. Calculate the cumulative survival probability from hire
Actives[order(EE_NO, Service), 
        cum_survival_pure_SP := cumprod(survival_pure_SP), 
        by = EE_NO]

# 3. Force the bookkeeping row to zero for members with a 1-year funding gap
Actives[funding_diff == 1 & Service == 0, cum_survival_pure_SP := 0]

# d. Zero-out the bookkeeping stub row for funding_diff == 1
Actives[funding_diff == 1 & Service == 0,
        cum_survival_pure := 0]

# e. Clean-up helper column
Actives[, decrement_start_service := NULL]

## --- Calculate Final PV Salary Columns (Post-Survival) --------
# This is here done now that cum_survival is correctly calculated.
Actives[, `:=`(
  PVSalary_fromHire = ValuationSalary * disc_since_Service * cum_survival,
  PVSalary_fromVal  = fifelse(is.na(disc_since_val), 0, #handles 0 rows in disc_since_val
                              ValuationSalary * disc_since_val * cum_survival)
)]
# --- End PV Salary Calculation 


## --- cap salaries at age‑70 and drop rows above 70 -----------------
Actives[EE_Age >= 70L, `:=`(PVSalary_fromHire = 0,     # hard‑code both PV cols to 0 at 70 (proval logic)
                            PVSalary_fromVal  = 0)]    # when age hits 70
Actives <- Actives[EE_Age <= 70L]                      # strip 71+ rows




# --- End ProVal Logic ---



#— Done: Actives now has:
#    qx_EE, qx_SP,
#    qx_ret_EE,
#    qx_term_EE,
#    health_trend,
#    Premium_EE, Premium_SP,
#    Avg_Age_Adj




#### — Create new adjusted‐premium columns
# 1. Calculate cumulative health trend, anchored to the valuation date
val_year <- year(valuation_date)
Actives[
  order(EE_NO, Year),
  cum_health_trend := {
    r <- fifelse(abs(health_trend) > 1, health_trend / 100, health_trend)
    cf <- cumprod(1 + r)
    anchor_idx <- match(val_year, Year)
    if (is.na(anchor_idx)) NA_real_ else cf / cf[anchor_idx]
  },
  by = EE_NO
]




#####################################LOOK HERE#########################


# 3. Create base premium columns for clarity and QA
Actives[, Premium_EE_base := defaults$premium_ee]
Actives[, Premium_SP_base := defaults$premium_ee] # make sure this points to spouse assumption 

# 4. Calculate premiums and apply toggled SOA adjustment rules

Actives[, `netPRE-EE` := Premium_EE_base * cum_health_trend]
Actives[, `netPRE-SP` := Premium_SP_base * cum_health_trend]

# First, calculate the standard adjusted premium for everyone
Actives[, Premium_EE_adj := `netPRE-EE` * Avg_Age_Adj]
Actives[, Premium_SP_adj := `netPRE-SP` * Avg_Age_Adj]

# THEN, if the toggle is on, apply the special ProVal rules
if (toggle_proval_soa_adj) {
  Actives[EE_Age <= 39, `:=` (Premium_EE_adj = 0, Premium_SP_adj = 0)]
  Actives[EE_Age >= 65, `:=` (Premium_EE_adj = 0, Premium_SP_adj = 0)]
}






## ---------- ELIGIBILITY FLAG ----------
elig_cut <- as.Date("2011-07-01")

# helper lambdas 
###### needs to be dynamically linked to assumptions 
elig_pre  <- function(a, s) a >= 60 |
  (a >= 50 & s >= 10) |
  (s >= 20 & (a + s) >= 70)
elig_post <- function(a, s) a >= 65 |
  (a >= 60 & s >= 30)
###### left join
Actives[, Eligibility :=
          fifelse(
            (DOH <  elig_cut & elig_pre (EE_Age, Service)) |
              (DOH >= elig_cut & elig_post(EE_Age, Service)),
            1L, 0L)
]






# --- Error Handling before producing output---
stopifnot(
  "One or more required assumptions (i_rate, salary_increase, elect_prob) are missing from the defaults list." =
    all(c("i_rate", "salary_increase", "elect_prob") %in% names(defaults)),
  
  "NA values were found in the projected ValuationSalary column in the Actives table." =
    !anyNA(Actives$ValuationSalary)
)


## ---- 0A · guarantee a clean Year column -------------------------------
if (!"Year" %in% names(Actives)) {
  yr_alt <- grep("^Year\\s*$", names(Actives), value = TRUE, ignore.case = TRUE)[1]
  stopifnot("Could not find any Year‑like column in Actives" = length(yr_alt) == 1)
  setnames(Actives, yr_alt, "Year")   # strip weird spaces, unify name
}



## ================================================================
##  PHASE 2 :  PFV  &  PV‑BENEFITS   (bullet‑proof version)
## ================================================================

library(data.table)
library(lubridate)
library(writexl)

## ---- parameters -------------------------------------------------------
benefit_end_age <- if (exists("defaults") &&
                       !is.null(defaults$max_benefit_age)) {
  defaults$max_benefit_age    # e.g. 65
} else 65                     # fallback
val_year <- year(valuation_date)

##  mid‑year payment‑timing discount applied to PFV pieces
pay_timing_disc <- (1 + defaults$i_rate)^(-0.5)   # ≈0.982235

## ---- 0 · prune EEs that lack a valuation‑date row ---------------------
val_EEs    <- Actives[Year == val_year, unique(EE_NO)]
missing_EE <- setdiff(Actives[, unique(EE_NO)], val_EEs)

if (length(missing_EE)) {
  cat("⚠  Skipping members with no val‑date row:\n",
      Actives[EE_NO %in% missing_EE,
              unique(paste(EE_NO, Name, sep = " – "))],
      sep = "\n")
  Actives <- Actives[!EE_NO %in% missing_EE]
}





## ---- 1 · eligibility anchors -----------------------------------------
elig_anchors <- Actives[
  Eligibility == 1 &
    EE_Age < benefit_end_age &
    qx_ret  < 1,
  .(EE_NO, Name,
    anchor_year = Year,
    anchor_age  = EE_Age)
]





## ---- 2 · build pay‑stream exploded rows with PFV payment stream --------------
setkey(elig_anchors, EE_NO)      # x‑table key
setkey(Actives,       EE_NO)     # i‑table key

# 1) plain equi‑join on EE_NO
pay_stream <- Actives[elig_anchors,
                      allow.cartesian = TRUE,
                      nomatch = 0L]

# 2) keep only rows meeting the timing rules
pay_stream <- pay_stream[
  Year      >= anchor_year &      # future (or anchor) years only
    EE_Age    <  benefit_end_age    # cap PFV window (<65)
]







## ---- 3 · pull anchor constants & merge --------------------------------
# Create a lookup table with the specific PFV decrement values from each anchor year.
anchor_info <- Actives[
  elig_anchors,
  on = .(EE_NO, Year = anchor_year),
  nomatch = 0L,
  .(EE_NO, anchor_year,
    disc_anchor_srv  = fifelse(is.na(disc_since_Service), 0, disc_since_Service),
    surv_pure_anchor = cum_survival_pure, #gets used as anchor for active PFV pieces
    surv_full_anchor = cum_survival, 
    surv_pure_anchor_SP = cum_survival_pure_SP, #gets used as anchor for spouse PFV pieces
    surv_joint_anchor = cum_survival * cum_survival_pure_SP, ## <-- Joint probability spouse alive and active retired&alive
    qx_ret_anchor    = qx_ret,
    disc_anchor_val  = fifelse(is.na(disc_since_val), 0, disc_since_val)) #replace NA with 0 so pre‑valdate  anchors zero‑out
]


# Merge these constant anchor values back into the exploded payment stream.
pay_stream <- merge(
  pay_stream,                       # The main table of future payment rows
  anchor_info,                      # The lookup table of anchor-specific constants
  by = c("EE_NO", "anchor_year"),   # Join on the unique anchor identifier
  all.x = TRUE,                     # Keep all rows from pay_stream, even if no match is found
  sort = FALSE                      # Preserve original order for efficiency
)





## ---- 4 · drop rows missing anchor data (safety + console log) ---------
bad_rows <- pay_stream[
  is.na(disc_anchor_srv) | is.na(surv_pure_anchor),
  .(EE_NO, Name)
][unique(.SD), on = .(EE_NO, Name)]          # unique pairs only

if (nrow(bad_rows)) {
  cat("⚠  Dropping", nrow(bad_rows), "rows with missing anchor data for:\n")
  bad_rows[, cat(sprintf(" • %s – %s\n", EE_NO, Name))]
}

pay_stream <- pay_stream[
  !is.na(disc_anchor_srv) & !is.na(surv_pure_anchor)
]





## ---- 5 · PFV piece‑level factors (PFV pieces sum to the PFV lumpsum at an anchor age) -------------------------------------
pay_stream[, `:=`(
  int_factor       = disc_since_Service / disc_anchor_srv,
  surv_pure_factor = fifelse(surv_pure_anchor == 0, 0,
                             cum_survival_pure / surv_pure_anchor),
  ## ---- SPOUSE NEW ----
  surv_pure_factor_SP = fifelse(surv_pure_anchor_SP == 0, 0,
                                cum_survival_pure_SP / surv_pure_anchor_SP)
)]

pay_stream[, pfv_piece       := Premium_EE_adj * int_factor *
             surv_pure_factor       * pay_timing_disc]
pay_stream[, pfv_piece_net   := `netPRE-EE`    * int_factor *
             surv_pure_factor       * pay_timing_disc]

## ---- SPOUSE NEW ----
pay_stream[, pfv_piece_SP     := Premium_SP_adj * int_factor *
             surv_pure_factor_SP   * pay_timing_disc]
pay_stream[, pfv_piece_net_SP := `netPRE-SP`    * int_factor *
             surv_pure_factor_SP   * pay_timing_disc]





## ---- 6 · collapse to PFV lump sums at anchor -> 1 row per anchor per member --
pfv_anchor <- pay_stream[, .(
  PFV_lumpsum       = sum(pfv_piece, na.rm = TRUE),     # The  gross pfv lumpsum
  PFV_lumpsum_net   = sum(pfv_piece_net, na.rm = TRUE), # The new net pfv lumpsum
  disc_anchor_srv   = first(disc_anchor_srv),           
  disc_anchor_val   = first(disc_anchor_val),                 ##decrements to get pfv_lumpsums to valuation date or service start date 
  qx_ret_anchor     = first(qx_ret_anchor),
  surv_full_anchor  = first(surv_full_anchor),
  # ---- spouse duplicates ----
  PFV_lumpsum_SP      = sum(pfv_piece_SP,     na.rm = TRUE),
  PFV_lumpsum_net_SP  = sum(pfv_piece_net_SP, na.rm = TRUE),
  surv_joint_anchor = first(surv_joint_anchor) 
), by = .(EE_NO, Name, anchor_year, anchor_age)]




## ---- 7 · bring PFV back to valuation date -----------------------------
# Create a lookup table for survival‑probability-of‑member (EE) **and** spouse
surv_lookup <- Actives[Year == val_year,
                       .(EE_NO,
                         cum_survival_val     = cum_survival,          # EE survival at valuation date
                         cum_surv_pure_SP_val = cum_survival_pure_SP   ## <-- SPOUSE NEW
                       )
]

# Merge valuation‑date survival probabilities into the anchor table.
pfv_anchor <- merge(
  pfv_anchor,
  surv_lookup,
  by = "EE_NO",      # Join by employee ID
  all.x = TRUE,      # Keep all anchor rows
  sort = FALSE
)[!is.na(cum_survival_val)]   # Safety check: drop anchors with no val‑date row.

# Add fraction‑married flag
pfv_anchor <- merge(                                   ## <-- SPOUSE NEW
  pfv_anchor,
  Actives[, .(EE_NO, frac_married)],                   # pre‑calculated earlier
  by = "EE_NO",
  all.x = TRUE,
  sort = FALSE
)

# Pre‑compute the joint EE‑and‑spouse survival at valuation date
pfv_anchor[, surv_joint_val :=
             cum_survival_val * cum_surv_pure_SP_val]  ## <-- SPOUSE NEW


#  Calculate PV Benefits
# ---------- MEMBER ----------
pfv_anchor[, PV_Benefits :=
             PFV_lumpsum * disc_anchor_val *
             (surv_full_anchor / cum_survival_val) *
             qx_ret_anchor * defaults$elect_prob]

pfv_anchor[, PV_Benefits_net :=
             PFV_lumpsum_net * disc_anchor_val *
             (surv_full_anchor / cum_survival_val) *
             qx_ret_anchor * defaults$elect_prob]

# funding‑age present values (member)
pfv_anchor[, PV_Benefits_funding :=
             PFV_lumpsum * disc_anchor_srv * surv_full_anchor *
             qx_ret_anchor * defaults$elect_prob]

pfv_anchor[, PV_Benefits_funding_net :=
             PFV_lumpsum_net * disc_anchor_srv * surv_full_anchor *
             qx_ret_anchor * defaults$elect_prob]

frac_married <- .3
# ---------- SPOUSE   -----------------------------------------  ## 
pfv_anchor[, PV_Benefits_SP :=
             PFV_lumpsum_SP * disc_anchor_val *
             (surv_joint_anchor / surv_joint_val) *
             qx_ret_anchor * defaults$elect_prob *
             frac_married]

pfv_anchor[, PV_Benefits_net_SP :=
             PFV_lumpsum_net_SP * disc_anchor_val *
             (surv_joint_anchor / surv_joint_val) *
             qx_ret_anchor * defaults$elect_prob *
             frac_married]

# funding‑age present values (spouse)                                      
pfv_anchor[, PV_Benefits_funding_SP :=
             PFV_lumpsum_SP * disc_anchor_srv * surv_joint_anchor *
             qx_ret_anchor * defaults$elect_prob *
             frac_married]

pfv_anchor[, PV_Benefits_funding_net_SP :=
             PFV_lumpsum_net_SP * disc_anchor_srv * surv_joint_anchor *
             qx_ret_anchor * defaults$elect_prob *
             frac_married]







## ---- 8 · salary PV benefits  ------------------------------------------
pv_salary_tbl_val <- Actives[, .(
  PV_salary_benefits_val = sum(PVSalary_fromVal,  na.rm = TRUE)
), by = EE_NO]

pv_salary_tbl_funding <- Actives[, .(
  PV_salary_benefits_funding = sum(PVSalary_fromHire, na.rm = TRUE)
), by = EE_NO]




###END EAN CALCS###




## ---- 9 · assemble report tabs -----------------------------------------
PFV_PV_Benefits <- merge(
  pfv_anchor[, .(EE_NO, Name,
                 anchor_age, anchor_year,
                 
                 ## --- interest & election factors (new cols) ---
                 disc_anchor_val,
                 disc_anchor_srv,
                 elect_prob = defaults$elect_prob,
                 
                 ## --- present values ---
                 PV_Benefits,
                 PV_Benefits_net,           
                 PV_Benefits_funding,
                 PV_Benefits_funding_net,   
                 PFV_lumpsum,
                 PFV_lumpsum_net,           
                  #spouses
                 PV_Benefits_SP, PV_Benefits_net_SP,
                 PV_Benefits_funding_SP, PV_Benefits_funding_net_SP,
                 PFV_lumpsum_SP, PFV_lumpsum_net_SP,
                 
                 ## --- decrement diagnostics ---
                 qx_ret_anchor,
                 surv_full_anchor,
                 cum_survival_val)],

  merge(pv_salary_tbl_val,                     # salary PVs (val‑date)
        pv_salary_tbl_funding,                 # salary PVs (fund‑age)
        by = "EE_NO"),
  
  by   = "EE_NO",
  all.x = TRUE,
  sort  = FALSE
)[order(Name, anchor_year)]


PFV_Under_Hood <- pay_stream[, .(
  EE_NO, Name,
  anchor_age, anchor_year,
  pay_year = Year,
  pay_age  = EE_Age,
  Premium_EE_adj,
  `netPRE-EE`,        # Add the source net premium column
  int_factor,
  surv_pure_factor,
  pfv_piece,
  pfv_piece_net,      # Add the resulting net piece column
  Premium_SP_adj, `netPRE-SP`,  #spouse stuff
  pfv_piece_SP, pfv_piece_net_SP
)]



## ---- 10 · write Excel -------------------------------------------------
out_path <- file.path(
  plan_spec$folder_path,
  "R code and output",
  "PFV_PV_Report_final.xlsx"
)


writexl::write_xlsx(
  list("PFV_PV_Benefits" = PFV_PV_Benefits,
       "PFV_Under_Hood" = PFV_Under_Hood),
  path = out_path
)

# 
# 
# 
# 
# — export census
# writexl::write_xlsx(Actives, "C:/Users/fleis/Documents/Rafi 2025/R code and output/Actives_v25.xlsx")








