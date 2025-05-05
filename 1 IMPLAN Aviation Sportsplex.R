### Aviation Sports Complex EIS Impacts - April 2025 ###

rm(list = ls()) 

# SET UP  -------------------------------------------------------
options(scipen = 999)  # no scientific notation

library(kableExtra)
library(dplyr)
library(tidyverse)
library(readxl)
library(knitr)
library(openxlsx)

## set working directory to desired file location

setwd("C:/Users/jyan/OneDrive - Econsult Solutions Inc/Aviation Sports Complex LLC/Analysis/Analysis/IMPLAN")

# Load References ---------------------------------------------------------------

output <-"output.csv"

## Reference Tables 
{
## NEW ##
conversion = read_xlsx("Emp_FTE_and_W&S_EC 528 Industries.xlsx", sheet = "2023") %>% 
  select(Implan528Index, ECtoWSInc, FTEperTotalEmp) %>% 
  rename(IndustryCode = Implan528Index)

industry_NAICS = read_xlsx("Results Aggregator NAICS Schemes 528.xlsx", sheet = "IMPLAN to NAICS")

industry_NAICS = industry_NAICS %>% 
  select(Implan528Index, `NAICS 2 Digit`) %>% 
  rename(NAICS_2 = `NAICS 2 Digit`,
         IndustryCode = Implan528Index)

}

# Load Functions ----------------------------------------------------------

{
## Helper Function: Generate Summary Table
generate_summary_table <- function(output, geography, geography_list, metric, group_vars) {
  df_list <- list()
  for (i in seq_along(geography_list)) {
    df <- output %>%
      filter(DestinationRegion %in% geography[[i]]) %>%
      group_by(across(all_of(group_vars))) %>%
      summarise(Value = sum(.data[[metric]], na.rm = TRUE))
    
    df[[geography_list[[i]]]] <- df$Value
    df <- df %>% select(-Value)
    df_list[[i]] <- df
  }
  
  Reduce(function(x, y) left_join(x, y, by = group_vars), df_list)
}
## Function to Create Output Tables
create_output_tables <- function(output, geography, geography_list) {
  impact_table <- generate_summary_table(output, geography, geography_list, "Output", c("EventGroup", "ImpactType"))
  fte_table <- generate_summary_table(output, geography, geography_list, "FTE", "EventGroup")
  employee_compensation_table <- generate_summary_table(output, geography, geography_list, "EmployeeCompensation", "EventGroup")
  total_wage_table <- generate_summary_table(output, geography, geography_list, "TotalWage", "EventGroup")
  value_added_table <- generate_summary_table(output, geography, geography_list, "ValueAdded", "EventGroup")
  
  ## more granular 
  fte_impacts <- generate_summary_table(output, geography, geography_list, "FTE", c("EventGroup", "ImpactType"))
  employee_compensation_impacts <- generate_summary_table(output, geography, geography_list, "EmployeeCompensation", c("EventGroup", "ImpactType"))
  total_wage_impacts <- generate_summary_table(output, geography, geography_list, "TotalWage", c("EventGroup", "ImpactType"))
  value_added_impacts <- generate_summary_table(output, geography, geography_list, "ValueAdded", c("EventGroup", "ImpactType"))
  employment_impacts <- generate_summary_table(output, geography, geography_list, "Employment", c("EventGroup", "ImpactType"))
  
  list(
    Impact = impact_table,
    FTE = fte_table,
    EmployeeCompensation = employee_compensation_table,
    TotalWage = total_wage_table,
    ValueAdded = value_added_table,
    ## additional details
    FTEimpacts = fte_impacts,
    EmpCompImpacts = employee_compensation_impacts,
    TotalWageImpacts = total_wage_impacts,
    ValueAddImpacts = value_added_impacts,
    EmploymentImpacts = employment_impacts
  )
}

## Function to Export Tables to Excel
export_to_excel <- function(tables, filename) {
  wb <- createWorkbook()
  
  for (name in names(tables)) {
    addWorksheet(wb, name)
    writeData(wb, name, tables[[name]])
  }
  
  saveWorkbook(wb, filename, overwrite = TRUE)
}

# Helper function to truncate sheet names if they exceed 31 characters (updated!)
shorten_sheet_name <- function(sheet_name, existing_names, suffix = "") {
  # Calculate the maximum allowable length after adding the suffix
  max_length <- 31 - nchar(suffix)
  
  # Truncate the sheet name if needed and append the suffix
  truncated_name <- substr(sheet_name, 1, max_length)
  unique_name <- paste0(truncated_name, suffix)
  
  # Ensure uniqueness by appending a counter if it already exists
  counter <- 1
  while (unique_name %in% existing_names) {
    counter <- counter + 1
    unique_name <- paste0(substr(truncated_name, 1, max_length - nchar(counter) - 1), "_", counter)
  }
  
  return(unique_name)
}

# Function to create both NAICS and Impact summaries and add them to workbook
process_event_group <- function(event_group_name) {
  
  # NAICS summary
  naics_summary <- output %>%
    filter(trimws(EventGroup) == event_group_name,
           ImpactType %in% c("Indirect", "Induced"),
           DestinationRegion %in% c("Cape May County, NJ (2023)")) %>% ## update as needed
    group_by(NAICS_2) %>%
    summarise(FTE = sum(FTE, na.rm = TRUE))
  
  # Impact summary
  impact_summary <- output %>%
    filter(trimws(EventGroup) == event_group_name,
           DestinationRegion %in% c("Cape May County, NJ (2023)")) %>% ## update as needed
    group_by(ImpactType) %>%
    summarise(FTE = sum(FTE, na.rm = TRUE))
  
  # Shorten sheet names if needed
  naics_sheet_name <- shorten_sheet_name(event_group_name, names(wb), suffix = "_NAICS")
  impact_sheet_name <- shorten_sheet_name(event_group_name, names(wb), suffix = "_Impact")
  
  # Add both summaries to the workbook
  addWorksheet(wb, naics_sheet_name)
  writeData(wb, sheet = naics_sheet_name, naics_summary)
  
  addWorksheet(wb, impact_sheet_name)
  writeData(wb, sheet = impact_sheet_name, impact_summary)
}

# Function to create both NAICS and Impact summaries and add them to workbook
state_process_event_group <- function(event_group_name) {
  
  # NAICS summary
  naics_summary <- output %>%
    filter(trimws(EventGroup) == event_group_name,
           ImpactType %in% c("Indirect", "Induced")) %>%
    group_by(NAICS_2) %>%
    summarise(FTE = sum(FTE, na.rm = TRUE))
  
  # Impact summary
  impact_summary <- output %>%
    filter(trimws(EventGroup) == event_group_name) %>%
    group_by(ImpactType) %>%
    summarise(FTE = sum(FTE, na.rm = TRUE))
  
  # Shorten sheet names if needed (updated!)
  naics_sheet_name <- shorten_sheet_name(event_group_name, names(wb1), suffix = "_NAICS")
  impact_sheet_name <- shorten_sheet_name(event_group_name, names(wb1), suffix = "_Impact")
  
  
  # Add both summaries to the workbook
  addWorksheet(wb1, naics_sheet_name)
  writeData(wb1, sheet = naics_sheet_name, naics_summary)
  
  addWorksheet(wb1, impact_sheet_name)
  writeData(wb1, sheet = impact_sheet_name, impact_summary)
}

}


# Process Data ------------------------------------------------------------

raw_output <- read.csv(output)

output <- left_join(raw_output, conversion, by = "IndustryCode") %>%
  mutate(
    TotalWage = EmployeeCompensation / ECtoWSInc,
    FTE = Employment * FTEperTotalEmp,
    ValueAdded = EmployeeCompensation + ProprietorIncome + TaxesOnProductionAndImports + OtherPropertyIncome,
    EventGroup = sub("-.*", "", EventName)
  )

## MUST UPDATE ##
geography_list <- c("Cape May", "New Jersey")

### update geography for each analysis; copy exactly based on DestinationRegion name
unique(output$DestinationRegion) 

# [1] "Cape May County, NJ (2023)" "NJ minus Cape May (2023)" 

geography <- list(
  c("Cape May County, NJ (2023)"), 
  c("Cape May County, NJ (2023)", "NJ minus Cape May (2023)")
)

output_tables <- create_output_tables(output, geography, geography_list)

# EXPORT 
export_to_excel(output_tables, "Processed/1 IMPLAN_Economic_Impacts_Output.xlsx")

## Industry Employment -----------------------------------------------------

output = left_join(raw_output, conversion, by = "IndustryCode") %>% 
  mutate(TotalWage = EmployeeCompensation / ECtoWSInc,
         FTE = Employment * FTEperTotalEmp,
         ValueAdded = EmployeeCompensation + ProprietorIncome + TaxesOnProductionAndImports + OtherPropertyIncome) %>% 
  mutate(EventGroup = sub("-.*", "", EventName))

output = left_join(output, industry_NAICS, by = "IndustryCode" )

unique(output$EventGroup) # check for updating captions below

# [1] "CapEx "      "Employees "     "Operations"    "Spend "

## city-wide employment ##

## Initialize workbook
wb <- createWorkbook()

# Define the EventGroups and corresponding captions
event_groups <- trimws(c("CapEx ", "Employees ", "Operations ", "Spend "
                         ))

# Loop through event groups and process them
for (event_group in event_groups) {
  process_event_group(event_group)
}

# Save workbook to Excel
saveWorkbook(wb, "Processed/City_Wide_Employment_Summary.xlsx", overwrite = TRUE)



## State-Wide Employment ##

## Initialize workbook
wb1 <- createWorkbook()

# Define the EventGroups and corresponding captions
#event_groups <- trimws(c("CapEx ", "Employees ", "Operations ", "Spend "
#                         ))

# Loop through event groups and process them
for (event_group in event_groups) {
  state_process_event_group(event_group)
}

# Save workbook to Excel
saveWorkbook(wb1, "Processed/Statewide_Employment_Summary.xlsx", overwrite = TRUE)


