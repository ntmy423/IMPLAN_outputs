### Innovation Space EIS  - December 2025 ###

rm(list = ls()) ## clear environments


# SET UP  -------------------------------------------------------
options(scipen = 999)  # no scientific notation

# Load Libraries
{
library(here) ## must have any other rscripts closed before loading!
library(tidyverse)
library(readxl)
library(openxlsx)
}

# Load Reference Data

raw_output <- read.csv(here("output-DE.csv")) ## DELAWARE IMPACTS

# raw_output <- read.csv(here("output-MSA.csv")) ## PHILA MSA IMPACTS

## Reference Files --> must have copies stored in working directory location; load just one!

## IMPLAN 528 (2023- )
{
  conversion <- read_xlsx(
    here("Emp_FTE_and_W&S_EC_528_Industries.xlsx"), sheet = "2023") %>%
    select(Implan528Index, ECtoWSInc, FTEperTotalEmp) %>%
    rename(IndustryCode = Implan528Index)

  industry_NAICS <- read_xlsx( 
    here("Results Aggregator NAICS Schemes 528.xlsx"), sheet = "IMPLAN to NAICS") %>%
    select(Implan528Index, `NAICS 2 Digit`) %>%
    rename(NAICS_2 = `NAICS 2 Digit`, IndustryCode = Implan528Index)
}




# UPDATE: DEFINITIONS -----------------------------------------------------

# UPDATE: Define geography column header titles; should match report names for each geography
geography_list <- c("New Castle County", "Delaware", "United States")

# geography_list <- c("New Castle County", "MSA", "United States") 


unique(raw_output$DestinationRegion) ## Delaware
# [1] "Delaware minus New Castle Co (2024)" "New Castle County, DE (2024)"       
# [3] "US minus DE (2024)"       

## UPDATE: Match exactly to DestinationRegion ## Delaware
direct_geography <- "New Castle County, DE (2024)"
indirect_geography1 <- "Delaware minus New Castle Co (2024)"
indirect_geography2 <- "US minus DE (2024)"



# unique(raw_output$DestinationRegion) ## MSA
# [1] "New Castle County, DE (2024)"    "Phila-Cam-Wilm MSA minus New Castle Co DE (2024)"
# [3] "US minus Phila-Cam-Wilm MSA (2024)"    

## UPDATE: Match exactly to DestinationRegion ## MSA
# direct_geography <- "New Castle County, DE (2024)"
# indirect_geography1 <- "Phila-Cam-Wilm MSA minus New Castle Co DE (2024)"
# indirect_geography2 <- "US minus Phila-Cam-Wilm MSA (2024)"


geography <- list(
  c(direct_geography), 
  c(direct_geography, indirect_geography1) ,
  c(direct_geography, indirect_geography1, indirect_geography2)
)

## event grouping 
unique((raw_output %>% mutate(EventName = sub("-.*", "", EventName)))$EventName) # replace below with output from console
#  "CapEx " "Ops "   

# UPDATE: Define the EventGroups and corresponding captions from above; copy and paste exactly
event_groups <- trimws(c("CapEx ", "Ops "
                         ))

# Functions ----------------------------------------------------
{
## Generate impact summary tables
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

## Create output tables
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
    ## all outputs by impact, typically needed for adjusting fiscal impacts
    FTEimpacts = fte_impacts,
    EmpCompImpacts = employee_compensation_impacts,
    TotalWageImpacts = total_wage_impacts,
    ValueAddImpacts = value_added_impacts,
    EmploymentImpacts = employment_impacts
  )
}

## Export outputs
export_to_excel <- function(tables, filename) {
  wb <- createWorkbook()
  for (name in names(tables)) {
    addWorksheet(wb, name)
    writeData(wb, name, tables[[name]])
  }
  saveWorkbook(wb, filename, overwrite = TRUE)
}

## Shorten sheet names
shorten_sheet_name <- function(sheet_name, existing_names, suffix = "") {
  max_length <- 31 - nchar(suffix)
  truncated_name <- substr(sheet_name, 1, max_length)
  unique_name <- paste0(truncated_name, suffix)
  
  counter <- 1
  while (unique_name %in% existing_names) {
    counter <- counter + 1
    unique_name <- paste0(substr(truncated_name, 1, max_length - nchar(counter) - 1), "_", counter)
  }
  return(unique_name)
}


# INDUSTRY EMPLOYMENT PIE CHARTS - DIRECT VS. INDIRECT/INDUCED
## Process industry employment impacts to smallest geography (typically city-wide)
process_subgeog_industry_employment <- function(event_group_name) {
  
  # NAICS summary for smallest (direct) geography
  naics_summary <- output %>%
    filter(trimws(EventGroup) == event_group_name,
           ImpactType %in% c("Indirect", "Induced"),
           DestinationRegion %in% direct_geography) %>% # update if different sub-geography level is wanted
    group_by(NAICS_2) %>%
    summarise(FTE = sum(FTE, na.rm = TRUE))
  
  # Impact summary for smallest (direct) geography
  impact_summary <- output %>%
    filter(trimws(EventGroup) == event_group_name,
           DestinationRegion %in% direct_geography) %>% # update if different sub-geography level is wanted
    group_by(ImpactType) %>%
    summarise(FTE = sum(FTE, na.rm = TRUE))
  
  # Shorten sheet names
  naics_sheet_name <- shorten_sheet_name(event_group_name, names(wb_direct), suffix = "_NAICS")
  impact_sheet_name <- shorten_sheet_name(event_group_name, names(wb_direct), suffix = "_Impact")
  
  # Add summaries to direct geography workbook
  addWorksheet(wb_direct, naics_sheet_name)
  writeData(wb_direct, sheet = naics_sheet_name, naics_summary)
  
  addWorksheet(wb_direct, impact_sheet_name)
  writeData(wb_direct, sheet = impact_sheet_name, impact_summary)
}

## Process industry employment impacts to largest geography (typically state-wide)
process_industry_employment <- function(event_group_name) {
  
  # NAICS summary for largest geography
  naics_summary <- output %>%
    filter(trimws(EventGroup) == event_group_name,
           ImpactType %in% c("Indirect", "Induced")) %>%
    group_by(NAICS_2) %>%
    summarise(FTE = sum(FTE, na.rm = TRUE))
  
  # Impact summary for largest geography
  impact_summary <- output %>%
    filter(trimws(EventGroup) == event_group_name) %>%  
    group_by(ImpactType) %>%
    summarise(FTE = sum(FTE, na.rm = TRUE))
  
  # Shorten sheet names
  naics_sheet_name <- shorten_sheet_name(event_group_name, names(wb_indirect), suffix = "_NAICS")
  impact_sheet_name <- shorten_sheet_name(event_group_name, names(wb_indirect), suffix = "_Impact")
  
  # Add summaries to state-wide workbook
  addWorksheet(wb_indirect, naics_sheet_name)
  writeData(wb_indirect, sheet = naics_sheet_name, naics_summary)
  
  addWorksheet(wb_indirect, impact_sheet_name)
  writeData(wb_indirect, sheet = impact_sheet_name, impact_summary)
}

}

# Data Processing ---------------------------------------------------------

## IMPLAN OUTPUTS - IMPACT, FTE, EMPLOYEE COMPENSATION, TOTAL WAGE, VALUE ADDED ##

output <- left_join(raw_output, conversion, by = "IndustryCode") %>%
  mutate(
    TotalWage = EmployeeCompensation / ECtoWSInc,
    FTE = Employment * FTEperTotalEmp,
    ValueAdded = EmployeeCompensation + ProprietorIncome + TaxesOnProductionAndImports + OtherPropertyIncome,
    EventGroup = sub("-.*", "", EventName)
  )

output <- left_join(output, industry_NAICS, by = "IndustryCode")

output_tables <- create_output_tables(output, geography, geography_list)


export_to_excel(output_tables,here("Processed","1 IMPLAN_Economic_Impacts_Output_DE.xlsx"))

# export_to_excel(output_tables,here("Processed","1 IMPLAN_Economic_Impacts_Output_MSA.xlsx"))



# Industry Employment - Sub-Geography -----------------------------------
## typically used for determining indirect & induced industry employment impacts at the city-level

wb_direct <- createWorkbook()

for (event_group in event_groups) {
  process_subgeog_industry_employment(event_group)
}

saveWorkbook(
  wb_direct,
  here("Processed", "SubGeography_Industry_Employment_Summary_DE.xlsx"),
  overwrite = TRUE
)


# saveWorkbook(
#   wb_direct,
#   here("Processed", "SubGeography_Industry_Employment_Summary_MSA.xlsx"),
#   overwrite = TRUE
# )


# Industry Employment - Largest Geography ---------------------------------------------------

wb_indirect <- createWorkbook()

for (event_group in event_groups) {
  process_industry_employment(event_group)
}

saveWorkbook(
  wb_direct,
  here("Processed", "Industry_Employment_Summary_DE.xlsx"),
  overwrite = TRUE
)


# saveWorkbook(
#   wb_direct,
#   here("Processed", "Industry_Employment_Summary_MSA.xlsx"),
#   overwrite = TRUE
# )





