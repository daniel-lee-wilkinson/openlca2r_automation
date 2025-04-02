# What this code does
## Process and clean the direct impact contribution sheet of OpenLCA 2 Product System export file

# Steps:
# 1. Define file details
# 2. Read Impacts sheet
# 3. Read Information sheet (headers)
# 4. Read Process data sheet
# 5. Ensure equal number of rows between datasets
# 6. Combine and reshape data (pivot_longer)
# 7. Separate composite header into columns
# 8. Perform data integrity checks

# Libraries
library(readxl)
library(cellranger)
library(tidyverse)


# File details
file_path <- file.path(getwd(), "3__Cooling.xlsx") # example Product System results file
sheet <- "Direct impact contributions"

# Read data
impact_df <- read_excel(file_path, sheet, range = cell_limits(c(5, 2), c(NA, 4)))

process_header <- read_excel(file_path, sheet, range = cell_limits(c(2, 5), c(4, NA)), col_names = FALSE) %>%
  apply(2, function(x) paste(replace_na(as.character(x), ""), collapse = "|"))

process_data_df <- read_excel(file_path, sheet, range = cell_limits(c(5, 5), c(NA, NA)), col_names = FALSE)
colnames(process_data_df) <- process_header

# Align rows
common_rows <- min(nrow(impact_df), nrow(process_data_df))
impact_df <- impact_df[1:common_rows, ]
process_data_df <- process_data_df[1:common_rows, ]

# Pivot and clean data
long_df <- bind_cols(impact_df, process_data_df) %>%
  pivot_longer(cols = -(1:ncol(impact_df)), names_to = "composite_header", values_to = "process_value") %>%
  separate(composite_header, into = c("process_uuid", "process", "location"), sep = "\\|", fill = "right") %>%
  filter(!process_uuid %in% c("Process UUID", NA, ""))

# Integrity check
cat("Expected cells:", nrow(process_data_df) * ncol(process_data_df), "\n")
cat("Actual cells:", nrow(long_df), "\n")
cat("Non-NA (wide):", sum(!is.na(process_data_df)), "\n")
cat("Non-NA (long):", sum(!is.na(long_df$process_value)), "\n")

# Preview result
print(head(long_df))
View(long_df)
