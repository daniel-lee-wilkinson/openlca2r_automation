library(readxl)
library(dplyr)
library(tidyr)

# Define file details
wd <- getwd()
excel_file_name <- "project result.xlsx"
file_path <- file.path(wd, excel_file_name)
sheet_name <- "Info"

# Read the entire "Info" sheet without predefined column names.
# (This file appears to have 6 columns; they will be named ...1, ...2, etc.)
info_raw <- read_excel(file_path, sheet = sheet_name, col_names = FALSE)
print(info_raw)

### 1. Extract General Info Section ###
# Based on your description and the sample:
# - Keys (e.g. "Name:", "Description:", "LCIA Method:") are in column A (raw ...1) in rows 2–4.
# - Corresponding values are in column B (raw ...2) in rows 2–4.
general_info <- tibble(
  label = info_raw[[1]][2:4] %>% as.character(),
  value = info_raw[[2]][2:4] %>% as.character()
)
cat("General Info:\n")
print(general_info)

### 2. Locate Section Markers ###
# Find the rows in column A (raw ...1) where "Variants" and "Parameters" appear.
variants_row <- which(info_raw[[1]] == "Variants")
parameters_row <- which(info_raw[[1]] == "Parameters")

if (length(variants_row) == 0) stop("No 'Variants' marker found.")
if (length(parameters_row) == 0) stop("No 'Parameters' marker found.")

### 3. Extract Variants Section ###
# The row immediately after the "Variants" marker (row 7 if Variants is in row 6) is the header.
variants_header <- info_raw[variants_row + 1, 1:6] %>% unlist() %>% as.character()

# Data for variants starts on the next row and continues until two rows before the "Parameters" marker.
# (Because the row immediately above "Parameters" is empty.)
variants_data <- info_raw[(variants_row + 2):(parameters_row - 2), 1:6]
variants_df <- as_tibble(variants_data)
colnames(variants_df) <- variants_header

cat("Variants:\n")
print(variants_df)

### 4. Extract Parameters Section ###
# Use the "Parameters" marker row as the header.
parameters_header <- info_raw[parameters_row, 1:6] %>% unlist() %>% as.character()

# Data starts on the next row (immediately below the empty row above, if any) and continues to the end.
parameters_data <- info_raw[(parameters_row + 1):nrow(info_raw), 1:6]
parameters_df <- as_tibble(parameters_data)
colnames(parameters_df) <- parameters_header

# Remove any columns that are entirely NA or empty strings using base R.
parameters_df <- parameters_df[, sapply(parameters_df, function(x) {
  ! (all(is.na(x)) || all(x == ""))
})]

cat("Parameters:\n")
print(parameters_df)

### 5. Combine the Cleaned Sections into a List (Optional)
info_clean <- list(
  general_info = general_info,
  variants = variants_df,
  parameters = parameters_df
)

# Preview the cleaned info
info_clean
