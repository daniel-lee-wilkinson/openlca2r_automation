library(readxl)
library(dplyr)
library(tidyr)

## Define file details
wd <- getwd()
excel_file_name <- "project result.xlsx"
file_path <- file.path(wd, excel_file_name)
sheet_name <- "LCI Results"

## Read the entire sheet without predefined column names.
raw_data <- read_excel(file_path, sheet = sheet_name, col_names = FALSE)
print(raw_data)

### Locate the "Outputs" Marker ###
# "Outputs" is in column A (raw_data[[1]]). We trim whitespace and use a case-insensitive search.
outputs_idx <- grep("output", tolower(trimws(raw_data[[1]])))
if(length(outputs_idx) == 0) stop("No 'Outputs' marker found in column A.")
outputs_row <- outputs_idx[1]
cat("Found 'Outputs' marker at row:", outputs_row, "\n")

### Determine Total Columns ###
n_cols <- ncol(raw_data)

### Extract Inputs Block ###
# For Inputs:
# - Columns A–F header comes from row 4.
# - Columns G onward header comes from row 3.
input_header_main <- raw_data[3, 1:5] %>% unlist() %>% as.character()
scenario_header_inputs <- raw_data[2, 6:n_cols] %>% unlist() %>% as.character()
input_header <- c(input_header_main, scenario_header_inputs)
cat("Inputs Header:\n")
print(input_header)

# Extract input data from row 5 up to (but not including) the Outputs marker, for columns A through the last column.
input_data <- raw_data[4:(outputs_row - 1), 1:n_cols]
inputs_df <- as_tibble(input_data)
colnames(inputs_df) <- input_header

### Extract Outputs Block ###
# For Outputs:
# - The scenario header (columns G onward) is in the Outputs marker row (row outputs_row).
scenario_header_outputs <- raw_data[outputs_row, 6:n_cols] %>% unlist() %>% as.character()
# - The main outputs header for columns A–F is in the row immediately after the Outputs marker.
output_header_main <- raw_data[outputs_row + 1, 1:5] %>% unlist() %>% as.character()
output_header <- c(output_header_main, scenario_header_outputs)

cat("Outputs Header:\n")
print(output_header)

# Extract output data: from the row after the outputs header to the end, for columns A onward.
output_data <- raw_data[(outputs_row + 2):nrow(raw_data), 1:n_cols]
outputs_df <- as_tibble(output_data)
colnames(outputs_df) <- output_header

### Data Integrity Checks (Optional) ###
expected_input_cells <- nrow(input_data) * ncol(input_data)
actual_input_cells   <- nrow(inputs_df) * ncol(inputs_df)
non_na_input_raw     <- sum(!is.na(as.matrix(input_data)))
non_na_input_df      <- sum(!is.na(as.matrix(inputs_df)))
cat("Inputs Block:\n")
cat("  Expected cells:", expected_input_cells, "\n")
cat("  Actual cells:", actual_input_cells, "\n")
cat("  Non-NA count (raw):", non_na_input_raw, "\n")
cat("  Non-NA count (df):", non_na_input_df, "\n\n")

expected_output_cells <- nrow(output_data) * ncol(output_data)
actual_output_cells   <- nrow(outputs_df) * ncol(outputs_df)
non_na_output_raw     <- sum(!is.na(as.matrix(output_data)))
non_na_output_df      <- sum(!is.na(as.matrix(outputs_df)))
cat("Outputs Block:\n")
cat("  Expected cells:", expected_output_cells, "\n")
cat("  Actual cells:", actual_output_cells, "\n")
cat("  Non-NA count (raw):", non_na_output_raw, "\n")
cat("  Non-NA count (df):", non_na_output_df, "\n\n")

### Preview the Results ###
cat("Inputs Data:\n")
print(head(inputs_df))
cat("\nOutputs Data:\n")
print(head(outputs_df))

# make all scenario values numerical
inputs_df <- inputs_df %>%
  mutate(across(6:ncol(inputs_df), as.numeric))

outputs_df <- outputs_df %>%
  mutate(across(6:ncol(outputs_df), as.numeric))
