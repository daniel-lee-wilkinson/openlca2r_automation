### Product Systems - Direct inventory contributions sheet ###

library(readxl)
library(cellranger)
library(dplyr)
library(tidyr)

# Define file details
wd <- getwd()
excel_file_name <- "3__Cooling.xlsx"  # Adjust if needed
file_path <- file.path(wd, excel_file_name)
sheet_name <- "Direct inventory contributions"

### 1. Read the Emissions Block (Columns B:F)
# Header is in row 5 (B5:F5) and data starts at row 6
emissions_df <- read_excel(file_path, 
                           sheet = sheet_name,
                           range = cell_limits(ul = c(5, 2), lr = c(NA, 6)),
                           col_names = TRUE)

### 2. Read and Process the System Inputs Header (Columns G onward, Rows 2–4)
system_header_df <- read_excel(file_path, 
                               sheet = sheet_name, 
                               range = cell_limits(ul = c(2, 7), lr = c(4, NA)),
                               col_names = FALSE)

# Create composite header: force characters, replace logical/NA with empty string, and pad to 3 components if needed
system_header <- apply(system_header_df, 2, function(x) {
  x <- sapply(x, function(y) {
    if (is.logical(y) || is.na(y)) "" else as.character(y)
  })
  if (length(x) < 3) x <- c(x, rep("", 3 - length(x)))
  paste(x, collapse = "|")
})
print(system_header)

### 3. Read the System Inputs Data (Columns G onward, from Row 6)
system_data_df <- read_excel(file_path, 
                             sheet = sheet_name, 
                             range = cell_limits(ul = c(6, 7), lr = c(NA, NA)),
                             col_names = FALSE)
colnames(system_data_df) <- system_header

### 4. Combine Emissions and System Inputs Data
combined_df <- cbind(emissions_df, system_data_df)

### 5. Pivot System Inputs from Wide to Long Format
# Assume emissions_df has 5 columns; pivot all system input columns
long_df <- combined_df %>%
  pivot_longer(
    cols = (ncol(emissions_df) + 1):ncol(combined_df),
    names_to = "composite_header",
    values_to = "system_input_value",
    values_drop_na = FALSE
  )

### 6. Separate the Composite Header into Components
long_df <- long_df %>%
  separate(composite_header, into = c("process_uuid", "process", "location"), 
           sep = "\\|", fill = "right")

### 7. (Optional) Remove Rows Containing Repeated Header Text
long_df <- long_df %>%
  filter(!(process_uuid %in% c("Process UUID", NA, "")))

### 8. Check the Final Cleaned Long DataFrame
print("Cleaned Long DataFrame:")
print(head(long_df))
View(long_df)

### Data Integrity Checks
expected_cells <- nrow(system_data_df) * ncol(system_data_df)
actual_cells <- nrow(long_df)
cat("Expected cells:", expected_cells, "\n")
cat("Actual cells in long_df:", actual_cells, "\n")

non_na_wide <- sum(!is.na(as.matrix(system_data_df)))
non_na_long <- sum(!is.na(long_df$system_input_value))
cat("Non-NA cells in wide data:", non_na_wide, "\n")
cat("Non-NA cells in long data:", non_na_long, "\n")

# Examine unique values in the header components
cat("Unique Process UUIDs:\n")
print(unique(long_df$process_uuid))
cat("\nUnique Process Names:\n")
print(unique(long_df$process))
cat("\nUnique Locations:\n")
print(unique(long_df$location))


# convert process input value to numerical
long_df <- long_df %>%
  mutate(across(contains("value"), as.numeric))
