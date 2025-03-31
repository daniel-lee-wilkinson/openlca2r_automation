### Product System - Direct impact contributions sheet ###

library(readxl)
library(cellranger)
library(dplyr)
library(tidyr)

# Define file details
wd <- getwd()
excel_file_name <- "3__Cooling.xlsx"  # Adjust if needed
file_path <- file.path(wd, excel_file_name)
sheet_name <- "Direct impact contributions"

### 1. Read the Impact Category Block (Columns B:D)
# Impact header is in row 5; data starts at row 6
impact_df <- read_excel(file_path,
                        sheet = sheet_name,
                        range = cell_limits(ul = c(5, 2), lr = c(NA, 4)),
                        col_names = TRUE)

### 2. Read and Process the Process Information Header (Columns E onward, Rows 2–4)
process_header_df <- read_excel(file_path,
                                sheet = sheet_name,
                                range = cell_limits(ul = c(2, 5), lr = c(4, NA)),
                                col_names = FALSE)

# Create composite header:
# - Force all elements to character
# - Replace logical/NA values with an empty string
# - Pad to three components if needed, then collapse with a pipe
process_header <- apply(process_header_df, 2, function(x) {
  x <- sapply(x, function(y) {
    if (is.logical(y) || is.na(y)) "" else as.character(y)
  })
  if (length(x) < 3) x <- c(x, rep("", 3 - length(x)))
  paste(x, collapse = "|")
})
print(process_header)

### 3. Read the Process Data (Columns E onward, starting from Row 5)
process_data_df <- read_excel(file_path,
                              sheet = sheet_name,
                              range = cell_limits(ul = c(5, 5), lr = c(NA, NA)),
                              col_names = FALSE)
colnames(process_data_df) <- process_header

### 4. Ensure Both Blocks Have the Same Number of Rows
n_impact <- nrow(impact_df)
n_process <- nrow(process_data_df)
if (n_impact != n_process) {
  warning("Impact data has ", n_impact, " rows and process data has ", n_process, 
          " rows. Subsetting to the minimum common rows.")
  common_rows <- min(n_impact, n_process)
  impact_df <- impact_df[1:common_rows, ]
  process_data_df <- process_data_df[1:common_rows, ]
}

### 5. Combine the Impact and Process Data
combined_df <- cbind(impact_df, process_data_df)

### 6. Pivot the Process Data from Wide to Long Format
# Impact data is in the first few columns (B:D); pivot all columns from E onward.
long_df <- combined_df %>%
  pivot_longer(
    cols = (ncol(impact_df) + 1):ncol(combined_df),
    names_to = "composite_header",
    values_to = "process_value",
    values_drop_na = FALSE
  )

### 7. Separate the Composite Header into Its Components
# This splits the composite header (e.g., "Process UUID|Process|Location")
# into three columns: process_uuid, process, and location.
long_df <- long_df %>%
  separate(composite_header, into = c("process_uuid", "process", "location"),
           sep = "\\|", fill = "right")

### 8. (Optional) Remove Rows Containing Repeated Header Text
long_df <- long_df %>%
  filter(!(process_uuid %in% c("Process UUID", NA, "")))

### 9. Data Integrity Checks and Final Output
expected_cells <- nrow(process_data_df) * ncol(process_data_df)
actual_cells <- nrow(long_df)
cat("Expected cells:", expected_cells, "\n")
cat("Actual cells in long_df:", actual_cells, "\n")

non_na_wide <- sum(!is.na(as.matrix(process_data_df)))
non_na_long <- sum(!is.na(long_df$process_value))
cat("Non-NA cells in wide data:", non_na_wide, "\n")
cat("Non-NA cells in long data:", non_na_long, "\n")

cat("Unique Process UUIDs:\n")
print(unique(long_df$process_uuid))
cat("\nUnique Process Names:\n")
print(unique(long_df$process))
cat("\nUnique Locations:\n")
print(unique(long_df$location))

print("Cleaned Long DataFrame:")
print(head(long_df))
View(long_df)

