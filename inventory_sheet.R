library(readxl)

##### Product System - Inventory Sheet ######

# import and clean up data for inventory sheet

# Define your working directory and file name
wd <- getwd()  # get or specify your working directory, e.g., "C:/Users/yourname/Documents"
excel_file_name <- "3__Cooling.xlsx"

# Build the full file path
file_path <- file.path(wd, excel_file_name)

# Read the Inventory sheet from cell B2:N1801 without treating any row as headers
inventory_df <- read_excel(file_path, 
                           sheet = "Inventory", 
                           col_names = FALSE, 
                           range = cell_limits(ul = c(2, 2), lr = c(NA, 14)))

# Extract the first row as header (this row corresponds to row 2 in the original Excel)
new_header <- as.character(inventory_df[2, ])

# Remove the first row from the data and set it as the column names
inventory_df <- inventory_df[-1, ]
names(inventory_df) <- new_header

# Append "_input" to headers in columns 1-7 and "_output" to headers in columns 8-13
names(inventory_df)[1:7] <- paste0(names(inventory_df)[1:7], "_input")
names(inventory_df)[8:13] <- paste0(names(inventory_df)[8:13], "_output")

# Delete the 7th column
inventory_df <- inventory_df[, -7]

# delete 1st row of headings (excess)
inventory_df <- inventory_df[-1,]

# Check the modified headers
print(names(inventory_df))

# make any columns containing Result numerical
inventory_df <- inventory_df %>%
  mutate(across(contains("Result"), as.numeric))

