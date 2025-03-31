library(readxl)

# Define your working directory and file name
wd <- getwd()
excel_file_name <- "3__Cooling.xlsx"  # ensure the file name matches exactly
file_path <- file.path(wd, excel_file_name)

# Read the cell C8 from the "Calculation setup" sheet
impact_methodology <- read_excel(file_path, 
                            sheet = "Calculation setup", 
                            range = "C8", 
                            col_names = FALSE)

# Convert the cell value to a string
impact_methodology <- as.character(cell_value_df[[1, 1]])
print(impact_methodology)
