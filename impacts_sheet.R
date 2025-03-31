library(readxl)
library(tidyverse)

##### Product Systems - Impacts Sheet ######

# import and clean up data for Impacts sheet

# Define your working directory and file name
wd <- getwd()  # get or specify your working directory, e.g., "C:/Users/yourname/Documents"
excel_file_name <- "3__Cooling.xlsx"

# Build the full file path
file_path <- file.path(wd, excel_file_name)

# Read the Inventory sheet from cell B2:Hx without treating any row as headers
impacts_df <- read_excel(file_path, 
                         sheet = "Impacts")


#### Research Questions ####
## which impact categories are most important?

pareto_df <- impacts_df %>%
  arrange(desc(Normalized)) %>% #normalized is exported with the raw data because of the ReCiPe methodology
  mutate(CumPercent = cumsum(Normalized) / sum(Normalized) * 100)

idx <- which(pareto_df$CumPercent >= 80)[1] # select index of first row to add up to at least 80%
# Select all rows up to that index
selected_categories <- pareto_df %>% slice(1:idx)
# The most important impact categories are:
cat("The most important impact categories are: ",
    paste(selected_categories$`Impact category`, collapse = ", "))


