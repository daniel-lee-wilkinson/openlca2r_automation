library(readxl)
library(dplyr)
library(tidyr)

## Define file details
wd <- getwd()
excel_file_name <- "project result.xlsx"
file_path <- file.path(wd, excel_file_name)
sheet_name <- "LCIA Results"

## Read the entire sheet without predefined column names.
raw_data_lcia <- read_excel(file_path, sheet = sheet_name, col_names = FALSE)
print(raw_data_lcia)

### Construct the Header ###
# Assume the following:
# - Columns: B = col1, C = col2, D = col3, E = col4, F = col5, etc.
# - Main LCIA headers (for columns B–D) are in row 3 (columns 1:3).
# - Scenario headings (for columns E onward) are in row 2 (columns 4 onward).

header_main <- raw_data_lcia[3, 1:3] %>% unlist() %>% as.character()
scenario_header <- raw_data_lcia[2, 4:ncol(raw_data_lcia)] %>% unlist() %>% as.character()
lcia_header <- c(header_main, scenario_header)
cat("LCIA Results Header:\n")
print(lcia_header)

### Extract the LCIA Results Data ###
# Data is from row 4 to the end, from column B onward.
lcia_data <- raw_data_lcia[4:nrow(raw_data_lcia), 1:ncol(raw_data_lcia)]
lcia_df <- as_tibble(lcia_data)
colnames(lcia_df) <- lcia_header

### Preview the LCIA Results Data Frame ###
cat("LCIA Results Data:\n")
print(head(lcia_df))

# convert scenario values to numerical
lcia_df <- lcia_df %>%
  mutate(across(4:ncol(lcia_df), as.numeric))



### Visualisation ####

# 1. heatmap of each scenario by impact category
lcia_long <- lcia_df %>%
  pivot_longer(
    cols = 4:5,   # or cols = c(`3. Cooling`, `3. Cooling (2)`)
    names_to = "Scenario",
    values_to = "Value"
  )

# For each impact category, compute a z-score so that the midpoint is 0.
# This way, a higher than average value gets a positive z-score (red),
# and a lower than average value gets a negative z-score (blue).
lcia_long <- lcia_long %>%
  group_by(`Impact category`) %>%
  mutate(Z = (Value - mean(Value, na.rm = TRUE)) / sd(Value, na.rm = TRUE)) %>%
  ungroup()

# Now create the heatmap using a diverging color scale.
ggplot(lcia_long, aes(x = Scenario, y = `Impact category`, fill = Z)) +
  geom_tile(color = "white") +
  scale_fill_gradient2(low = "blue", mid = "white", high = "red", midpoint = 0) +
  labs(
    title = "Diverging Heatmap of LCIA Results by Scenario",
    x = "Scenario",
    y = "Impact Category",
    fill = "Z-Score"
  ) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))
