# LICA Data Automation Project -  openlca2r_automation

This repository contains scripts and workflows for extracting, cleaning, and analysing Life Cycle Inventory (LCI) and Life Cycle Impact Assessment (LCIA) results for an example cooling system product system exported from OpenLCA 2. It automates reading from Excel exports, restructuring data into tidy formats, performing quality checks, and generating basic visualisations.

## Table of Contents

- [Description](#description)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Data Files](#data-files)
- [Scripts Overview](#scripts-overview)
- [Usage](#usage)
- [Visualisation](#visualisation)
- [Methodology](#methodology)
- [Technical Approach](#technical-approach)
- [Contributing](#contributing)
- [License](#licence)

## Description

The cooling project workflows perform the following tasks:

1. **Extract general project info** (name, description, LCIA method, variants, parameters)  
2. **Parse LCI Results** into inputs and outputs tables with integrity checks  
3. **Parse LCIA Results** and reshape for analysis  
4. **Analyse Impacts** (Pareto analysis of impact categories)  
5. **Clean Inventory Data** for material and emissions flows  
6. **Process Direct Impact Contributions** and reshape into long format  
7. **Process Direct Inventory Contributions** similarly  
8. **Read Impact Methodology** from setup sheet  

All scripts assume the corresponding Excel workbooks are in the working directory.

## Prerequisites

- **R** (version ≥ 3.6)  
- **R packages**:  
  - `tidyverse`  
  - `cellranger`  
  - `readxl`

Install missing packages with:

```r
install.packages(c("readxl", "tidyverse", "cellranger"))
```

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/daniel-lee-wilkinson/openlca2r_automation.git
   cd openlca2r_automation
   ```
2. Place the Excel files in the project root:
   - `project result.xlsx` (sheets: `Info`, `LCI Results`, `LCIA Results`)  
   - `3__Cooling.xlsx` (sheets: `Impacts`, `Inventory`, `Direct impact contributions`, `Direct inventory contributions`, `Calculation setup`)  # example output from OpenLCA 2, Excel worksheet names are standard

## Data Files

- **project result.xlsx**  
  - `Info`: general project metadata, variants, parameters  
  - `LCI Results`: input/output flows by scenario  
  - `LCIA Results`: impact assessment results by category & scenario  

- **3__Cooling.xlsx**  
  - `Impacts`: normalised impact shares for Pareto analysis  
  - `Inventory`: raw material/emission inventory flows  
  - `Direct impact contributions`: process-level impact contributions  
  - `Direct inventory contributions`: process-level emission/inventory contributions  
  - `Calculation setup`: LCIA methodology cell (C8)  

## Scripts Overview

| Script                                    | Purpose                                                           |
|-------------------------------------------|-------------------------------------------------------------------|
| `project_info.R`                          | Load and clean general info, variants & parameters               |
| `project_lci_results.R`                  | Read `LCI Results`, split inputs/outputs, perform integrity checks |
| `project_lcia_results.R`                 | Read `LCIA Results`, reshape for analysis, generate heatmap      |
| `impacts_sheet.R`                        | Pareto analysis on `Impacts` sheet                                |
| `inventory_sheet.R`                      | Clean and restructure `Inventory` sheet                           |
| `direct_impact_contributions.R`          | Clean and pivot `Direct impact contributions`                     |
| `direct_inventory_contributions_sheet.R` | Clean and pivot `Direct inventory contributions`                  |
| `impact_methodology.R`                   | Read LCIA methodology from `Calculation setup` sheet              |

## Usage

Run each script from R or command line:

```bash
Rscript project_info.R
Rscript project_lci_results.R
Rscript project_lcia_results.R
Rscript impacts_sheet.R
Rscript inventory_sheet.R
Rscript direct_impact_contributions.R
Rscript direct_inventory_contributions_sheet.R
Rscript impact_methodology.R
```

After running, check console output for previews and any integrity check messages.

## visualisation

This project uses **ggplot2** for all graphing and charting needs, ensuring reproducible, high-quality plots. Key visualisations include:

- **Diverging Heatmap** (in `project_lcia_results.R`):
  - Displays z-scored LCIA results per scenario and impact category.
  - Built with `geom_tile()`, a custom diverging color scale (`scale_fill_gradient2()`), and thematic adjustments (`theme_minimal()`).

- **Pareto Bar Chart** (in `impacts_sheet.R`):
  - Shows cumulative impact contribution by category.
  - Created with `geom_bar()` for individual shares and a `geom_line()` overlay for cumulative percentage.
  - Annotated with `geom_text()` to highlight thresholds (e.g., 80% cumulative share).

- **Inventory Flow Plots** (optional; extendable in `inventory_sheet.R`):
  - Users can visualize top material/emission flows using `geom_col()` or `geom_point()`.
  - Faceting (`facet_wrap()`) allows comparison across process stages or scenarios.

- **Contribution Waterfall Chart** (optional; in `direct_impact_contributions.R` or `direct_inventory_contributions_sheet.R`):
  - Illustrates stepwise impact or inventory contributions by process.
  - Implemented via `geom_segment()` and `geom_rect()`, or by using the `waterfall` package integration.

## Methodology

- Impact methodology is read directly from cell C8 in the `Calculation setup` sheet by `impact_methodology.R`.
- Ensure the LCIA method name is correctly placed in the workbook.

## Technical Approach

### Steps, Techniques, and Tools

1. **Project Metadata Extraction**
   - **Purpose**: Load and clean general project information (name, description, LCIA method, variants, parameters).
   - **Techniques**: `readxl::read_excel` to import the `Info` sheet; `dplyr::select`, `rename`, and `mutate` for tidying and type conversion.
   - **Tools**: R, `readxl`, `dplyr`.

2. **LCI Results Parsing**
   - **Purpose**: Read the `LCI Results` sheet, then split and validate input/output flows.
   - **Techniques**: Use `cellranger::cell_limits` to define data ranges; reshape with `tidyr::pivot_longer` and `pivot_wider`; perform integrity checks via `dplyr::filter` and summary statistics.
   - **Tools**: R, `readxl`, `dplyr`, `tidyr`, `cellranger`.

3. **LCIA Results Reshaping & visualisation**
   - **Purpose**: Process the `LCIA Results` sheet and create comparative impact heatmaps.
   - **Techniques**: Normalize (z-score) impact values for comparability; reshape with `tidyr::pivot_longer`; visualize using `ggplot2` with a diverging color scale.
   - **Tools**: R, `tidyr`, `dplyr`, `ggplot2`.

4. **Pareto Analysis of Impact Categories**
   - **Purpose**: Identify the most significant impact categories from the `Impacts` sheet.
   - **Techniques**: Calculate cumulative percentages; filter top contributors until cumulative share threshold (e.g., 80%); use `dplyr::arrange` and `cumsum`.
   - **Tools**: R, `dplyr`.

5. **Inventory Data Cleaning**
   - **Purpose**: Standardize and tidy material/emissions flows from the `Inventory` sheet.
   - **Techniques**: String manipulation with `stringr::str_replace_all`; reshape into long format using `tidyr::pivot_longer`; label flows consistently.
   - **Tools**: R, `tidyverse` (`stringr`, `dplyr`, `tidyr`).

6. **Direct Contribution Processing**
   - **Purpose**: Parse process-level contributions for both impacts and inventory.
   - **Techniques**: `tidyr::pivot_longer` to convert wide tables into key-value pairs; separate category labels; join back to master lists.
   - **Tools**: R, `tidyr`, `dplyr`.

7. **LCIA Methodology Extraction**
   - **Purpose**: Retrieve the selected LCIA method from the `Calculation setup` sheet.
   - **Techniques**: Single-cell import via `readxl::read_excel` with a specific `range`; minimal transformation of the raw string.
   - **Tools**: R, `readxl`.

### Underlying Framework
- **Language**: R (version ≥ 3.6)
- **Automation**: Scripts executed via `Rscript` for reproducibility and integration into pipelines.
- **Version Control**: Git for collaboration, branching, and change tracking.

## Contributing

1. Fork the repo  
2. Create a feature branch: `git checkout -b feature/my-feature`  
3. Commit your changes  
4. Push to your branch  
5. Open a Pull Request  

Please include clear descriptions of any new scripts or changes to existing workflows.


