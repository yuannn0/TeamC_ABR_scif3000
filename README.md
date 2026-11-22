# TeamC_ABR_scif3000
# SCIF3000 | NHMRC Funding Analysis, 2020–2025

This analysis summarises cumulative funding by research area for awards announced by Australia's National Health and Medical Research Council (NHMRC) from 2020 to 2025. It also reproduces the NHMRC funding figure used in the SCIF3000 final report for Australian BioResources.

## Data source and preparation

The source data came from the NHMRC funding outcomes data. The project team manually cleaned the source data in Excel before analysis. The resulting workbook, `2020_2025_NHMRC.xlsx`, is the analysis input. The workflow below starts from this prepared workbook. The R script reads the only worksheet, `Sheet1`, which contains 4,386 award records across all six announcement years. It uses the `APP ID`, `Date Announced`, `Total` and `Field of Research` columns.

## Reproduce the analysis

1. Download `nhmrc_full_count.R` and `2020_2025_NHMRC.xlsx` and place them in the same folder. Keep both filenames as shown.
2. Install R. In the R console, install `readxl` once:

   ```r
   install.packages("readxl")
   ```

3. In RStudio, open `nhmrc_full_count.R` and click **Source**. Alternatively, open a terminal in the folder containing both files and run:

   ```sh
   Rscript "nhmrc_full_count.R"
   ```

### Before running the uploaded script

The chart function calls `mtext(subtitle, ...)`, but this version of the script does not define `subtitle`. Add the following line immediately below `main_title` near the top of the script to use a blank subtitle:

```r
subtitle <- ""
```

Without this edit, a fresh R session stops at the first chart export with `object 'subtitle' not found`. Some summary and audit files may already have been written by that point.

## How funding was counted

The script selects awards by announcement year and reports funding in millions of Australian dollars. It assigns each workbook record's full award amount once to each research area that it matches. The script parses the `Field of Research` values and groups related labels into immunology and neuroscience. Because an award can contribute to more than one area, the totals across areas can include the same funding more than once. The chart ranks 15 predefined categories by cumulative funding. The grey bar represents records whose research field was marked “Not Applicable”.

## Outputs

The script writes these files beside the R script using its default settings.

| File | Contents |
| --- | --- |
| `nhmrc_cumulative_bars.png` | Cumulative funding chart |
| `nhmrc_cumulative_bars.pdf` | Vector version of the chart |
| `displayed_funding.csv` | Funding values displayed in the chart |
| `NHMRC_full_count_summary.csv` | Funding summary for the 15 chart categories |
| `audit/` | All-field and annual summaries, classification mappings and record-level funding allocations |

To also generate `NHMRC_full_count_summary.xlsx`, install the optional `writexl` package with `install.packages("writexl")` before running the script.
