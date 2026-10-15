# CR Site Availability Report Automation Tool

An Excel VBA automation tool for evaluating GSM, UMTS, and LTE cell availability during a configurable commissioning window. The tool combines raw KPI data, evaluates site-level availability, identifies degraded cells, generates a summary report, and exports individual reports for sites that pass the configured criteria.

## Overview

The tool automates the following tasks:

- Import GSM, UMTS, and LTE raw KPI data.
- Combine cell-level availability data into a unified dataset.
- Evaluate site availability against configurable thresholds.
- Check hourly KPI continuity and cell data completeness.
- Identify degraded cells and failed sites.
- Generate site-level summary and PASS site data.
- Export individual CR commissioning reports using an Excel template.
- Format worksheets for easier analysis and reporting.

## Requirements

- Microsoft Excel with VBA macro support enabled.
- The Excel workbook must contain a worksheet named `Control`.
- The CR report template must be available at `Template\CR Report.xlsx`.
- Input files must contain the expected GSM, UMTS, and LTE KPI columns.
- The user must have permission to create files in the export directory.

## Configuration

The settings form allows users to configure the evaluation criteria before processing.

- **Window Hours:** Number of hours to evaluate.
- **Availability Threshold:** Availability threshold used to assess hourly performance and average availability.
- **Minimum Availability:** Minimum permitted cell availability. Values below this limit cause the site to fail.
- **Maximum Degraded Hours:** Maximum number of degraded hours permitted during the evaluation window.

Availability values are expected as percentages from `0` to `100`.

## Input Data

The tool processes raw KPI data for three technologies.

### GSM

- Column A: Time
- Column D: Cell Name
- Column G: Cell Availability

### UMTS

- Column A: Time
- Column D: Cell Name
- Column G: Cell Availability

### LTE

- Column A: Time
- Column D: Cell Name
- Column H: Cell Availability

The input files must follow the expected column structure for the corresponding technology.

## Processing Workflow

1. Open the settings form and enter the evaluation criteria.
2. Select the source KPI files.
3. Import GSM, UMTS, and LTE data.
4. Combine the imported data into the `Combine Data` worksheet.
5. Extract the site identifier and technology from each cell name.
6. Evaluate site performance against the configured availability criteria.
7. Generate the `Summary` worksheet with site-level results.
8. Create the `PASS Site Data` worksheet containing the details of sites that pass.
9. Export an individual commissioning report for each PASS site.
10. Apply the required worksheet formatting.

## Evaluation Logic

The tool evaluates site availability using the configured evaluation window and KPI thresholds.

### Hourly Data Validation

- Check that hourly timestamps are consecutive.
- Check that the expected cell data is available for each hour.
- Identify cells whose hourly availability falls below the configured availability threshold.

### Site PASS/FAIL Criteria

A site passes when all configured conditions are satisfied:

- The required hourly evaluation window is available.
- The required cell data is present for every evaluated hour.
- No cell availability value falls below the minimum availability limit.
- The number of degraded hours does not exceed the configured maximum.
- The site's average availability meets the configured availability threshold.

A site that violates any required condition is marked as `FAIL`.

## Output Worksheets

### Control

Contains the control worksheet used by the workbook.

### Combine Data

Contains the combined KPI dataset with the following fields:

- Time
- Cell Name
- Cell Availability (%)
- Site
- Tech

### Summary

Contains the site-level evaluation results, including:

- Site ID
- Technology
- Final Status
- Total Cell Count
- Start Time
- End Time
- Total Window (hrs)
- Remark
- Degraded Cells

### PASS Site Data

Contains cell-level information for sites that pass the evaluation, including the derived band and cell information.

## Exported Reports

The tool exports a separate commissioning report for each site that passes the evaluation.

**Filename format:**

`<SiteID>_KPI_Commission Report.xlsx`

**Output directory:**

`Export`

The source template is:

`Template\CR Report.xlsx`

The template is expected to contain the following worksheets:

- `2G RAW`
- `3G RAW`
- `4G RAW`
- `3 Day KPI Summary`
- `Cell Level Hourly KPI Graph`
- `2G Pivot`
- `3G Pivot`
- `4G Pivot`

The final exported reports remove the raw-data and pivot worksheets specified by the automation logic.

## How to Use

1. Open the macro-enabled Excel workbook.
2. Confirm that the required `Control` worksheet exists.
3. Verify that `Template\CR Report.xlsx` is available.
4. Prepare the source KPI files using the required column structure.
5. Run the main `Process` macro.
6. Configure the evaluation settings in the form.
7. Select the source files when prompted.
8. Wait for the processing and site evaluation to finish.
9. Review the `Summary` worksheet and check the final site status.
10. Open the `Export` folder to access the reports generated for PASS sites.

## Important Notes

- **Back up your workbook before running the macro.** The processing routine deletes worksheets other than `Control` before rebuilding the output.
- The file picker may allow multiple selections, but the current import logic uses the first selected filename containing `cr`. Confirm that the intended source file is selected.
- Ensure that timestamps are consistent and suitable for hourly evaluation.
- Check that cell names follow the expected naming convention because site, band, and cell identifiers are derived from those names.
- Ensure that availability values use the expected percentage scale from `0` to `100`, rather than decimal fractions such as `0.95`.
- Confirm that the template worksheets and folder paths match the workbook's expected structure.
- Review failed sites and their remarks before finalizing the commissioning report.

## Troubleshooting

**No reports are exported**

Check the `Summary` worksheet to confirm whether any sites passed the configured evaluation criteria.

**Incorrect site identification**

Verify the cell naming convention used by each technology. GSM site identification uses the first six characters, while UMTS and LTE use the substring beginning at character five with a length of six characters.

**Unexpected FAIL status**

Review the hourly data continuity, cell data completeness, minimum availability, average availability, and degraded-hour criteria.

**Template or export errors**

Confirm that `Template\CR Report.xlsx` exists, that the required worksheets are present, and that the workbook has permission to write to the `Export` directory.

## Technology

- Microsoft Excel
- Excel VBA
- Excel worksheets and formulas
- Automated KPI evaluation and report generation

## Purpose

This tool helps streamline CR site availability assessment by standardizing KPI validation, reducing manual Excel processing, and automatically generating commissioning reports for sites that meet the configured acceptance criteria.
