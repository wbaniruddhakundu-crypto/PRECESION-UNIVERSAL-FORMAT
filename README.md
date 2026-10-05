# Universal Precision Validation

Separate Streamlit app for precision-study population.

Inputs:
- Precision Excel template
- INTRA PDF(s)
- Optional INTER PDF(s)

Logic:
- INTRA: 20 replicates per QC level
- INTER Run 1: first 5 INTRA replicates
- INTER Run 2+: one run per uploaded INTER PDF, in upload order
- Mean, SD and CV% formulas are rebuilt dynamically
- Excel rows are expanded when more INTER runs are uploaded
- Uploaded template workbook is used as the starting workbook

This is a separate app and does not modify the existing Tata 1MG Precision app.
