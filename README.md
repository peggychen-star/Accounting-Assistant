# Accounting Excel Preview — filters, deleted columns, pivot summaries, Excel export

A Streamlit app for reviewing accounting Excel worksheets, filtering records, hiding columns, creating pivot summaries (including optional split-by categories), and exporting results as an Excel workbook.

## Features

- **Upload Excel**: Select a local `.xlsx` or `.xls` file and a worksheet. The first row contains the headers.
- **Filter**: Each click creates a filter setup tab. Choose a column, then select multiple values. Choices within one filter use **OR**; separate filter tabs combine using **AND**.
- **Delete Column**: Each click creates a delete-column tab. Hide any number of columns from the processed preview/export without affecting the original worksheet or previously chosen filters.
- **Create New Tab → Pivot**: Each click creates a pivot setup tab. Select:
  1. **Categorical column (group by)** — such as `Name`.
  2. **Continuous / numeric column (sum)** — such as `Value`.
  3. **Split by (optional)** — such as `Month`; then select **one or more categories** (such as `Jan` and `Feb`) in the multiselect that appears.

  Without a split, the summary adds a results tab named **Pivot Name_Value**. With a split, it adds **Pivot Name_Value_Separated_by_Month**. Each selected split category becomes a separate sum column; absent combinations are **0**. The final **Total** column sums the *selected* split categories, not unselected ones.

- **Pivot data source**: Pivots use every matching row after filters, including source columns hidden from the Processed Data preview. Changes to row filters automatically recalculate pivots.
- **Scrollable results**: The processed dataset and pivot result tables scroll vertically, showing about 10 rows at a time. Exports include **all** rows, not just visible ones.
- **Export**: Download one `.xlsx` workbook with the **Processed Data** worksheet (filtered rows and remaining columns) and all completed pivot summaries on separate worksheets. Export has a frozen header row, spreadsheet filters, and readable widths. The pivot worksheets contain summary values, not interactive Excel PivotTable objects.
- **Back**: Undo the latest action, including creating a filter/deletion/pivot, changing a dropdown, or changing selected split categories.

The original uploaded file is not modified. Uploading a different file or selecting a different worksheet starts a fresh workspace.

## Example: Pivot with Split by

**Source**

| Name | Value | Month |
| --- | ---: | --- |
| Adam | 2 | Jan |
| Adam | 3 | Jan |
| Adam | 2 | Feb |
| Eve | 1 | Jan |
| Eve | 2 | Jan |

Select `Name` for the categorical column, `Value` for the numeric column, `Month` for **Split by**, and select `Jan` and `Feb` in the category multiselect.

**Result tab: `Pivot Name_Value_Separated_by_Month`**

| Name | Jan | Feb | Total |
| --- | ---: | ---: | ---: |
| Adam | 5 | 2 | 7 |
| Eve | 3 | 0 | 3 |

**Excel worksheet naming:** Excel limits worksheet names to **31 characters**, so it cannot use the full 35-character example title. The **Streamlit tab displays the full requested title**, but the exported Excel worksheet uses a readable abbreviation, **`Pivot Name_Value_by_Month`**. Very long or repeated names are shortened and made unique automatically.

## Run on Windows PowerShell

1. Extract the ZIP, then open PowerShell inside the folder containing `app.py`.
2. Install the requirements if needed:

   ```powershell
   python -m pip install -r requirements.txt
   ```

3. Start the app:

   ```powershell
   python -m streamlit run app.py
   ```

4. Open `http://localhost:8501` if a browser doesn't open automatically.

**Note:** Export creates a new processed Excel workbook; it does not preserve original workbook formatting, charts, formulas, or hidden sheets.
