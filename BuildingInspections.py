import openpyxl
import pandas as pd
import datetime
import subprocess

def excel_range_to_styled_html(file_path, range_name, output_html_path):
  """Opens an Excel file, reads a specific named range,

  cleans empty cells, and generates a custom-styled HTML table.
  """
  # 1. Load the workbook
  wb = openpyxl.load_workbook(file_path, data_only=True)

  if range_name not in wb.defined_names:
    raise ValueError(f"Named range '{range_name}' not found in the workbook.")

  destinations = list(wb.defined_names[range_name].destinations)
  if not destinations:
    raise ValueError(
        f"Named range '{range_name}' does not point to any valid cells."
    )

  sheet_title, cell_coordinate = destinations[0]
  ws = wb[sheet_title]

  # 2. Extract all rows into a raw list of lists
  raw_data = []
  for row in ws[cell_coordinate]:
    raw_data.append([cell.value for cell in row])

  if not raw_data:
    print("The named range is empty.")
    return

  # 3. Replace None and NaN values with empty strings
  cleaned_data = []
  for row in raw_data:
    cleaned_row = []
    for cell in row:
      if cell is None or (isinstance(cell, float) and np.isnan(cell)):
        cleaned_row.append("")
      elif isinstance(cell, datetime.datetime):
        # <<< CHANGE IS HERE: Convert datetime to date string only >>>
        cleaned_row.append(cell.strftime("%Y-%m-%d"))
      elif isinstance(cell, datetime.date):
        cleaned_row.append(cell.strftime("%Y-%m-%d"))
      else:
        cleaned_row.append(cell)
    cleaned_data.append(cleaned_row)

  if len(cleaned_data) > 2:
      cleaned_data.pop(2)

  # 4. Build the HTML string with inline CSS
  html = [
      '<table style="border-collapse: collapse; font-family: Arial, sans-serif;'
      ' font-size: 14px; width: 100%;">'
  ]

  for r_idx, row in enumerate(cleaned_data):
    html.append("  <tr>")
    # Both Row 1 (index 0) and Row 2 (index 1) act as headers
    is_header = r_idx in (0, 2)

    for c_idx, val in enumerate(row):
      tag = "th" if is_header else "td"

      # Base styles mimicking Excel gridlines
      styles = [
          "border: 1px solid #d3d3d3",
          "padding: 8px 12px",
          "text-align: left",
      ]

      if is_header:
        styles.append("font-weight: bold")
        styles.append("color: black")

        # Column-specific header backgrounds (Columns A, B, C = 0, 1, 2; E, F, G = 4, 5, 6)
        if c_idx in [0, 1, 2]:
          styles.append("background-color: #b3d1ff")  # Light Blue
        elif c_idx in [4, 5, 6]:
          styles.append("background-color: #ffb3b3")  # Light Red
        elif c_idx == 3:
          styles.append("background-color: #eaeaea")  # Spacer column header
      else:
        # Normal cells
        styles.append("color: black")
        styles.append("background-color: #ffffff")
        if c_idx == 3:
          styles.append("background-color: #fafafa")  # Subtle spacer background

      style_str = "; ".join(styles)
      html.append(f'    <{tag} style="{style_str}">{val}</{tag}>')

    html.append("  </tr>")
  html.append("</table>")

  # 5. Save to HTML file
  html_output = "\n".join(html)
  with open(output_html_path, "w", encoding="utf-8") as f:
    f.write(html_output)

  print(f"Styled HTML table successfully saved to: {output_html_path}")


# --- Example Usage ---
if __name__ == "__main__":
  file_path = r"C:\Users\jhoj\City of Placentia\Inspection Log - Documents\INSPECTION REQUEST LOG.xlsx"
  range_name = "Daily_Schedule"  # Replace with your actual named range
  output_html_path = r"C:\Users\jhoj\Scripts\Building_Inspections\output_table.html"

  try:
    excel_range_to_styled_html(file_path, range_name, output_html_path)

    print("Pushing update to GitHub...")

    subprocess.run(["git", "add", output_html_path], check = True)

    subprocess.run(
        [
            "git",
            "commit",
            "-m",
            "Automated update: refreshed table data",
        ],
        check = True,
    )

    subprocess.run(["git", "push", "origin", "main"], check = True)
    print("Successfully updated and pushed data to GitHub!")
  except Exception as e:
    print(f"Error: {e}")