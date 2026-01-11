import os
import sys

path_to_add = os.path.abspath(os.path.join(os.path.dirname(__file__), '../'))
if path_to_add not in sys.path:
    sys.path.insert(0, path_to_add)

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from portfolio_tracker import config
from collections import defaultdict

def clean_value(val):
    if isinstance(val, str):
        # Remove $, commas, and whitespace
        val = val.replace('$', '').replace(',', '').strip()
    try:
        return float(val)
    except:
        return 0
    
# Read the CSV file
# The header has one less comma than data rows, so we'll handle this
# Read the file line by line
with open(config.INPUT_FILE, 'r') as f:
    lines = f.readlines()

# Add comma to the first line (header) if it doesn't end with comma
if lines[0].strip() and not lines[0].strip().endswith(','):
    lines[0] = lines[0].strip() + ',\n'

# Write the fixed version
with open(config.TEMP_FILE, 'w') as f:
    f.writelines(lines)

print("Fixed header row - added comma at the end")
print()

# Now read the fixed CSV
df = pd.read_csv(config.TEMP_FILE)

# Find the Symbol column
symbol_col = None
for col in df.columns:
    if 'symbol' in str(col).lower() or 'ticker' in str(col).lower():
        symbol_col = col
        break

# Find the Value column
value_col = None
for col in df.columns:
    if 'value' in str(col).lower() or 'amount' in str(col).lower() or 'total' in str(col).lower():
        value_col = col
        break

# Display results
print("PORTFOLIO HOLDINGS")
print("=" * 50)

portfolio_data = defaultdict(list)

if symbol_col and value_col:

    for index, row in df.iterrows():
        symbol = row[symbol_col]
        value = row[value_col]
        portfolio_data[symbol].append(clean_value(value))
        print(f"{symbol:<15} {value:>20}")

    portfolio_data = {key: sum(values) for key, values in portfolio_data.items()}

    print(portfolio_data)
    # Calculate total if possible
    # try:
    #     total = pd.to_numeric(df[value_col], errors='coerce').sum()
    #     print(f"{'TOTAL:':<15} ${total:>19,.2f}")
    # except:
    #     print("Could not calculate total")

    
    # Prepare data for chart (convert values to numbers)
    chart_data = df[[symbol_col, value_col]].copy()
    # Convert to DataFrame for chart
    chart_data = pd.DataFrame(list(portfolio_data.items()), columns=['Symbol', 'Value'])

    print(chart_data)
    total = chart_data.iloc[1:, 1].sum()  # Sum of values column
    chart_data['Percentage'] = ((chart_data.iloc[:, 1] / total))
    
    # Export to Excel
    chart_data.to_excel(config.OUTPUT_FILE, index=False, sheet_name='Portfolio')
    
    # Add pie chart
    wb = load_workbook(config.OUTPUT_FILE)
    ws = wb['Portfolio']
    for row in range(2, len(chart_data) + 2):
        cell = ws[f'C{row}']
        cell.number_format = '0.00%'  # Excel percentage format
    # Create pie chart
    pie = PieChart()
    labels = Reference(ws, min_col=1, min_row=2, max_row=len(chart_data)+1)  # Column A: Symbols
    data = Reference(ws, min_col=3, min_row=2, max_row=len(chart_data)+1)    # Column C: Percentage

    pie.add_data(data, titles_from_data=False)
    pie.set_categories(labels)
    pie.title = "Portfolio Distribution"

    # Add percentages to the chart
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showVal = True          # Show the percentage values
    pie.dataLabels.showCatName = False     # Don't show symbol name
    pie.dataLabels.showPercent = False     # Don't calculate percent (we already have it)
    pie.dataLabels.showSerName = False     # Don't show series name (this removes "Series")
    pie.dataLabels.showLeaderLines = False
    pie.dataLabels.showLegendKey = False
    # Position labels outside the pie slices to prevent overlap
    from openpyxl.chart.label import DataLabelList
    #pie.dataLabels.position = 'bestFit'  # or 'outEnd' to put all labels outside
    pie.dataLabels.position = 'bestFit'
    # Make the chart bigger to fit labels better
    # Make the chart much bigger to fit labels better
    pie.width = 20  # Width in cm
    pie.height = 15  # Height in cm
        # Add chart to sheet
# Remove chart title (we'll add it as a cell instead)
    pie.title = None

    # Add title as a cell above the chart
    ws['E2'] = 'Portfolio Distribution'
    ws['E2'].font = ws['E2'].font.copy(size=16, bold=True)
    ws.add_chart(pie, "E4")
    
    # Save
    wb.save(config.OUTPUT_FILE)

    print(f"✓ Excel file created: {config.OUTPUT_FILE}")
    print("  - Data is in columns A & B")
    print("  - Pie chart is on the right")
    
else:
    print("Error: Could not find Symbol or Value columns")
    print(f"Available columns: {list(df.columns)}")