import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList

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
input_file = r'C:\Josh\Stock\Portfolio_Positions_Jan-10-2026.csv'
temp_file = r'C:\Josh\Stock\Book1_fixed.csv'
output_excel = r'C:\Josh\Stock\Portfolio_Chart.xlsx'
# Read the file line by line
with open(input_file, 'r') as f:
    lines = f.readlines()

# Add comma to the first line (header) if it doesn't end with comma
if lines[0].strip() and not lines[0].strip().endswith(','):
    lines[0] = lines[0].strip() + ',\n'

# Write the fixed version
with open(temp_file, 'w') as f:
    f.writelines(lines)

print("Fixed header row - added comma at the end")
print()

# Now read the fixed CSV
df = pd.read_csv(temp_file)

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

if symbol_col and value_col:
    print(f"{symbol_col:<15} {value_col:>20}")
    print("-" * 50)
    
    for index, row in df.iterrows():
        symbol = row[symbol_col]
        value = row[value_col]
        print(f"{symbol:<15} {value:>20}")
    
    print("=" * 50)
    
    # Calculate total if possible
    try:
        total = pd.to_numeric(df[value_col], errors='coerce').sum()
        print(f"{'TOTAL:':<15} ${total:>19,.2f}")
    except:
        print("Could not calculate total")
    
    # Export to Excel with pie chart
    print()
    print("Creating Excel file with pie chart...")
    
    # Prepare data for chart (convert values to numbers)
    chart_data = df[[symbol_col, value_col]].copy()
    print(chart_data[value_col])


    chart_data[value_col] = chart_data[value_col].apply(clean_value)
    #chart_data[value_col] = pd.to_numeric(chart_data[value_col], errors='coerce')
    print(chart_data[value_col])
    total = chart_data.iloc[1:, 1].sum()  # Sum of values column
    chart_data['Percentage'] = ((chart_data.iloc[:, 1] / total))
    
    # Export to Excel
    chart_data.to_excel(output_excel, index=False, sheet_name='Portfolio')
    
    # Add pie chart
    wb = load_workbook(output_excel)
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
    wb.save(output_excel)
    
    print(f"✓ Excel file created: {output_excel}")
    print("  - Data is in columns A & B")
    print("  - Pie chart is on the right")
    
else:
    print("Error: Could not find Symbol or Value columns")
    print(f"Available columns: {list(df.columns)}")