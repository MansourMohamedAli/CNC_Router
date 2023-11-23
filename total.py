import openpyxl

# Function to calculate the sum of a column
def calculate_column_sum(sheet, column_name):
    column_values = [float(row[column_name].value) for row in sheet.iter_rows(values_only=True)]
    return sum(column_values)

# Open the Excel file
filename = 'BOM.xlsx'
workbook = openpyxl.load_workbook(filename)
sheet = workbook.active

# Calculate the sums
quantity_sum = calculate_column_sum(sheet, 'Quantity')
unit_price_sum = calculate_column_sum(sheet, 'Unit Price')
total_price_sum = calculate_column_sum(sheet, 'Total Price')

# Print the sums
print(f"Sum of Quantity: {quantity_sum}")
print(f"Sum of Unit Price: {unit_price_sum}")
print(f"Sum of Total Price: {total_price_sum}")

# Close the workbook
workbook.close()

