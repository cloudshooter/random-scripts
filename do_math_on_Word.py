from docx import Document
import pandas as pd

# Prompt the user for the filename
filename = input("Enter the path to your Word document: ")

# Load the Word document
doc = Document(filename)

# Initialize a list to store extracted data
data = []

# Iterate through each table in the document
for table in doc.tables:
    # Iterate through each row in the table
    for row in table.rows:
        cells = row.cells
        # Try to extract "Severity" and "Total Affected" assuming they are in the first two columns
        try:
            severity = cells[0].text.strip()
            total_affected = int(cells[1].text.strip())
            # Append the data if it looks like what we expect
            if severity and total_affected is not None:
                data.append({'Severity': severity, 'Total Affected': total_affected})
        except (IndexError, ValueError):
            # Ignore rows that don't match the expected format
            continue

# Convert the list of dictionaries to a DataFrame for easy manipulation
df = pd.DataFrame(data)

# Group by 'Severity' and sum 'Total Affected'
totals_by_severity = df.groupby('Severity')['Total Affected'].sum()

print(totals_by_severity)
