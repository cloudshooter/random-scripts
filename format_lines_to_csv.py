# This script will take a block of four lines, the bottom 2 of which are blank, and convert them into a csv column format
# Example input text:
# Text for line 1
# Text for Line 2
# <empty>
# <empty>
#
# Output will then look like
# Text for line 1,Text for Line2

import csv

def process_text_file(input_file, output_file):
    with open(input_file, 'r') as file:
        lines = file.readlines()

    data = []
    for i in range(0, len(lines), 4):
        text_line = lines[i].strip()
        date_line = lines[i+1].strip()
        data.append([text_line, date_line])

    with open(output_file, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['Text Line', 'Date'])
        csvwriter.writerows(data)

# Usage
input_file = 'your_text_file.txt'
output_file = 'output.csv'
process_text_file(input_file, output_file)
