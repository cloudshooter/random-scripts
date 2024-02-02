import csv

def process_file(input_file, output_file):
    # Read the content from the input file and remove duplicates
    with open(input_file, 'r') as file:
        lines = file.read().splitlines()
        host_set = set(lines)

    # Create a list of lists for each column
    columns = [[] for _ in range(3)]

    # Organize values into columns
    for index, value in enumerate(host_set):
        # Remove ':' and trailing characters
        cleaned_value = value.split(':')[0].strip()
        columns[index % 3].append(cleaned_value)

    # Add empty values to columns to account for the remaining values
    remainder = len(host_set) % 3
    if remainder > 0:
        for i in range(3 - remainder):
            columns[i].append('')

    # Write processed data to the output CSV file
    with open(output_file, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)

        # Write each row to the CSV file
        for row in zip(*columns):
            csvwriter.writerow(row)

    # Print the number of items in the list divided by 3
    print(f"Number of items in the list divided by 3: {len(host_set) // 3}")

# Example usage
process_file("input.txt", "output.csv")
