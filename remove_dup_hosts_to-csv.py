import csv

def process_file(input_file, output_file):
    # Read the content from the input file and remove duplicates
    with open(input_file, 'r') as file:
        lines = file.read().splitlines()
        host_set = set(lines)

    # Create a list of lists for each column
    columns = [[] for _ in range(3)]

    # Organize values into columns
    total_values = len(host_set)
    step = total_values / 3
    for index, value in enumerate(host_set):
        # Remove ':' and trailing characters
        cleaned_value = value.split(':')[0].strip()
        columns[int(index % 3)].append(cleaned_value)

    # Write processed data to the output CSV file
    with open(output_file, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)

        # Write each row to the CSV file
        for row in zip(*columns):
            csvwriter.writerow(row)

    # Print the number of items in the list divided by 3
    print(f"Number of items in the list divided by 3: {total_values // 3}")

# Example usage
process_file("input.txt", "output.csv")
