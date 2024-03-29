def process_file(input_file, output_file):
    # Read the content from the input file and remove duplicates
    with open(input_file, 'r') as file:
        lines = file.read().splitlines()
        host_set = set(lines)

    # Write processed data to the output file
    with open(output_file, 'w') as file:
        step = len(host_set) // 3
        count = 0
        for index, value in enumerate(host_set):
            # Remove ':' and trailing characters
            cleaned_value = value.split(':')[0].strip()
            file.write(cleaned_value + '\n')
            count += 1
            # Insert separator at the 1/3 and 2/3 marks, excluding the last iteration
            if count % step == 0 and count < len(host_set) and count < step * 2:
                file.write("===========\n")

    # Print the number of items in the list divided by 3
    print(f"Number of items in the list divided by 3: {len(host_set) // 3}")

# Example usage
process_file("input.txt", "output.txt")
