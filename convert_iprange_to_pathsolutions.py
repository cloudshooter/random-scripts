def format_range(ip_range):
    range_start, range_end = ip_range.strip().split('-')
    formatted_range = f"range,{range_start},{range_end},Default"
    return formatted_range

def convert_file(input_file, output_file):
    with open(input_file, 'r') as infile, open(output_file, 'w') as outfile:
        for line in infile:
            formatted_range = format_range(line)
            outfile.write(formatted_range + '\n')

# usage
convert_file('input.txt', 'output.txt')