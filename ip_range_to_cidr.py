import ipaddress

# Function to convert IP range to CIDR
def get_cidrs(start_ip, end_ip):
    cidrs = [ipaddr for ipaddr in ipaddress.summarize_address_range(
        ipaddress.IPv4Address(start_ip),
        ipaddress.IPv4Address(end_ip))]
    return cidrs

# Function to read input file, convert IP range to CIDR, and write to output file
def generate_cidr(input_file, output_file):
    with open(input_file, 'r') as file:
        ip_ranges = [line.strip() for line in file]
    with open(output_file, 'w') as file:
        for ip_range in ip_ranges:
            start_ip, end_ip = ip_range.split('-')
            cidrs = get_cidrs(start_ip, end_ip)
            for cidr in cidrs:
                file.write(str(cidr) + '\n')

# Call the function
generate_cidr('input.txt', 'output.txt')
