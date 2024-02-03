import csv

# Read in the file
with open('input.txt', 'r') as file:
    ips = [line.strip().split(':')[0] for line in file.readlines()]

# Keep only unique IPs, maintain order
ips = list(dict.fromkeys(ips))

# Sort IPs into groups of 3
ip_groups = [ips[i:i+3] for i in range(0, len(ips), 3)]

# Write to output.csv
with open('output.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    for group in ip_groups:
        # Add trailing commas if group has less than 3 members
        while len(group) < 3:
            group.append('')
        writer.writerow(group)
