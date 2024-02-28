def chunk(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

with open('input.txt', 'r') as f:
    ip_list = f.read().splitlines()

ip_set = set(ip_list)
ip_chunks = list(chunk(list(ip_set), 64))

for index, chunk in enumerate(ip_chunks):
    with open(f'output_{index}.txt', 'w') as f:
        for ip in chunk:
            f.write(f'{ip}\n')
