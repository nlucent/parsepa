import re
import xlwings as xw

xlout = []
curline = []
# open file
f = open('route_desc.txt')

name = ''
ip = ''
metric = ''
dest = ''
interface = ''

xlout.append(['name', 'nexthop', 'metric',  'interface', 'destination'])


for line in f:
    if line.strip().endswith('{'):
        match = re.search(r'(.+) \{', line)
        if match:
            if "nexthop" in match.group(1):
                next
            else:
                name = match.group(1)

    if line.strip().endswith(';'):
        match = re.match(r'\s+ip-address\s(.+);', line)
        if match:
            ip = match.group(1)
        
        match = re.match(r'\s+metric\s(.+);', line)
        if match:
            metric = match.group(1)

        match = re.match(r'\s+destination (.+);', line)
        if match:
            dest = match.group(1)
    
        match = re.match(r'\s+interface (.+);', line)
        if match:
            interface = match.group(1)

    match = re.match(r'}', line)
    if match:

# if line.strip().endswith('}'):
        xlout.append([name,ip, metric, interface, dest])
        name, ip, metric, interface, dest = '','','','',''





