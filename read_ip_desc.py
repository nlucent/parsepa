import re
import xlwings as xw

xlout = []
curline = []
# open file
f = open('ip_desc.txt')

name = ''
ip = ''
desc = ''

xlout.append(['name', 'ip-netmask', 'description'])
for line in f:
    if line.strip().endswith('{'):
        match = re.search(r'(.+) \{', line)
        if match:
            name = match.group(1)

    if line.strip().endswith(';'):
        match = re.match(r'\s+ip-netmask\s(.+);', line)
        if match:
            ip = match.group(1)
        
        match = re.match(r'\s+ip-range\s(.+);', line)
        if match:
            ip = match.group(1)

        match = re.match(r'\s+description (.+);', line)
        if match:
            desc = match.group(1)
    
    if line.strip().endswith('}'):
        xlout.append([name,ip,desc])
        name, ip, desc = '','',''





