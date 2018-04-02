import re
import sys
import xlwings as xw


f = open('security.txt', encoding='utf-8')
xlout = []
outdict = {}
curtitle = ''

all_lines = f.readlines()
for line in all_lines:
    match = re.search(r'^\s{16}(\w+) {', line)  # ex: DNS_Queries or other top level item
    
    if match:
        print('Group=' + match.group(1))
        curtitle = match.group(1)
        outdict[curtitle] = {}


    if line.strip().endswith(';'):
        #match = re.search(r'\s{18,20}(.+) \[ (.+) \];', line)
        # if match:
        #     outdict[curtitle][match.group(1)] = match.group(2)
        #     print(match.group(1) + '=' + match.group(2))

        # Match rule name
        match = re.search(r'\s{18,20}(\w+)\s(.+);', line)
        if match:
            outdict[curtitle][match.group(1)] = match.group(2)
            print(curtitle + ':' + match.group(1)+'='+match.group(2))

        # Match rule options
        match = re.search(r'\s{18,20}(\w+(?:-\w+)*)\s(?:-\[\s)(.+);', line)
        if match:
            outdict[curtitle][match.group(1)] = match.group(2)
            print(curtitle + ':' + match.group(1)+'='+match.group(2))

tmplist = []
max = 0
min = 0
maxname = ''
minname = ''

for key in sorted(outdict.keys()):
    tmplist = list(sorted(outdict[key].keys()))
    if min == 0:
        min = len(tmplist)
        minname = key

    if len(tmplist) > max:
        max = len(tmplist)
        maxname = key
    elif len(tmplist) < min:
        min = len(tmplist)
        minname = key

columns = list(sorted(outdict[maxname].keys()))
xlout.append(columns)
row = []

for key in sorted(outdict.keys()):
    row = [key]
    for col in columns:
        try:
            row.append(outdict[key][col])
        except  KeyError:
            row.append('')
    xlout.append(row)

wb = xw.Book()
rownum = 1

for i in range(0, len(xlout)):
    xw.Range('Sheet1', 'A' + str(rownum)).value = xlout[i]
    rownum +=1
