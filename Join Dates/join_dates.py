import datetime

formatted_lines = []
checked_lines = []
write_lines = 'Pernr,Start,End,Hours\n'
key = 0

file = open('dates.csv', 'r')
csv_lines = file.readlines()
file.close()

# Remove the header row
csv_lines.pop(0)

# Remove line breaks and make into a list
for csv_line in csv_lines:
    csv_line = csv_line.replace('\n', '')
    formatted_lines.append(csv_line.split(','))

# Convert date strings to datetime objects and back to a sortable string
for formatted_line in formatted_lines:
    formatted_line[1] = datetime.datetime.strptime(formatted_line[1], '%d/%m/%Y').strftime('%Y/%m/%d')
    formatted_line[2] = datetime.datetime.strptime(formatted_line[2], '%d/%m/%Y').strftime('%Y/%m/%d')

formatted_lines.sort()

# Check if the start date of each sorted line is one day on from the end date of the previous line; combine them, if so
for formatted_line in formatted_lines:
    if len(checked_lines) == 0:
        checked_lines.append(formatted_line)
    elif formatted_line[0] == checked_lines[key][0] and datetime.datetime.strptime(formatted_line[1], '%Y/%m/%d') == \
            datetime.datetime.strptime(checked_lines[key][2], '%Y/%m/%d') + datetime.timedelta(days=1):
        checked_lines[key][2] = formatted_line[2]
        checked_lines[key][3] = str(int(checked_lines[key][3]) + int(formatted_line[3]))
    else:
        checked_lines.append(formatted_line)
        key += 1

# Reformat the dates before outputting them to a new csv
for checked_line in checked_lines:
    checked_line[1] = datetime.datetime.strptime(checked_line[1], '%Y/%m/%d').strftime('%d/%m/%Y')
    checked_line[2] = datetime.datetime.strptime(checked_line[2], '%Y/%m/%d').strftime('%d/%m/%Y')
    write_lines += ','.join(checked_line) + '\n'

new_file = open('checked_dates.csv', 'w')
new_file.write(write_lines)
new_file.close()
