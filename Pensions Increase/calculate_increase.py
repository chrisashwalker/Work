import datetime

fs_types = {}
ca_types = {}
gmp_types = {}
fs_multipliers = {}
ca_multipliers = {}
new_rates_lines = []
new_csv_lines = []
new_gmp_lines = []
# TODO: Input dates and file paths
increase_date = datetime.datetime.strptime('10/04/2021', '%d/%m/%Y')
increase_end_date = datetime.datetime.strptime('31/12/9999', '%d/%m/%Y')
gmp_date = datetime.datetime.strptime('06/04/2021', '%d/%m/%Y')
gmp_end_date = datetime.datetime.strptime('31/12/9999', '%d/%m/%Y')
gmp_rate = 0.01
written_lines = ''
problems = []
gmp_run = False
pens_inc_run = False
gmp_done = False
pens_inc_done = False

if increase_date > gmp_date:
    increase_end_date = datetime.datetime.strptime('31/12/9999', '%d/%m/%Y')
    gmp_end_date = increase_date - datetime.timedelta(days=1)
    gmp_run = True
elif increase_date < gmp_date:
    increase_end_date = gmp_date - datetime.timedelta(days=1)
    gmp_end_date = datetime.datetime.strptime('31/12/9999', '%d/%m/%Y')
    pens_inc_run = True

f = open('multipliers.csv', 'r')
rates_lines = f.readlines()
f.close()

# Remove the header row
rates_lines.pop(0)

# Remove line breaks and make into a new list
for rates_line in rates_lines:
    rates_line = rates_line.replace('\n', '').split(',')
    new_rates_lines.append(rates_line)

f = open('wage_types.csv', 'r')
types_lines = f.readlines()
f.close()

# Remove the header row
types_lines.pop(0)

# Remove line breaks and make into a list; store increasing FS & CA types in separate lists
for types_line in types_lines:
    types_line = types_line.replace('\n', '').split(',')
    if types_line[1] == 'FS' and types_line[2] != '0':
        fs_types[types_line[0]] = types_line[2]
    elif types_line[1] == 'CA' and types_line[2] != '0':
        ca_types[types_line[0]] = types_line[2]
    elif types_line[1] == 'GMP' and types_line[2] != '0':
        gmp_types[types_line[0]] = types_line[2]

f = open('pay_records.csv', 'r')
csv_lines = f.readlines()
f.close()

# Remove the header row
csv_lines.pop(0)

# Remove line breaks and make into a new list
for csv_line in csv_lines:
    csv_line = csv_line.replace('\n', '')
    new_csv_lines.append(csv_line.split(','))

while not pens_inc_done or not gmp_done:
    if pens_inc_run and not pens_inc_done:
        # Convert amount strings into integers, apply the multiplier where relevant,
        # back to strings and then write the new line
        for new_csv_line in new_csv_lines:
            inc_list = []

            # Build dictionaries of IDs against FS or CA multipliers
            for rates_line in new_rates_lines:

                if datetime.datetime.strptime(rates_line[0], '%d/%m/%Y') <= \
                        datetime.datetime.strptime(new_csv_line[19], '%d/%m/%Y') <= \
                        datetime.datetime.strptime(rates_line[1], '%d/%m/%Y'):
                    fs_multipliers[new_csv_line[0]] = float(rates_line[2])
                if datetime.datetime.strptime(rates_line[0], '%d/%m/%Y') <= \
                        datetime.datetime.strptime(new_csv_line[20], '%d/%m/%Y') <= \
                        datetime.datetime.strptime(rates_line[1], '%d/%m/%Y'):
                    ca_multipliers[new_csv_line[0]] = float(rates_line[2])

            # Store the increase amounts in a list to be entered later; column no. against increase amount
            for j in range(3, 18, 2):
                wt_found = False
                blank_wt = 99
                if new_csv_line[j] in fs_types.keys():
                    for j2 in range(3, 18, 2):
                        if new_csv_line[j2] == fs_types[new_csv_line[j]]:
                            inc_list.append((j2 + 1, float(new_csv_line[j+1]) * fs_multipliers[new_csv_line[0]] / 100))
                            wt_found = True
                        if not new_csv_line[j2]:
                            blank_wt = j2
                            break
                    if not wt_found:
                        if blank_wt == 99:  # Too many wage types, can't input a new one
                            problems.append(new_csv_line[0])
                        else:
                            new_csv_line[blank_wt] = fs_types[new_csv_line[j]]
                            new_csv_line[blank_wt + 1] = str(0.00)
                            inc_list.append((blank_wt + 1, float(new_csv_line[j + 1]) *
                                             fs_multipliers[new_csv_line[0]] / 100))
                elif new_csv_line[j] in ca_types.keys():
                    for j2 in range(3, 18, 2):
                        if new_csv_line[j2] == ca_types[new_csv_line[j]]:
                            inc_list.append((j2 + 1, float(new_csv_line[j+1]) * ca_multipliers[new_csv_line[0]] / 100))
                            wt_found = True
                        if not new_csv_line[j2]:
                            blank_wt = j2
                            break
                    if not wt_found:
                        if blank_wt == 99:  # Too many wage types, can't input a new one
                            problems.append(new_csv_line[0])
                        else:
                            new_csv_line[blank_wt] = ca_types[new_csv_line[j]]
                            new_csv_line[blank_wt + 1] = str(0.00)
                            inc_list.append((blank_wt + 1, float(new_csv_line[j + 1]) *
                                             ca_multipliers[new_csv_line[0]] / 100))

            # Add the increase amounts into the new csv lines
            for inc in inc_list:
                new_csv_line[inc[0]] = str(float(new_csv_line[inc[0]]) + inc[1])

            if new_csv_line[0] in problems:
                new_csv_line[1] = 'PROBLEM'
                new_csv_line[2] = 'PROBLEM'
            else:
                new_csv_line[1] = datetime.datetime.strftime(increase_date, '%d/%m/%Y')
                new_csv_line[2] = datetime.datetime.strftime(increase_end_date, '%d/%m/%Y')

            written_lines += ','.join(new_csv_line) + '\n'
            pens_inc_done = True
            gmp_run = True

    if gmp_run and not gmp_done:
        # Convert amount strings into integers, apply the multiplier where relevant,
        # back to strings and then write the new line
        for new_csv_line in new_csv_lines:
            inc_list = []

            # Store the increase amounts in a list to be entered later; column no against increase amount
            for j in range(3, 18, 2):
                wt_found = False
                blank_wt = 99
                if new_csv_line[j] in gmp_types.keys():
                    for j2 in range(3, 18, 2):
                        if new_csv_line[j2] == gmp_types[new_csv_line[j]]:
                            inc_list.append(
                                (j2 + 1, float(new_csv_line[j + 1]) * gmp_rate))
                            wt_found = True
                        if not new_csv_line[j2]:
                            blank_wt = j2
                            break
                    if not wt_found:
                        if blank_wt == 99:  # Too many wage types, can't input a new one
                            problems.append(new_csv_line[0])
                        else:
                            new_csv_line[blank_wt] = gmp_types[new_csv_line[j]]
                            new_csv_line[blank_wt + 1] = str(0.00)
                            inc_list.append(
                                (blank_wt + 1, float(new_csv_line[j + 1]) * gmp_rate))

            # Add the increase amounts into the new csv lines
            for inc in inc_list:
                new_csv_line[inc[0]] = str(float(new_csv_line[inc[0]]) + inc[1])

            if new_csv_line[0] in problems:
                new_csv_line[1] = 'PROBLEM'
                new_csv_line[2] = 'PROBLEM'
            else:
                new_csv_line[1] = datetime.datetime.strftime(gmp_date, '%d/%m/%Y')
                new_csv_line[2] = datetime.datetime.strftime(gmp_end_date, '%d/%m/%Y')

            written_lines += ','.join(new_csv_line) + '\n'
            gmp_done = True
            pens_inc_run = True

w = open('new_pay_records.csv', 'w')
w.write(written_lines)
w.close()
