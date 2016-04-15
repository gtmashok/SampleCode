# The "xmlSpshtReader.py" is designed to read XML Spreadsheet files that have been converted
# from a spreadsheet format (.xlsx, .xls, etc.). You can fork this code and make appropriate 
# adjustments to the code below (especially the tags).

# Included are comments next to each line explaining key sections of the code.

import xml.etree.ElementTree as ET

file_in = raw_input('Enter file name: ')

fh = open(file_in)          # 'fh' is the file handle
data = fh.read()
#print type(data)                               # Optional information to print
#print 'Retrieved',len(data),'characters'       # Optional information to print

tree = ET.parse(file_in).getroot()
#tree = ET.fromstring(data)                     # Another way of obtaining the tags.

ftag = '{urn:schemas-microsoft-com:office:spreadsheet}'     # Create a string with the first part of 
                                                            # the xml tag since this code is designed
                                                            # to read XML Spreadsheet files
rows = tree.findall('.//'+ftag+'Data')              

print 'Name count:', len(rows)/5-1              # To count the names, the length of the 'row' is divided by 5 
                                                # because each row contains 5 pieces of attributes with the 
                                                # 'Data' xml tag for each person and the header (which is the
                                                # very first row of the sample file called "SampleList").         

c = 0
fname = list()
lname = list()
email = list()
for item in rows:
    c = c+1
    if not (c % 5 == 2 or c % 5 == 3 or c % 5 == 0):        # Determining the data items (representing column numbers) to include.
        continue
    elif c % 5 == 2:
        fname.append(item.text)
    elif c % 5 == 3:
        lname.append(item.text)
    elif c % 5 == 0:
        email.append(item.text)

# All the first & last names and emails are printed in a formatted fashion below.
for v in range(len(rows)/5):                        # v is the counter variable for each row of data outputted.
    if v == 0:
        print '{:^15}  {:^15}  {:^15}'.format(fname[v],lname[v],email[v])   # The first row of headers is printed in a center-aligned format.
    else:
        print '{:<15}  {:<15}  {:<15}'.format(fname[v],lname[v],email[v])
